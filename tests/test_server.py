"""The 15 tools end to end through a real MCP client, in-process, against the fake Graph in conftest.py."""

from __future__ import annotations

from pathlib import Path

import httpx
import pytest
from mcp.client import Client

from conftest import Call, FakeGraph, make_graph
from graph_mcp.auth import NotSignedIn
from graph_mcp.graph import Graph
from graph_mcp.server import build_server, calendar_range

pytestmark = pytest.mark.anyio

TOOLS = {
    "whoami",
    "mail_folders",
    "mail_search",
    "mail_list",
    "mail_get",
    "mail_attachment",
    "onedrive_list",
    "onedrive_search",
    "onedrive_shared",
    "onedrive_read",
    "onenote_notebooks",
    "onenote_pages",
    "onenote_read",
    "calendar_events",
    "contacts_search",
}

# ---- the read-only guarantees ----


async def test_exactly_these_tools_all_read_only(fake: FakeGraph) -> None:
    async with Client(build_server(make_graph(fake)), mode="legacy") as client:
        tools = (await client.list_tools()).tools
    assert {t.name for t in tools} == TOOLS  # a new tool must be added here on purpose, after review
    assert all(t.annotations and t.annotations.read_only_hint for t in tools)


async def test_whoami_sends_bearer_token(fake: FakeGraph, call: Call) -> None:
    err, text = await call("whoami", graph=make_graph(fake, "secret-token"))
    assert not err and text == "Test User (test@example.com)"
    assert fake.requests[0].headers["authorization"] == "Bearer secret-token"


# ---- errors ----


async def test_graph_errors_become_tool_errors(call: Call) -> None:
    err, text = await call("mail_get", message_id="missing")
    assert err and "404" in text and "itemNotFound" in text


async def test_not_signed_in_is_a_tool_error(fake: FakeGraph, call: Call) -> None:
    def no_token() -> str:
        raise NotSignedIn("Not signed in. Run `graph-mcp login`.")

    graph = Graph(no_token, httpx.Client(transport=httpx.MockTransport(fake)))
    err, text = await call("whoami", graph=graph)
    assert err and "graph-mcp login" in text
    assert fake.requests == []


async def test_ids_are_url_quoted(fake: FakeGraph, call: Call) -> None:
    await call("mail_get", message_id="a/b?c")
    assert fake.requests[0].url.raw_path.startswith(b"/v1.0/me/messages/a%2Fb%3Fc")


# ---- mail ----


async def test_mail_folders(call: Call) -> None:
    err, text = await call("mail_folders")
    assert not err and text == "- Inbox: 2 unread / 9 total (id: f1)"


async def test_mail_search(fake: FakeGraph, call: Call) -> None:
    err, text = await call("mail_search", query="invoice")
    assert not err
    assert "Shop <shop@example.com> | Your invoice [unread, attachments]" in text
    assert "id: m1" in text and "Thanks for your order" in text
    assert fake.requests[0].url.params["$search"] == '"invoice"'


async def test_mail_search_strips_quotes(fake: FakeGraph, call: Call) -> None:
    await call("mail_search", query='subject:"tax return"')
    assert fake.requests[0].url.params["$search"] == '"subject:tax return"'  # can't break out of the quoted search


async def test_mail_list_filters(fake: FakeGraph, call: Call) -> None:
    err, text = await call("mail_list", unread_only=True, since="2026-11-01", top=500)
    assert not err and "Your invoice" in text
    params = fake.requests[0].url.params
    assert fake.requests[0].url.path == "/v1.0/me/mailFolders/inbox/messages"
    assert params["$filter"] == "isRead eq false and receivedDateTime ge 2026-11-01T00:00:00Z"
    assert params["$top"] == "50"  # capped


@pytest.mark.parametrize(
    ("since", "expected"),
    [
        ("2026-11-10", "2026-11-10T05:00:00Z"),  # local midnight, EST
        ("2026-11-01", "2026-11-01T04:00:00Z"),  # the day summer time ends: midnight is still EDT
        ("2026-11-10T09:30", "2026-11-10T14:30:00Z"),
        ("2026-11-10T09:30+01:00", "2026-11-10T08:30:00Z"),  # an explicit offset wins
    ],
)
async def test_mail_list_since_is_local(fake: FakeGraph, call: Call, since: str, expected: str) -> None:
    err, _ = await call("mail_list", server_tz="America/New_York", since=since)
    assert not err
    assert fake.requests[0].url.params["$filter"] == f"receivedDateTime ge {expected}"


async def test_mail_list_rejects_anything_but_a_date(fake: FakeGraph, call: Call) -> None:
    err, text = await call("mail_list", since="2026-01-01 or true")
    assert err and "since must be an ISO date" in text
    assert fake.requests == []  # never reaches the filter


async def test_mail_get_lists_attachments(fake: FakeGraph, call: Call) -> None:
    err, text = await call("mail_get", message_id="m1")
    assert not err
    assert "Cc: Sam <sam@example.com>" in text
    assert "Total: $12" in text and "invoice.txt" in text and "id: a1" in text
    assert fake.requests[0].headers["prefer"] == 'outlook.body-content-type="text"'


async def test_mail_attachment_is_converted_in_memory(
    call: Call, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.chdir(tmp_path)
    err, text = await call("mail_attachment", message_id="m1", attachment_id="a1")
    assert not err
    assert text.startswith("invoice.txt (11 bytes).") and "Invoice #42" in text
    assert list(tmp_path.iterdir()) == []  # nothing written anywhere by default


async def test_mail_attachment_is_saved_when_downloads_set(call: Call, tmp_path: Path) -> None:
    err, text = await call("mail_attachment", downloads=tmp_path, message_id="m1", attachment_id="a1")
    assert not err
    saved = tmp_path / "mail" / "invoice.txt"
    assert saved.read_bytes() == b"Invoice #42"
    assert str(saved) in text and "Invoice #42" in text


# ---- OneDrive ----


async def test_onedrive_list_root(call: Call) -> None:
    err, text = await call("onedrive_list")
    assert not err and text.startswith("- [folder] /Documents (4 items")


async def test_onedrive_list_by_path(call: Call) -> None:
    err, text = await call("onedrive_list", path="/Documents/")
    assert not err
    assert "[folder] /Documents/Taxes (3 items" in text
    assert "[file] /Documents/notes.txt (1,234 bytes, modified 2026-05-02) id: i1" in text


async def test_onedrive_search(call: Call) -> None:
    err, text = await call("onedrive_search", query="budget")
    assert not err and "[file] /Money/budget.xlsx (2,048 bytes" in text


async def test_onedrive_search_escapes_quotes(fake: FakeGraph, call: Call) -> None:
    await call("onedrive_search", query="O'Brien')/x")
    # quotes are doubled and the rest URL-encoded: it can't end the search expression or change the path
    assert fake.requests[0].url.raw_path.startswith(b"/v1.0/me/drive/root/search(q='O%27%27Brien%27%27%29%2Fx')")


async def test_onedrive_shared(call: Call) -> None:
    err, text = await call("onedrive_shared")
    assert not err and text == "- [file] Plan.txt (shared by Alex) id: r1 drive_id: d1"


async def test_onedrive_read_by_id(call: Call) -> None:
    err, text = await call("onedrive_read", item_id="i1")
    assert not err and "buy mulch now" in text


async def test_onedrive_read_by_path(fake: FakeGraph, call: Call) -> None:
    err, text = await call("onedrive_read", path="/Documents/notes.txt")
    assert not err and "buy mulch now" in text
    assert fake.requests[0].url.path == "/v1.0/me/drive/root:/Documents/notes.txt:"


async def test_onedrive_read_shared_file(call: Call) -> None:
    err, text = await call("onedrive_read", item_id="r1", drive_id="d1")
    assert not err and "rent 1200" in text


async def test_onedrive_read_refuses_folders_and_nothing(call: Call) -> None:
    err, text = await call("onedrive_read", item_id="f2")
    assert err and "That is a folder" in text
    err, text = await call("onedrive_read")
    assert err and "Give item_id or path" in text


# ---- OneNote ----


async def test_onenote_notebooks(call: Call) -> None:
    err, text = await call("onenote_notebooks")
    assert not err and text == "- Home (modified 2026-09-09)\n    - Yard (section id: s1)"


async def test_onenote_pages_expands_section_name(fake: FakeGraph, call: Call) -> None:
    err, text = await call("onenote_pages")
    assert not err and "Garden plan | Yard | 2026-10-01" in text
    assert fake.requests[0].url.params["$expand"] == "parentSection($select=displayName)"


async def test_onenote_pages_in_a_section_by_title(fake: FakeGraph, call: Call) -> None:
    err, text = await call("onenote_pages", section_id="s1", title_contains="O'Neil")
    assert not err and "Garden plan" in text
    assert fake.requests[0].url.path == "/v1.0/me/onenote/sections/s1/pages"
    assert fake.requests[0].url.params["$filter"] == "contains(tolower(title),'o''neil')"  # quote doubled


async def test_onenote_read(call: Call) -> None:
    err, text = await call("onenote_read", page_id="p1")
    assert not err and "Tomatoes in May" in text


# ---- calendar and contacts ----


async def test_calendar_dates_are_in_the_server_time_zone(fake: FakeGraph, call: Call) -> None:
    err, text = await call("calendar_events", server_tz="America/New_York", start="2026-11-16", end="2026-11-16")
    assert not err and "2026-11-16 09:30 → 10:15 | Dentist @ Main St" in text
    request = fake.requests[0]
    # Graph reads times without an offset as UTC: they must carry the local offset, and a date-only end the whole day.
    assert request.url.params["startDateTime"] == "2026-11-16T00:00:00-05:00"
    assert request.url.params["endDateTime"] == "2026-11-17T00:00:00-05:00"
    assert request.headers["prefer"] == 'outlook.timezone="America/New_York"'


async def test_calendar_shows_end_dates_and_all_day_events(call: Call) -> None:
    err, text = await call("calendar_events", start="2026-11-01", end="2026-12-31")
    assert not err
    assert "- 2026-12-07 06:00 → 2026-12-11 21:00 | Trip" in text  # the end date used to be dropped
    assert "- 2026-11-17 | Sign-up deadline (all day)" in text
    assert "- 2026-12-24 → 2026-12-26 | Holidays (all day)" in text  # Graph's end is midnight after the last day


async def test_times_are_shown_in_the_server_time_zone(call: Call) -> None:
    tz = "America/New_York"
    _, text = await call("mail_list", server_tz=tz)
    assert text.startswith("- 2026-11-10 03:30 | Shop")  # received 08:30 UTC
    _, text = await call("mail_get", server_tz=tz, message_id="m1")
    assert "Date: 2026-11-10 03:30" in text
    _, text = await call("onedrive_list", server_tz=tz, path="/Documents")
    assert "notes.txt (1,234 bytes, modified 2026-05-01)" in text  # 00:00 UTC on May 2 is still May 1 here
    _, text = await call("onenote_notebooks", server_tz=tz)
    assert "Home (modified 2026-09-08)" in text


async def test_instructions_name_the_time_zone(fake: FakeGraph) -> None:
    server = build_server(make_graph(fake), timezone="Asia/Tokyo")
    assert "All dates and times, shown and given, are in Asia/Tokyo" in (server.instructions or "")


async def test_calendar_rejects_a_bad_date(fake: FakeGraph, call: Call) -> None:
    err, text = await call("calendar_events", start="next week", end="2026-11-20")
    assert err and "ISO date" in text
    assert fake.requests == []


@pytest.mark.parametrize(
    ("start", "end", "expected"),
    [
        ("2026-11-16", "2026-11-23", ("2026-11-16T00:00:00-05:00", "2026-11-24T00:00:00-05:00")),
        ("2026-07-01", "2026-07-01", ("2026-07-01T00:00:00-04:00", "2026-07-02T00:00:00-04:00")),  # summer time
        ("2026-11-16 09:00", "2026-11-16 10:00", ("2026-11-16T09:00:00-05:00", "2026-11-16T10:00:00-05:00")),
        ("2026-11-16T09:00Z", "2026-11-16T10:00+01:00", ("2026-11-16T09:00:00+00:00", "2026-11-16T10:00:00+01:00")),
    ],
)
def test_calendar_range(start: str, end: str, expected: tuple[str, str]) -> None:
    assert calendar_range(start, end, "America/New_York") == expected


async def test_contacts_search_escapes_quotes(fake: FakeGraph, call: Call) -> None:
    err, text = await call("contacts_search", query="O'Brien")
    assert not err and text == "- Pat O'Brien | pat@example.com | 555"
    assert fake.requests[0].url.params["$filter"] == (
        "startswith(displayName,'O''Brien') or startswith(givenName,'O''Brien') or startswith(surname,'O''Brien')"
    )
