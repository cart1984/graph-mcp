"""The Graph client (GET only, errors, timeouts) and file conversion."""

from __future__ import annotations

from pathlib import Path

import httpx
import pytest

from conftest import FakeGraph, make_graph
from graph_mcp.graph import Graph, GraphError, convert, looks_binary, unique_path

# ---- the client ----


def test_graph_client_has_no_write_methods() -> None:
    public = {n for n in dir(Graph) if not n.startswith("_")}
    assert public == {"get", "get_bytes"}


def test_get_and_get_bytes(fake: FakeGraph) -> None:
    graph = make_graph(fake, "tok")
    assert graph.get("/me")["displayName"] == "Test User"
    assert graph.get_bytes("/me/messages/m1/attachments/a1/$value") == b"Invoice #42"
    assert graph.get("https://graph.microsoft.com/v1.0/me")["mail"] == "test@example.com"  # full nextLink-style URLs
    assert all(r.headers["authorization"] == "Bearer tok" for r in fake.requests)


def test_error_messages() -> None:
    def broken(request: httpx.Request) -> httpx.Response:
        if request.url.path.endswith("/json"):
            return httpx.Response(403, json={"error": {"code": "accessDenied", "message": "no"}})
        return httpx.Response(502, text="<html>Bad gateway</html>")

    graph = make_graph(broken)
    with pytest.raises(GraphError, match=r"\(403\) accessDenied: no"):
        graph.get("/json")
    with pytest.raises(GraphError, match=r"\(502\) <html>Bad gateway"):
        graph.get("/html")


def test_timeout_is_retried_once(fake: FakeGraph) -> None:
    calls = 0

    def flaky(request: httpx.Request) -> httpx.Response:
        nonlocal calls
        calls += 1
        if calls == 1:
            raise httpx.ReadTimeout("slow", request=request)
        return fake(request)

    assert make_graph(flaky).get("/me")["displayName"] == "Test User"
    assert calls == 2


def test_repeated_timeout_is_an_error() -> None:
    def stuck(request: httpx.Request) -> httpx.Response:
        raise httpx.ReadTimeout("slow", request=request)

    with pytest.raises(GraphError, match="didn't answer in time"):
        make_graph(stuck).get("/me")


def test_network_error_is_an_error() -> None:
    def down(request: httpx.Request) -> httpx.Response:
        raise httpx.ConnectError("no route", request=request)

    with pytest.raises(GraphError, match="Couldn't reach Microsoft Graph"):
        make_graph(down).get("/me")


# ---- conversion ----


def test_convert_in_memory_writes_nothing(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.chdir(tmp_path)
    text = convert(b"<h1>Plan</h1><p>Plant tomatoes</p>", "page.html")
    assert text.startswith("page.html (") and "Plant tomatoes" in text
    assert list(tmp_path.iterdir()) == []


def test_convert_saves_when_asked(tmp_path: Path) -> None:
    text = convert(b"hello", "notes.txt", tmp_path / "mail")
    assert (tmp_path / "mail" / "notes.txt").read_bytes() == b"hello"
    assert text.startswith(f"Saved to {tmp_path / 'mail' / 'notes.txt'}") and "hello" in text


def test_convert_refuses_binary() -> None:
    # MarkItDown's type detection takes these bytes for text, so they used to come back as raw characters
    assert convert(b"\x00\x01\x02", "blob.bin").endswith("[Can't convert .bin to text: it looks like a binary file.]")


@pytest.mark.parametrize(
    ("text", "binary"),
    [
        ("Grüße aus dem Garten 🌱\n\tTabelle\r\n", False),
        ("", False),
        ("ok\x00rest", True),
        ("\x01\x02\x03\x04 mostly control characters", True),
    ],
)
def test_looks_binary(text: str, binary: bool) -> None:
    assert looks_binary(text) is binary


def test_convert_too_large(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr("graph_mcp.graph.MAX_CONVERT_BYTES", 4)
    assert convert(b"12345", "big.txt") == "big.txt (5 bytes). Too large to convert to text."


def test_same_names_do_not_overwrite(tmp_path: Path) -> None:
    convert(b"first", "invoice.txt", tmp_path)
    convert(b"second", "invoice.txt", tmp_path)
    convert(b"first", "invoice.txt", tmp_path)  # the same file again reuses its path
    assert sorted(p.name for p in tmp_path.iterdir()) == ["invoice (2).txt", "invoice.txt"]
    assert (tmp_path / "invoice (2).txt").read_bytes() == b"second"


@pytest.mark.parametrize("name", ["../../etc/passwd", "..", "/abs/path.txt", "a\\b.txt"])
def test_unique_path_stays_in_the_folder(tmp_path: Path, name: str) -> None:
    assert unique_path(tmp_path, name, b"x").parent == tmp_path
