"""Shared test setup: a fake Microsoft Graph that only answers GET, and a way to call tools through an MCP client."""

from __future__ import annotations

import json
from collections.abc import Awaitable, Callable
from pathlib import Path

import httpx
import pytest
from mcp.client import Client
from mcp.types import TextContent

from graph_mcp.graph import Graph
from graph_mcp.server import build_server

INBOX_MESSAGE = {
    "id": "m1",
    "subject": "Your invoice",
    "from": {"emailAddress": {"name": "Shop", "address": "shop@example.com"}},
    "receivedDateTime": "2026-11-10T08:30:00Z",
    "isRead": False,
    "hasAttachments": True,
    "bodyPreview": "Thanks for your order",
}

ROUTES: dict[str, object] = {
    "/v1.0/me": {"displayName": "Test User", "mail": "test@example.com"},
    # mail
    "/v1.0/me/mailFolders": {
        "value": [{"id": "f1", "displayName": "Inbox", "unreadItemCount": 2, "totalItemCount": 9}]
    },
    "/v1.0/me/mailFolders/inbox/messages": {"value": [INBOX_MESSAGE]},
    "/v1.0/me/messages": {"value": [INBOX_MESSAGE]},
    "/v1.0/me/messages/m1": {
        "subject": "Your invoice",
        "from": {"emailAddress": {"name": "Shop", "address": "shop@example.com"}},
        "toRecipients": [{"emailAddress": {"address": "test@example.com"}}],
        "ccRecipients": [{"emailAddress": {"name": "Sam", "address": "sam@example.com"}}],
        "receivedDateTime": "2026-11-10T08:30:00Z",
        "body": {"contentType": "text", "content": "Total: $12"},
        "hasAttachments": True,
        "webLink": "https://outlook.live.com/x",
    },
    "/v1.0/me/messages/m1/attachments": {
        "value": [{"id": "a1", "name": "invoice.txt", "size": 11, "contentType": "text/plain"}]
    },
    "/v1.0/me/messages/m1/attachments/a1": {"name": "invoice.txt"},
    "/v1.0/me/messages/m1/attachments/a1/$value": b"Invoice #42",
    # OneDrive
    "/v1.0/me/drive/root/children": {
        "value": [
            {
                "id": "f1",
                "name": "Documents",
                "folder": {"childCount": 4},
                "lastModifiedDateTime": "2026-04-01T00:00:00Z",
                "parentReference": {"path": "/drive/root:"},
            }
        ]
    },
    "/v1.0/me/drive/root:/Documents:/children": {
        "value": [
            {
                "id": "f2",
                "name": "Taxes",
                "folder": {"childCount": 3},
                "lastModifiedDateTime": "2026-04-01T00:00:00Z",
                "parentReference": {"path": "/drive/root:/Documents"},
            },
            {
                "id": "i1",
                "name": "notes.txt",
                "size": 1234,
                "file": {},
                "lastModifiedDateTime": "2026-05-02T00:00:00Z",
                "parentReference": {"path": "/drive/root:/Documents"},
            },
        ]
    },
    "/v1.0/me/drive/root/search(q='budget')": {
        "value": [
            {
                "id": "i2",
                "name": "budget.xlsx",
                "size": 2048,
                "file": {},
                "lastModifiedDateTime": "2026-03-03T00:00:00Z",
                "parentReference": {"path": "/drive/root:/Money"},
            }
        ]
    },
    "/v1.0/me/drive/sharedWithMe": {
        "value": [
            {
                "name": "Plan.txt",
                "remoteItem": {
                    "id": "r1",
                    "file": {},
                    "parentReference": {"driveId": "d1"},
                    "shared": {"owner": {"user": {"displayName": "Alex"}}},
                },
            }
        ]
    },
    "/v1.0/drives/d1/items/r1": {"name": "Plan.txt", "size": 9},
    "/v1.0/drives/d1/items/r1/content": b"rent 1200",
    "/v1.0/me/drive/items/i1": {"name": "notes.txt", "size": 13},
    "/v1.0/me/drive/items/i1/content": b"buy mulch now",
    "/v1.0/me/drive/items/f2": {"name": "Taxes", "folder": {"childCount": 3}},
    "/v1.0/me/drive/root:/Documents/notes.txt:": {"name": "notes.txt", "size": 13},
    "/v1.0/me/drive/root:/Documents/notes.txt:/content": b"buy mulch now",
    # OneNote
    "/v1.0/me/onenote/notebooks": {
        "value": [
            {
                "id": "nb1",
                "displayName": "Home",
                "lastModifiedDateTime": "2026-09-09T00:00:00Z",
                "sections": [{"id": "s1", "displayName": "Yard"}],
            }
        ]
    },
    "/v1.0/me/onenote/pages": {
        "value": [
            {
                "id": "p1",
                "title": "Garden plan",
                "lastModifiedDateTime": "2026-10-01T10:00:00Z",
                "parentSection": {"displayName": "Yard"},
            }
        ]
    },
    "/v1.0/me/onenote/sections/s1/pages": {
        "value": [
            {
                "id": "p1",
                "title": "Garden plan",
                "lastModifiedDateTime": "2026-10-01T10:00:00Z",
                "parentSection": {"displayName": "Yard"},
            }
        ]
    },
    "/v1.0/me/onenote/pages/p1/content": b"<html><body><h1>Garden plan</h1><p>Tomatoes in May</p></body></html>",
    # calendar and contacts
    "/v1.0/me/calendarView": {
        "value": [
            {
                "subject": "Dentist",
                "start": {"dateTime": "2026-11-16T09:30:00.0000000"},
                "end": {"dateTime": "2026-11-16T10:15:00.0000000"},
                "location": {"displayName": "Main St"},
                "isAllDay": False,
            },
            {
                "subject": "Trip",
                "start": {"dateTime": "2026-12-07T06:00:00.0000000"},
                "end": {"dateTime": "2026-12-11T21:00:00.0000000"},
                "isAllDay": False,
            },
            {
                "subject": "Sign-up deadline",
                "start": {"dateTime": "2026-11-17T00:00:00.0000000"},
                "end": {"dateTime": "2026-11-18T00:00:00.0000000"},
                "isAllDay": True,
            },
            {
                "subject": "Holidays",
                "start": {"dateTime": "2026-12-24T00:00:00.0000000"},
                "end": {"dateTime": "2026-12-27T00:00:00.0000000"},
                "isAllDay": True,
            },
        ]
    },
    "/v1.0/me/contacts": {
        "value": [
            {"displayName": "Pat O'Brien", "emailAddresses": [{"address": "pat@example.com"}], "mobilePhone": "555"}
        ]
    },
}


class FakeGraph:
    """Answers from ROUTES by path; anything else is a Graph-style 404. Any method but GET fails the test."""

    def __init__(self) -> None:
        self.requests: list[httpx.Request] = []

    def __call__(self, request: httpx.Request) -> httpx.Response:
        assert request.method == "GET", f"the server sent {request.method} {request.url}"
        self.requests.append(request)
        body = ROUTES.get(request.url.path)
        if body is None:
            return httpx.Response(404, json={"error": {"code": "itemNotFound", "message": "nope"}})
        if isinstance(body, bytes):
            return httpx.Response(200, content=body)
        return httpx.Response(200, content=json.dumps(body).encode(), headers={"content-type": "application/json"})


def make_graph(handler: Callable[[httpx.Request], httpx.Response], token: str = "tok") -> Graph:
    return Graph(lambda: token, httpx.Client(transport=httpx.MockTransport(handler)))


@pytest.fixture
def fake() -> FakeGraph:
    return FakeGraph()


Call = Callable[..., Awaitable[tuple[bool, str]]]


@pytest.fixture
def call(fake: FakeGraph) -> Call:
    """`await call(tool, **args)`: runs one tool through an MCP client. Options: `graph=` (default: the fake),
    `downloads=` (default: keep nothing), `server_tz=` (default: UTC)."""

    async def run(
        tool: str, *, graph: Graph | None = None, downloads: Path | None = None, server_tz: str = "UTC", **args: object
    ) -> tuple[bool, str]:
        server = build_server(graph or make_graph(fake), downloads, server_tz)
        async with Client(server, mode="legacy") as client:
            result = await client.call_tool(tool, args)
        content = result.content[0]
        assert isinstance(content, TextContent)
        return bool(result.is_error), content.text

    return run
