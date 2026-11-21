from __future__ import annotations

import pytest
from starlette.types import Message, Receive, Scope, Send

from graph_mcp.httpserve import BearerAuth, check_token

pytestmark = pytest.mark.anyio
TOKEN = "t" * 40


def test_check_token() -> None:
    check_token("127.0.0.1", None, allow_no_auth=True)
    check_token("0.0.0.0", TOKEN, allow_no_auth=False)
    with pytest.raises(ValueError, match="needs a token"):
        check_token("0.0.0.0", None, allow_no_auth=True)
    with pytest.raises(ValueError, match="too short"):
        check_token("127.0.0.1", "short", allow_no_auth=False)
    with pytest.raises(ValueError, match="needs a token"):
        check_token("127.0.0.1", None, allow_no_auth=False)  # loopback still needs --no-auth
    check_token("::1", None, allow_no_auth=True)


async def _status(app: BearerAuth, path: str, auth: str | None, kind: str = "http") -> int:
    headers = [(b"authorization", auth.encode())] if auth else []
    scope: Scope = {"type": kind, "path": path, "headers": headers}
    sent: list[Message] = []

    async def receive() -> Message:
        return {"type": "http.request"}

    async def send(message: Message) -> None:
        sent.append(message)

    await app(scope, receive, send)
    return sent[0]["status"]


async def _ok(scope: Scope, receive: Receive, send: Send) -> None:
    await send({"type": "http.response.start", "status": 200, "headers": []})
    await send({"type": "http.response.body", "body": b""})


async def test_bearer_auth() -> None:
    app = BearerAuth(_ok, TOKEN)
    assert await _status(app, "/mcp", None) == 401
    assert await _status(app, "/mcp", "Bearer wrong") == 401
    assert await _status(app, "/mcp", f"Bearer {TOKEN}") == 200
    assert await _status(app, "/healthz", None) == 200
    assert await _status(app, "/mcp", f"bearer {TOKEN}") == 200  # the scheme is case-insensitive
    assert await _status(app, "/mcp", f"Basic {TOKEN}") == 401
    assert await _status(app, "/mcp", f"Bearer {TOKEN}x") == 401


async def test_lifespan_passes_through_without_a_token() -> None:
    assert await _status(BearerAuth(_ok, TOKEN), "", None, kind="lifespan") == 200


async def test_no_token_lets_everything_through() -> None:
    assert await _status(BearerAuth(_ok, None), "/mcp", None) == 200  # only allowed on loopback, see check_token


async def test_health_reports_the_sign_in() -> None:
    assert await _status(BearerAuth(_ok, TOKEN, signed_in=lambda: True), "/healthz", None) == 200
    assert await _status(BearerAuth(_ok, TOKEN, signed_in=lambda: False), "/healthz", None) == 503


async def test_health_with_an_unreadable_cache_is_503() -> None:
    def broken() -> bool:
        raise OSError("keychain locked")

    assert await _status(BearerAuth(_ok, TOKEN, signed_in=broken), "/healthz", None) == 503
