"""Serving over Streamable HTTP behind a Bearer token, for agents that can't start a local process."""

from __future__ import annotations

import hmac
import json
from collections.abc import Callable
from typing import Any

import anyio
from mcp.server.mcpserver import MCPServer
from starlette.types import ASGIApp, Receive, Scope, Send

LOOPBACK = ("127.0.0.1", "localhost", "::1")
MIN_TOKEN_LENGTH = 32
MCP_PATH = "/mcp"
HEALTH_PATH = "/healthz"


class BearerAuth:
    """Rejects HTTP requests without the right `Authorization: Bearer <token>`. `/healthz` stays open: 200 when
    `signed_in()` says there's a sign-in, else 503, so a container whose sign-in is gone shows as unhealthy."""

    def __init__(self, app: ASGIApp, token: str | None, signed_in: Callable[[], bool] | None = None) -> None:
        self.app = app
        self.token = token.encode() if token else None
        self.signed_in = signed_in

    async def __call__(self, scope: Scope, receive: Receive, send: Send) -> None:
        if scope["type"] != "http":
            await self.app(scope, receive, send)  # lifespan events
            return
        if scope["path"] == HEALTH_PATH:
            if await self._has_sign_in():
                await _respond(send, 200, {"status": "ok"})
            else:
                await _respond(send, 503, {"status": "not signed in"})
            return
        if self.token is not None and not self._authorized(scope):
            await _respond(send, 401, {"error": "unauthorized"}, {"www-authenticate": "Bearer"})
            return
        await self.app(scope, receive, send)

    async def _has_sign_in(self) -> bool:
        if self.signed_in is None:
            return True
        try:  # reads the keychain or the cache file: off the event loop
            return await anyio.to_thread.run_sync(self.signed_in)
        except Exception:  # an unreadable cache counts as no sign-in; details stay out of an open endpoint
            return False

    def _authorized(self, scope: Scope) -> bool:
        assert self.token is not None
        for name, value in scope.get("headers", []):
            if name == b"authorization" and value[:7].lower() == b"bearer ":
                return hmac.compare_digest(value[7:].strip(), self.token)
        return False


async def _respond(send: Send, status: int, body: dict[str, Any], headers: dict[str, str] | None = None) -> None:
    payload = json.dumps(body).encode()
    raw = [(b"content-type", b"application/json"), (b"content-length", str(len(payload)).encode())]
    raw += [(k.encode(), v.encode()) for k, v in (headers or {}).items()]
    await send({"type": "http.response.start", "status": status, "headers": raw})
    await send({"type": "http.response.body", "body": payload})


def check_token(host: str, token: str | None, allow_no_auth: bool) -> None:
    if token is None:
        if host in LOOPBACK and allow_no_auth:
            return
        raise ValueError(
            "HTTP needs a token: set GRAPH_MCP_TOKEN (e.g. `openssl rand -hex 32`). Only on a loopback address can "
            "you go without one, with --no-auth."
        )
    if len(token) < MIN_TOKEN_LENGTH:
        raise ValueError(f"GRAPH_MCP_TOKEN is too short; use at least {MIN_TOKEN_LENGTH} characters.")


def http_app(server: MCPServer, token: str | None, host: str, signed_in: Callable[[], bool] | None = None) -> ASGIApp:
    return BearerAuth(server.streamable_http_app(streamable_http_path=MCP_PATH, host=host), token, signed_in)


def serve_http(
    server: MCPServer, host: str, port: int, token: str | None, signed_in: Callable[[], bool] | None = None
) -> None:
    import uvicorn

    uvicorn.run(http_app(server, token, host, signed_in), host=host, port=port, log_level="info")
