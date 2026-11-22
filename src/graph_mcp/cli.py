"""Command line entry point: `graph-mcp <command>`.

Settings come from environment variables, so every MCP client can pass them the same way:
  GRAPH_MCP_CLIENT_ID     Application (client) ID of your Entra app registration (required)
  GRAPH_MCP_TENANT        consumers (default: personal accounts), common, organizations, or a tenant ID
  GRAPH_MCP_TOKEN_CACHE   auto (default: keychain if there is one, else a file), keyring, or file
  GRAPH_MCP_CACHE_FILE    token cache file (default ~/.config/graph-mcp/token-cache.json)
  GRAPH_MCP_DOWNLOADS     keep the originals of files, attachments and pages here (default: keep nothing)
  GRAPH_MCP_TIMEZONE      IANA time zone for calendar times (default: $TZ, else this machine's)
  GRAPH_MCP_TOKEN         Bearer token clients must send when serving over HTTP
"""

from __future__ import annotations

import argparse
import logging
import os
import sys
from pathlib import Path
from typing import cast
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from graph_mcp import __version__
from graph_mcp.auth import Auth, AuthError, CacheKind, NotSignedIn, make_store


def _env(name: str, default: str = "") -> str:
    return os.environ.get(f"GRAPH_MCP_{name}") or default


def system_timezone() -> str | None:
    """The system's zone name from the /etc/localtime link (Linux, macOS), or None."""
    target = os.path.realpath("/etc/localtime")
    return target.split("zoneinfo/", 1)[1] if "zoneinfo/" in target else None


def resolve_timezone() -> str:
    """GRAPH_MCP_TIMEZONE, else TZ, else the system's zone, else UTC. Unknown names are an error, not a silent UTC."""
    name = _env("TIMEZONE") or os.environ.get("TZ", "").lstrip(":") or system_timezone() or "UTC"
    try:
        ZoneInfo(name)
    except (ZoneInfoNotFoundError, ValueError) as e:
        raise ValueError(
            f"Unknown time zone {name!r}: use a name like America/New_York (GRAPH_MCP_TIMEZONE or TZ)."
        ) from e
    return name


def make_auth() -> Auth:
    kind = cast(CacheKind, _env("TOKEN_CACHE", "auto"))
    if kind not in ("auto", "keyring", "file"):
        raise AuthError(f"GRAPH_MCP_TOKEN_CACHE must be auto, keyring or file, not {kind!r}.")
    file = Path(_env("CACHE_FILE")).expanduser() if _env("CACHE_FILE") else None
    return Auth(client_id=_env("CLIENT_ID"), tenant=_env("TENANT", "consumers"), store=make_store(kind, file))


def quiet_logs() -> None:
    """No log line per request to Graph (the MCP server turns on INFO logging): they'd fill a container's log with
    message and file ids. Warnings and errors still show."""
    logging.getLogger("httpx").setLevel(logging.WARNING)


def cmd_serve(args: argparse.Namespace) -> int:
    from graph_mcp.graph import Graph
    from graph_mcp.server import build_server

    quiet_logs()
    auth = make_auth()
    downloads = Path(_env("DOWNLOADS")).expanduser() if _env("DOWNLOADS") else None
    server = build_server(Graph(auth.token), downloads, resolve_timezone())
    if args.transport == "stdio":
        server.run("stdio")
        return 0
    from graph_mcp.httpserve import check_token, serve_http

    token = _env("TOKEN") or None
    check_token(args.host, token, args.no_auth)
    serve_http(server, args.host, args.port, token, signed_in=lambda: auth.account() is not None)
    return 0


def cmd_login(args: argparse.Namespace) -> int:
    auth = make_auth()
    user = auth.login_browser() if args.browser else auth.login()
    print(f"Signed in as {user}. Token cache: {auth.store.describe()}.")
    return 0


def cmd_logout(_: argparse.Namespace) -> int:
    auth = make_auth()
    auth.logout()
    print(f"Signed out; removed the token cache ({auth.store.describe()}).")
    return 0


def cmd_status(_: argparse.Namespace) -> int:
    auth = make_auth()
    print(f"Token cache: {auth.store.describe()}")
    user = auth.account()
    if not user:
        print("Not signed in. Run `graph-mcp login`.")
        return 1
    try:
        auth.token()
    except NotSignedIn:
        print(f"Signed in as {user}, but the sign-in expired. Run `graph-mcp login`.")
        return 1
    print(f"Signed in as {user}; token is valid.")
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="graph-mcp",
        description="Read-only MCP server for Microsoft Graph.",
        epilog=__doc__.split("\n", 2)[2] if __doc__ else None,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--version", action="version", version=f"graph-mcp {__version__}")
    sub = parser.add_subparsers(dest="command", metavar="<command>")

    serve = sub.add_parser("serve", help="run the MCP server (the default command)")
    serve.add_argument("--transport", choices=["stdio", "http"], default="stdio")
    serve.add_argument("--host", default="127.0.0.1", help="HTTP bind address (default: loopback only)")
    serve.add_argument("--port", type=int, default=int(_env("PORT", "8766")), help="default: $GRAPH_MCP_PORT or 8766")
    serve.add_argument("--no-auth", action="store_true", help="allow HTTP without a token, on loopback only")
    serve.set_defaults(func=cmd_serve)

    login = sub.add_parser("login", help="sign in with a device code (or --browser)")
    login.add_argument(
        "--browser",
        action="store_true",
        help="sign in through a browser on this machine instead (needs redirect URI http://localhost)",
    )
    login.set_defaults(func=cmd_login)
    sub.add_parser("logout", help="forget the cached sign-in").set_defaults(func=cmd_logout)
    sub.add_parser("status", help="show who is signed in and where the token is kept").set_defaults(func=cmd_status)
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    if not args.command:  # bare `graph-mcp` serves over stdio, so client configs stay short
        args = parser.parse_args(["serve"])
    try:
        return args.func(args)
    except (AuthError, ValueError) as e:
        print(f"graph-mcp: {e}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
