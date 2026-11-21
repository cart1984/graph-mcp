from __future__ import annotations

from pathlib import Path

import pytest

from graph_mcp import cli
from graph_mcp.auth import Auth, FileStore, NotSignedIn


@pytest.fixture(autouse=True)
def clean_env(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv("GRAPH_MCP_TIMEZONE", raising=False)
    monkeypatch.delenv("TZ", raising=False)
    monkeypatch.setattr(cli, "system_timezone", lambda: "Europe/Berlin")


def test_timezone_setting_wins(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("GRAPH_MCP_TIMEZONE", "Asia/Tokyo")
    monkeypatch.setenv("TZ", "America/Chicago")
    assert cli.resolve_timezone() == "Asia/Tokyo"


def test_timezone_falls_back_to_tz(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("TZ", ":America/Chicago")  # the leading colon is allowed in TZ
    assert cli.resolve_timezone() == "America/Chicago"


def test_timezone_falls_back_to_system_then_utc(monkeypatch: pytest.MonkeyPatch) -> None:
    assert cli.resolve_timezone() == "Europe/Berlin"
    monkeypatch.setattr(cli, "system_timezone", lambda: None)
    assert cli.resolve_timezone() == "UTC"


def test_unknown_timezone_stops_serve(monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]) -> None:
    monkeypatch.setenv("GRAPH_MCP_TIMEZONE", "Mars/Olympus")
    monkeypatch.setenv("GRAPH_MCP_CLIENT_ID", "id")
    assert cli.main(["serve"]) == 2
    assert "Unknown time zone 'Mars/Olympus'" in capsys.readouterr().err


class StubAuth:
    def __init__(self, user: str | None, valid: bool = True) -> None:
        self.user, self.valid, self.logged_out = user, valid, False
        self.store = type("Store", (), {"describe": lambda self: "file /x/cache.json"})()

    def account(self) -> str | None:
        return self.user

    def token(self) -> str:
        if not self.valid:
            raise NotSignedIn("expired")
        return "t"

    def logout(self) -> None:
        self.logged_out = True


@pytest.mark.parametrize(
    ("auth", "code", "message"),
    [
        (StubAuth("me@outlook.com"), 0, "Signed in as me@outlook.com; token is valid."),
        (StubAuth(None), 1, "Not signed in. Run `graph-mcp login`."),
        (StubAuth("me@outlook.com", valid=False), 1, "but the sign-in expired"),
    ],
)
def test_status(
    monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str], auth: StubAuth, code: int, message: str
) -> None:
    monkeypatch.setattr(cli, "make_auth", lambda: auth)
    assert cli.main(["status"]) == code
    out = capsys.readouterr().out
    assert "Token cache: file /x/cache.json" in out and message in out


def test_logout(monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]) -> None:
    auth = StubAuth("me@outlook.com")
    monkeypatch.setattr(cli, "make_auth", lambda: auth)
    assert cli.main(["logout"]) == 0
    assert auth.logged_out and "Signed out" in capsys.readouterr().out


def test_login_browser_flag(monkeypatch: pytest.MonkeyPatch) -> None:
    called: list[str] = []
    monkeypatch.setattr(Auth, "login", lambda self, out=None: called.append("device") or "u")
    monkeypatch.setattr(Auth, "login_browser", lambda self, out=None: called.append("browser") or "u")
    monkeypatch.setenv("GRAPH_MCP_CLIENT_ID", "id")
    assert cli.main(["login", "--browser"]) == 0
    assert cli.main(["login"]) == 0
    assert called == ["browser", "device"]


def test_bad_token_cache_setting(monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]) -> None:
    monkeypatch.setenv("GRAPH_MCP_TOKEN_CACHE", "cloud")
    assert cli.main(["status"]) == 2
    assert "GRAPH_MCP_TOKEN_CACHE must be auto, keyring or file" in capsys.readouterr().err


def test_make_auth_reads_settings(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    monkeypatch.setenv("GRAPH_MCP_CLIENT_ID", "abc")
    monkeypatch.setenv("GRAPH_MCP_TENANT", "organizations")
    monkeypatch.setenv("GRAPH_MCP_TOKEN_CACHE", "file")
    monkeypatch.setenv("GRAPH_MCP_CACHE_FILE", str(tmp_path / "c.json"))
    auth = cli.make_auth()
    assert (auth.client_id, auth.tenant) == ("abc", "organizations")
    assert isinstance(auth.store, FileStore) and auth.store.path == tmp_path / "c.json"


def test_http_without_token_is_refused(monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]) -> None:
    monkeypatch.setenv("GRAPH_MCP_CLIENT_ID", "id")
    monkeypatch.delenv("GRAPH_MCP_TOKEN", raising=False)
    assert cli.main(["serve", "--transport", "http", "--host", "0.0.0.0"]) == 2
    assert "HTTP needs a token" in capsys.readouterr().err
