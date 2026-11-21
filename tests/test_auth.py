from __future__ import annotations

import io
import stat
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import pytest

import graph_mcp.auth as auth_module
from graph_mcp.auth import SCOPES, Auth, AuthError, FileStore, KeyringStore, NotSignedIn, make_store


def test_scopes_are_read_only() -> None:
    assert all(s.endswith((".Read", ".Read.All", ".Read.Shared")) for s in SCOPES)
    assert not any("Write" in s for s in SCOPES)


def test_file_store_round_trip_is_private(tmp_path: Path) -> None:
    store = FileStore(tmp_path / "sub" / "cache.json")
    assert store.load() is None
    store.save('{"a": 1}')
    assert store.load() == '{"a": 1}'
    assert stat.S_IMODE(store.path.stat().st_mode) == 0o600
    assert stat.S_IMODE(store.path.parent.stat().st_mode) == 0o700
    store.clear()
    assert store.load() is None
    store.clear()  # clearing twice is fine


def test_make_store_choices(tmp_path: Path) -> None:
    assert isinstance(make_store("file", tmp_path / "c.json"), FileStore)
    assert isinstance(make_store("keyring"), KeyringStore)


def test_make_store_auto_falls_back_to_a_file(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(auth_module, "keyring_usable", lambda: False)
    store = make_store("auto", tmp_path / "c.json")
    assert isinstance(store, FileStore) and store.path == tmp_path / "c.json"
    monkeypatch.setattr(auth_module, "keyring_usable", lambda: True)
    assert isinstance(make_store("auto"), KeyringStore)


def test_keyring_store_round_trip(monkeypatch: pytest.MonkeyPatch) -> None:
    saved: dict[tuple[str, str], str] = {}
    monkeypatch.setattr(auth_module.keyring, "get_password", lambda svc, user: saved.get((svc, user)))
    monkeypatch.setattr(auth_module.keyring, "set_password", lambda svc, user, pw: saved.__setitem__((svc, user), pw))

    def delete(svc: str, user: str) -> None:
        if (svc, user) not in saved:
            raise auth_module.keyring.errors.PasswordDeleteError("missing")
        del saved[(svc, user)]

    monkeypatch.setattr(auth_module.keyring, "delete_password", delete)
    store = KeyringStore()
    assert store.load() is None
    store.save("{}")
    assert saved == {("graph-mcp", "token-cache"): "{}"} and store.load() == "{}"
    store.clear()
    store.clear()  # clearing twice is fine
    assert store.load() is None


def test_missing_client_id(tmp_path: Path) -> None:
    auth = Auth(client_id="", tenant="consumers", store=FileStore(tmp_path / "c.json"))
    with pytest.raises(AuthError, match="GRAPH_MCP_CLIENT_ID"):
        auth.token()


class FakeCache:
    def __init__(self, changed: bool = True) -> None:
        self.has_state_changed = changed

    def serialize(self) -> str:
        return '{"cached": true}'


class FakeApp:
    """Stands in for msal.PublicClientApplication: every sign-in or refresh returns `result`."""

    def __init__(self, result: dict | None, accounts: list[dict] | None = None, flow: dict | None = None) -> None:
        self.result = result
        self.accounts = accounts if accounts is not None else [{"username": "me@outlook.com"}]
        self.flow = flow if flow is not None else {"user_code": "ABC", "message": "Enter ABC at microsoft.com/link"}
        self.calls: list[dict] = []

    def get_accounts(self) -> list[dict]:
        return self.accounts

    def acquire_token_silent(self, scopes: list[str], account: dict) -> dict | None:
        return self.result

    def acquire_token_interactive(self, scopes: list[str], **kwargs: object) -> dict | None:
        self.calls.append({"scopes": scopes, **kwargs})
        return self.result

    def initiate_device_flow(self, scopes: list[str]) -> dict:
        self.calls.append({"scopes": scopes})
        return self.flow

    def acquire_token_by_device_flow(self, flow: dict) -> dict | None:
        return self.result


def _auth_with(app: FakeApp, tmp_path: Path, monkeypatch: pytest.MonkeyPatch, changed: bool = True) -> Auth:
    auth = Auth(client_id="id", tenant="consumers", store=FileStore(tmp_path / "c.json"))
    monkeypatch.setattr(auth, "_app", lambda: (app, FakeCache(changed)))
    return auth


SIGNED_IN = {"access_token": "t", "id_token_claims": {"preferred_username": "me@outlook.com"}}


def test_token_refreshes_and_saves(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    auth = _auth_with(FakeApp(SIGNED_IN), tmp_path, monkeypatch)
    assert auth.token() == "t"
    assert auth.store.load() == '{"cached": true}'


def test_token_saves_only_when_the_cache_changed(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    auth = _auth_with(FakeApp(SIGNED_IN), tmp_path, monkeypatch, changed=False)
    assert auth.token() == "t"
    assert auth.store.load() is None


@pytest.mark.parametrize(("result", "accounts"), [(SIGNED_IN, []), (None, None), ({"error": "invalid_grant"}, None)])
def test_token_not_signed_in(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, result: dict | None, accounts: list | None
) -> None:
    auth = _auth_with(FakeApp(result, accounts), tmp_path, monkeypatch)
    with pytest.raises(NotSignedIn, match="graph-mcp login"):
        auth.token()


def test_account_and_logout(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    auth = _auth_with(FakeApp(SIGNED_IN), tmp_path, monkeypatch)
    assert auth.account() == "me@outlook.com"
    auth.store.save("{}")
    auth.logout()
    assert auth.store.load() is None
    assert _auth_with(FakeApp(SIGNED_IN, []), tmp_path, monkeypatch).account() is None


def test_device_login(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    app = FakeApp(SIGNED_IN)
    auth = _auth_with(app, tmp_path, monkeypatch)
    out = io.StringIO()
    assert auth.login(out=out) == "me@outlook.com"
    assert "Enter ABC" in out.getvalue()
    assert app.calls[0]["scopes"] == SCOPES
    assert auth.store.load() == '{"cached": true}'


def test_device_login_cannot_start(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    auth = _auth_with(FakeApp(SIGNED_IN, flow={"error_description": "bad client"}), tmp_path, monkeypatch)
    with pytest.raises(AuthError, match="Could not start sign-in: bad client"):
        auth.login(out=io.StringIO())


def test_browser_login_saves_cache(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    app = FakeApp(SIGNED_IN)
    auth = _auth_with(app, tmp_path, monkeypatch)
    assert auth.login_browser(out=io.StringIO()) == "me@outlook.com"
    assert app.calls[0]["scopes"] == SCOPES
    assert auth.store.load() == '{"cached": true}'


def test_browser_login_failure(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    auth = _auth_with(FakeApp({"error_description": "user cancelled"}), tmp_path, monkeypatch)
    with pytest.raises(AuthError, match="user cancelled"):
        auth.login_browser(out=io.StringIO())
    assert auth.store.load() is None


def test_file_store_concurrent_saves(tmp_path: Path) -> None:
    store = FileStore(tmp_path / "cache.json")
    payloads = [f'{{"n": {i}}}' for i in range(20)]
    with ThreadPoolExecutor(max_workers=20) as pool:
        list(pool.map(store.save, payloads))  # used to fail: every save renamed the same .tmp file
    assert store.load() in payloads
    assert [p.name for p in tmp_path.iterdir()] == ["cache.json"]  # no temp files left behind
    assert stat.S_IMODE(store.path.stat().st_mode) == 0o600


class SlowRefreshApp:
    """Counts how many token refreshes run at the same time."""

    def __init__(self) -> None:
        self.running = 0
        self.max_running = 0
        self.guard = threading.Lock()

    def get_accounts(self) -> list[dict]:
        return [{"username": "me@outlook.com"}]

    def acquire_token_silent(self, scopes: list[str], account: dict) -> dict:
        with self.guard:
            self.running += 1
            self.max_running = max(self.max_running, self.running)
        time.sleep(0.02)
        with self.guard:
            self.running -= 1
        return {"access_token": "t"}


def test_parallel_token_calls_refresh_one_at_a_time(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    app = SlowRefreshApp()
    auth = Auth(client_id="id", tenant="consumers", store=FileStore(tmp_path / "c.json"))
    monkeypatch.setattr(auth, "_app", lambda: (app, FakeCache()))
    with ThreadPoolExecutor(max_workers=8) as pool:
        tokens = list(pool.map(lambda _: auth.token(), range(8)))
    assert tokens == ["t"] * 8
    assert app.max_running == 1
    assert auth.store.load() == '{"cached": true}'
