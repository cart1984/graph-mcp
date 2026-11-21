"""Signing in to Microsoft (device code, or a browser on this machine) and keeping the token cache.

The cache holds a refresh token, so it lives in the OS keychain where there is one (macOS Keychain, Windows
Credential Manager, Secret Service on Linux desktops). In containers and on headless servers, where there is no
keychain, it falls back to a file readable only by the owner.
"""

from __future__ import annotations

import contextlib
import os
import sys
import tempfile
import threading
from dataclasses import dataclass, field
from pathlib import Path
from typing import Literal, Protocol

import keyring
import keyring.backends.fail
import keyring.errors
import msal

# Read-only scopes only. `offline_access` (for the refresh token) is added by MSAL itself.
# OneNote needs `Notes.Read`, not `Notes.Read.All`: the OneNote API answers every request with
# 401 (40001) when a personal account's token carries `Notes.Read.All`.
SCOPES = [
    "User.Read",
    "Mail.Read",
    "Files.Read.All",
    "Notes.Read",
    "Calendars.Read",
    "Calendars.Read.Shared",
    "Contacts.Read",
]
KEYRING_SERVICE = "graph-mcp"
KEYRING_USER = "token-cache"

CacheKind = Literal["auto", "keyring", "file"]


class NotSignedIn(Exception):
    pass


class AuthError(Exception):
    pass


class TokenStore(Protocol):
    def load(self) -> str | None: ...
    def save(self, data: str) -> None: ...
    def clear(self) -> None: ...
    def describe(self) -> str: ...


class KeyringStore:
    def load(self) -> str | None:
        return keyring.get_password(KEYRING_SERVICE, KEYRING_USER)

    def save(self, data: str) -> None:
        keyring.set_password(KEYRING_SERVICE, KEYRING_USER, data)

    def clear(self) -> None:
        with contextlib.suppress(keyring.errors.PasswordDeleteError):
            keyring.delete_password(KEYRING_SERVICE, KEYRING_USER)

    def describe(self) -> str:
        return f"keychain ({type(keyring.get_keyring()).__name__}, service '{KEYRING_SERVICE}')"


class FileStore:
    """A plain JSON file, mode 0600 in a 0700 folder. Not encrypted: protect the folder (or volume) it lives in."""

    def __init__(self, path: Path) -> None:
        self.path = path

    def load(self) -> str | None:
        return self.path.read_text(encoding="utf-8") if self.path.is_file() else None

    def save(self, data: str) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True, mode=0o700)
        # A temp file of its own (mode 0600) per save, so two writers never rename the same file away.
        fd, tmp = tempfile.mkstemp(dir=self.path.parent, prefix=f".{self.path.name}.", suffix=".tmp")
        try:
            with os.fdopen(fd, "w", encoding="utf-8") as f:
                f.write(data)
            os.replace(tmp, self.path)
        except BaseException:
            with contextlib.suppress(FileNotFoundError):
                os.unlink(tmp)
            raise

    def clear(self) -> None:
        self.path.unlink(missing_ok=True)

    def describe(self) -> str:
        return f"file {self.path}"


def default_cache_file() -> Path:
    base = os.environ.get("XDG_CONFIG_HOME") or "~/.config"
    return Path(base).expanduser() / "graph-mcp" / "token-cache.json"


def keyring_usable() -> bool:
    return not isinstance(keyring.get_keyring(), keyring.backends.fail.Keyring)


def make_store(kind: CacheKind, file: Path | None = None) -> TokenStore:
    if kind == "keyring" or (kind == "auto" and keyring_usable()):
        return KeyringStore()
    return FileStore(file or default_cache_file())


@dataclass
class Auth:
    client_id: str
    tenant: str
    store: TokenStore
    # Tools run in worker threads: without the lock, parallel calls each refresh an expired token and save at once.
    _lock: threading.Lock = field(default_factory=threading.Lock, init=False, repr=False, compare=False)

    def _app(self) -> tuple[msal.PublicClientApplication, msal.SerializableTokenCache]:
        if not self.client_id:
            raise AuthError("GRAPH_MCP_CLIENT_ID is not set: the Application (client) ID of your app registration.")
        cache = msal.SerializableTokenCache()
        data = self.store.load()
        if data:
            cache.deserialize(data)
        app = msal.PublicClientApplication(
            self.client_id, authority=f"https://login.microsoftonline.com/{self.tenant}", token_cache=cache
        )
        return app, cache

    def _persist(self, cache: msal.SerializableTokenCache) -> None:
        if cache.has_state_changed:
            self.store.save(cache.serialize())

    def token(self) -> str:
        """An access token, refreshed silently from the cache. Raises NotSignedIn if a new sign-in is needed."""
        with self._lock:  # the second caller then finds the token the first one refreshed
            app, cache = self._app()
            accounts = app.get_accounts()
            result = app.acquire_token_silent(SCOPES, account=accounts[0]) if accounts else None
            self._persist(cache)
        if not result or "access_token" not in result:
            raise NotSignedIn(
                "Not signed in to Microsoft (or the sign-in expired). Run `graph-mcp login` in a terminal."
            )
        return result["access_token"]

    def account(self) -> str | None:
        app, _ = self._app()
        accounts = app.get_accounts()
        return accounts[0].get("username") if accounts else None

    def login(self, out=sys.stdout) -> str:
        """Device-code sign-in: prints a code to enter at microsoft.com/link, waits, returns the username."""
        app, cache = self._app()
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise AuthError(f"Could not start sign-in: {flow.get('error_description', flow)}")
        print(flow["message"], file=out, flush=True)
        return self._finish(cache, app.acquire_token_by_device_flow(flow))

    def login_browser(self, out=sys.stdout, timeout: int = 300) -> str:
        """Browser sign-in: opens Microsoft's sign-in page and catches the answer on http://localhost.

        Needs a browser on this machine and the redirect URI `http://localhost` under the app's
        "Mobile and desktop applications" platform. Returns the username.
        """
        app, cache = self._app()
        print("Opening Microsoft sign-in in your browser...", file=out, flush=True)
        result = app.acquire_token_interactive(SCOPES, prompt="select_account", timeout=timeout)
        return self._finish(cache, result)

    def _finish(self, cache: msal.SerializableTokenCache, result: dict) -> str:
        if "access_token" not in result:
            raise AuthError(f"Sign-in failed: {result.get('error_description', result)}")
        self._persist(cache)
        return result.get("id_token_claims", {}).get("preferred_username", "?")

    def logout(self) -> None:
        self.store.clear()
