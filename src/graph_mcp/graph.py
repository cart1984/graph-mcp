"""A Microsoft Graph client that can only read: it has no method other than GET."""

from __future__ import annotations

import io
import re
from collections.abc import Callable
from pathlib import Path
from typing import Any

import httpx

GRAPH = "https://graph.microsoft.com/v1.0"
MAX_TEXT = 60_000  # characters returned inline; the full file stays on disk
MAX_CONVERT_BYTES = 25 * 1024 * 1024


class GraphError(Exception):
    pass


class Graph:
    def __init__(self, token: Callable[[], str], http: httpx.Client | None = None) -> None:
        self._token = token
        self._http = http or httpx.Client(timeout=60, follow_redirects=True)

    def _get(self, path: str, params: dict[str, Any] | None, headers: dict[str, str] | None) -> httpx.Response:
        url = path if path.startswith("https://") else GRAPH + path
        h = {"Authorization": f"Bearer {self._token()}", **(headers or {})}
        r = self._send(url, params, h)
        if r.status_code >= 400:
            try:
                err = r.json().get("error", {})
                msg = f"{err.get('code')}: {err.get('message')}"
            except ValueError:
                msg = r.text[:500]
            raise GraphError(f"Graph request failed ({r.status_code}) {msg}")
        return r

    def _send(self, url: str, params: dict[str, Any] | None, headers: dict[str, str]) -> httpx.Response:
        """One GET, retried once on a timeout: OneNote in particular can stall on its first request."""
        for attempt in range(2):
            try:
                return self._http.get(url, params=params, headers=headers)
            except httpx.TimeoutException as e:
                if attempt:
                    raise GraphError("Microsoft Graph didn't answer in time (twice). Try again in a minute.") from e
            except httpx.TransportError as e:
                raise GraphError(f"Couldn't reach Microsoft Graph: {e}") from e
        raise AssertionError("unreachable")

    def get(self, path: str, params: dict[str, Any] | None = None, headers: dict[str, str] | None = None) -> dict:
        return self._get(path, params, headers).json()

    def get_bytes(self, path: str) -> bytes:
        return self._get(path, None, None).content


def clip(text: str) -> str:
    if len(text) <= MAX_TEXT:
        return text
    return text[:MAX_TEXT] + f"\n\n[... truncated, {len(text) - MAX_TEXT:,} more characters]"


def safe_name(name: str) -> str:
    return re.sub(r"[^\w.\- ]", "_", name).strip(" .") or "file"


def to_text(data: bytes, name: str) -> str:
    """Convert a file's bytes (docx, xlsx, pptx, pdf, html, txt, ...) to Markdown, in memory. The name picks the
    converter by its extension."""
    from markitdown import MarkItDown, StreamInfo

    suffix = Path(name).suffix.lower()
    try:
        info = StreamInfo(extension=suffix or None, filename=name)
        text = MarkItDown().convert_stream(io.BytesIO(data), stream_info=info).text_content
    except Exception as e:  # unsupported or broken file
        return f"[Could not convert {suffix or 'the file'} to text: {e}]"
    if looks_binary(text):  # MarkItDown falls back to plain text for unknown types, e.g. a .zip
        return f"[Can't convert {suffix or 'the file'} to text: it looks like a binary file.]"
    return text


def looks_binary(text: str, sample: int = 4000) -> bool:
    """True if text has NUL characters or is more than 10% unprintable (tabs and line breaks count as printable)."""
    head = text[:sample]
    if "\x00" in head:
        return True
    unprintable = sum(1 for c in head if not (c.isprintable() or c in "\n\r\t"))
    return unprintable > len(head) * 0.1


def unique_path(folder: Path, name: str, data: bytes) -> Path:
    """`folder/name`, or `name (2)`, `name (3)`, ... if a different file already has that name. The same file saved
    again reuses its path."""
    path = folder / safe_name(name)
    n = 2
    while path.exists() and path.read_bytes() != data:
        stem, suffix = Path(safe_name(name)).stem, Path(safe_name(name)).suffix
        path = folder / f"{stem} ({n}){suffix}"
        n += 1
    return path


def convert(data: bytes, name: str, save_to: Path | None = None) -> str:
    """The file's text. With `save_to`, the original is also saved there (for local use, see GRAPH_MCP_DOWNLOADS)."""
    head = f"{name} ({len(data):,} bytes)."
    if save_to is not None:
        save_to.mkdir(parents=True, exist_ok=True)
        path = unique_path(save_to, name, data)
        path.write_bytes(data)
        head = f"Saved to {path} ({len(data):,} bytes)."
    if len(data) > MAX_CONVERT_BYTES:
        return head + " Too large to convert to text."
    return head + "\n\n" + clip(to_text(data, name))
