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

# Pictures pasted into text as data URIs: thousands of characters of base64 that no reader can use.
DATA_URI = re.compile(r"data:[\w.+-]+/[\w.+-]+;base64,[A-Za-z0-9+/=]+")
# Characters that show as nothing: newsletters pad their preview text with them (zero-width spaces and joiners,
# the combining grapheme joiner, soft hyphens, byte order marks).
INVISIBLE = re.compile("[\u00ad\u034f\u200b\u200c\u200e\u200f\u2060\ufeff]")
NOT_A_VALUE = {"NaN", "nan", "None"}  # how pandas, behind MarkItDown's Excel converter, writes empty cells
UNNAMED = re.compile(r"Unnamed: \d+")  # ... and headers of columns without one


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
    if suffix in (".xlsx", ".xls"):
        text = tidy_tables(text)
    text = tidy(text)
    if not text.strip():
        if suffix == ".pdf":
            return "[No text in this PDF: it's probably scanned, or made of pictures.]"
        return "[No text in this file.]"
    return text


def tidy(text: str) -> str:
    """Text without pasted-in pictures (data URIs) and invisible padding characters."""
    return INVISIBLE.sub("", DATA_URI.sub("image", text))


def _cells(line: str) -> list[str]:
    return [c.strip() for c in re.split(r"(?<!\\)\|", line.strip()[1:-1])]


def _table(lines: list[str]) -> list[str]:
    """A Markdown table without empty cells spelled NaN, unnamed column headers, and empty rows and columns."""
    rows = [_cells(line) for line in lines]
    rule = len(rows) > 1 and all(re.fullmatch(r":?-+:?", c) for c in rows[1] if c)
    if rule:
        del rows[1]
    rows = [["" if c in NOT_A_VALUE or UNNAMED.fullmatch(c) else c for c in r] for r in rows]
    width = max(len(r) for r in rows)
    rows = [r + [""] * (width - len(r)) for r in rows]
    rows = rows[:1] + [r for r in rows[1:] if any(r)]
    if not any(rows[0]) and len(rows) > 1:  # no header at all (the sheet's headings sit lower): use the first row
        rows = rows[1:]
    keep = [i for i in range(width) if any(r[i] for r in rows)]
    if not keep:
        return []
    out = ["| " + " | ".join(r[i] for i in keep) + " |" for r in rows]
    if rule:
        out.insert(1, "| " + " | ".join("---" for _ in keep) + " |")
    return out


def tidy_tables(text: str) -> str:
    """Every Markdown table in `text` tidied (see `_table`); the rest stays as it is."""
    out: list[str] = []
    block: list[str] = []
    for line in [*text.splitlines(), ""]:
        if line.startswith("|") and line.rstrip().endswith("|") and len(line.strip()) > 1:
            block.append(line)
            continue
        if block:
            out += _table(block)
            block = []
        out.append(line)
    return "\n".join(out[:-1])


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
