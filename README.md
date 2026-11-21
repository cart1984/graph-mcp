# graph-mcp

A small **read-only** MCP server for Microsoft Graph. It gives any MCP client (Claude Code, Claude Desktop, Codex, VS Code, Cursor, Gemini CLI, or your own agent) access to a Microsoft account's:

- **Outlook mail:** search, list, read messages, read attachments
- **OneDrive:** browse, search, read files, including files others shared with you
- **OneNote:** notebooks, sections, pages
- **Calendar** events and **contacts**

Word, Excel, PowerPoint and PDF files are converted to text in memory (with [MarkItDown](https://github.com/microsoft/markitdown)); nothing is written to disk unless you ask for it. All dates and times, shown and given, are in your time zone.

It works with **personal Microsoft accounts** (outlook.com, hotmail.com, live.com) and, with a different tenant setting, with work or school accounts.

Status: early, personal project. Licensed under MIT.

## Read-only, twice over

1. **The sign-in only grants read permissions:** `User.Read`, `Mail.Read`, `Files.Read.All`, `Notes.Read`, `Calendars.Read`, `Calendars.Read.Shared`, `Contacts.Read` (plus `offline_access` for staying signed in). Microsoft refuses any write with this token, whatever the code does.
2. **The code only reads:** the Graph client has no method except GET, and every tool is marked read-only for MCP clients.

Content of mails, files and notes reaches the agent as data. The server's instructions tell agents to treat it as data, never as instructions, but as with any tool that reads untrusted text, keep that in mind when you let an agent act on what it read.

## Setup

### 1. Register an app with Microsoft (once, about 5 minutes)

Both ways of running it need your own app registration, which gives you a client ID. It's free.

1. Sign in at [portal.azure.com](https://portal.azure.com) and open **Microsoft Entra ID**.
   - With a personal account you may see *"Selected user account does not exist in tenant 'Microsoft Services'"*: your account has no directory of its own. Create a [free Azure account](https://azure.microsoft.com/free) with the same Microsoft account; that gives you a *Default Directory*. App registrations cost nothing; you're only charged if you upgrade to pay-as-you-go and create paid resources.
2. **App registrations → New registration.** Name: anything (e.g. `graph-mcp`). Supported account types: **Personal Microsoft accounts only** (or "any organizational directory and personal accounts"). Leave the redirect URI empty for now. **Register.**
3. **Authentication:** set **Allow public client flows** to **Yes**. Then **Add a platform → Mobile and desktop applications** with the custom redirect URI `http://localhost`, for the browser sign-in (see the sign-in steps below). Save.
4. **API permissions → Add a permission → Microsoft Graph → Delegated permissions:** `Mail.Read`, `Files.Read.All`, `Notes.Read`, `Calendars.Read`, `Calendars.Read.Shared`, `Contacts.Read`, `offline_access` (`User.Read` is already there). Take `Notes.Read`, not `Notes.Read.All`: OneNote rejects every request from a personal account whose token has `Notes.Read.All`. With a personal account there's no admin consent; you approve them when you sign in.
5. Copy the **Application (client) ID** from **Overview**. It's an identifier, not a secret. Don't create a client secret: this is a public client.

### 2. Choose how to run it

| | **A. On this machine** | **B. In Docker** |
|---|---|---|
| For | Agents on the same computer | A home server that agents on several machines use |
| Agents talk to it over | stdio (each agent starts `graph-mcp`) | HTTP with a Bearer token |
| The sign-in lives in | Your OS keychain | A Docker volume |
| Steps | A1–A3 below | B1–B5 below |

You can use both; each keeps its own sign-in.

### Option A: on this machine

**A1. Install** with [uv](https://docs.astral.sh/uv/), as a command on your `PATH`:

```bash
uv tool install git+https://github.com/cart1984/graph-mcp
```

Or from a local checkout (edits take effect right away): `uv tool install --editable ~/projects/graph-mcp`.

**A2. Sign in** (once per machine):

```bash
GRAPH_MCP_CLIENT_ID=<your client id> graph-mcp login
```

It prints a code; enter it at [microsoft.com/devicelogin](https://microsoft.com/devicelogin) and approve the permissions.

If the page rejects the code (*"The code you entered has expired"* or *"That code didn't work"* right after you enter it), sign in through the browser instead. It opens Microsoft's sign-in page and catches the answer on `http://localhost` (the redirect URI from step 1.3):

```bash
GRAPH_MCP_CLIENT_ID=<your client id> graph-mcp login --browser
```

`graph-mcp status` shows who is signed in. The sign-in stays valid as long as it's used at least every 90 days or so; if it lapses, tools answer with "run `graph-mcp login`". All agents on the machine share this one sign-in.

**A3. Add it to your agents.** Every client starts the same command, `graph-mcp`, with `GRAPH_MCP_CLIENT_ID` set. Personal accounts need nothing else.

Claude Code:

```bash
claude mcp add --scope user graph -e GRAPH_MCP_CLIENT_ID=<id> -- graph-mcp
```

Claude Desktop, Cursor, Gemini CLI (`claude_desktop_config.json`, `~/.cursor/mcp.json`, `~/.gemini/settings.json`):

```json
{
  "mcpServers": {
    "graph": {
      "command": "graph-mcp",
      "env": { "GRAPH_MCP_CLIENT_ID": "<id>" }
    }
  }
}
```

GUI apps on macOS may not see your shell's `PATH`; if the server doesn't start, use the full path from `which graph-mcp` (usually `~/.local/bin/graph-mcp`).

VS Code (user or workspace `mcp.json`):

```json
{
  "servers": {
    "graph": {
      "type": "stdio",
      "command": "graph-mcp",
      "env": { "GRAPH_MCP_CLIENT_ID": "<id>" }
    }
  }
}
```

Codex (`~/.codex/config.toml`):

```toml
[mcp_servers.graph]
command = "graph-mcp"
env = { GRAPH_MCP_CLIENT_ID = "<id>" }
```

### Option B: in Docker

`compose.yaml` runs the HTTP server in a locked-down container: read-only file system, no capabilities, a non-root user, the port on one address only, and the sign-in in a named volume. Files are converted in memory and never kept.

**B1. Get the code and write the settings** on the machine that will run it:

```bash
git clone https://github.com/cart1984/graph-mcp && cd graph-mcp
echo "GRAPH_MCP_CLIENT_ID=<your client id>" > .env
echo "GRAPH_MCP_TOKEN=$(openssl rand -hex 32)" >> .env
echo "GRAPH_MCP_BIND=<this machine's address on your clients' network>" >> .env
echo "TZ=America/New_York" >> .env   # your time zone
chmod 600 .env
```

`GRAPH_MCP_TOKEN` is what agents must send; anyone who has it can read your mail. `GRAPH_MCP_BIND` is required: the port is published on that address only, never on every interface.

**B2. Build** (the image is about 1 GB, mostly the document converters' dependencies):

```bash
docker compose build
```

**B3. Sign in** (once). With a device code:

```bash
docker compose run --rm graph login    # enter the code at microsoft.com/devicelogin
```

If the page rejects the code, sign in through a browser on a computer that has graph-mcp installed (step A1), into a file instead of the keychain, then copy that file into the volume:

```bash
# on the computer with the browser:
GRAPH_MCP_CLIENT_ID=<id> GRAPH_MCP_TOKEN_CACHE=file graph-mcp login --browser   # writes ~/.config/graph-mcp/token-cache.json
# put that file into the graph-mcp folder on the server (scp, or cp on the same machine), then there:
docker compose run --rm -T --entrypoint sh graph -c 'umask 077 && cat > /data/token-cache.json' < token-cache.json
rm token-cache.json   # and the one in ~/.config/graph-mcp/ on the other computer: keep no copies around
```

**B4. Start it and check:**

```bash
docker compose up -d
docker compose run --rm graph status   # "Signed in as ...; token is valid."
docker compose ps                      # "healthy" once it's up
```

The container counts as healthy only while it has a sign-in (`/healthz` answers 503 "not signed in" otherwise), so `docker compose ps` also tells you when the sign-in is gone.

**B5. Add it to your agents.** They connect to `http://<GRAPH_MCP_BIND>:8766/mcp` with the header `Authorization: Bearer <GRAPH_MCP_TOKEN>`. Claude Code:

```bash
claude mcp add --scope user --transport http graph http://<host>:8766/mcp --header "Authorization: Bearer <token>"
```

Other clients that support remote (Streamable HTTP) servers take the same URL and header in their config; for a client that only speaks stdio, bridge it with [`mcp-remote`](https://www.npmjs.com/package/mcp-remote).

**This is plain HTTP: the token and everything you read travel unencrypted.** That's fine on a network you trust and inside a VPN. Anywhere else, put a TLS proxy in front (Caddy, for example) or don't expose it.

**Updating:** `git pull && docker compose up -d --build`. The sign-in stays in the volume.

## Other ways to run

**HTTP without Docker**, e.g. on a server without containers:

```bash
GRAPH_MCP_CLIENT_ID=<id> GRAPH_MCP_TOKEN=$(openssl rand -hex 32) graph-mcp serve --transport http --host <address>
```

The same URL, header and plain-HTTP warning apply as in B5. A token is required except on loopback, where `--no-auth` is allowed.

**The sign-in as a file.** Where there's no keychain (containers, headless Linux), the sign-in is kept in a file, `~/.config/graph-mcp/token-cache.json`, mode 0600. It holds a refresh token and isn't encrypted: anyone who can read it can read your mail, so keep it (or the Docker volume) where only you can.

**Keeping the originals.** By default, files, attachments and OneNote pages are only converted to text. Set `GRAPH_MCP_DOWNLOADS` to a folder to also keep the originals there, for an agent on the same machine that needs the file itself (for example a long document whose text was cut off).

## Settings

All settings are environment variables:

| Variable | Default | What |
|---|---|---|
| `GRAPH_MCP_CLIENT_ID` | (required) | Application (client) ID of your app registration |
| `GRAPH_MCP_TENANT` | `consumers` | `consumers` for personal accounts; `organizations`, `common` or a tenant ID for work/school |
| `GRAPH_MCP_TOKEN_CACHE` | `auto` | `auto` (keychain if there is one, else a file), `keyring` or `file` |
| `GRAPH_MCP_CACHE_FILE` | `~/.config/graph-mcp/token-cache.json` | The sign-in file, when a file is used |
| `GRAPH_MCP_DOWNLOADS` | (none: keep nothing) | Folder to keep the originals of files, attachments and pages in |
| `GRAPH_MCP_TIMEZONE` | `TZ`, else the system's | IANA time zone (e.g. `America/New_York`) for every date shown and given; an unknown name stops the server |
| `GRAPH_MCP_TOKEN` | (none) | Bearer token HTTP clients must send; at least 32 characters |
| `GRAPH_MCP_PORT` | `8766` | HTTP port |
| `GRAPH_MCP_BIND` | (required in Docker) | The address `compose.yaml` publishes the port on |

## Tools

| Tool | What it does |
|---|---|
| `whoami` | Which account is signed in |
| `mail_folders` | Folders with unread/total counts |
| `mail_search` | Search all mail, Outlook syntax (`from:amazon`, `subject:"tax"`, `hasattachments:true`) |
| `mail_list` | Newest messages in a folder, optionally unread only or since a date |
| `mail_get` | One message in full, with its attachments listed |
| `mail_attachment` | One attachment as text |
| `onedrive_list` | List a folder by path |
| `onedrive_search` | Search by name and content |
| `onedrive_shared` | Files and folders shared with you |
| `onedrive_read` | One file as text, by id or path |
| `onenote_notebooks` | Notebooks and their sections |
| `onenote_pages` | Pages, newest first, optionally per section or by title |
| `onenote_read` | One page as text |
| `calendar_events` | Events in a date range (both days included), multi-day and all-day events shown with their days |
| `contacts_search` | Contacts by name |

## Development

```bash
uv sync
uv run pytest
uv run ruff check && uv run ruff format --check
uv run pyright
```

Tests run against a fake Graph and never touch the network; the fake fails any test whose code sends anything but GET. CI also builds the Docker image, starts the compose stack and checks its lockdown.
