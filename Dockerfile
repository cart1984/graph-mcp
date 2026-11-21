# graph-mcp in a container: Streamable HTTP on port 8766, the sign-in kept in the /data volume. Files and
# attachments are converted in memory; nothing is kept (GRAPH_MCP_DOWNLOADS is unset).
#
#   docker compose run --rm graph login    # once: device-code sign-in, writes /data/token-cache.json
#   (or sign in with `graph-mcp login --browser` on a desktop and copy the cache file in: see README)
#   docker compose up -d
#
# The health check (/healthz) is healthy only while there's a sign-in: `docker ps` shows "unhealthy" until you
# sign in, and again if the sign-in is removed.

FROM python:3.14-slim

COPY --from=ghcr.io/astral-sh/uv:0.13.7 /uv /bin/uv

RUN apt-get update \
    && apt-get install -y --no-install-recommends tzdata \
    && rm -rf /var/lib/apt/lists/* \
    && useradd --uid 1000 --create-home --shell /usr/sbin/nologin graph

ENV UV_COMPILE_BYTECODE=1 \
    UV_LINK_MODE=copy \
    UV_PROJECT_ENVIRONMENT=/app/.venv \
    PATH=/app/.venv/bin:$PATH \
    GRAPH_MCP_TOKEN_CACHE=file \
    GRAPH_MCP_CACHE_FILE=/data/token-cache.json \
    GRAPH_MCP_PORT=8766

WORKDIR /app
COPY pyproject.toml uv.lock README.md LICENSE ./
RUN uv sync --locked --no-dev --no-install-project
COPY src ./src
RUN uv sync --locked --no-dev --no-editable

RUN mkdir -p /data && chown graph:graph /data && chmod 700 /data
USER graph

EXPOSE 8766
HEALTHCHECK --interval=30s --timeout=5s --start-period=20s \
    CMD ["python", "-c", "import os, urllib.request; urllib.request.urlopen(f'http://127.0.0.1:{os.environ[\"GRAPH_MCP_PORT\"]}/healthz', timeout=4)"]

ENTRYPOINT ["graph-mcp"]
CMD ["serve", "--transport", "http", "--host", "0.0.0.0"]
