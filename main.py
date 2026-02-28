from __future__ import annotations

import argparse

from docx_mcp.server import run_server


def main() -> None:
    parser = argparse.ArgumentParser(description="Run DOCX MCP server.")
    parser.add_argument(
        "--transport",
        choices=("stdio", "sse", "streamable-http"),
        default="stdio",
        help="MCP transport type.",
    )
    args = parser.parse_args()
    run_server(transport=args.transport)


if __name__ == "__main__":
    main()

