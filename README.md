# docx-mcp

A local MCP server for reading, editing, and basic styling of `.docx` files.

## Current Status
- Phase A-C completed: read/write/style tools are implemented.
- Phase D compatibility completed: tool inputs now accept legacy string-style values (e.g. `"2"`, `"true"`, `"1,2,3"`, JSON matrix strings).
- Test suite is in place with read/write/style/compat coverage.

## Project Layout
- `main.py`: CLI entry to run MCP server.
- `docx_mcp/server.py`: FastMCP server setup and tool registration.
- `docx_mcp/tools/`: MCP tool wrappers and compatibility parsers.
- `docx_mcp/services/`: core DOCX read/write/style operations.
- `tests/`: pytest-based unit and integration-style tests.

## Run Locally
```bash
uv sync --all-groups
uv run python main.py --transport stdio
```

## Run Tests
```bash
uv run pytest -q
```

## Available MCP Tools
- Read: `list_available_documents`, `get_document_info`, `get_document_text`, `get_document_outline`, `find_text_in_document`
- Write: `search_and_replace`, `add_paragraph`, `add_heading`, `add_table`
- Style: `format_table`, `set_paragraph_format`

## Notes
- The server returns structured errors in `{ ok: false, error: { code, message, details } }` format.
- For production usage, prefer sanitized test fixtures and avoid committing sensitive documents.
