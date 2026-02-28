# Repository Guidelines

## Project Structure & Module Organization
This repository is an MCP server for DOCX read/write/style operations.

- `main.py`: MCP CLI entrypoint (`--transport stdio|sse|streamable-http`).
- `docx_mcp/server.py`: server creation and tool registration.
- `docx_mcp/services/`: core document IO and business logic (`read_ops`, `write_ops`, `style_ops`).
- `docx_mcp/tools/`: MCP-facing wrappers and compatibility parsing (`compat.py`).
- `tests/`: pytest suite (`test_read_ops.py`, `test_write_ops.py`, `test_style_ops.py`, `test_tool_compat.py`, plus new module tests).
- `.github/workflows/ci.yml`: CI pipeline (lint + typecheck + tests).

## Progress Snapshot
Completed:

1. MCP read/write/style baseline tools.
2. Legacy-compatible input parsing layer (string/bool/number/list coercion).
3. Generic document-level style engine (`apply_document_style`) for page layout + paragraph/heading baseline formatting.

In progress / just added:

1. Generic section header/footer operation (`set_headers_footers`) for:
   - section targeting,
   - header/footer text,
   - PAGE field insertion,
   - section start page number,
   - odd/even + first-page header/footer switches.

Next priority after current step:

1. Numbering and cross-reference foundations (heading numbering + TOC field management).
2. Advanced paragraph pagination controls (keep-with-next, widow/orphan, page-break-before).
3. Table structure controls (merge/split and deterministic width policies).

## Build, Test, and Development Commands
- `uv sync --all-groups`: install runtime + dev dependencies.
- `uv run python main.py --transport stdio`: run local MCP server.
- `uv add <package>`: add a dependency and update lockfile.
- `uv run pytest -q`: run tests.
- `uv run ruff check .`: lint/import-order checks.
- `uv run mypy`: type checks for `docx_mcp/` and `main.py`.
- `./scripts/check.sh`: run all required quality gates in one command.

## Lint & Check Policy
All commits must pass the following quality gates:

1. `ruff` (`uv run ruff check .`)
2. `mypy` (`uv run mypy`)
3. `pytest` (`uv run pytest -q`)

Failure in any gate blocks merge and must be fixed before pushing.

## Commit Hook Enforcement
A repository-level pre-commit hook is provided at `.githooks/pre-commit` and executes `./scripts/check.sh`.

- Enable hooks once per clone:
  - `git config core.hooksPath .githooks`
- Normal commit path:
  - `git add . && git commit -m "<type>: <message>"`
- Bypass (`--no-verify`) is only for emergency/debugging and must be justified in PR notes.

## Coding Style & Naming Conventions
- Target Python `>=3.13`.
- Follow PEP 8 with 4-space indentation.
- Use `snake_case` for functions/variables/modules and `PascalCase` for classes/dataclasses.
- Add type hints for new code; prefer `Path` over raw string paths where practical.
- Keep IO (file loading, CLI) near `main()`; keep parsing/statistics logic in small pure functions.

## Testing Guidelines
- Use `pytest` for unit and regression tests.
- Name tests as `tests/test_<feature>.py`.
- For parser/style changes, include at least one regression-style DOCX case.
- Prioritize compatibility tests for string-style MCP inputs (`"1,2,3"`, `"true"`, JSON matrix strings).

## Commit & Pull Request Guidelines
Use Conventional Commits:
- `feat: ...`, `fix: ...`, `docs: ...`, `test: ...`, `refactor: ...`
- `ci: ...`, `chore: ...`

PRs should include:
- clear summary of behavior changes,
- commands run locally (`ruff`, `mypy`, `pytest`),
- sample output or tool call payload/response for MCP behavior changes.

## Roadmap to Thesis-Grade DOCX Styling
Current implementation is functional but not yet “paper-grade”. Major gaps:

1. Full OOXML style graph handling (`styles.xml`, `numbering.xml`, inheritance resolution).
2. Section/page layout control completeness (multi-type section breaks and richer header/footer variants).
3. Auto numbering and cross-references (headings, figures, tables, TOC/LOT/LOF fields).
4. Advanced paragraph controls (widow/orphan, keep-with-next, page-break-before, tab stops).
5. Chinese academic typography conventions (mixed CJK/Latin spacing and punctuation behavior).
6. Complex table layout (merge/split cells, precise widths/heights, repeated header rows across pages).
7. Floating object handling (images/textboxes positioning and wrapping).
8. Transaction-safe editing and minimal-diff writeback for high reliability.
9. Broader regression corpus based on real thesis templates and rendering consistency checks.

## Security & Configuration Tips
- Do not commit private or sensitive source documents.
- Use sanitized fixtures for tests.
- Keep `.venv/` and temporary outputs untracked.
