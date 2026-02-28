# Repository Guidelines

## Project Structure & Module Organization
This repository is currently a small `uv`-managed Python project for DOCX parsing experiments.

- `main.py`: executable entry point; contains parsing, summarization, and reporting logic.
- `pyproject.toml`: project metadata and dependencies.
- `uv.lock`: pinned dependency lockfile (commit this file).
- `README.md`: reserved for user-facing overview.

As the project grows, move runtime code into a package such as `docx_mcp/` and keep `main.py` as a thin CLI wrapper. Add tests under `tests/` and fixture documents under `tests/fixtures/`.

## Build, Test, and Development Commands
- `uv sync`: create/update the virtual environment from `pyproject.toml` and `uv.lock`.
- `uv run python main.py /path/to/file.docx`: run the parser on a target document.
- `uv run python main.py`: run with the script’s default DOCX path.
- `uv add <package>`: add a dependency and update lockfile.
- `uv run pytest`: run tests (once test files are added).

## Coding Style & Naming Conventions
- Target Python `>=3.13`.
- Follow PEP 8 with 4-space indentation.
- Use `snake_case` for functions/variables/modules and `PascalCase` for classes/dataclasses.
- Add type hints for new code; prefer `Path` over raw string paths where practical.
- Keep IO (file loading, CLI) near `main()`; keep parsing/statistics logic in small pure functions.

## Testing Guidelines
- Use `pytest` for unit and regression tests.
- Name tests as `tests/test_<feature>.py`.
- Add fixture-based tests for:
  - paragraph/style extraction,
  - table structure counting,
  - Chinese text handling and Unicode stability.
- For parser changes, include at least one regression test against a real `.docx` fixture.

## Commit & Pull Request Guidelines
There is no established commit history yet; use Conventional Commits going forward:
- `feat: ...`, `fix: ...`, `docs: ...`, `test: ...`, `refactor: ...`

PRs should include:
- clear summary of behavior changes,
- commands run locally (e.g., `uv run python main.py ...`, `uv run pytest`),
- sample output for parser/reporting changes.

## Security & Configuration Tips
- Do not commit private or sensitive source documents.
- Use sanitized fixtures for tests.
- Keep `.venv/` and temporary outputs untracked.
