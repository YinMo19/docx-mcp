#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT_DIR"

if ! command -v uv >/dev/null 2>&1; then
  echo "[check] uv is required but not found in PATH" >&2
  exit 1
fi

echo "[check] Running ruff..."
uv run ruff check .

echo "[check] Running mypy..."
uv run mypy

echo "[check] Running pytest..."
uv run pytest -q

echo "[check] All checks passed."
