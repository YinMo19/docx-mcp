from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.document import Document as DocxDocument

from docx_mcp.errors import DocxMCPError


def resolve_directory(directory: str | None) -> Path:
    base = Path.cwd() if directory is None else Path(directory)
    resolved = base.expanduser().resolve()
    if not resolved.exists():
        raise DocxMCPError(
            code="DIRECTORY_NOT_FOUND",
            message=f"Directory does not exist: {resolved}",
        )
    if not resolved.is_dir():
        raise DocxMCPError(
            code="NOT_A_DIRECTORY",
            message=f"Path is not a directory: {resolved}",
        )
    return resolved


def resolve_docx_path(filename: str) -> Path:
    path = Path(filename).expanduser()
    if not path.is_absolute():
        path = Path.cwd() / path
    resolved = path.resolve()
    if not resolved.exists():
        raise DocxMCPError(
            code="FILE_NOT_FOUND",
            message=f"DOCX file does not exist: {resolved}",
        )
    if not resolved.is_file():
        raise DocxMCPError(
            code="NOT_A_FILE",
            message=f"Path is not a file: {resolved}",
        )
    if resolved.suffix.lower() != ".docx":
        raise DocxMCPError(
            code="INVALID_EXTENSION",
            message=f"Only .docx files are supported: {resolved}",
        )
    return resolved


def list_docx_files(directory: Path) -> list[Path]:
    return sorted(
        [p for p in directory.iterdir() if p.is_file() and p.suffix.lower() == ".docx"],
        key=lambda p: p.name.lower(),
    )


def open_document(path: Path) -> DocxDocument:
    try:
        return Document(str(path))
    except Exception as exc:  # pragma: no cover - depends on external file validity
        raise DocxMCPError(
            code="DOCX_OPEN_FAILED",
            message=f"Failed to open DOCX: {path}",
            details={"error": str(exc)},
        ) from exc


def resolve_output_docx_path(path: str | Path) -> Path:
    output = Path(path).expanduser()
    if not output.is_absolute():
        output = Path.cwd() / output
    resolved = output.resolve()
    if resolved.suffix.lower() != ".docx":
        raise DocxMCPError(
            code="INVALID_EXTENSION",
            message=f"Output must be a .docx file: {resolved}",
        )
    if not resolved.parent.exists():
        raise DocxMCPError(
            code="DIRECTORY_NOT_FOUND",
            message=f"Output directory does not exist: {resolved.parent}",
        )
    return resolved


def resolve_for_write(filename: str, output_filename: str | None = None) -> tuple[Path, Path]:
    src_path = resolve_docx_path(filename)
    dst_path = src_path if output_filename is None else resolve_output_docx_path(output_filename)
    return src_path, dst_path


def save_document(doc: DocxDocument, output_path: Path) -> None:
    try:
        doc.save(str(output_path))
    except Exception as exc:  # pragma: no cover - depends on fs/docx internals
        raise DocxMCPError(
            code="DOCX_SAVE_FAILED",
            message=f"Failed to save DOCX: {output_path}",
            details={"error": str(exc)},
        ) from exc
