"""Utilities to convert Word documents (``.doc``/``.docx``) into PDF files.

This module offers a small command line interface that searches for Word
documents inside a directory and converts each one into a PDF by using the
``libreoffice``/``soffice`` headless converter.  When some documents cannot be
converted because of read-only or security restrictions, the converter will try
to bypass the issue by copying the file to a temporary location, removing
restrictive file permissions and re-running the conversion from there.  This is
often enough to get around the “This document is read-only/restricted” prompts
that stop the automated conversion.

Example
-------

```
python convert_doc_to_pdf.py ./documentos --output ./pdfs
```

The command above will convert every ``.doc`` (and ``.docx``) file found inside
``./documentos`` and place the resulting PDFs into ``./pdfs`` while preserving
the directory structure.  Conversion errors are logged and the script continues
with the remaining files.

The conversion relies on ``libreoffice``/``soffice`` being available in the
system path.  On Windows the LibreOffice installation usually exposes the
binary as ``soffice.exe``.  Feel free to pass the explicit path through the
``--office-binary`` argument if the executable cannot be found automatically.
"""

from __future__ import annotations

import argparse
import logging
import os
import shutil
import stat
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List


LOGGER = logging.getLogger(__name__)


class ConversionError(RuntimeError):
    """Raised when a document cannot be converted into a PDF."""


@dataclass
class ConversionResult:
    """Stores the outcome of a single document conversion."""

    source: Path
    destination: Path | None
    succeeded: bool
    error: str | None = None


def discover_documents(root: Path, recursive: bool = True) -> List[Path]:
    """Return a sorted list with all doc/docx files inside *root*.

    Parameters
    ----------
    root:
        Directory where the search starts.
    recursive:
        Whether to scan subdirectories.
    """

    patterns = ("*.doc", "*.DOC", "*.docx", "*.DOCX")
    if recursive:
        files = [
            path
            for pattern in patterns
            for path in root.rglob(pattern)
            if path.is_file()
        ]
    else:
        files = [
            path
            for pattern in patterns
            for path in root.glob(pattern)
            if path.is_file()
        ]

    # ``sorted`` keeps the execution order deterministic which is nice for
    # logging and testing purposes.
    return sorted(files)


def ensure_office_binary(binary_name: str) -> str:
    """Return the path to the LibreOffice binary or fail with a clear message."""

    resolved = shutil.which(binary_name)
    if resolved:
        return resolved

    raise FileNotFoundError(
        "Não foi possível encontrar o executável do LibreOffice/soffice. "
        "Informe o caminho com --office-binary ou instale o LibreOffice."
    )


def run_libreoffice(
    office_binary: str,
    document: Path,
    output_dir: Path,
) -> None:
    """Invoke LibreOffice in headless mode to convert a document to PDF."""

    # ``--convert-to pdf"`` writes the resulting PDF into *output_dir*.
    command = [
        office_binary,
        "--headless",
        "--convert-to",
        "pdf",
        "--outdir",
        str(output_dir),
        str(document),
    ]

    LOGGER.debug("Executando comando: %s", " ".join(command))
    subprocess.run(command, check=True, capture_output=True)


def make_destination(original: Path, source_root: Path, target_root: Path) -> Path:
    """Return the destination path for the converted PDF preserving the tree."""

    relative = original.relative_to(source_root).with_suffix(".pdf")
    destination = target_root / relative
    destination.parent.mkdir(parents=True, exist_ok=True)
    return destination


def unlock_file(path: Path) -> None:
    """Attempt to remove read-only restrictions from the file.

    The function updates the POSIX permissions and, on Windows, removes the
    read-only flag if present.
    """

    try:
        current_mode = path.stat().st_mode
        desired_mode = current_mode | stat.S_IWRITE
        if desired_mode != current_mode:
            path.chmod(desired_mode)
    except OSError as exc:  # pragma: no cover - best effort only
        LOGGER.debug("Não foi possível ajustar permissões de %s: %s", path, exc)

    # Remove the Windows read-only attribute when possible.  This requires
    # ctypes but keeps the dependency optional as it only runs on Windows.
    if os.name == "nt":  # pragma: no cover - Windows specific behaviour
        import ctypes

        FILE_ATTRIBUTE_NORMAL = 0x80
        FILE_ATTRIBUTE_READONLY = 0x01
        SetFileAttributes = ctypes.windll.kernel32.SetFileAttributesW
        SetFileAttributes.argtypes = [ctypes.c_wchar_p, ctypes.c_uint32]
        SetFileAttributes.restype = ctypes.c_bool

        if not SetFileAttributes(str(path), FILE_ATTRIBUTE_NORMAL):
            # If we cannot reset to normal, try at least removing the
            # read-only flag.  Errors are ignored as there is nothing else we
            # can do automatically.
            SetFileAttributes(str(path), FILE_ATTRIBUTE_NORMAL | FILE_ATTRIBUTE_READONLY)


def convert_single(
    office_binary: str,
    document: Path,
    destination: Path,
) -> None:
    """Convert *document* to PDF and move it to *destination*."""

    run_libreoffice(office_binary, document, destination.parent)
    generated_pdf = document.with_suffix(".pdf")
    candidate = destination.parent / generated_pdf.name
    if not candidate.exists():
        raise ConversionError(
            f"O LibreOffice não gerou o PDF esperado para {document}."
        )
    candidate.replace(destination)


def convert_with_restriction_bypass(
    office_binary: str,
    document: Path,
    destination: Path,
) -> None:
    """Try to convert the document, applying workarounds for restricted files."""

    try:
        convert_single(office_binary, document, destination)
        return
    except subprocess.CalledProcessError as exc:
        LOGGER.warning(
            "Falha ao converter %s diretamente (%s). Tentando bypass...",
            document,
            exc,
        )
    except ConversionError:
        LOGGER.warning(
            "O arquivo %s não foi gerado corretamente. Tentando bypass...",
            document,
        )
    else:
        return

    with tempfile.TemporaryDirectory() as tmp_dir:
        temp_document = Path(tmp_dir) / document.name
        shutil.copy2(document, temp_document)
        unlock_file(temp_document)

        try:
            convert_single(office_binary, temp_document, destination)
        except subprocess.CalledProcessError as exc:
            raise ConversionError(
                f"Não foi possível converter {document} nem após o bypass: {exc}"
            ) from exc


def convert_documents(
    office_binary: str,
    documents: Iterable[Path],
    source_root: Path,
    output_dir: Path,
) -> List[ConversionResult]:
    """Convert all *documents*, storing PDFs inside *output_dir*."""

    results: List[ConversionResult] = []
    for document in documents:
        destination = make_destination(document, source_root, output_dir)
        try:
            convert_with_restriction_bypass(office_binary, document, destination)
        except Exception as exc:  # pragma: no cover - logging/IO heavy
            LOGGER.error("Falha ao converter %s: %s", document, exc)
            results.append(
                ConversionResult(
                    source=document,
                    destination=None,
                    succeeded=False,
                    error=str(exc),
                )
            )
        else:
            LOGGER.info("Convertido: %s -> %s", document, destination)
            results.append(
                ConversionResult(
                    source=document,
                    destination=destination,
                    succeeded=True,
                )
            )
    return results


def parse_args(argv: List[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Converta todos os arquivos .doc/.docx de uma pasta para PDF",
    )
    parser.add_argument(
        "source",
        type=Path,
        help="Diretório onde os documentos serão procurados.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("pdf"),
        help="Diretório de saída para os PDFs (padrão: ./pdf).",
    )
    parser.add_argument(
        "--office-binary",
        default="libreoffice",
        help=(
            "Nome ou caminho completo do executável do LibreOffice/soffice. "
            "Use, por exemplo, 'soffice' no Windows."
        ),
    )
    parser.add_argument(
        "--no-recursive",
        action="store_true",
        help="Não procurar documentos em subpastas.",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Exibir logs detalhados do processo de conversão.",
    )
    return parser.parse_args(argv)


def configure_logging(verbose: bool) -> None:
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(levelname)s: %(message)s",
    )


def main(argv: List[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    configure_logging(args.verbose)

    if not args.source.exists() or not args.source.is_dir():
        LOGGER.error("O diretório fonte %s não existe ou não é uma pasta.", args.source)
        return 1

    try:
        office_binary = ensure_office_binary(args.office_binary)
    except FileNotFoundError as exc:
        LOGGER.error("%s", exc)
        return 2

    documents = discover_documents(args.source, recursive=not args.no_recursive)
    if not documents:
        LOGGER.warning("Nenhum arquivo .doc/.docx encontrado em %s", args.source)
        return 0

    args.output.mkdir(parents=True, exist_ok=True)
    convert_documents(office_binary, documents, args.source, args.output)
    return 0


if __name__ == "__main__":  # pragma: no cover - CLI entry point
    sys.exit(main())
