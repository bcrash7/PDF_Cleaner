"""
Clean DOC/DOCX files for Salesforce Agentforce / Data Cloud ingestion.

Uses docling for layout-aware parsing of DOCX files. Legacy .doc files are
auto-converted to .docx first via LibreOffice (must be installed separately).

USAGE IN PYCHARM:
    1. Drop your Word documents into the ./docs folder (created on first run).
    2. Right-click this file → Run 'main'.
    3. Cleaned markdown will appear in ./cleaned_docs.

USAGE FROM TERMINAL:
    python clean_docs.py                              # uses defaults
    python clean_docs.py ./docs ./cleaned_docs

INSTALL:
    pip install docling

LEGACY .DOC SUPPORT (optional):
    Install LibreOffice from https://www.libreoffice.org/download/
    On Windows, the script auto-detects it at:
        C:\\Program Files\\LibreOffice\\program\\soffice.exe
    If installed elsewhere, set LIBREOFFICE_PATH below.
"""

import argparse
import os
import re
import shutil
import subprocess
import sys
import tempfile
from collections import Counter
from pathlib import Path

# Suppress the harmless Hugging Face symlinks warning on Windows
os.environ["HF_HUB_DISABLE_SYMLINKS_WARNING"] = "1"

# ---------------------------------------------------------------------------
# Defaults — edit these or pass them as arguments / Run Configuration params.
# ---------------------------------------------------------------------------
DEFAULT_INPUT_DIR = Path("./docs")
DEFAULT_OUTPUT_DIR = Path("./cleaned_docs")

# Set this if LibreOffice isn't in the standard install location.
# Leave as None to let the script auto-detect.
LIBREOFFICE_PATH = None  # e.g. r"C:\Program Files\LibreOffice\program\soffice.exe"


def find_libreoffice() -> Path | None:
    """Locate the LibreOffice executable, returning None if not found."""
    if LIBREOFFICE_PATH:
        p = Path(LIBREOFFICE_PATH)
        return p if p.exists() else None

    for name in ("soffice", "soffice.exe", "libreoffice"):
        found = shutil.which(name)
        if found:
            return Path(found)

    candidates = [
        Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
        Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
        Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"),
        Path("/usr/bin/libreoffice"),
        Path("/usr/bin/soffice"),
    ]
    for path in candidates:
        if path.exists():
            return path
    return None


def convert_doc_to_docx(doc_path: Path, output_dir: Path) -> Path:
    """Convert a legacy .doc to .docx using LibreOffice. Returns the new path."""
    soffice = find_libreoffice()
    if not soffice:
        raise RuntimeError(
            f"Cannot convert {doc_path.name}: LibreOffice not found. "
            "Install it from https://www.libreoffice.org/download/ or set "
            "LIBREOFFICE_PATH at the top of this script."
        )

    result = subprocess.run(
        [
            str(soffice),
            "--headless",
            "--convert-to",
            "docx",
            "--outdir",
            str(output_dir),
            str(doc_path),
        ],
        capture_output=True,
        text=True,
        timeout=120,
    )
    if result.returncode != 0:
        raise RuntimeError(
            f"LibreOffice conversion failed: {result.stderr.strip() or result.stdout.strip()}"
        )

    converted = output_dir / f"{doc_path.stem}.docx"
    if not converted.exists():
        raise RuntimeError(f"LibreOffice didn't produce the expected file: {converted}")
    return converted


def convert_docx_to_markdown(docx_path: Path, converter) -> str:
    """Convert a DOCX to markdown using docling's parser."""
    result = converter.convert(str(docx_path))
    return result.document.export_to_markdown()


def remove_repeated_lines(md_text: str) -> str:
    """
    Remove headers/footers by detecting short lines that repeat throughout
    the document. Word docs often have running headers and footers that
    docling preserves as repeated text.
    """
    lines = md_text.split("\n")

    counts = Counter()
    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue
        if stripped.startswith("#"):  # Preserve markdown headings
            continue
        if len(stripped) > 120:  # Long lines are real content
            continue
        counts[stripped] += 1

    # Drop lines that appear 4+ times — almost certainly a header/footer
    noise = {line for line, count in counts.items() if count >= 4}

    return "\n".join(line for line in lines if line.strip() not in noise)


def clean_doc_markdown(md_text: str) -> str:
    """
    Clean up Word-converted markdown.

    Word docs have these common noise patterns:
    - Page numbers and "Page X of Y" footers
    - Confidentiality / copyright boilerplate
    - Tracked-changes artifacts
    - Empty list bullets
    - Form field placeholders
    """
    # Strip page-number-only lines
    md_text = re.sub(r"^\s*\d{1,4}\s*$", "", md_text, flags=re.MULTILINE)

    # "Page X of Y" patterns
    md_text = re.sub(
        r"^\s*Page\s+\d+(\s+of\s+\d+)?\s*$",
        "",
        md_text,
        flags=re.MULTILINE | re.IGNORECASE,
    )

    # Empty bullets
    md_text = re.sub(r"^\s*[-*]\s*$", "", md_text, flags=re.MULTILINE)

    # Separator lines
    md_text = re.sub(r"^\s*[-_.]{3,}\s*$", "", md_text, flags=re.MULTILINE)

    # Common boilerplate footers (case-insensitive)
    boilerplate_patterns = [
        r"^\s*confidential\s*$",
        r"^\s*proprietary\s+and\s+confidential\s*$",
        r"^\s*internal\s+use\s+only\s*$",
        r"^\s*do\s+not\s+distribute\s*$",
        r"^\s*©\s*\d{4}.*$",
        r"^\s*copyright\s+\d{4}.*$",
        r"^\s*all\s+rights\s+reserved\.?\s*$",
    ]
    for pattern in boilerplate_patterns:
        md_text = re.sub(pattern, "", md_text, flags=re.MULTILINE | re.IGNORECASE)

    # Common form field placeholders Word leaves behind
    md_text = re.sub(
        r"\[?(click here to enter text|click or tap here to enter text|enter date|select an item)\.?\]?",
        "",
        md_text,
        flags=re.IGNORECASE,
    )

    # Run repeated-line removal
    md_text = remove_repeated_lines(md_text)

    # Trim trailing whitespace and collapse blank lines
    md_text = re.sub(r"[ \t]+$", "", md_text, flags=re.MULTILINE)
    md_text = re.sub(r"\n{3,}", "\n\n", md_text)

    return md_text.strip() + "\n"


def clean_doc(docx_path: Path, output_path: Path, converter) -> dict:
    """Process a single DOCX file. Returns stats for reporting."""
    md_text = convert_docx_to_markdown(docx_path, converter)
    original_len = len(md_text)

    md_text = clean_doc_markdown(md_text)

    output_path.write_text(md_text, encoding="utf-8")

    return {
        "original_chars": original_len,
        "cleaned_chars": len(md_text),
        "reduction_pct": round((1 - len(md_text) / original_len) * 100, 1)
        if original_len
        else 0,
    }


def main():
    parser = argparse.ArgumentParser(
        description="Clean DOC/DOCX files for Salesforce Agentforce ingestion."
    )
    parser.add_argument(
        "input_dir",
        type=Path,
        nargs="?",
        default=DEFAULT_INPUT_DIR,
        help=f"Directory of Word documents (default: {DEFAULT_INPUT_DIR})",
    )
    parser.add_argument(
        "output_dir",
        type=Path,
        nargs="?",
        default=DEFAULT_OUTPUT_DIR,
        help=f"Where to write cleaned files (default: {DEFAULT_OUTPUT_DIR})",
    )
    args = parser.parse_args()

    if not args.input_dir.exists():
        args.input_dir.mkdir(parents=True, exist_ok=True)
        print(
            f"Created input directory: {args.input_dir.resolve()}\n"
            f"Drop your .doc/.docx files there and run again."
        )
        return

    if not args.input_dir.is_dir():
        sys.exit(f"Input path is not a directory: {args.input_dir}")

    args.output_dir.mkdir(parents=True, exist_ok=True)

    docx_files = sorted(args.input_dir.glob("*.docx"))
    doc_files = sorted(args.input_dir.glob("*.doc"))
    # Filter out .docx matches from the .doc glob (some systems include them)
    doc_files = [f for f in doc_files if f.suffix.lower() == ".doc"]
    all_files = docx_files + doc_files

    if not all_files:
        print(
            f"No .doc or .docx files found in {args.input_dir.resolve()}\n"
            f"Drop your Word documents there and run again."
        )
        return

    try:
        from docling.document_converter import DocumentConverter
    except ImportError:
        sys.exit("docling is not installed. Run: pip install docling")

    converter = DocumentConverter()

    print(
        f"Processing {len(all_files)} Word document(s) from {args.input_dir.resolve()}"
    )
    if doc_files:
        print(f"  ({len(doc_files)} legacy .doc file(s) will be auto-converted)")
    print()

    total_reduction = 0
    failures = []

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)

        for doc_file in all_files:
            try:
                if doc_file.suffix.lower() == ".doc":
                    print(f"  Converting {doc_file.name} → .docx...")
                    docx_file = convert_doc_to_docx(doc_file, tmp_path)
                else:
                    docx_file = doc_file

                md_output = args.output_dir / f"{doc_file.stem}.md"
                stats = clean_doc(docx_file, md_output, converter)
                total_reduction += stats["reduction_pct"]

                print(
                    f"  {doc_file.name}: "
                    f"{stats['original_chars']:,} → {stats['cleaned_chars']:,} chars "
                    f"({stats['reduction_pct']}% reduction)"
                )

            except Exception as e:
                failures.append((doc_file.name, str(e)))
                print(f"  {doc_file.name}: FAILED — {e}")

    print(f"\nDone. Cleaned files written to {args.output_dir.resolve()}")
    successes = len(all_files) - len(failures)
    if successes:
        avg = total_reduction / successes
        print(f"Average reduction: {avg:.1f}%")
    if failures:
        print(f"\n{len(failures)} file(s) failed:")
        for name, err in failures:
            print(f"  - {name}: {err}")


if __name__ == "__main__":
    main()
