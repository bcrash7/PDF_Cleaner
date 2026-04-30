"""
Clean PPT/PPTX files for Salesforce Agentforce / Data Cloud ingestion.

This version includes verbose diagnostics to help debug why slides may
not be processing as expected.

USAGE IN PYCHARM:
    1. Drop your slide decks into the ./slides folder (created on first run).
    2. Right-click this file → Run 'main'.
    3. Cleaned markdown will appear in ./cleaned_slides.

USAGE FROM TERMINAL:
    python clean_slides.py                              # uses defaults
    python clean_slides.py ./slides ./cleaned_slides
    python clean_slides.py ./slides ./cleaned_slides --debug

INSTALL:
    pip install docling pypdf

LEGACY .PPT SUPPORT (optional):
    Install LibreOffice from https://www.libreoffice.org/download/
"""

import argparse
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

# Suppress the harmless Hugging Face symlinks warning on Windows
os.environ["HF_HUB_DISABLE_SYMLINKS_WARNING"] = "1"

# ---------------------------------------------------------------------------
# Defaults — edit these or pass them as arguments / Run Configuration params.
# ---------------------------------------------------------------------------
DEFAULT_INPUT_DIR = Path("./slides")
DEFAULT_OUTPUT_DIR = Path("./cleaned_slides")

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


def convert_ppt_to_pptx(ppt_path: Path, output_dir: Path) -> Path:
    """Convert a legacy .ppt to .pptx using LibreOffice."""
    soffice = find_libreoffice()
    if not soffice:
        raise RuntimeError(
            f"Cannot convert {ppt_path.name}: LibreOffice not found. "
            "Install it from https://www.libreoffice.org/download/"
        )

    result = subprocess.run(
        [str(soffice), "--headless", "--convert-to", "pptx",
         "--outdir", str(output_dir), str(ppt_path)],
        capture_output=True, text=True, timeout=120,
    )
    if result.returncode != 0:
        raise RuntimeError(
            f"LibreOffice conversion failed: {result.stderr.strip() or result.stdout.strip()}"
        )

    converted = output_dir / f"{ppt_path.stem}.pptx"
    if not converted.exists():
        raise RuntimeError(f"LibreOffice didn't produce: {converted}")
    return converted


def convert_pptx_to_markdown(pptx_path: Path, converter, debug: bool = False) -> str:
    """Convert a PPTX to markdown using docling's parser."""
    result = converter.convert(str(pptx_path))
    md = result.document.export_to_markdown()
    if debug:
        print(f"    [debug] docling extracted {len(md):,} chars of markdown")
        print(f"    [debug] first 300 chars: {md[:300]!r}")
    return md


def clean_slide_markdown(md_text: str, debug: bool = False) -> str:
    """
    Conservative cleanup for slide-converted markdown. Only removes obvious
    noise and won't drop slide sections even if they look 'empty', because
    docling's PPTX output structure varies and we'd rather keep too much
    than too little.
    """
    before_len = len(md_text)

    # Strip slide-number-only lines
    md_text = re.sub(r"^\s*\d{1,4}\s*$", "", md_text, flags=re.MULTILINE)

    # "Slide X of Y" patterns
    md_text = re.sub(
        r"^\s*Slide\s+\d+(\s+of\s+\d+)?\s*$",
        "", md_text, flags=re.MULTILINE | re.IGNORECASE,
    )

    # Empty bullets
    md_text = re.sub(r"^\s*[-*]\s*$", "", md_text, flags=re.MULTILINE)

    # Separator-only lines
    md_text = re.sub(r"^\s*[-_.]{3,}\s*$", "", md_text, flags=re.MULTILINE)

    # Common boilerplate footers
    boilerplate_patterns = [
        r"^\s*confidential\s*$",
        r"^\s*proprietary\s+and\s+confidential\s*$",
        r"^\s*internal\s+use\s+only\s*$",
        r"^\s*©\s*\d{4}.*$",
        r"^\s*copyright\s+\d{4}.*$",
        r"^\s*all\s+rights\s+reserved\.?\s*$",
    ]
    for pattern in boilerplate_patterns:
        md_text = re.sub(pattern, "", md_text, flags=re.MULTILINE | re.IGNORECASE)

    # Normalize whitespace
    md_text = re.sub(r"[ \t]+$", "", md_text, flags=re.MULTILINE)
    md_text = re.sub(r"\n{3,}", "\n\n", md_text)

    cleaned = md_text.strip() + "\n"

    if debug:
        print(f"    [debug] cleanup: {before_len:,} → {len(cleaned):,} chars")

    return cleaned


def clean_slide_deck(pptx_path: Path, output_path: Path, converter,
                     debug: bool = False) -> dict:
    """Process a single PPTX file. Returns stats for reporting."""
    md_text = convert_pptx_to_markdown(pptx_path, converter, debug=debug)
    original_len = len(md_text)

    md_text = clean_slide_markdown(md_text, debug=debug)

    output_path.write_text(md_text, encoding="utf-8")

    return {
        "original_chars": original_len,
        "cleaned_chars": len(md_text),
        "reduction_pct": round((1 - len(md_text) / original_len) * 100, 1)
        if original_len else 0,
    }


def main():
    parser = argparse.ArgumentParser(
        description="Clean PPT/PPTX files for Salesforce Agentforce ingestion."
    )
    parser.add_argument("input_dir", type=Path, nargs="?", default=DEFAULT_INPUT_DIR)
    parser.add_argument("output_dir", type=Path, nargs="?", default=DEFAULT_OUTPUT_DIR)
    parser.add_argument("--debug", action="store_true",
                        help="Print verbose diagnostics for each file.")
    args = parser.parse_args()

    print(f"Input dir:  {args.input_dir.resolve()}")
    print(f"Output dir: {args.output_dir.resolve()}\n")

    if not args.input_dir.exists():
        args.input_dir.mkdir(parents=True, exist_ok=True)
        print(
            f"Created input directory. Drop your .ppt/.pptx files in "
            f"{args.input_dir.resolve()} and run again."
        )
        return

    if not args.input_dir.is_dir():
        sys.exit(f"Input path is not a directory: {args.input_dir}")

    args.output_dir.mkdir(parents=True, exist_ok=True)

    # Show EVERYTHING in the input dir so we can debug missing-file issues
    all_in_dir = sorted(args.input_dir.iterdir())
    print(f"Found {len(all_in_dir)} item(s) in input directory:")
    for item in all_in_dir:
        marker = "/" if item.is_dir() else ""
        print(f"  - {item.name}{marker}")
    print()

    # Case-insensitive matching for .pptx and .ppt
    pptx_files = sorted(
        f for f in args.input_dir.iterdir()
        if f.is_file() and f.suffix.lower() == ".pptx"
    )
    ppt_files = sorted(
        f for f in args.input_dir.iterdir()
        if f.is_file() and f.suffix.lower() == ".ppt"
    )
    all_files = pptx_files + ppt_files

    if not all_files:
        print(
            "No .ppt or .pptx files found. Make sure your files have those "
            "extensions and aren't in a subfolder.\n"
            f"Looking in: {args.input_dir.resolve()}"
        )
        return

    try:
        from docling.document_converter import DocumentConverter
    except ImportError:
        sys.exit("docling is not installed. Run: pip install docling")

    converter = DocumentConverter()

    print(f"Processing {len(all_files)} slide deck(s):")
    print(f"  - {len(pptx_files)} .pptx file(s)")
    print(f"  - {len(ppt_files)} .ppt file(s) (will be auto-converted)")
    print()

    total_reduction = 0
    failures = []

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)

        for slide_file in all_files:
            print(f"  Processing: {slide_file.name}")
            try:
                if slide_file.suffix.lower() == ".ppt":
                    print(f"    Converting .ppt → .pptx via LibreOffice...")
                    pptx_file = convert_ppt_to_pptx(slide_file, tmp_path)
                else:
                    pptx_file = slide_file

                md_output = args.output_dir / f"{slide_file.stem}.md"
                stats = clean_slide_deck(
                    pptx_file, md_output, converter, debug=args.debug
                )
                total_reduction += stats["reduction_pct"]

                print(
                    f"    Result: {stats['original_chars']:,} → "
                    f"{stats['cleaned_chars']:,} chars "
                    f"({stats['reduction_pct']}% reduction)"
                )
                print(f"    Saved: {md_output.resolve()}")

            except Exception as e:
                failures.append((slide_file.name, str(e)))
                print(f"    FAILED: {type(e).__name__}: {e}")
                if args.debug:
                    import traceback
                    traceback.print_exc()
            print()

    print(f"Done. Cleaned files written to {args.output_dir.resolve()}")
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
