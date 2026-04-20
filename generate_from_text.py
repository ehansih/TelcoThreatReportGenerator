"""
Generate PDF and/or DOCX output directly from a plain text manifest file.

Example:
  python generate_from_text.py --input report_input.txt --pdf out.pdf
  python generate_from_text.py --input report_input.txt --pdf out.pdf --docx out.docx
"""

from __future__ import annotations

import argparse
from pathlib import Path

from generate_report_gui import build_docx, build_pdf, dump_text_manifest, parse_text_manifest


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Generate threat intel report files from a plain text manifest."
    )
    parser.add_argument("--input", required=True, help="Path to the .txt input manifest.")
    parser.add_argument("--pdf", help="Output PDF path. Overrides the output field in the text file.")
    parser.add_argument("--docx", help="Output DOCX path. Overrides the output_docx field in the text file.")
    parser.add_argument(
        "--write-template",
        action="store_true",
        help="Write a blank text template to the input path and exit.",
    )
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    input_path = Path(args.input)

    if args.write_template:
        input_path.write_text(dump_text_manifest({}), encoding="utf-8")
        print(f"Blank text template written to: {input_path}")
        return 0

    data = parse_text_manifest(input_path.read_text(encoding="utf-8"))

    pdf_path = args.pdf or data.get("output", "").strip()
    docx_path = args.docx or data.get("output_docx", "").strip()

    if not pdf_path and not docx_path:
        parser.error("Provide --pdf and/or --docx, or set output/output_docx in the text file.")

    if pdf_path:
        build_pdf(data, pdf_path, progress_cb=lambda message: print(message))

    if docx_path:
        build_docx(data, docx_path, progress_cb=lambda message: print(message))

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
