"""Application entry point for PDF conversion tools."""

import argparse
import os
import sys
from pathlib import Path


TABLE_COMMANDS = {"table", "excel", "xlsx", "pdf-to-excel"}
WORD_COMMANDS = {"word", "docx", "pdf-to-word"}


def ensure_standard_streams():
    """Make prints safe when the app is packaged as a Windows windowed exe."""
    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w", encoding="utf-8")
    if sys.stderr is None:
        sys.stderr = open(os.devnull, "w", encoding="utf-8")


def process_pdf(*args, **kwargs):
    """Backward-compatible alias for the original table-to-Excel converter."""
    from converters.table_to_excel import convert_table_pdf_to_excel

    return convert_table_pdf_to_excel(*args, **kwargs)


def run_table_conversion(pdf_file, output=None, dpi=300):
    from converters.table_to_excel import convert_table_pdf_to_excel

    output = output or f"{Path(pdf_file).stem}.xlsx"
    convert_table_pdf_to_excel(pdf_file, output, dpi=dpi)


def run_word_conversion(pdf_file, output=None, dpi=300):
    from converters.pdf_to_word import convert_pdf_to_word

    output = output or f"{Path(pdf_file).stem}.docx"
    convert_pdf_to_word(pdf_file, output, dpi=dpi)


def run_gui():
    from PyQt5.QtWidgets import QApplication

    from ui.main_window import MainWindow

    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


def self_test_ocr():
    from rapidocr import RapidOCR

    RapidOCR()
    print("RapidOCR OK")


def parse_table_args(argv, description="PDF 表格转 Excel（支持扫描件）"):
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument("pdf_file", nargs="?", default="Scan.pdf", help="PDF 文件路径")
    parser.add_argument("-o", "--output", default=None, help="输出 Excel 文件路径")
    parser.add_argument("--dpi", type=int, default=300, help="OCR 分辨率（默认300）")
    parser.add_argument("--self-test-ocr", action="store_true", help=argparse.SUPPRESS)
    args = parser.parse_args(argv)
    if args.self_test_ocr:
        self_test_ocr()
        return
    run_table_conversion(args.pdf_file, args.output, args.dpi)


def parse_word_args(argv):
    parser = argparse.ArgumentParser(description="PDF 转 Word")
    parser.add_argument("pdf_file", help="PDF 文件路径")
    parser.add_argument("-o", "--output", default=None, help="输出 Word 文件路径")
    parser.add_argument("--dpi", type=int, default=300, help="OCR 分辨率（默认300）")
    args = parser.parse_args(argv)
    run_word_conversion(args.pdf_file, args.output, dpi=args.dpi)


def main(argv=None):
    ensure_standard_streams()
    argv = list(sys.argv[1:] if argv is None else argv)

    if not argv:
        run_gui()
        return

    if "--self-test-ocr" in argv:
        self_test_ocr()
        return

    command = argv[0].lower()
    if command in WORD_COMMANDS:
        parse_word_args(argv[1:])
    elif command in TABLE_COMMANDS:
        parse_table_args(argv[1:])
    else:
        # Legacy behavior: python main.py input.pdf -o output.xlsx
        parse_table_args(argv)


if __name__ == "__main__":
    main()
