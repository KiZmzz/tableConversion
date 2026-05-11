"""Build the Windows executable with PyInstaller.

Keeping this in a normal Python file avoids fragile quoting/indentation issues
when GitHub Actions runs commands through PowerShell.
"""

import sys
import os

import PyInstaller.__main__
from PyInstaller.utils.hooks import collect_dynamic_libs


def build():
    sys.setrecursionlimit(5000)

    extra_args = []
    for package in ("onnxruntime",):
        for src, dest in collect_dynamic_libs(package):
            extra_args.extend(["--add-binary", f"{src}{os.pathsep}{dest}"])

    PyInstaller.__main__.run(
        [
            "--noconfirm",
            "--onedir",
            "--windowed",
            "--name",
            "TableConversion",
            "--icon",
            "app.ico",
            "--hidden-import",
            "PyQt5.QtPrintSupport",
            "--hidden-import",
            "qfluentwidgets",
            "--hidden-import",
            "rapidocr",
            "--hidden-import",
            "onnxruntime",
            "--hidden-import",
            "openpyxl",
            "--hidden-import",
            "pdfplumber",
            "--hidden-import",
            "fitz",
            "--hidden-import",
            "docx",
            "--hidden-import",
            "shapely",
            "--hidden-import",
            "ui.main_window",
            "--hidden-import",
            "converters.table_to_excel",
            "--hidden-import",
            "converters.pdf_to_word",
            "--collect-all",
            "qfluentwidgets",
            "--collect-all",
            "rapidocr",
            "--collect-all",
            "onnxruntime",
            "--collect-all",
            "pdfplumber",
            "--collect-all",
            "shapely",
            "--copy-metadata",
            "PyQt-Fluent-Widgets",
            "--copy-metadata",
            "rapidocr",
            "--copy-metadata",
            "onnxruntime",
            "--copy-metadata",
            "PyMuPDF",
            "--copy-metadata",
            "python-docx",
            *extra_args,
            "main.py",
        ]
    )


if __name__ == "__main__":
    build()
