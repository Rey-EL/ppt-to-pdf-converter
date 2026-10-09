"""Smoke tests for ppt_to_pdf_converter.py.

The module drives PowerPoint through Windows COM automation, so it cannot
be imported or exercised on Linux. These tests verify the file compiles
cleanly and still defines the entry points the GUI expects.
"""

import ast
import py_compile
from pathlib import Path

MODULE = Path(__file__).resolve().parent.parent / "ppt_to_pdf_converter.py"


def test_module_compiles():
    py_compile.compile(str(MODULE), doraise=True)


def test_expected_functions_defined():
    tree = ast.parse(MODULE.read_text(encoding="utf-8"))
    names = {node.name for node in ast.walk(tree) if isinstance(node, ast.FunctionDef)}
    assert {"convert_ppt_to_pdf", "main_process"} <= names


def test_app_class_defined():
    tree = ast.parse(MODULE.read_text(encoding="utf-8"))
    names = {node.name for node in ast.walk(tree) if isinstance(node, ast.ClassDef)}
    assert "App" in names
