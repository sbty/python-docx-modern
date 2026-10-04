"""Static type-checking gate: the `docx` package passes mypy and pyright in strict mode."""

from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path

import pytest

pytest.importorskip("mypy")

PROJECT_ROOT = Path(__file__).parent.parent


class DescribeTypeChecking:
    """Runs mypy and pyright with the project configuration in `pyproject.toml`.

    `strict` applies to the whole `docx` package (`files` under `[tool.mypy]` and `strict`
    under `[tool.pyright]`).
    """

    def it_passes_mypy_strict(self):
        result = subprocess.run(
            [sys.executable, "-m", "mypy", "--no-incremental"],
            cwd=PROJECT_ROOT,
            capture_output=True,
            text=True,
        )

        assert result.returncode == 0, result.stdout + result.stderr

    def it_passes_pyright_strict(self):
        pytest.importorskip("pyright")
        result = subprocess.run(
            [sys.executable, "-m", "pyright", "--outputjson", "src/docx"],
            cwd=PROJECT_ROOT,
            capture_output=True,
            text=True,
        )
        # -- pyright can fail before reporting (e.g. no Node.js available); show why --
        assert result.stdout.lstrip().startswith("{"), result.stdout + result.stderr
        report = json.loads(result.stdout)
        # -- count only diagnostics in the package itself; pyright can also report on the
        # -- typeshed stubs it bundles, which this project does not control --
        src_dir = os.path.normcase(str(PROJECT_ROOT / "src" / "docx"))
        errors = [
            f"{d['file']}:{d['range']['start']['line'] + 1}: {d['message']}"
            for d in report["generalDiagnostics"]
            if d["severity"] == "error" and os.path.normcase(d["file"]).startswith(src_dir)
        ]

        assert errors == []
