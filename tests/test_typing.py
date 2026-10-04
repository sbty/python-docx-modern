"""Static type-checking gate: the `docx` package passes `mypy --strict`."""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import pytest

pytest.importorskip("mypy")

PROJECT_ROOT = Path(__file__).parent.parent


class DescribeTypeChecking:
    """Runs mypy with the project configuration in `pyproject.toml`.

    `strict` applies to the whole `docx` package (`files` under `[tool.mypy]`).
    """

    def it_passes_mypy_strict(self):
        result = subprocess.run(
            [sys.executable, "-m", "mypy", "--no-incremental"],
            cwd=PROJECT_ROOT,
            capture_output=True,
            text=True,
        )

        assert result.returncode == 0, result.stdout + result.stderr
