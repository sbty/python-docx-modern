"""Static type-checking gate for modules migrated to `mypy --strict`."""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import pytest

pytest.importorskip("mypy")

PROJECT_ROOT = Path(__file__).parent.parent


class DescribeTypeChecking:
    """Runs mypy with the project configuration in `pyproject.toml`.

    The set of modules held to `strict` is the ratchet list under `[tool.mypy.overrides]`.
    """

    def it_passes_mypy_strict_on_migrated_modules(self):
        result = subprocess.run(
            [sys.executable, "-m", "mypy", "--no-incremental"],
            cwd=PROJECT_ROOT,
            capture_output=True,
            text=True,
        )

        assert result.returncode == 0, result.stdout + result.stderr
