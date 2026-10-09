"""Compatibility entry point for the LatexToWord command-line program.

The implementation lives in ``src/latex_to_word``. Keeping this small wrapper
means the Word macro can run directly from a cloned or downloaded repository
without requiring an installation step.
"""

from __future__ import annotations

import importlib
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(PROJECT_ROOT / "src"))

main = importlib.import_module("latex_to_word.cli").main


if __name__ == "__main__":
    raise SystemExit(main())
