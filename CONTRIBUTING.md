# Contributing

Thank you for helping improve LatexToWord.

## Development setup

```powershell
py -3 -m venv .venv
.\.venv\Scripts\python -m pip install -e ".[dev]"
ruff check .
pytest
```

## Pull requests

- Keep changes focused and explain the user-facing reason.
- Add or update tests for Python behavior.
- Update the README and changelog when behavior changes.
- Never commit Word lock files, logs, generated output, virtual environments,
  personal paths, or private document text.
- If VBA changes, test it manually in desktop Word using a copy of a document.
- Record the Word, Windows, Python, and LatexToWord versions used for manual tests.

## Reporting equation compatibility problems

Reduce the problem to the smallest LaTeX expression that reproduces it. Include
the expected result and the actual Word result, but remove private information.
