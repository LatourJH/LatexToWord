# LatexToWord

LatexToWord converts selected LaTeX text into Microsoft Word OMath equations.
It combines a small Word VBA macro with a dependency-free Python parser and is
intended for academic papers, lab reports, and other equation-heavy documents.

> [!IMPORTANT]
> This project is in beta. Work on a copy of an important document and review
> the result before saving. VBA behavior is tested manually in desktop Word.

## What changed in version 0.2

- Converts only the selected text instead of resetting the entire document.
- Removes hard-coded usernames and filesystem paths.
- Supports `$...$`, `$$...$$`, `\(...\)`, and `\[...\]` delimiters.
- Treats separate non-empty lines as equations when delimiters are omitted.
- Reports malformed delimiters and Python failures without replacing the selection.
- Groups each conversion into a single Word Undo action when Word supports it.
- Requires no third-party Python packages for normal use.

## Requirements

- Windows with desktop Microsoft Word
- Python 3.10 or newer, installed with the Windows `py` launcher
- Permission to run macros in a trusted document or template

The conversion program itself has no runtime package dependencies.

## Install

1. Download the source ZIP from GitHub or clone this repository.
2. Keep `PythonToLatexMainFile.py` and the `src` folder together.
3. In Word, press `Alt+F11` to open the Visual Basic Editor.
4. Choose **File > Import File** and import `LatexToWordMain.bas`.
5. Save the document as a macro-enabled `.docm` file, or save the macro in a
   trusted `.dotm` template if you want it available across documents.
6. Optional: assign `MainSequence` to a ribbon button or keyboard shortcut.

The macro first looks for `PythonToLatexMainFile.py` beside the active document
and beside the template that contains the macro. If it cannot find the script,
it asks you to locate it. If the repository has a `.venv`, the macro uses that
environment; otherwise it runs `py -3`.

### Optional development environment

```powershell
py -3 -m venv .venv
.\.venv\Scripts\python -m pip install -e ".[dev]"
```

## Use

1. Select one LaTeX expression or a block containing several expressions.
2. Run the `MainSequence` macro.
3. LatexToWord replaces the selection with professionally formatted OMath equations.

Select equations only. For example, this selection contains two delimited equations:

```tex
$A_v = \frac{V_{out}}{V_{in}}$
\(P = V \cdot I\)
```

For a block of equations, delimiters are optional:

```tex
V = I \cdot R
P = V \cdot I
I = \tfrac{V}{R}
```

The selected block is replaced. Formatting outside the selection is not reset.
Use Word's **Undo** command if the result is not what you expected.
The macro rejects non-whitespace text outside math delimiters so surrounding
prose cannot be removed accidentally.

## Supported input

The parser recognizes:

- Inline math: `$...$` and `\(...\)`
- Display math: `$$...$$` and `\[...\]`
- Escaped dollar signs such as `\$`
- Multiple equations in one selection
- Plain, non-empty equation lines without delimiters
- Common compatibility substitutions including `\tfrac`, `\dfrac`, `\cdot`,
  `\times`, `\div`, `\pm`, `\approx`, `\leq`, `\geq`, and `\sqrt`

Word's equation engine ultimately determines which remaining LaTeX or linear
math commands can be built into professional notation. Unsupported commands may
remain visible as text.

## Command-line use

The Python converter can be tested without Word:

```powershell
py -3 PythonToLatexMainFile.py `
  --text '$x \approx \tfrac{1}{2}$' `
  --output latex_output.txt
```

Or pass a UTF-8 input file:

```powershell
py -3 PythonToLatexMainFile.py `
  --text-file equations.txt `
  --output latex_output.txt `
  --log latex_processing.log
```

The command returns exit code `0` on success and `2` for invalid input or file
errors. Output contains one converted equation per line.

## Troubleshooting

### Word cannot find Python

Open PowerShell and run `py -3 --version`. If Windows cannot find the command,
install Python from python.org and enable the Python launcher during setup.

### Word asks for the Python script every time

Place `PythonToLatexMainFile.py` and `src` beside the active Word document, or
store them beside the `.dotm` template containing the macro.

### A conversion fails

The selection is not replaced when Python reports an error. The message includes
the path of a temporary diagnostic log. The log records status and error details,
not the selected equation contents.

### Equations remain in linear form

Word may not recognize every LaTeX command. Reduce the expression to a minimal
example and open a GitHub issue with your Word, Windows, Python, and
LatexToWord versions. Do not attach private document content.

## Development

```powershell
python -m pip install -e ".[dev]"
ruff check .
pytest
```

Python tests run on Windows and Linux in GitHub Actions. Changes to the VBA
module also require a manual test in desktop Word using a copy of a document.

See [CONTRIBUTING.md](CONTRIBUTING.md) for the pull-request checklist and
[SECURITY.md](SECURITY.md) for private vulnerability reporting guidance.

## License

LatexToWord is available under the [MIT License](LICENSE).
