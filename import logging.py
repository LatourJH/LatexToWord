import logging
import re
from sympy.parsing.latex import parse_latex
import os

# Set up logging
logging.basicConfig(filename=r'C:\\Users\\latou\\Desktop\\latex_processing.log', level=logging.DEBUG,
                    format='%(asctime)s %(levelname)s %(message)s')

# Function to track and validate the braces in a LaTeX string
def track_braces(latex_string):
    brace_stack = []
    for i, char in enumerate(latex_string):
        if char == '{':
            brace_stack.append(i)
        elif char == '}':
            if brace_stack:
                brace_stack.pop()
            else:
                logging.error(f"Unmatched closing brace at position {i}")
                return False
    if brace_stack:
        logging.error(f"Unmatched opening brace at positions {brace_stack}")
        return False
    return True

# Function to clean LaTeX text
def clean_latex_text(latex_string):
    latex_string = re.sub(r"[^a-zA-Z0-9{}^_=+\-().,/\\ ]", "", latex_string)
    logging.debug(f"Cleaned LaTeX string: {latex_string}")

    replacements = {
        r"\alpha": "α", r"\beta": "β", r"\gamma": "γ", r"\delta": "δ", r"\epsilon": "ε",
        r"\phi": "φ", r"\omega": "ω", r"\Omega": "Ω", r"\cdot": "·", r"\times": "×",
        r"\div": "÷", r"\pm": "±", r"\leq": "≤", r"\geq": "≥", r"\approx": "≈", r"\sqrt": "√",
        r"\infty": "∞", r"\neq": "≠"
    }

    for latex_symbol, replacement in replacements.items():
        latex_string = latex_string.replace(latex_symbol, replacement)

    logging.debug(f"Replaced symbols LaTeX string: {latex_string}")
    return latex_string

# Function to clean up hidden characters
def clean_latex_input(latex_string):
    latex_string = latex_string.replace(chr(160), " ")
    latex_string = latex_string.replace("\n", " ")
    latex_string = latex_string.replace("\t", " ")
    latex_string = re.sub(r'[\u200B-\u200D\uFEFF]', '', latex_string)
    logging.debug(f"Cleaned hidden characters from LaTeX input: {latex_string}")
    return latex_string

# Function to split the continuous LaTeX string into individual equations
def split_latex_equations(latex_string):
    equations = re.split(r'(?<!\\)\$', latex_string)
    logging.debug(f"Split equations: {equations}")
    return [eq.strip() for eq in equations if eq.strip()]

# Function to convert LaTeX to MathML
def convert_latex_to_mathml(latex_string):
    try:
        expr = parse_latex(latex_string)
        mathml_output = expr._repr_mathml_()
        logging.debug(f"MathML output: {mathml_output}")
        return mathml_output
    except Exception as e:
        logging.error(f"Error in LaTeX to MathML conversion: {e}")
        return None

# Paths to input file (dynamically passed from VBA)
input_file_path = r'C:\\Users\\latou\\Desktop\\LatexInput.txt'

if not os.path.exists(input_file_path):
    logging.error(f"Input file not found: {input_file_path}")
    exit()

with open(input_file_path, 'r', encoding='utf-8') as f:
    latex_string = f.read()

logging.debug(f"Read LaTeX input: {latex_string}")

# Clean and process LaTeX
latex_string = clean_latex_input(latex_string)
equations = split_latex_equations(latex_string)

results = []

# Process each equation
for eq in equations:
    if not track_braces(eq):
        logging.error(f"Error: Unmatched braces detected in equation: {eq}")
    else:
        cleaned_latex = clean_latex_text(eq)
        mathml_output = convert_latex_to_mathml(cleaned_latex)

        if not mathml_output:
            logging.warning(f"MathML conversion failed for equation: {eq}, using cleaned LaTeX.")
            mathml_output = cleaned_latex

        results.append(mathml_output)

# Log final results
logging.debug(f"Final results: {results}")
logging.debug("Python script finished.")
