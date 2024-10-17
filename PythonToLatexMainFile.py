import os
import re
from docx import Document
from pylatexenc.latex2text import LatexNodes2Text
import logging

# Set up logging for debugging
logging.basicConfig(filename=r'C:\\Users\\latou\\Desktop\\LatexToWordProject\\latex_processing.log', level=logging.DEBUG,
                    format='%(asctime)s %(levelname)s %(message)s')

# Function to clean LaTeX text and manually handle fractions
def clean_latex_text(latex_string):
    logging.debug(f"Original LaTeX string: {latex_string}")

    # Manual replacements for common LaTeX symbols (keep \frac intact)
    replacements = {
        r'\cdot': '·',      # Multiplication dot
        r'\times': '×',     # Multiplication cross
        r'\div': '÷',       # Division symbol
        r'\pm': '±',        # Plus-minus
        r'\approx': '≈',    # Approximately equal
        r'\leq': '≤',       # Less than or equal to
        r'\geq': '≥',       # Greater than or equal to
        r'\sqrt': '√',      # Square root
    }

    # Replace LaTeX symbols with Unicode equivalents
    for latex_symbol, unicode_char in replacements.items():
        latex_string = latex_string.replace(latex_symbol, unicode_char)

    # Ensure that fractions are preserved as \frac{numerator}{denominator}
    # Convert all the rest of LaTeX except \frac{}{} using pylatexenc
    latex_text = LatexNodes2Text().latex_to_text(latex_string)

    # Replace any modified \frac text back to LaTeX fraction format
    latex_text = re.sub(r'@@FRAC@@(.*?)@@END_FRAC@@', r'\\frac{\1}', latex_text)

    logging.debug(f"Cleaned LaTeX string after restoring fractions: {latex_text}")
    return latex_text

# Function to clean up hidden characters
def clean_latex_input(latex_string):
    latex_string = latex_string.replace(chr(160), " ")  # Replace non-breaking spaces
    latex_string = latex_string.replace("\n", " ")      # Replace paragraph breaks with spaces
    latex_string = latex_string.replace("\t", " ")      # Replace any tabs with spaces
    latex_string = re.sub(r'[\u200B-\u200D\uFEFF]', '', latex_string)  # Remove zero-width spaces
    logging.debug(f"Cleaned hidden characters from LaTeX input: {latex_string}")
    return latex_string

# Function to split the continuous LaTeX string into individual equations
def split_latex_equations(latex_string):
    equations = re.split(r'(?<!\\)\$', latex_string)  # Split by $ symbols while avoiding escaped \$ signs
    logging.debug(f"Split equations: {equations}")
    return [eq.strip() for eq in equations if eq.strip()]  # Filter out empty strings

# Ensure that the Python script is pulling input directly from the document
doc_path = r'C:\\Users\\latou\\Desktop\\LatexToWordProject\\LatexTestWord.docx'
output_file_path = r'C:\\Users\\latou\\Desktop\\LatexToWordProject\\latex_output.txt'  # Corrected output path

if not os.path.exists(doc_path):
    logging.error(f"Document not found: {doc_path}")
    exit()

doc = Document(doc_path)
latex_string = ""

# Extract text from the Word document
for para in doc.paragraphs:
    latex_string += para.text + "\n"

logging.debug(f"Extracted LaTeX from Word document: {latex_string}")

# Clean the LaTeX input to remove hidden characters
latex_string = clean_latex_input(latex_string)

# Split the continuous LaTeX input into separate equations
equations = split_latex_equations(latex_string)

# Create a list to hold the final results
results = []

# Process each equation separately
for eq in equations:
    cleaned_latex = clean_latex_text(eq)
    results.append(cleaned_latex)

# Join the results with newline characters and prepare the output string
final_output = '\n'.join(results)

# Output the final cleaned LaTeX directly (for VBA to capture)
try:
    with open(output_file_path, 'w', encoding='utf-8') as f:
        f.write(final_output)
    logging.debug(f"Final LaTeX/MathML results written to {output_file_path}")
except Exception as e:
    logging.error(f"Error writing to file: {e}")

logging.debug("Python script finished.")
