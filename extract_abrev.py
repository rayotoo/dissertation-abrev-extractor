import re
from collections import OrderedDict
from docx import Document
import fitz  # PyMuPDF for PDF files
import pandas as pd

def extract_text_from_docx(file_path):
    doc = Document(file_path)
    return '\n'.join([para.text for para in doc.paragraphs])

def extract_text_from_pdf(file_path):
    doc = fitz.open(file_path)
    text = ""
    for page in doc:
        text += page.get_text()
    return text

def extract_abbreviations(text):
    abbrev_pattern = re.compile(r'\b[A-Z]{2,6}s?\b')
    all_matches = abbrev_pattern.findall(text)

    order = OrderedDict()
    for match in all_matches:
        clean_match = match.rstrip('s') if match.endswith('s') and len(match) > 3 else match
        if clean_match not in order:
            order[clean_match] = 1
        else:
            order[clean_match] += 1
    return order

def find_definitions(text, abbrevs):
    definitions = {}
    for abbr in abbrevs:
        # Match phrases like "Linear Discriminant Analysis (LDA)"
        pattern = re.compile(r'([A-Z][a-z]+(?:\s+[A-Z][a-z]+){0,5})\s+\(\b{}\b\)'.format(re.escape(abbr)))
        match = pattern.search(text)
        definitions[abbr] = match.group(1) if match else ""
    return definitions

def generate_abbreviation_excel(abbrevs, definitions, output_file="abbreviations.xlsx"):
    data = []
    for abbr, count in abbrevs.items():
        definition = definitions.get(abbr, "")
        data.append({
            "Abbreviation": abbr,
            "Definition": definition,
            "Count": count
        })
    df = pd.DataFrame(data)
    df.to_excel(output_file, index=False)
    print(f"✅ Abbreviation table saved to: {output_file}")

# === Main Usage ===
file_path = "/Users/raymondotoo/Desktop/dissertation/Dissertation_Committee Submission_edting.docx"  # or .pdf

# Extract full text
if file_path.endswith(".docx"):
    full_text = extract_text_from_docx(file_path)
elif file_path.endswith(".pdf"):
    full_text = extract_text_from_pdf(file_path)
else:
    raise ValueError("Unsupported file type. Please use .docx or .pdf")

# Run abbreviation logic
abbrevs = extract_abbreviations(full_text)
definitions = find_definitions(full_text, abbrevs)
generate_abbreviation_excel(abbrevs, definitions, "dissertation_abbreviations.xlsx")
