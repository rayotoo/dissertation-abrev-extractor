# Dissertation Abbreviation Extractor

This Python script helps dissertation writers and researchers automatically extract potential abbreviations and their definitions from Microsoft Word (.docx) and PDF (.pdf) documents. It then compiles this information into an organized Excel spreadsheet, which can significantly aid in creating a comprehensive "List of Abbreviations" section.

## Features

* **Supports Multiple Formats:** Processes both `.docx` (Microsoft Word) and `.pdf` files.
* **Automatic Abbreviation Detection:** Identifies sequences of 2-6 uppercase letters (optionally followed by 's') as potential abbreviations.
* **Definition Matching:** Attempts to find the full definition of an abbreviation by searching for patterns like "Full Term (ABBR)" within the text.
* **Usage Count:** Tracks the frequency of each extracted abbreviation.
* **Excel Output:** Generates a structured Excel file (`.xlsx`) with "Abbreviation," "Definition," and "Count" columns.

## Getting Started

These instructions will get you a copy of the project up and running on your local machine for development and testing purposes.

### Prerequisites

* Python 3.x

### Installation and Usage

Follow these steps to set up and run the script:

1.  **Navigate to your Project Directory:**
    Open your terminal or command prompt and change your current directory to where your dissertation file is located (e.g., your Desktop).

    ```bash
    cd ~/Desktop/dissertation
    ```

2.  **Create a Virtual Environment (Recommended):**
    A virtual environment helps manage project dependencies without affecting your global Python installation.

    ```bash
    python3 -m venv venv
    ```

3.  **Activate the Virtual Environment:**

    * **On macOS/Linux:**
        ```bash
        source venv/bin/activate
        ```
    * **On Windows (Command Prompt):**
        ```bash
        venv\Scripts\activate.bat
        ```
    * **On Windows (PowerShell):**
        ```bash
        venv\Scripts\Activate.ps1
        ```

4.  **Install Required Libraries:**
    Install the necessary Python packages.

    ```bash
    pip install python-docx PyMuPDF pandas openpyxl
    ```

5.  **Place Your Dissertation File:**
    Ensure your dissertation file (e.g., `Dissertation_Committee Submission_edting.docx` or `your_dissertation.pdf`) is in the same directory where you plan to run the script.

6.  **Update the Script with Your File Path:**
    Open the `extract_abrev.py` script and modify the `file_path` variable to point to your dissertation file:

    ```python
    # === Main Usage ===
    file_path = "/Users/raymondotoo/Desktop/dissertation/Dissertation_Committee Submission_edting.docx"  # <--- UPDATE THIS LINE
    # For a PDF file, it would look like:
    # file_path = "/Users/raymondotoo/Desktop/dissertation/your_dissertation.pdf"
    ```

7.  **Run the Script:**
    Execute the Python script from your terminal.

    ```bash
    python extract_abrev.py
    ```

    Upon successful execution, you will see a confirmation message, and an Excel file named `dissertation_abbreviations.xlsx` will be created in your current directory.

8.  **Deactivate the Virtual Environment:**
    Once you are done, you can exit the virtual environment.

    ```bash
    deactivate
    ```

## How it Works

The script performs the following main operations:

1.  **Text Extraction:**
    * For `.docx` files, it uses `python-docx` to read all paragraphs.
    * For `.pdf` files, it uses `PyMuPDF` (fitz) to extract text page by page.
2.  **Abbreviation Identification:**
    * It employs a regular expression `\b[A-Z]{2,6}s?\b` to find sequences of 2 to 6 uppercase letters, optionally followed by an 's' (to handle plural forms like "LODs").
    * It uses `collections.OrderedDict` to maintain the order of appearance and count the frequency of each unique abbreviation.
3.  **Definition Discovery:**
    * For each identified abbreviation, it searches the entire text for a common pattern: `Full Term (ABBR)`.
    * It captures the "Full Term" preceding the abbreviation in parentheses.
4.  **Excel Generation:**
    * Utilizes the `pandas` library to create a DataFrame from the extracted abbreviations, their definitions, and counts.
    * Exports the DataFrame to an Excel `.xlsx` file using `openpyxl`.

## Customization

* **`file_path`**: Easily change the input file to your specific dissertation document.
* **`abbrev_pattern`**: If the current regex for abbreviation detection (`r'\b[A-Z]{2,6}s?\b'`) doesn't perfectly suit your document's style, you can modify it in the `extract_abbreviations` function.
* **`pattern` for Definitions**: The regex used to find definitions `r'([A-Z][a-z]+(?:\s+[A-Z][a-z]+){0,5})\s+\(\b{}\b\)'` can be adjusted if your dissertation uses a different convention for defining terms.
* **`output_file`**: Change the name of the output Excel file in the `generate_abbreviation_excel` function.

## Contributing

If you have suggestions for improving this script, feel free to open an issue or submit a pull request!

## License

This project is open-source and available under the MIT License.
