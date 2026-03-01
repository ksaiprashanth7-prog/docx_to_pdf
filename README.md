📄 DOCX to PDF Converter

  A simple Python command-line tool to convert .docx files into .pdf format.
  Supports both single file conversion and batch folder conversion.

🚀 Features

  Convert a single .docx file to PDF
  Convert all .docx files inside a folder
  Optional custom output location
  Simple CLI interface
  Error handling for invalid paths and formats

🛠 Requirements

  Python 3.7+
  Microsoft Word (required by docx2pdf)
  Python package: pip install docx2pdf

⚠️ Note: docx2pdf requires Microsoft Word to be installed because it uses Word’s internal conversion engine.

📦 Installation

Clone or download this repository.

Install dependency:pip install docx2pdf

▶️ Usage 

Run the script:
python your_script_name.py
You will see:

DOCX to PDF Converter
---------------------
1. Convert Single File
2. Convert Folder
🔹 Option 1: Convert Single File

  Enter full path of the .docx file
  (Optional) Enter output PDF path
  Leave blank to save in the same folder

🔹 Option 2: Convert Folder

  Enter folder path containing .docx files
  (Optional) Enter output folder path  
  Leave blank to save in the same folder

📁 Example

  Single file input example:
  C:\Users\John\Documents\file.docx
  Folder input example:  
  C:\Users\John\Documents\Reports
  ❌ Error Handling

  The program checks for:
  File existence
  Folder existence
  Correct .docx format
  Invalid menu options
  Conversion failures

🧠 How It Works
  The script uses the docx2pdf Python package to convert DOCX files to PDF using Microsoft Word.

📌 Notes
  Works on Windows and macOS  
  Linux is not officially supported
  Microsoft Word must be installed

📄 License
  This project is open-source and free to use.



