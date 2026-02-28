import os
import sys
from docx2pdf import convert


def convert_single_file(input_path, output_path=None):
    """
    Convert a single DOCX file to PDF.
    """
    if not os.path.exists(input_path):
        print("❌ Error: Input file does not exist.")
        return

    if not input_path.lower().endswith(".docx"):
        print("❌ Error: File must be .docx format.")
        return

    try:
        if output_path:
            convert(input_path, output_path)
        else:
            convert(input_path)
        print("✅ Conversion successful.")
    except Exception as e:
        print("❌ Conversion failed:", str(e))


def convert_folder(input_folder, output_folder=None):
    """
    Convert all DOCX files in a folder to PDF.
    """
    if not os.path.isdir(input_folder):
        print("❌ Error: Input folder does not exist.")
        return

    try:
        if output_folder:
            os.makedirs(output_folder, exist_ok=True)
            convert(input_folder, output_folder)
        else:
            convert(input_folder)

        print("✅ Folder conversion completed.")
    except Exception as e:
        print("❌ Conversion failed:", str(e))


def main():
    print("DOCX to PDF Converter")
    print("---------------------")
    print("1. Convert Single File")
    print("2. Convert Folder")
    choice = input("Choose option (1 or 2): ").strip()

    if choice == "1":
        input_file = input("Enter full path of DOCX file: ").strip()
        output_file = input("Enter output PDF path (leave empty for same folder): ").strip()

        if output_file == "":
            output_file = None

        convert_single_file(input_file, output_file)

    elif choice == "2":
        input_folder = input("Enter folder path containing DOCX files: ").strip()
        output_folder = input("Enter output folder path (leave empty for same folder): ").strip()

        if output_folder == "":
            output_folder = None

        convert_folder(input_folder, output_folder)

    else:
        print("❌ Invalid option selected.")


if __name__ == "__main__":
    main()
