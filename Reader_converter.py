import os
import shutil
import subprocess

# Paths to executables (adjust if necessary on your PC)
SOFFICE = r"C:\Program Files\LibreOffice\program\soffice.exe"
EBOOK_CONVERT = r"C:\Program Files\Calibre2\ebook-convert.exe"

def convert_docx_to_pdf(input_file, output_folder):
    """Converts a DOCX to PDF using LibreOffice."""
    try:
        subprocess.run(
            [SOFFICE, "--headless", "--convert-to", "pdf", input_file, "--outdir", output_folder],
            check=True
        )
        print(f"DOCX converted with LibreOffice: {input_file}")
    except Exception as e:
        print(f"Error converting {input_file} with LibreOffice: {e}")

def convert_ebook_to_pdf(input_file, output_file):
    """Converts FB2, LRF or MOBI to PDF using Calibre."""
    try:
        subprocess.run([EBOOK_CONVERT, input_file, output_file], check=True)
        print(f"Ebook converted with Calibre: {input_file} -> {output_file}")
    except Exception as e:
        print(f"Error converting {input_file} with Calibre: {e}")

def process_folder(folder):
    """Searches for DOC, FB2, LRF and MOBI files in the folder (without subfolders).
       Converts to PDF and moves originals to 'originals'."""
    extensions = ["docx", "fb2", "lrf", "mobi"]

    originals_folder = os.path.join(folder, "originals")
    os.makedirs(originals_folder, exist_ok=True)

    for file in os.listdir(folder):
        input_path = os.path.join(folder, file)

        if os.path.isfile(input_path):
            extension = file.lower().split(".")[-1]

            if extension in extensions:
                output_name = os.path.splitext(file)[0] + ".pdf"
                output_path = os.path.join(folder, output_name)

                print(f"Found: {input_path}")

                if extension == "docx":
                    convert_docx_to_pdf(input_path, folder)
                else:
                    convert_ebook_to_pdf(input_path, output_path)

                # move original to "originals"
                try:
                    shutil.move(input_path, os.path.join(originals_folder, file))
                    print(f"Moved original to: {originals_folder}")
                except Exception as e:
                    print(f"Could not move {input_path}: {e}")

def main():
    initial_folder = input("Enter the path of the folder to process: ").strip()

    if not os.path.isdir(initial_folder):
        print("The path is not valid.")
        return

    process_folder(initial_folder)

    print("\nProcess completed. Originals are in 'originals' and PDFs in the main folder.")

if __name__ == "__main__":
    main()
