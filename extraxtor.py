import os
import shutil
from pyunpack import Archive

def extract_compressed_file(compressed_path, temp_folder):
    """
    Extracts a compressed file (.zip, .rar, .7z, etc.) to temp_folder.
    Returns the extraction path or None if it fails.
    """
    try:
        folder_name = os.path.splitext(os.path.basename(compressed_path))[0]
        destination_path = os.path.join(temp_folder, folder_name)
        os.makedirs(destination_path, exist_ok=True)

        Archive(compressed_path).extractall(destination_path)
        return destination_path
    except Exception as e:
        print(f"Could not extract {compressed_path}: {e}")
        return None


def traverse_folder(folder, result_folder, temp_folder):
    """
    Traverses the folder looking for PDF/EPUB files.
    If it finds compressed files, it extracts and traverses them as well.
    """
    for root, _, files in os.walk(folder):
        for file in files:
            full_path = os.path.join(root, file)
            extension = file.lower().split('.')[-1]

            if extension in ["pdf", "epub"]:
                print(f"Found: {full_path}")
                try:
                    shutil.copy2(full_path, result_folder)
                except Exception as e:
                    print(f"Could not copy {full_path}: {e}")

            elif extension in ["zip", "rar", "7z"]:
                print(f"Extracting: {full_path}")
                new_folder = extract_compressed_file(full_path, temp_folder)
                if new_folder:  # only if extraction was successful
                    traverse_folder(new_folder, result_folder, temp_folder)


def main():
    # Initial folder
    initial_folder = input("Enter the path of the folder to traverse: ").strip()

    if not os.path.isdir(initial_folder):
        print("The path is not valid.")
        return

    # Results folder
    result_folder = os.path.join(initial_folder, "resultado")
    os.makedirs(result_folder, exist_ok=True)

    # Temporary folder for extracting compressed files
    temp_folder = os.path.join(initial_folder, "temp_zip")
    os.makedirs(temp_folder, exist_ok=True)

    # Traverse the folder
    traverse_folder(initial_folder, result_folder, temp_folder)

    print("\nProcess completed. Files saved in:", result_folder)


if __name__ == "__main__":
    main()
