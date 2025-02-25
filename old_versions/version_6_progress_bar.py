import os
import hashlib
import argparse
import openpyxl
from tqdm import tqdm


def get_all_files(folder):
    """Recursively get all file paths in a folder, relative to the base folder."""
    file_paths = []
    for root, _, files in os.walk(folder):
        for file in files:
            relative_path = os.path.relpath(os.path.join(root, file), folder)
            file_paths.append(relative_path)
    return set(file_paths)


def get_file_hash(filepath):
    """Generate SHA256 hash of a file."""
    hash_sha256 = hashlib.sha256()
    with open(filepath, "rb") as f:
        for chunk in iter(lambda: f.read(4096), b""):
            hash_sha256.update(chunk)
    return hash_sha256.hexdigest()


def compare_folders(folder1, folder2):
    """Compare two folders recursively and generate an Excel report."""
    print(f"\n🔍 Comparing folders:\n  📁 {folder1}\n  📁 {folder2}")

    files1 = get_all_files(folder1)
    files2 = get_all_files(folder2)

    only_in_folder1 = sorted(files1 - files2)
    only_in_folder2 = sorted(files2 - files1)
    common_files = sorted(files1 & files2)

    # Display missing or extra files
    if only_in_folder1 or only_in_folder2:
        print("\n❌ Missing or Extra Files:")

        if only_in_folder1:
            print(f"\n📂 Files present only in: {folder1}")
            for file in only_in_folder1:
                print(f"   ❌ {file}")

        if only_in_folder2:
            print(f"\n📂 Files present only in: {folder2}")
            for file in only_in_folder2:
                print(f"   ❌ {file}")
    else:
        print("\n✅ All expected files are present in both folders.")

    # Compare file contents
    different_files = []
    identical_files = []

    print("\n🔄 Comparing file contents...")
    for file in tqdm(common_files, desc="Processing files"):
        file1_path = os.path.join(folder1, file)
        file2_path = os.path.join(folder2, file)

        hash1 = get_file_hash(file1_path)
        hash2 = get_file_hash(file2_path)

        if hash1 != hash2:
            different_files.append(file)
        else:
            identical_files.append(file)

    # Show content comparison results
    if different_files:
        print("\n⚠️ Files that differ in content:")
        for file in different_files:
            print(f"   ❌ {file}")

    if identical_files:
        print("\n✅ Identical files:")
        for file in identical_files:
            print(f"   ✔️ {file}")

    # Save results to Excel
    save_to_excel(folder1, folder2, only_in_folder1,
                  only_in_folder2, different_files, identical_files)


def save_to_excel(folder1, folder2, only_in_folder1, only_in_folder2, different_files, identical_files):
    """Save comparison results to an Excel file."""
    workbook = openpyxl.Workbook()
    ws = workbook.active
    ws.title = "File Comparison"

    # Set up headers
    ws.append(["Status", "File Path"])

    # Write missing/extra files
    for file in only_in_folder1:
        ws.append(["❌ Missing in Folder 2", file])
    for file in only_in_folder2:
        ws.append(["❌ Missing in Folder 1", file])

    # Write different and identical files
    for file in different_files:
        ws.append(["⚠️ Different Content", file])
    for file in identical_files:
        ws.append(["✅ Identical", file])

    # Save the workbook
    output_file = "comparison_results.xlsx"
    workbook.save(output_file)
    print(f"\n📊 Results saved to {output_file}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Compare two folders recursively and generate an Excel report.")
    parser.add_argument("folder1", help="Path to the first folder")
    parser.add_argument("folder2", help="Path to the second folder")
    args = parser.parse_args()

    compare_folders(args.folder1, args.folder2)
