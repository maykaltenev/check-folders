import os
import hashlib
import argparse


def get_all_files(folder):
    """Recursively get all file paths in a folder."""
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
    """Compare two folders recursively."""
    files1 = get_all_files(folder1)
    files2 = get_all_files(folder2)

    # Check for missing or extra files
    only_in_folder1 = files1 - files2
    only_in_folder2 = files2 - files1

    if only_in_folder1 or only_in_folder2:
        print("\n❌ Differences in file structure:")
        if only_in_folder1:
            print(f"Only in {folder1}: {only_in_folder1}")
        if only_in_folder2:
            print(f"Only in {folder2}: {only_in_folder2}")
    else:
        print("\n✅ File structure matches.")

    # Compare file contents
    for file in files1 & files2:
        file1_path = os.path.join(folder1, file)
        file2_path = os.path.join(folder2, file)

        hash1 = get_file_hash(file1_path)
        hash2 = get_file_hash(file2_path)

        if hash1 != hash2:
            print(f"❌ {file} differs in content.")
        else:
            print(f"✅ {file} is identical.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Compare two folders recursively.")
    parser.add_argument("folder1", help="Path to the first folder")
    parser.add_argument("folder2", help="Path to the second folder")
    args = parser.parse_args()

    compare_folders(args.folder1, args.folder2)
