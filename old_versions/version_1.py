import os
import hashlib


def get_file_hash(filepath):
    """Generate SHA256 hash of a file."""
    hash_sha256 = hashlib.sha256()
    with open(filepath, "rb") as f:
        for chunk in iter(lambda: f.read(4096), b""):
            hash_sha256.update(chunk)
    return hash_sha256.hexdigest()


def compare_folders(folder1, folder2):
    """Compare the contents of two folders."""
    files1 = set(os.listdir(folder1))
    files2 = set(os.listdir(folder2))

    # Check for missing or extra files
    only_in_folder1 = files1 - files2
    only_in_folder2 = files2 - files1

    if only_in_folder1 or only_in_folder2:
        print("Differences in file names:")
        if only_in_folder1:
            print(f"Only in {folder1}: {only_in_folder1}")
        if only_in_folder2:
            print(f"Only in {folder2}: {only_in_folder2}")
    else:
        print("File names match.")

    # Compare file contents
    for file in files1 & files2:
        file1_path = os.path.join(folder1, file)
        file2_path = os.path.join(folder2, file)

        if os.path.isfile(file1_path) and os.path.isfile(file2_path):
            hash1 = get_file_hash(file1_path)
            hash2 = get_file_hash(file2_path)

            if hash1 != hash2:
                print(f"Files {file} differ in content.")
            else:
                print(f"Files {file} are identical.")


folder1 = "/Users/begemoth/Documents/test_1"
folder2 = "/Users/begemoth/Documents/test_2"

compare_folders(folder1, folder2)
