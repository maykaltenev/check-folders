import os
import hashlib
import argparse


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
    """Compare two folders recursively."""
    files1 = get_all_files(folder1)
    files2 = get_all_files(folder2)

    only_in_folder1 = files1 - files2
    only_in_folder2 = files2 - files1
    common_files = files1 & files2

    # Show missing files
    if only_in_folder1 or only_in_folder2:
        print("\n❌ Missing or Extra Files:")

        if only_in_folder1:
            print(f"\n📂 Files present only in: {folder1}")
            for file in sorted(only_in_folder1):
                print(f"   ❌ {os.path.join(folder1, file)}")

        if only_in_folder2:
            print(f"\n📂 Files present only in: {folder2}")
            for file in sorted(only_in_folder2):
                print(f"   ❌ {os.path.join(folder2, file)}")
    else:
        print("\n✅ All expected files are present in both folders.")

    # Compare file contents
    different_files = []
    identical_files = []

    for file in sorted(common_files):
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


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Compare two folders recursively with detailed output.")
    parser.add_argument("folder1", help="Path to the first folder")
    parser.add_argument("folder2", help="Path to the second folder")
    args = parser.parse_args()

    compare_folders(args.folder1, args.folder2)
