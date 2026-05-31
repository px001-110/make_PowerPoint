import os
import time

FOLDERS = [
    "uploads",
    "generated",
    "temp"
]

# 秒
MAX_AGE = 60 * 60 * 24  # 24時間


def cleanup_old_files():
    now = time.time()

    for folder in FOLDERS:

        if not os.path.exists(folder):
            continue

        for filename in os.listdir(folder):

            if filename.startswith("."):
                continue

            filepath = os.path.join(folder, filename)

            if not os.path.isfile(filepath):
                continue

            file_age = now - os.path.getmtime(filepath)

            if file_age > MAX_AGE:
                try:
                    os.remove(filepath)
                    print(f"Deleted: {filepath}")

                except Exception as e:
                    print(f"Error deleting {filepath}: {e}")