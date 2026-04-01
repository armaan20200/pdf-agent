"""
Cleanup script for temporary files older than 1 hour
Run periodically or on startup
"""

import os
import time
from pathlib import Path

TEMP_DIR = Path("temporary_files")
MAX_AGE_SECONDS = 3600  # 1 hour


def cleanup_old_files():
    if not TEMP_DIR.exists():
        return

    now = time.time()
    removed = 0

    for file_path in TEMP_DIR.iterdir():
        if file_path.is_file():
            file_age = now - file_path.stat().st_mtime
            if file_age > MAX_AGE_SECONDS:
                try:
                    file_path.unlink()
                    removed += 1
                except Exception as e:
                    print(f"Could not remove {file_path}: {e}")

    if removed:
        print(f"Cleaned up {removed} temporary files")


if __name__ == "__main__":
    cleanup_old_files()
