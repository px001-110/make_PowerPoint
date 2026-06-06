from pathlib import Path
from werkzeug.utils import secure_filename
import uuid
import time


def save_upload_file(file, upload_dir):

    original_name = file.filename

    ext = Path(original_name).suffix

    safe_name = f"{uuid.uuid4()}{ext}"

    save_path = upload_dir / safe_name

    file.save(str(save_path))

    return save_path


def cleanup_old_files(
    upload_dir=None,
    output_dir=None,
    expire_seconds=3600
):
    """
    1時間以上経過したファイルを削除
    """

    now = time.time()

    targets = []

    if upload_dir:
        targets.append(upload_dir)

    if output_dir:
        targets.append(output_dir)

    for directory in targets:

        if not directory.exists():
            continue

        for file in directory.iterdir():

            if not file.is_file():
                continue

            age = now - file.stat().st_mtime

            if age > expire_seconds:
                try:
                    file.unlink()
                except Exception as e:
                    print(
                        f"削除失敗: {file.name} - {e}"
                    )