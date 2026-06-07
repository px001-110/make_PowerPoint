from flask import (
    Blueprint,
    render_template,
    request,
    send_file,
    current_app
)
from pptx import Presentation
from urllib.parse import quote
from pathlib import Path
from datetime import datetime
import re

from .services.generator import make_powerpoint
from .services.file_service import (
    save_upload_file,
    cleanup_old_files
)

main = Blueprint("main", __name__)


@main.before_app_request
def auto_cleanup():
    cleanup_old_files()


@main.route("/")
def index():
    return render_template("index.html")


@main.route("/upload", methods=["POST"])
def upload():

    upload_files = request.files.getlist("files")

    output_name = request.form.get("output_name", "").strip()

    if not output_name:
        today = datetime.now().strftime("%Y-%m-%d")
        output_name = f"{today}-自動生成"

    output_name = re.sub(r'[\\/*?:"<>|]', "", output_name)

    if not output_name:
        output_name = "自動生成"

    BASE_DIR = Path(__file__).resolve().parent.parent

    template_path = ( BASE_DIR / "app" / "templates" / "template.pptx")
    print(template_path.exists())
    print(template_path)

    prs = Presentation(str(template_path))

    for file in upload_files:

        save_path = save_upload_file(
            file,
            current_app.config["UPLOAD_DIR"]
        )

        try:
            make_powerpoint(
                str(save_path),
                prs
            )

        except Exception as e:
            current_app.logger.error(
                f"PowerPoint生成エラー: {e}"
            )

    output_path = (
        current_app.config["OUTPUT_DIR"]
        / f"{output_name}.pptx"
    )

    output_path = output_path.resolve()

    print(output_path)

    prs.save(output_path)

    print(output_path.exists())

    encoded_name = quote(f"{output_name}.pptx")

    response = send_file(
        str(output_path),
        as_attachment=True,
        download_name=f"{output_name}.pptx"
    )

    response.headers["Content-Disposition"] = (
        f"attachment; filename*=UTF-8''{encoded_name}"
    )

    return response