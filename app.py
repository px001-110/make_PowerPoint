from flask import Flask, render_template, request, send_file
from pathlib import Path
from werkzeug.utils import secure_filename
from pptx import Presentation
from generator import PowerPointGenerator
import uuid
import re
from urllib.parse import quote
from cleanup import cleanup_old_files
from datetime import datetime


app = Flask(__name__)

cleanup_old_files()

BASE_DIR = Path(__file__).parent

UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "generated"

UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

@app.before_request
def auto_cleanup():
    cleanup_old_files()

# 操作画面
@app.route('/')
def index():
    return render_template('index.html')

# ファイルのアップロード
@app.route('/upload', methods=["POST"])
def upload():

    upload_files = request.files.getlist('files')

    output_name = request.form.get("output_name", "")
    print(output_name)
 
    # 空白を除去
    output_name = output_name.strip()
 
    #　未入力対策
    if output_name == "":
        today = datetime.now().strftime("%Y-%m-%d")
        output_name = f"{today}-自動生成"

    print(output_name)
    # 禁止文字を除去
    output_name = re.sub(r'[\\/*?:"<>|]', '', output_name)

    if output_name == "":
        output_name = "自動生成"

    prs = Presentation("templates/template.pptx")

    generator = PowerPointGenerator()

    for file in upload_files:

        original_name = file.filename

        ext = Path(original_name).suffix

        safe_name = f"{uuid.uuid4()}{ext}"

        save_path = UPLOAD_DIR / safe_name

        file.save(str(save_path))

        try:            
            generator.make_powerpoint(str(save_path), prs)

        except Exception as e:
            app.logger.error(str(e))

    output_path = OUTPUT_DIR / f"{output_name}.pptx"

    prs.save(output_path)

    encoded_name = quote(f"{output_name}.pptx")

    response = send_file(
        output_path,
        as_attachment=True,
        download_name=f"{output_name}.pptx"
    )

    response.headers["Content-Disposition"] = f"attachment; filename*=UTF-8''{encoded_name}"

    return response


if __name__ == '__main__':
    app.run(debug=True)