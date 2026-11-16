# -*- coding: utf-8 -*-
"""Flask web demo: upload an image-based PDF, download an editable PPTX.

Run:  python app.py            (default http://127.0.0.1:7860)
      PORT=8000 python app.py
"""
import os
import tempfile
import traceback
import uuid

from flask import Flask, render_template, request, send_file, jsonify

from converter import convert_pdf_to_pptx

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 32 * 1024 * 1024  # 32 MB


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/convert", methods=["POST"])
def convert():
    f = request.files.get("pdf")
    if not f or not f.filename.lower().endswith(".pdf"):
        return jsonify({"error": "请上传 PDF 文件"}), 400
    dpi = int(request.form.get("dpi", 150))
    with tempfile.TemporaryDirectory() as td:
        in_path = os.path.join(td, "input.pdf")
        out_path = os.path.join(td, "output.pptx")
        f.save(in_path)
        try:
            manifest = convert_pdf_to_pptx(in_path, out_path, dpi=dpi)
        except Exception:
            traceback.print_exc()
            return jsonify({"error": "转换失败: " + traceback.format_exc(limit=1)}), 500
        n_txt = sum(len(p["texts"]) for p in manifest)
        data = open(out_path, "rb").read()
    from io import BytesIO
    name = os.path.splitext(os.path.basename(f.filename))[0] + "_editable.pptx"
    resp = send_file(BytesIO(data), as_attachment=True, download_name=name,
                     mimetype="application/vnd.openxmlformats-officedocument."
                              "presentationml.presentation")
    resp.headers["X-Stats"] = f"slides={len(manifest)};texts={n_txt}"
    return resp


if __name__ == "__main__":
    # 魔搭 Docker 创空间有两级探测：supervisor 按 README app_port(8080) 等待，
    # 功能实例健康检查固定探 7860。同时监听两个端口才能通过。
    import threading
    from werkzeug.serving import make_server

    ports = sorted({int(os.environ.get("PORT", 7860)), 7860, 8080})
    servers = [make_server("0.0.0.0", p, app) for p in ports]
    for s in servers[1:]:
        threading.Thread(target=s.serve_forever, daemon=True).start()
    print(f"listening on ports {ports}")
    servers[0].serve_forever()
