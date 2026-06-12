import os
import re
import json
import io
import uuid
from threading import Thread, Lock
from concurrent.futures import ThreadPoolExecutor, as_completed

from flask import Flask, render_template, request, send_file, redirect, flash, jsonify
from werkzeug.utils import secure_filename

import openpyxl
import pdfplumber
import docx

from openai import OpenAI

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-key")

UPLOAD_FOLDER = "uploads"
ALLOWED_EXTENSIONS = {"pdf", "docx", "doc", "txt"}

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "")
client = OpenAI(api_key=OPENAI_API_KEY) if OPENAI_API_KEY else None

MAX_WORKERS = 4

tasks = {}
tasks_lock = Lock()


# ========================
# HELPERS
# ========================
def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def extract_text_from_pdf(filepath):
    text = ""
    try:
        with pdfplumber.open(filepath) as pdf:
            for page in pdf.pages[:5]:
                t = page.extract_text()
                if t:
                    text += t + "\n"
    except:
        pass
    return text


def extract_text_from_docx(filepath):
    text = ""
    try:
        doc = docx.Document(filepath)
        for para in doc.paragraphs[:200]:
            text += para.text + "\n"
    except:
        pass
    return text


def extract_text(filepath, ext):
    if ext == "pdf":
        return extract_text_from_pdf(filepath)
    elif ext in ("doc", "docx"):
        return extract_text_from_docx(filepath)
    return ""


# ========================
# IA
# ========================
def analyze_resume(text, vaga):
    if not client:
        return {"nome": "Sem API", "score": "-", "status": "-", "justificativa": "API não configurada"}

    prompt = f"""
VAGA:
{vaga}

CURRICULO:
{text[:8000]}

Retorne JSON: nome, whatsapp, email, score, status, justificativa
"""

    try:
        res = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}],
            timeout=20
        )

        content = res.choices[0].message.content.strip()

        content = re.sub(r"```.*?```", "", content, flags=re.DOTALL).strip()

        data = json.loads(content)

        return data
    except:
        return {"nome": "Erro IA", "score": "-", "status": "Erro", "justificativa": "Falha IA"}


# ========================
# BACKGROUND PROCESS
# ========================
def background_task(task_id, filepaths, vaga):
    results = []

    def process_path(path):
        try:
            ext = path.split(".")[-1].lower()
            text = extract_text(path, ext)

            os.remove(path)

            if not text.strip():
                return {"nome": "Erro leitura", "arquivo": os.path.basename(path), "score": "-"}

            result = analyze_resume(text, vaga)
            result["arquivo"] = os.path.basename(path)

            return result

        except Exception as e:
            return {"nome": "Erro", "arquivo": os.path.basename(path), "score": "-", "status": "Erro"}

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = [executor.submit(process_path, p) for p in filepaths]

        total = len(futures)
        done = 0

        for future in as_completed(futures):
            results.append(future.result())
            done += 1

            with tasks_lock:
                tasks[task_id]["progress"] = int((done / total) * 100)

    results.sort(
        key=lambda x: float(x.get("score", 0))
        if str(x.get("score", "")).isdigit() else -1,
        reverse=True
    )

    with tasks_lock:
        tasks[task_id]["done"] = True
        tasks[task_id]["results"] = results


# ========================
# ROTAS
# ========================
@app.route("/")
def index():
    return render_template("index.html")


@app.route("/processar", methods=["POST"])
def processar():
    vaga = request.form.get("job_profile", "")
    files = request.files.getlist("resumes")

    if not vaga:
        flash("Descreva a vaga")
        return redirect("/")

    if not files or all(f.filename == "" for f in files):
        flash("Envie arquivos")
        return redirect("/")

    task_id = str(uuid.uuid4())

    saved_files = []

    # 🚀 SALVA ARQUIVOS ANTES DA THREAD (CORREÇÃO DO BUG)
    for file in files:
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            path = os.path.join(UPLOAD_FOLDER, f"{task_id}_{filename}")
            file.save(path)
            saved_files.append(path)

    with tasks_lock:
        tasks[task_id] = {"progress": 0, "done": False, "results": []}

    thread = Thread(target=background_task, args=(task_id, saved_files, vaga))
    thread.start()

    return redirect(f"/status/{task_id}")


@app.route("/status/<task_id>")
def status(task_id):
    return render_template("status.html", task_id=task_id)


@app.route("/progresso/<task_id>")
def progresso(task_id):
    with tasks_lock:
        task = tasks.get(task_id)

    if not task:
        return jsonify({"error": "not found"}), 404

    return jsonify({
        "progress": task["progress"],
        "done": task["done"]
    })


@app.route("/resultado/<task_id>")
def resultado(task_id):
    with tasks_lock:
        task = tasks.get(task_id)

    if not task or not task["done"]:
        return redirect(f"/status/{task_id}")

    return render_template("resultado.html", results=task["results"])


@app.route("/exportar", methods=["POST"])
def exportar():
    results = json.loads(request.form.get("results_json"))

    wb = openpyxl.Workbook()
    ws = wb.active

    ws.append(["Nome", "Score", "Status", "Arquivo"])

    for r in results:
        ws.append([r.get("nome"), r.get("score"), r.get("status"), r.get("arquivo")])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(output, as_attachment=True, download_name="candidatos.xlsx")


if __name__ == "__main__":
    app.run(debug=False)
