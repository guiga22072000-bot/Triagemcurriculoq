import os
import re
import json
import io
from concurrent.futures import ThreadPoolExecutor, as_completed

from flask import Flask, render_template, request, send_file, redirect, url_for, flash
from werkzeug.utils import secure_filename

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

import pdfplumber
import docx

from openai import OpenAI

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-key")

UPLOAD_FOLDER = "uploads"
ALLOWED_EXTENSIONS = {"pdf", "docx", "doc", "txt"}

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024  # Limite 10MB

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Config API
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "")
OPENAI_BASE_URL = os.environ.get("OPENAI_BASE_URL", "https://api.openai.com/v1")
OPENAI_MODEL = os.environ.get("OPENAI_MODEL", "gpt-4o-mini")

client = OpenAI(api_key=OPENAI_API_KEY, base_url=OPENAI_BASE_URL) if OPENAI_API_KEY else None

# 🔥 CONTROLE DE CONCORRÊNCIA (ajuste entre 3–5)
MAX_WORKERS = 4


def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


# 🔥 OTIMIZADO: limita páginas do PDF (evita travamento)
def extract_text_from_pdf(filepath):
    text = ""
    try:
        with pdfplumber.open(filepath) as pdf:
            for page in pdf.pages[:5]:  # limite de páginas
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
    except Exception as e:
        text = f"[Erro ao ler PDF: {e}]"
    return text


def extract_text_from_docx(filepath):
    text = ""
    try:
        doc = docx.Document(filepath)
        for para in doc.paragraphs[:200]:  # limite
            text += para.text + "\n"
    except Exception as e:
        text = f"[Erro ao ler DOCX: {e}]"
    return text


def extract_text_from_txt(filepath):
    try:
        with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
            return f.read()
    except Exception as e:
        return f"[Erro ao ler TXT: {e}]"


def extract_text(filepath, ext):
    if ext == "pdf":
        return extract_text_from_pdf(filepath)
    elif ext in ("docx", "doc"):
        return extract_text_from_docx(filepath)
    elif ext == "txt":
        return extract_text_from_txt(filepath)
    return ""


def regex_fallback_extract(text):
    email_match = re.search(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", text)
    phone_match = re.search(r"(\+?\d{1,3}[\s.-]?)?\(?\d{2}\)?[\s.-]?\d{4,5}[\s.-]?\d{4}", text)
    first_lines = [l.strip() for l in text.split("\n") if l.strip()]
    name = first_lines[0] if first_lines else "Não identificado"

    return {
        "nome": name[:80],
        "whatsapp": phone_match.group(0) if phone_match else "Não encontrado",
        "email": email_match.group(0) if email_match else "Não encontrado",
    }


# 🔥 OTIMIZADO: menor input + timeout
def analyze_resume_with_ai(resume_text, job_profile):
    if not client:
        fallback = regex_fallback_extract(resume_text)
        fallback.update({
            "score": "N/A",
            "status": "Sem API",
            "justificativa": "API não configurada",
        })
        return fallback

    prompt = f"""
PERFIL:
{job_profile}

CURRÍCULO:
{resume_text[:8000]}

Retorne JSON com:
nome, whatsapp, email, score (0-100), status, justificativa
"""

    try:
        response = client.chat.completions.create(
            model=OPENAI_MODEL,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.2,
            timeout=20
        )

        content = response.choices[0].message.content.strip()
        content = re.sub(r"```.*?```", "", content, flags=re.DOTALL).strip()

        data = json.loads(content)

        for key in ["nome", "whatsapp", "email", "score", "status", "justificativa"]:
            if key not in data:
                data[key] = "N/A"

        return data

    except Exception as e:
        fallback = regex_fallback_extract(resume_text)
        fallback.update({
            "score": "Erro",
            "status": "Erro",
            "justificativa": str(e)[:150],
        })
        return fallback


# 🔥 PROCESSAMENTO INDIVIDUAL (PARALELO)
def process_file(file, job_profile):
    try:
        if not allowed_file(file.filename):
            return None

        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)

        file.save(filepath)

        ext = filename.rsplit(".", 1)[1].lower()
        text = extract_text(filepath, ext)

        os.remove(filepath)

        if not text.strip():
            return {
                "arquivo": filename,
                "nome": "Erro leitura",
                "whatsapp": "-",
                "email": "-",
                "score": "-",
                "status": "Erro",
                "justificativa": "Arquivo vazio ou inválido",
            }

        analysis = analyze_resume_with_ai(text, job_profile)
        analysis["arquivo"] = filename

        return analysis

    except Exception as e:
        return {
            "arquivo": file.filename,
            "nome": "Erro",
            "whatsapp": "-",
            "email": "-",
            "score": "-",
            "status": "Erro",
            "justificativa": str(e)[:150],
        }


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/processar", methods=["POST"])
def processar():
    job_profile = request.form.get("job_profile", "").strip()
    files = request.files.getlist("resumes")

    if not job_profile:
        flash("Descreva a vaga.")
        return redirect(url_for("index"))

    if not files or all(f.filename == "" for f in files):
        flash("Envie currículos.")
        return redirect(url_for("index"))

    results = []

    # 🔥 EXECUÇÃO PARALELA CONTROLADA
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = [executor.submit(process_file, f, job_profile) for f in files]

        for future in as_completed(futures):
            result = future.result()
            if result:
                results.append(result)

    def sort_key(r):
        try:
            return float(r.get("score", 0))
        except:
            return -1

    results.sort(key=sort_key, reverse=True)

    return render_template("resultado.html", results=results, job_profile=job_profile)


@app.route("/exportar", methods=["POST"])
def exportar():
    results_json = request.form.get("results_json")
    results = json.loads(results_json)

    wb = openpyxl.Workbook()
    ws = wb.active

    headers = ["Nome", "WhatsApp", "Email", "Score", "Status", "Justificativa", "Arquivo"]
    ws.append(headers)

    for r in results:
        ws.append([
            r.get("nome", ""),
            r.get("whatsapp", ""),
            r.get("email", ""),
            r.get("score", ""),
            r.get("status", ""),
            r.get("justificativa", ""),
            r.get("arquivo", ""),
        ])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="candidatos.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
