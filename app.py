import os
import re
import json
import io
import uuid
import threading
import time
from flask import Flask, render_template, request, send_file, redirect, url_for, flash, jsonify
from werkzeug.utils import secure_filename
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

# Leitura de PDFs e DOCX
import pdfplumber
import docx

from openai import OpenAI

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-key")

UPLOAD_FOLDER = "uploads"
ALLOWED_EXTENSIONS = {"pdf", "docx", "doc", "txt"}
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Configuração da API (OpenAI ou compatível)
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "")
OPENAI_BASE_URL = os.environ.get("OPENAI_BASE_URL", "https://api.openai.com/v1")
OPENAI_MODEL = os.environ.get("OPENAI_MODEL", "gpt-4o-mini")

client = OpenAI(api_key=OPENAI_API_KEY, base_url=OPENAI_BASE_URL) if OPENAI_API_KEY else None

# Armazenamento em memória dos jobs de processamento em background
# Estrutura: { job_id: { "total": int, "done": int, "results": [...], "job_profile": str, "status": "processing"|"done"|"error", "error": str|None } }
JOBS = {}
JOBS_LOCK = threading.Lock()


def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def extract_text_from_pdf(filepath):
    text = ""
    try:
        with pdfplumber.open(filepath) as pdf:
            for page in pdf.pages:
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
        for para in doc.paragraphs:
            text += para.text + "\n"
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    text += cell.text + " "
                text += "\n"
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
    """Extração simples por regex como fallback, caso a IA falhe."""
    email_match = re.search(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", text)
    phone_match = re.search(
        r"(\+?\d{1,3}[\s.-]?)?\(?\d{2}\)?[\s.-]?\d{4,5}[\s.-]?\d{4}", text
    )
    # Nome: tenta pegar a primeira linha não vazia significativa
    first_lines = [l.strip() for l in text.split("\n") if l.strip()]
    name = first_lines[0] if first_lines else "Não identificado"

    return {
        "nome": name[:80],
        "whatsapp": phone_match.group(0) if phone_match else "Não encontrado",
        "email": email_match.group(0) if email_match else "Não encontrado",
    }


def analyze_resume_with_ai(resume_text, job_profile):
    """
    Usa IA para extrair dados do candidato e avaliar compatibilidade com a vaga.
    Retorna um dicionário com nome, whatsapp, email, score, status, justificativa.
    """
    if not client:
        # Sem chave de API configurada -> usa fallback simples
        fallback = regex_fallback_extract(resume_text)
        fallback.update({
            "score": "N/A",
            "status": "Configurar API Key",
            "justificativa": "Chave da API de IA não configurada no servidor.",
        })
        return fallback

    prompt = f"""
Você é um analista de RH especializado em recrutamento e seleção.

PERFIL DA VAGA (requisitos definidos pela empresa):
\"\"\"
{job_profile}
\"\"\"

CURRÍCULO DO CANDIDATO (texto extraído de um arquivo, pode conter ruídos de formatação):
\"\"\"
{resume_text[:12000]}
\"\"\"

Analise o currículo do candidato em relação ao perfil da vaga e retorne SOMENTE um JSON válido (sem markdown, sem texto adicional) com os seguintes campos:

{{
  "nome": "Nome completo do candidato (apenas o nome próprio da pessoa, sem prefixos como 'Contato', 'Currículo de', rótulos de seção ou textos de cabeçalho/rodapé)",
  "whatsapp": "Número de telefone/WhatsApp do candidato (ou 'Não encontrado')",
  "email": "Email do candidato (ou 'Não encontrado')",
  "score": "Número de 0 a 100 representando o percentual de compatibilidade do candidato com a vaga",
  "status": "Recomendado ou Não recomendado, com base no score (>=60 = Recomendado)",
  "justificativa": "Breve justificativa de 1 a 3 frases explicando o motivo do score, citando pontos fortes e lacunas em relação à vaga"
}}

Seja criterioso, justo e objetivo na análise. Considere experiência, habilidades técnicas, formação e aderência ao perfil descrito.

Atenção especial ao extrair o "nome": o texto pode vir de um PDF exportado do LinkedIn ou de um modelo com colunas, onde palavras como "Contato", "Perfil", "Resumo" ou ícones de seção podem aparecer coladas ao nome. Extraia APENAS o nome próprio da pessoa (ex: "Roseni Leão", não "Contato Roseni Leão").
"""

    try:
        response = client.chat.completions.create(
            model=OPENAI_MODEL,
            messages=[
                {"role": "system", "content": "Você responde apenas com JSON válido, sem formatação markdown."},
                {"role": "user", "content": prompt},
            ],
            temperature=0.2,
        )
        content = response.choices[0].message.content.strip()

        # Remove possíveis blocos de markdown
        content = re.sub(r"^```(json)?", "", content).strip()
        content = re.sub(r"```$", "", content).strip()

        data = json.loads(content)

        # Garante campos obrigatórios
        for key in ["nome", "whatsapp", "email", "score", "status", "justificativa"]:
            if key not in data:
                data[key] = "N/A"

        return data

    except Exception as e:
        fallback = regex_fallback_extract(resume_text)
        fallback.update({
            "score": "Erro",
            "status": "Erro na análise",
            "justificativa": f"Erro ao processar com IA: {str(e)[:200]}",
        })
        return fallback
        def process_resumes_background(job_id, filepaths_exts, job_profile):
    """Processa os currículos em background, atualizando o job em JOBS."""
    results = []
    total = len(filepaths_exts)

    with JOBS_LOCK:
        JOBS[job_id]["total"] = total
        JOBS[job_id]["done"] = 0
        JOBS[job_id]["status"] = "processing"

    for i, (filepath, ext, original_name) in enumerate(filepaths_exts):
        try:
            text = extract_text(filepath, ext)
            result = analyze_resume_with_ai(text, job_profile)
            result["arquivo"] = original_name
        except Exception as e:
            result = {
                "arquivo": original_name,
                "nome": "Erro",
                "whatsapp": "Erro",
                "email": "Erro",
                "score": "Erro",
                "status": "Erro",
                "justificativa": str(e)[:200],
            }
        finally:
            try:
                os.remove(filepath)
            except Exception:
                pass

        results.append(result)

        with JOBS_LOCK:
            JOBS[job_id]["done"] = i + 1
            JOBS[job_id]["results"] = results

        time.sleep(0.1)  # Pequena pausa para não sobrecarregar a API

    with JOBS_LOCK:
        JOBS[job_id]["status"] = "done"


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        job_profile = request.form.get("job_profile", "").strip()
        files = request.files.getlist("resumes")

        if not job_profile:
            flash("Por favor, descreva o perfil da vaga.", "warning")
            return redirect(url_for("index"))

        if not files or all(f.filename == "" for f in files):
            flash("Nenhum arquivo selecionado.", "warning")
            return redirect(url_for("index"))

        filepaths_exts = []
        for f in files:
            if f and allowed_file(f.filename):
                filename = secure_filename(f.filename)
                unique_name = f"{uuid.uuid4().hex}_{filename}"
                filepath = os.path.join(app.config["UPLOAD_FOLDER"], unique_name)
                f.save(filepath)
                ext = filename.rsplit(".", 1)[1].lower()
                filepaths_exts.append((filepath, ext, f.filename))

        if not filepaths_exts:
            flash("Nenhum arquivo válido enviado (aceitos: PDF, DOCX, TXT).", "danger")
            return redirect(url_for("index"))

        # Cria o job e dispara thread
        job_id = uuid.uuid4().hex
        with JOBS_LOCK:
            JOBS[job_id] = {
                "total": len(filepaths_exts),
                "done": 0,
                "results": [],
                "job_profile": job_profile,
                "status": "processing",
                "error": None,
            }

        thread = threading.Thread(
            target=process_resumes_background,
            args=(job_id, filepaths_exts, job_profile),
            daemon=True,
        )
        thread.start()

        return redirect(url_for("progresso", job_id=job_id))

    return render_template("index.html")


@app.route("/progresso/<job_id>")
def progresso(job_id):
    with JOBS_LOCK:
        job = JOBS.get(job_id)
    if not job:
        flash("Job não encontrado.", "danger")
        return redirect(url_for("index"))
    return render_template("progresso.html", job_id=job_id)


@app.route("/status/<job_id>")
def status(job_id):
    with JOBS_LOCK:
        job = JOBS.get(job_id)
    if not job:
        return jsonify({"error": "Job não encontrado"}), 404
    return jsonify({
        "total": job["total"],
        "done": job["done"],
        "status": job["status"],
        "error": job["error"],
    })


@app.route("/resultado/<job_id>")
def resultado(job_id):
    with JOBS_LOCK:
        job = JOBS.get(job_id)
    if not job or job["status"] != "done":
        flash("Resultado ainda não disponível.", "warning")
        return redirect(url_for("progresso", job_id=job_id))

    results = job["results"]
    # Ordena por score decrescente
    def sort_key(r):
        try:
            return -int(r.get("score", 0))
        except (ValueError, TypeError):
            return 0

    results_sorted = sorted(results, key=sort_key)
    results_json = json.dumps(results_sorted)
    return render_template(
        "resultado.html",
        results=results_sorted,
        results_json=results_json,
        job_profile=job["job_profile"],
    )


@app.route("/exportar", methods=["POST"])
def exportar():
    results_json = request.form.get("results_json", "[]")
    try:
        results = json.loads(results_json)
    except Exception:
        results = []

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Triagem de Currículos"

    headers = ["Arquivo", "Nome", "WhatsApp", "Email", "Score (%)", "Status", "Justificativa"]
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)

    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    for row_idx, r in enumerate(results, 2):
        ws.cell(row=row_idx, column=1, value=r.get("arquivo", ""))
        ws.cell(row=row_idx, column=2, value=r.get("nome", ""))
        ws.cell(row=row_idx, column=3, value=r.get("whatsapp", ""))
        ws.cell(row=row_idx, column=4, value=r.get("email", ""))
        ws.cell(row=row_idx, column=5, value=r.get("score", ""))
        status_cell = ws.cell(row=row_idx, column=6, value=r.get("status", ""))
        ws.cell(row=row_idx, column=7, value=r.get("justificativa", ""))

        fill = green_fill if "Recomendado" in str(r.get("status", "")) else red_fill
        for col in range(1, 8):
            ws.cell(row=row_idx, column=col).fill = fill

    col_widths = [30, 25, 18, 30, 12, 20, 60]
    for col, width in enumerate(col_widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = width

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="triagem_curriculos.xlsx",
    )


if __name__ == "__main__":
    app.run(debug=True)
