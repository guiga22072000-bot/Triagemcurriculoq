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

Seja extremamente criterioso e crítico na análise.

Atribua notas mais realistas (evite inflar scores).
Candidates medianos devem ficar entre 50-70.
Somente perfis realmente fortes devem ultrapassar 80.

Considere:
- aderência técnica real
- experiência prática comprovada
- profundidade das habilidades
- coerência profissional

Se faltarem requisitos importantes, reduza significativamente o score.

Evite avaliações genéricas ou superficiais.
Justifique de forma objetiva os pontos fortes e as lacunas.

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
        def process_job(job_id, filepaths, job_profile):
    """
    Função que roda em background (thread) para processar os currículos
    """
    try:
        from concurrent.futures import ThreadPoolExecutor

def process_job(job_id, filepaths, job_profile):
    try:
        def process_single(filepath):
            ext = filepath.rsplit(".", 1)[1].lower()
            text = extract_text(filepath, ext)

            result = analyze_resume_with_ai(text, job_profile)
            result["arquivo"] = os.path.basename(filepath)

            with JOBS_LOCK:
                JOBS[job_id]["results"].append(result)
                JOBS[job_id]["done"] += 1

        # 🔥 roda até 3 currículos ao mesmo tempo (evita travar API)
        with ThreadPoolExecutor(max_workers=3) as executor:
            executor.map(process_single, filepaths)

        def safe_score(x):
            try:
                return int(x.get("score", 0))
            except:
                return 0

        with JOBS_LOCK:
            JOBS[job_id]["results"].sort(key=safe_score, reverse=True)
            JOBS[job_id]["status"] = "done"

    except Exception as e:
        with JOBS_LOCK:
            JOBS[job_id]["status"] = "error"
            JOBS[job_id]["error"] = str(e)
            ext = filepath.rsplit(".", 1)[1].lower()
            text = extract_text(filepath, ext)

            result = analyze_resume_with_ai(text, job_profile)
            result["arquivo"] = os.path.basename(filepath)

            with JOBS_LOCK:
                JOBS[job_id]["results"].append(result)
                JOBS[job_id]["done"] += 1

        # Ordena por score (melhor primeiro)
        def safe_score(x):
            try:
                return int(x.get("score", 0))
            except:
                return 0

        with JOBS_LOCK:
            JOBS[job_id]["results"].sort(key=safe_score, reverse=True)
            JOBS[job_id]["status"] = "done"

    except Exception as e:
        with JOBS_LOCK:
            JOBS[job_id]["status"] = "error"
            JOBS[job_id]["error"] = str(e)


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/processar", methods=["POST"])
def processar():
    job_profile = request.form.get("job_profile", "").strip()
    files = request.files.getlist("resumes")

    if not job_profile or not files or files[0].filename == "":
        flash("Preencha o perfil da vaga e selecione pelo menos um currículo.")
        return redirect(url_for("index"))

    saved_paths = []

    # Salva os arquivos
    for file in files:
        if file and allowed_file(file.filename):
            filename = f"{uuid.uuid4()}_{secure_filename(file.filename)}"
            filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
            file.save(filepath)
            saved_paths.append(filepath)

    if not saved_paths:
        flash("Nenhum arquivo válido foi enviado.")
        return redirect(url_for("index"))

    job_id = str(uuid.uuid4())

    with JOBS_LOCK:
        JOBS[job_id] = {
            "total": len(saved_paths),
            "done": 0,
            "results": [],
            "job_profile": job_profile,
            "status": "processing",
            "error": None,
        }

    # Cria thread em background
    thread = threading.Thread(
        target=process_job,
        args=(job_id, saved_paths, job_profile),
        daemon=True
    )
    thread.start()

    # Redireciona para página de progresso
    return redirect(url_for("progresso", job_id=job_id))


@app.route("/progresso/<job_id>")
def progresso(job_id):
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
        "status": job["status"]
    })


@app.route("/resultado/<job_id>")
def resultado(job_id):
    with JOBS_LOCK:
        job = JOBS.get(job_id)

    if not job:
        flash("Job não encontrado.")
        return redirect(url_for("index"))

    if job["status"] != "done":
        return redirect(url_for("progresso", job_id=job_id))

    return render_template(
        "resultado.html",
        results=job["results"],
        job_profile=job["job_profile"]
    )


@app.route("/exportar", methods=["POST"])
def exportar():
    results_json = request.form.get("results_json")
    results = json.loads(results_json)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Resultados"

    headers = ["Nome", "WhatsApp", "Email", "Score", "Status", "Justificativa", "Arquivo"]
    ws.append(headers)

    for r in results:
        ws.append([
            r.get("nome"),
            r.get("whatsapp"),
            r.get("email"),
            r.get("score"),
            r.get("status"),
            r.get("justificativa"),
            r.get("arquivo"),
        ])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        download_name="resultados.xlsx",
        as_attachment=True
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
