import os
import re
import json
import io
from flask import Flask, render_template, request, send_file, redirect, url_for, flash
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

# ConfiguraÃ§Ã£o da API (OpenAI ou compatÃ­vel)
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "")
OPENAI_BASE_URL = os.environ.get("OPENAI_BASE_URL", "https://api.openai.com/v1")
OPENAI_MODEL = os.environ.get("OPENAI_MODEL", "gpt-4o-mini")

client = OpenAI(api_key=OPENAI_API_KEY, base_url=OPENAI_BASE_URL) if OPENAI_API_KEY else None


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
    """ExtraÃ§Ã£o simples por regex como fallback, caso a IA falhe."""
    email_match = re.search(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", text)
    phone_match = re.search(
        r"(\+?\d{1,3}[\s.-]?)?\(?\d{2}\)?[\s.-]?\d{4,5}[\s.-]?\d{4}", text
    )
    # Nome: tenta pegar a primeira linha nÃ£o vazia significativa
    first_lines = [l.strip() for l in text.split("\n") if l.strip()]
    name = first_lines[0] if first_lines else "NÃ£o identificado"

    return {
        "nome": name[:80],
        "whatsapp": phone_match.group(0) if phone_match else "NÃ£o encontrado",
        "email": email_match.group(0) if email_match else "NÃ£o encontrado",
    }


def analyze_resume_with_ai(resume_text, job_profile):
    """
    Usa IA para extrair dados do candidato e avaliar compatibilidade com a vaga.
    Retorna um dicionÃ¡rio com nome, whatsapp, email, score, status, justificativa.
    """
    if not client:
        # Sem chave de API configurada -> usa fallback simples
        fallback = regex_fallback_extract(resume_text)
        fallback.update({
            "score": "N/A",
            "status": "Configurar API Key",
            "justificativa": "Chave da API de IA nÃ£o configurada no servidor.",
        })
        return fallback

    prompt = f"""
VocÃª Ã© um analista de RH especializado em recrutamento e seleÃ§Ã£o.

PERFIL DA VAGA (requisitos definidos pela empresa):
\"\"\"
{job_profile}
\"\"\"

CURRÃCULO DO CANDIDATO (texto extraÃ­do de um arquivo, pode conter ruÃ­dos de formataÃ§Ã£o):
\"\"\"
{resume_text[:12000]}
\"\"\"

Analise o currÃ­culo do candidato em relaÃ§Ã£o ao perfil da vaga e retorne SOMENTE um JSON vÃ¡lido (sem markdown, sem texto adicional) com os seguintes campos:

{{
  "nome": "Nome completo do candidato (apenas o nome prÃ³prio da pessoa, sem prefixos como 'Contato', 'CurrÃ­culo de', rÃ³tulos de seÃ§Ã£o ou textos de cabeÃ§alho/rodapÃ©)",
  "whatsapp": "NÃºmero de telefone/WhatsApp do candidato (ou 'NÃ£o encontrado')",
  "email": "Email do candidato (ou 'NÃ£o encontrado')",
  "score": "NÃºmero de 0 a 100 representando o percentual de compatibilidade do candidato com a vaga",
  "status": "Recomendado ou NÃ£o recomendado, com base no score (>=60 = Recomendado)",
  "justificativa": "Breve justificativa de 1 a 3 frases explicando o motivo do score, citando pontos fortes e lacunas em relaÃ§Ã£o Ã  vaga"
}}

Seja criterioso, justo e objetivo na anÃ¡lise. Considere experiÃªncia, habilidades tÃ©cnicas, formaÃ§Ã£o e aderÃªncia ao perfil descrito.

AtenÃ§Ã£o especial ao extrair o "nome": o texto pode vir de um PDF exportado do LinkedIn ou de um modelo com colunas, onde palavras como "Contato", "Perfil", "Resumo" ou Ã­cones de seÃ§Ã£o podem aparecer coladas ao nome. Extraia APENAS o nome prÃ³prio da pessoa (ex: "Roseni LeÃ£o", nÃ£o "Contato Roseni LeÃ£o").
"""

    try:
        response = client.chat.completions.create(
            model=OPENAI_MODEL,
            messages=[
                {"role": "system", "content": "VocÃª responde apenas com JSON vÃ¡lido, sem formataÃ§Ã£o markdown."},
                {"role": "user", "content": prompt},
            ],
            temperature=0.2,
        )
        content = response.choices[0].message.content.strip()

        # Remove possÃ­veis blocos de markdown
        content = re.sub(r"^```(json)?", "", content).strip()
        content = re.sub(r"```$", "", content).strip()

        data = json.loads(content)

        # Garante campos obrigatÃ³rios
        for key in ["nome", "whatsapp", "email", "score", "status", "justificativa"]:
            if key not in data:
                data[key] = "N/A"

        return data

    except Exception as e:
        fallback = regex_fallback_extract(resume_text)
        fallback.update({
            "score": "Erro",
            "status": "Erro na anÃ¡lise",
            "justificativa": f"Erro ao processar com IA: {str(e)[:200]}",
        })
        return fallback


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/processar", methods=["POST"])
def processar():
    job_profile = request.form.get("job_profile", "").strip()
    files = request.files.getlist("resumes")

    if not job_profile:
        flash("Por favor, descreva o perfil da vaga.")
        return redirect(url_for("index"))

    if not files or all(f.filename == "" for f in files):
        flash("Por favor, anexe ao menos um currÃ­culo.")
        return redirect(url_for("index"))

    results = []

    for file in files:
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
            file.save(filepath)

            ext = filename.rsplit(".", 1)[1].lower()
            text = extract_text(filepath, ext)

            if not text.strip():
                results.append({
                    "arquivo": filename,
                    "nome": "NÃ£o foi possÃ­vel ler o arquivo",
                    "whatsapp": "-",
                    "email": "-",
                    "score": "-",
                    "status": "Erro de leitura",
                    "justificativa": "O texto nÃ£o pÃ´de ser extraÃ­do (arquivo pode estar escaneado/imagem).",
                })
            else:
                analysis = analyze_resume_with_ai(text, job_profile)
                analysis["arquivo"] = filename
                results.append(analysis)

            try:
                os.remove(filepath)
            except OSError:
                pass

    # Ordena por score (maior primeiro), tratando valores nÃ£o numÃ©ricos
    def sort_key(r):
        try:
            return float(r.get("score", 0))
        except (ValueError, TypeError):
            return -1

    results.sort(key=sort_key, reverse=True)

    return render_template("resultado.html", results=results, job_profile=job_profile)


@app.route("/exportar", methods=["POST"])
def exportar():
    results_json = request.form.get("results_json")
    results = json.loads(results_json)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Candidatos"

    headers = ["Nome do Candidato", "WhatsApp", "Email", "Score (%)", "Status", "Justificativa", "Arquivo"]
    ws.append(headers)

    header_fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

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

        # Colorir linha conforme status
        row_idx = ws.max_row
        status_val = str(r.get("status", "")).lower()
        if "recomendado" in status_val and "nÃ£o" not in status_val and "nao" not in status_val:
            fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
        elif "nÃ£o recomendado" in status_val or "nao recomendado" in status_val:
            fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
        else:
            fill = None

        if fill:
            for cell in ws[row_idx]:
                cell.fill = fill

    # Ajusta largura das colunas
    widths = [28, 18, 28, 10, 18, 60, 25]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="candidatos_analisados.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)

