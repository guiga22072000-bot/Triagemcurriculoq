# 📋 Análise de Currículos com IA

App web que recebe vários currículos (PDF, DOCX, TXT), extrai os dados de cada candidato (nome, WhatsApp, email) e usa IA para avaliar a compatibilidade com uma vaga, gerando uma planilha Excel exportável.

---

## 🚀 Passo a passo completo

### 1. Criar conta na OpenAI e gerar a chave de API

1. Acesse https://platform.openai.com/ e crie uma conta (ou faça login).
2. Vá em **Settings → Billing** e adicione um cartão/crédito (mesmo modelos baratos como `gpt-4o-mini` exigem crédito mínimo).
3. Vá em **API Keys** (https://platform.openai.com/api-keys), clique em **Create new secret key** e copie a chave (algo como `sk-...`). Guarde — ela só aparece uma vez.

> Você pode usar outro provedor compatível com a API da OpenAI (ex: OpenRouter, Groq, etc). Nesse caso, defina `OPENAI_BASE_URL` apontando para o endpoint deles e `OPENAI_MODEL` com o nome do modelo suportado.

---

### 2. Subir o código no GitHub

1. Crie uma conta no https://github.com (se ainda não tiver).
2. Crie um novo repositório (botão **New repository**), dê um nome (ex: `analise-curriculos-ia`), deixe como **Public** ou **Private**, e clique em **Create repository**.
3. No seu computador, extraia o `.zip` deste projeto em uma pasta.
4. Abra o terminal dentro dessa pasta e rode:

```bash
git init
git add .
git commit -m "Primeira versão do app de análise de currículos"
git branch -M main
git remote add origin https://github.com/SEU_USUARIO/analise-curriculos-ia.git
git push -u origin main
```

(Substitua `SEU_USUARIO` pelo seu nome de usuário do GitHub.)

> Alternativa sem terminal: na página do repositório, clique em **uploading an existing file** e arraste todos os arquivos/pastas do projeto.

---

### 3. Criar o serviço no Render

1. Acesse https://render.com e crie uma conta (pode usar login do GitHub).
2. No painel, clique em **New → Web Service**.
3. Conecte sua conta do GitHub e selecione o repositório que você criou (`analise-curriculos-ia`).
4. O Render deve detectar automaticamente o arquivo `render.yaml`. Caso não detecte automaticamente, configure manualmente:
   - **Name**: `analise-curriculos-ia` (ou o nome que quiser)
   - **Region**: escolha a mais próxima (ex: Ohio ou São Paulo, se disponível)
   - **Branch**: `main`
   - **Runtime**: `Python 3`
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app`
   - **Instance Type**: Free (para começar) ou pago, se quiser mais performance

---

### 4. Configurar as variáveis de ambiente

Ainda na tela de configuração (ou em **Environment** depois de criado), adicione:

| Key | Value |
|---|---|
| `OPENAI_API_KEY` | A chave que você copiou no passo 1 (ex: `sk-...`) |
| `OPENAI_MODEL` | `gpt-4o-mini` (recomendado: bom custo/benefício) |
| `SECRET_KEY` | Qualquer texto aleatório (o Render pode gerar automaticamente) |

Clique em **Save Changes** (isso vai disparar um novo deploy automaticamente).

---

### 5. Fazer o deploy

1. Clique em **Create Web Service** (ou, se já criado, **Manual Deploy → Deploy latest commit**).
2. Aguarde o build terminar (acompanhe os logs). Quando aparecer algo como `Your service is live 🎉`, está pronto.
3. O Render vai te dar uma URL pública, algo como `https://analise-curriculos-ia.onrender.com`. Acesse essa URL no navegador.

---

### 6. Usando o app

1. No campo **"Perfil da vaga / requisitos"**, descreva detalhadamente a vaga: cargo, requisitos técnicos, experiência exigida, formação, soft skills, etc. Quanto mais detalhado, melhor a análise da IA.
2. Em **"Currículos"**, clique e selecione vários arquivos PDF/DOCX/TXT de uma vez (Ctrl ou Cmd + clique para selecionar múltiplos).
3. Clique em **"Analisar Currículos"**. Aguarde — cada currículo é processado individualmente pela IA (pode levar alguns segundos por arquivo).
4. Você verá uma tabela com: Nome, WhatsApp, Email, % de Compatibilidade, Status (Recomendado/Não recomendado) e uma justificativa.
5. Clique em **"Exportar para Excel (.xlsx)"** para baixar a planilha completa, já formatada e colorida por status.

---

## ⚙️ Detalhes técnicos

- **Backend**: Flask + Gunicorn
- **Extração de texto**: `pdfplumber` (PDF) e `python-docx` (Word)
- **IA**: API da OpenAI (`gpt-4o-mini` por padrão), retornando JSON estruturado por candidato
- **Exportação**: `openpyxl`, gera `.xlsx` com formatação e cores condicionais

## 🔒 Observações importantes

- Currículos em PDF escaneados (imagem, sem texto selecionável) não terão o texto extraído corretamente — nesse caso, o app indicará "Erro de leitura". Para suportar esses casos, seria necessário OCR (não incluído por padrão).
- Os arquivos enviados são processados e **excluídos imediatamente** após a leitura — nada é armazenado permanentemente no servidor.
- No plano gratuito do Render, o app "dorme" após período de inatividade e pode demorar ~30s para "acordar" na primeira requisição.
- Custos da IA: com `gpt-4o-mini`, o custo por currículo é muito baixo (centavos de dólar por currículo), mas depende do volume de uso — acompanhe seu uso em https://platform.openai.com/usage.

## 🛠️ Testando localmente (opcional)

```bash
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate
pip install -r requirements.txt
export OPENAI_API_KEY=sk-xxxxx   # Windows: set OPENAI_API_KEY=sk-xxxxx
python app.py
```

Acesse http://localhost:5000
