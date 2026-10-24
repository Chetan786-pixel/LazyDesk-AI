# ✦ LazyDesk AI

**A privacy-first, local AI document organization workspace powered by Gemma 3 and Ollama.**

LazyDesk AI helps turn scattered documents into an organized workspace. Upload PDF, DOCX, and TXT files, generate AI-powered summaries, categorize documents, suggest descriptive filenames, and export organized copies in one ZIP archive.

## ✨ Features

- **Smart categorization** — classify documents into meaningful categories.
- **AI summaries** — generate concise summaries from document content.
- **Suggested filenames** — create descriptive filenames for organized copies.
- **Interactive workspace** — review results, filter categories, sort documents, and edit AI-generated summaries.
- **ZIP export** — download categorized document copies and a JSON organization report.
- **Local-first AI** — run document analysis with Gemma 3 through Ollama.
- **Privacy-focused workflow** — document content is processed through the local Ollama service rather than a remote AI API.

## 🧰 Tech Stack

- Python
- Streamlit
- Ollama
- Gemma 3 (`gemma3:1b`)
- pypdf
- python-docx

## 📁 Supported Formats

| Format | Extension |
|---|---|
| PDF | `.pdf` |
| Microsoft Word | `.docx` |
| Plain text | `.txt` |

Maximum supported file size: 10 MB per file.

## 🚀 Getting Started

### 1. Clone the repository

```bash
git clone https://github.com/Chetan786-pixel/LazyDesk-AI.git
cd LazyDesk-AI
```

### 2. Create and activate a virtual environment

Windows PowerShell:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

### 3. Install dependencies

```bash
python -m pip install -r requirements.txt
```

### 4. Install and start Ollama

Install Ollama from [ollama.com](https://ollama.com/).

Pull the required model:

```bash
ollama pull gemma3:1b
```

Make sure the Ollama service is running before analyzing documents.

### 5. Start LazyDesk AI

```bash
python -m streamlit run app.py
```

Open the local URL printed in the terminal, usually `http://localhost:8501`.

## 🧪 How to Use

1. Start the app and verify that Ollama and Gemma 3 are available.
2. Upload one or more PDF, DOCX, or TXT documents.
3. Select **Analyze documents with AI**.
4. Review the generated categories, summaries, and suggested filenames.
5. Optionally edit and save the results.
6. Download the organized ZIP archive.

The ZIP includes categorized document copies and an `organization_report.json` file. Original uploaded files are not modified.

## 🏗️ How It Works

1. Streamlit provides the interactive document workspace.
2. Python extracts text from supported document formats.
3. The app sends extracted text to the locally running Gemma 3 model through Ollama.
4. The model returns structured categorization, summary, and filename suggestions.
5. Python packages organized copies and a JSON report into a downloadable ZIP archive.

## 🔐 Privacy

LazyDesk AI is designed for local-first document analysis. Document content is sent to the Ollama service configured on the machine, not to a remote AI API by this application. Use documents you are authorized to process and keep your local environment secure.

## ⚠️ Limitations

- Scanned PDFs without selectable text may require OCR before analysis.
- Results depend on the quality of extracted text and the model's output.
- Ollama and the configured Gemma 3 model must be available for AI analysis.
- This application is designed to run locally; a public deployment needs a suitable inference and privacy architecture.

## 🌱 Why Open Innovation Matters

LazyDesk AI demonstrates how open-weight models and local inference can make useful AI tools more accessible, customizable, and privacy-conscious. Developers can inspect the code, adapt the workflow, and build on the project without depending on a remote AI API for document analysis.

## 📌 Project

- **Repository:** [Chetan786-pixel/LazyDesk-AI](https://github.com/Chetan786-pixel/LazyDesk-AI)
- **Model:** Gemma 3 through Ollama
- **Challenge tag:** `#hf26challenge`

---

Built to make everyday document work simpler.
