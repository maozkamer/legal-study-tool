import os
import io
import logging
from flask import Flask, request, render_template, jsonify

import PyPDF2
from pptx import Presentation
from docx import Document
import anthropic

logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s")
log = logging.getLogger(__name__)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # 20 MB

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")

# ─────────────────────────────────────────────────────────────
#  Text extraction
# ─────────────────────────────────────────────────────────────

def extract_text_pdf(file_bytes: bytes) -> str:
    reader = PyPDF2.PdfReader(io.BytesIO(file_bytes))
    parts = []
    for page in reader.pages:
        text = page.extract_text()
        if text:
            parts.append(text)
    return "\n".join(parts)


def extract_text_pptx(file_bytes: bytes) -> str:
    prs = Presentation(io.BytesIO(file_bytes))
    parts = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    line = " ".join(run.text for run in para.runs).strip()
                    if line:
                        parts.append(line)
    return "\n".join(parts)


def extract_text_docx(file_bytes: bytes) -> str:
    doc = Document(io.BytesIO(file_bytes))
    parts = [para.text.strip() for para in doc.paragraphs if para.text.strip()]
    return "\n".join(parts)


def extract_text_txt(file_bytes: bytes) -> str:
    for encoding in ("utf-8", "windows-1255", "iso-8859-8", "latin-1"):
        try:
            return file_bytes.decode(encoding)
        except (UnicodeDecodeError, LookupError):
            continue
    return file_bytes.decode("utf-8", errors="replace")


# ─────────────────────────────────────────────────────────────
#  Document type detection
# ─────────────────────────────────────────────────────────────

VERDICT_KEYWORDS    = ["בית משפט", "נגד", "פסק דין", "השופט", "הנאשם", "המשיב", "המערער", "פסק-דין"]
LEGISLATION_KEYWORDS = ["חוק", "סעיף", "תקנות", "כנסת", "ספר החוקים", "תשכ", "תשנ", "תש\"", "רשומות"]


def detect_doc_type(text: str) -> str:
    sample = text[:3000].lower()
    verdict_hits    = sum(1 for kw in VERDICT_KEYWORDS    if kw in sample)
    legislation_hits = sum(1 for kw in LEGISLATION_KEYWORDS if kw in sample)

    if verdict_hits >= 2:
        return "verdict"
    if legislation_hits >= 2:
        return "legislation"
    return "study"


# ─────────────────────────────────────────────────────────────
#  Prompt builder
# ─────────────────────────────────────────────────────────────

def build_prompt(doc_type: str, text: str) -> str:
    truncated = text[:12000]

    if doc_type == "verdict":
        instructions = """אתה עוזר לימודים משפטי. סכם את פסק הדין הבא בעברית בפורמט המדויק הזה:

📋 **פרטי התיק:** [בית משפט + שנה + שמות הצדדים]

⚖️ **השאלה המשפטית:** [מה הייתה השאלה המשפטית שנדונה]

📖 **עובדות המקרה:** [עובדות עיקריות בנקודות קצרות]

💬 **טענות הצדדים:** [טענות התובע/מדינה מול הנתבע/נאשם]

🔨 **פסיקה:** [מה פסק בית המשפט ומדוע]

📌 **העיקרון המשפטי לבחינה:** [הכלל או ההלכה החשובים שעולים מפסק הדין]"""

    elif doc_type == "legislation":
        instructions = """אתה עוזר לימודים משפטי. סכם את החקיקה הבאה בעברית בפורמט המדויק הזה:

📋 **שם החוק ושנה:** [שם מלא + שנת חקיקה]

🎯 **מטרת החוק:** [מה החוק בא להסדיר ולמה]

📖 **סעיפים מרכזיים:** [הסעיפים החשובים ביותר עם הסבר קצר לכל אחד]

⚠️ **חריגים חשובים:** [חריגים, תנאים מיוחדים, הגנות]

🔗 **קשר לפסיקה רלוונטית:** [אם ניתן לזהות — פסיקה קשורה]

📌 **מה חשוב לדעת לבחינה:** [הנקודות הכי חשובות לשינון]"""

    else:
        instructions = """אתה עוזר לימודים משפטי. סכם את חומר הלימוד הבא בעברית בפורמט המדויק הזה:

🎯 **נושא המצגת/החומר:** [נושא ראשי]

📖 **רעיונות מרכזיים:**
1. [רעיון ראשון]
2. [רעיון שני]
3. [המשך לפי הצורך]

💡 **מושגים חשובים להבין:** [הגדרות ומושגי מפתח]

❓ **שאלות אפשריות לבחינה:**
- [שאלה אפשרית 1]
- [שאלה אפשרית 2]
- [שאלה אפשרית 3]

📌 **סיכום בשורה אחת:** [תמצית החומר כולו]"""

    return f"{instructions}\n\n---\n\n{truncated}"


# ─────────────────────────────────────────────────────────────
#  Claude summarization
# ─────────────────────────────────────────────────────────────

def summarize(doc_type: str, text: str) -> str:
    if not ANTHROPIC_API_KEY:
        raise RuntimeError("ANTHROPIC_API_KEY לא מוגדר")

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    prompt = build_prompt(doc_type, text)

    message = client.messages.create(
        model="claude-opus-4-5",
        max_tokens=2048,
        messages=[{"role": "user", "content": prompt}],
    )
    return message.content[0].text


# ─────────────────────────────────────────────────────────────
#  Routes
# ─────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    file = request.files.get("file")
    if not file or file.filename == "":
        return jsonify({"error": "לא נבחר קובץ"}), 400

    filename = file.filename.lower()
    file_bytes = file.read()

    try:
        if filename.endswith(".pdf"):
            text = extract_text_pdf(file_bytes)
        elif filename.endswith(".pptx"):
            text = extract_text_pptx(file_bytes)
        elif filename.endswith(".docx"):
            text = extract_text_docx(file_bytes)
        elif filename.endswith(".txt"):
            text = extract_text_txt(file_bytes)
        else:
            return jsonify({"error": "סוג קובץ לא נתמך. קבצים נתמכים: PDF, PPTX, DOCX, TXT"}), 400
    except Exception as exc:
        log.error("Extraction error: %s", exc)
        return jsonify({"error": f"שגיאה בקריאת הקובץ: {exc}"}), 500

    if not text.strip():
        return jsonify({"error": "לא ניתן היה לחלץ טקסט מהקובץ"}), 400

    doc_type = detect_doc_type(text)

    doc_type_labels = {
        "verdict":     {"label": "פסק דין", "icon": "⚖️"},
        "legislation": {"label": "חקיקה",   "icon": "📜"},
        "study":       {"label": "חומר לימוד", "icon": "📚"},
    }

    try:
        summary = summarize(doc_type, text)
    except Exception as exc:
        log.error("Summarize error: %s", exc)
        return jsonify({"error": f"שגיאה בסיכום: {exc}"}), 500

    return jsonify({
        "doc_type":  doc_type,
        "label":     doc_type_labels[doc_type]["label"],
        "icon":      doc_type_labels[doc_type]["icon"],
        "summary":   summary,
        "filename":  file.filename,
    })


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
