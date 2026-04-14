import os
import io
import uuid
import logging
from datetime import date
from flask import Flask, request, render_template, jsonify, send_file
import PyPDF2
from pptx import Presentation
from docx import Document
from docx.oxml import OxmlElement
import anthropic
from openai import OpenAI

logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s")
log = logging.getLogger(__name__)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB

client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))
openai_client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))

# In-memory text cache: token → extracted text (up to 50 entries)
_TEXT_CACHE: dict[str, str] = {}
_MAX_CACHE  = 50


@app.errorhandler(413)
def too_large(e):
    return jsonify({"error": "הקובץ גדול מדי — מקסימום 50MB"}), 413


# ─────────────────────────────────────────────────────────────
#  Text extraction
# ─────────────────────────────────────────────────────────────

def extract_pdf(data: bytes) -> str:
    reader = PyPDF2.PdfReader(io.BytesIO(data))
    return "\n".join(p.extract_text() or "" for p in reader.pages)


def extract_pptx(data: bytes) -> str:
    prs = Presentation(io.BytesIO(data))
    lines = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    line = " ".join(r.text for r in para.runs).strip()
                    if line:
                        lines.append(line)
    return "\n".join(lines)


def extract_docx(data: bytes) -> str:
    doc = Document(io.BytesIO(data))
    return "\n".join(p.text.strip() for p in doc.paragraphs if p.text.strip())


def extract_txt(data: bytes) -> str:
    for enc in ("utf-8", "windows-1255", "iso-8859-8", "latin-1"):
        try:
            return data.decode(enc)
        except (UnicodeDecodeError, LookupError):
            continue
    return data.decode("utf-8", errors="replace")


# ─────────────────────────────────────────────────────────────
#  Document type detection
# ─────────────────────────────────────────────────────────────

def detect_type(text: str) -> str:
    sample = text[:3000]
    verdict_hits     = sum(1 for w in ["בית משפט", "נגד", "השופט", "הנאשם", "המשיב", "פסק דין"] if w in sample)
    legislation_hits = sum(1 for w in ["חוק", "סעיף", "תקנות", "כנסת", "ספר החוקים"] if w in sample)
    if verdict_hits >= 2:
        return "verdict"
    if legislation_hits >= 2:
        return "legislation"
    return "study"


# ─────────────────────────────────────────────────────────────
#  Prompts
# ─────────────────────────────────────────────────────────────

PROMPTS = {
    "verdict": """\
אתה עוזר לימודים משפטי. סכם את פסק הדין הבא בעברית בפורמט הזה בדיוק:

📋 **פרטי התיק:** [בית משפט, שנה, שמות הצדדים]
⚖️ **השאלה המשפטית:** [מה נדון]
📖 **עובדות המקרה:** [עובדות עיקריות]
💬 **טענות הצדדים:** [תובע מול נתבע]
🔨 **פסיקה:** [מה פסק בית המשפט ומדוע]
📌 **העיקרון המשפטי לבחינה:** [ההלכה החשובה]

המסמך:
""",
    "legislation": """\
אתה עוזר לימודים משפטי. סכם את החקיקה הבאה בעברית בפורמט הזה בדיוק:

📋 **שם החוק ושנה:** [שם מלא + שנת חקיקה]
🎯 **מטרת החוק:** [מה החוק בא להסדיר]
📖 **סעיפים מרכזיים:** [הסעיפים החשובים]
⚠️ **חריגים חשובים:** [חריגים, הגנות, תנאים]
📌 **מה חשוב לדעת לבחינה:** [נקודות לשינון]

המסמך:
""",
    "study": """\
אתה עוזר לימודים משפטי. סכם את חומר הלימוד הבא בעברית בפורמט הזה בדיוק:

🎯 **נושא המסמך:** [נושא ראשי]
📖 **רעיונות מרכזיים:**
1. [רעיון ראשון]
2. [רעיון שני]
3. [המשך לפי הצורך]
💡 **מושגים חשובים:** [הגדרות ומושגי מפתח]
❓ **שאלות אפשריות לבחינה:**
- [שאלה 1]
- [שאלה 2]
- [שאלה 3]
📌 **סיכום:** [תמצית בשורה אחת]

המסמך:
""",
}

TYPE_LABELS = {
    "verdict":     ("⚖️", "פסק דין"),
    "legislation": ("📜", "חקיקה"),
    "study":       ("📚", "חומר לימוד"),
}


# ─────────────────────────────────────────────────────────────
#  Claude helper
# ─────────────────────────────────────────────────────────────

def _claude(prompt: str) -> str:
    response = client.messages.create(
        model="claude-opus-4-5",
        max_tokens=2048,
        messages=[{"role": "user", "content": prompt}],
    )
    return response.content[0].text


# ─────────────────────────────────────────────────────────────
#  Routes
# ─────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    file = request.files.get("file")
    if not file or not file.filename:
        return jsonify({"error": "לא נבחר קובץ"}), 400

    name = file.filename.lower()
    data = file.read()

    try:
        if name.endswith(".pdf"):
            text = extract_pdf(data)
        elif name.endswith(".pptx"):
            text = extract_pptx(data)
        elif name.endswith(".docx"):
            text = extract_docx(data)
        elif name.endswith(".txt"):
            text = extract_txt(data)
        else:
            return jsonify({"error": "סוג קובץ לא נתמך. קבצים נתמכים: PDF, PPTX, DOCX, TXT"}), 400
    except Exception as exc:
        log.error("Extraction failed: %s", exc)
        return jsonify({"error": f"שגיאה בקריאת הקובץ: {exc}"}), 500

    if not text.strip():
        return jsonify({"error": "לא ניתן לחלץ טקסט מהקובץ"}), 400

    doc_type = detect_type(text)
    prompt   = PROMPTS[doc_type] + text[:12000]

    try:
        summary = _claude(prompt)
    except Exception as exc:
        log.error("Claude error: %s", exc)
        return jsonify({"error": f"שגיאה בסיכום: {exc}"}), 500

    token = str(uuid.uuid4())
    if len(_TEXT_CACHE) >= _MAX_CACHE:
        oldest = next(iter(_TEXT_CACHE))
        del _TEXT_CACHE[oldest]
    _TEXT_CACHE[token] = text[:12000]

    icon, label = TYPE_LABELS[doc_type]
    return jsonify({
        "icon": icon, "label": label,
        "doc_type": doc_type,
        "summary": summary, "token": token,
    })


@app.route("/questions", methods=["POST"])
def questions():
    token = (request.json or {}).get("token", "")
    text  = _TEXT_CACHE.get(token)
    if not text:
        return jsonify({"error": "הסשן פג תוקף — אנא העלה את הקובץ מחדש"}), 400

    prompt = (
        "אתה עוזר לימודים משפטי. צור 10 שאלות לבחינה בעברית מהחומר הבא.\n"
        "5 שאלות אמריקאיות (עם 4 תשובות, אחת נכונה) ו-5 שאלות פתוחות.\n"
        "השתמש בפורמט הזה בדיוק:\n\n"
        "📝 שאלות לבחינה\n\n"
        "שאלות אמריקאיות:\n"
        "1. [שאלה]?\n"
        "א. [תשובה]\n"
        "ב. [תשובה]\n"
        "ג. [תשובה]\n"
        "ד. [תשובה]\n"
        "✅ תשובה נכונה: א\n\n"
        "(חזור על הפורמט לשאלות 2-5)\n\n"
        "שאלות פתוחות:\n"
        "1. [שאלה]?\n"
        "💡 תשובה מנחה: [תשובה קצרה]\n\n"
        "(חזור על הפורמט לשאלות 2-5)\n\n"
        "---\n\n" + text
    )
    try:
        return jsonify({"result": _claude(prompt)})
    except Exception as exc:
        log.error("Questions error: %s", exc)
        return jsonify({"error": f"שגיאה ביצירת שאלות: {exc}"}), 500


@app.route("/flashcards", methods=["POST"])
def flashcards():
    token = (request.json or {}).get("token", "")
    text  = _TEXT_CACHE.get(token)
    if not text:
        return jsonify({"error": "הסשן פג תוקף — אנא העלה את הקובץ מחדש"}), 400

    prompt = (
        "אתה עוזר לימודים משפטי. צור 10 כרטיסיות חזרה בעברית מהחומר הבא.\n"
        "כל כרטיסייה: שאלה קצרה מצד אחד, תשובה קצרה מצד שני.\n"
        "השתמש בפורמט הזה בדיוק:\n\n"
        "🃏 כרטיסיות חזרה\n\n"
        "כרטיסייה 1:\n"
        "❓ שאלה: [שאלה קצרה]\n"
        "💡 תשובה: [תשובה קצרה]\n\n"
        "כרטיסייה 2:\n"
        "❓ שאלה: [שאלה קצרה]\n"
        "💡 תשובה: [תשובה קצרה]\n\n"
        "(המשך עד כרטיסייה 10)\n\n"
        "---\n\n" + text
    )
    try:
        return jsonify({"result": _claude(prompt)})
    except Exception as exc:
        log.error("Flashcards error: %s", exc)
        return jsonify({"error": f"שגיאה ביצירת כרטיסיות: {exc}"}), 500


@app.route("/export-docx", methods=["POST"])
def export_docx():
    data     = request.json or {}
    summary  = data.get("summary", "")
    filename = data.get("filename", "סיכום")
    notes    = data.get("notes", "")

    doc = Document()

    # RTL helper
    def _rtl(paragraph):
        pPr = paragraph._p.get_or_add_pPr()
        bidi = OxmlElement("w:bidi")
        pPr.append(bidi)

    title_para = doc.add_heading(filename, level=0)
    _rtl(title_para)

    for line in summary.split("\n"):
        clean = line.strip("*").strip()
        if not clean:
            continue
        p = doc.add_paragraph(clean)
        _rtl(p)

    if notes.strip():
        doc.add_paragraph("")
        h = doc.add_heading("הערות אישיות", level=2)
        _rtl(h)
        for line in notes.split("\n"):
            if line.strip():
                p = doc.add_paragraph(line.strip())
                _rtl(p)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)

    safe_name = "".join(c for c in filename if c.isalnum() or c in " .-_()") or "summary"
    return send_file(
        buf,
        as_attachment=True,
        download_name=f"{safe_name}.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@app.route("/transcribe", methods=["POST"])
def transcribe():
    audio_file = request.files.get("audio")
    duration   = request.form.get("duration", "")
    mime_type  = request.form.get("content_type", "audio/webm")

    if not audio_file:
        return jsonify({"error": "לא התקבל קובץ אודיו"}), 400

    data = audio_file.read()
    if not data:
        return jsonify({"error": "קובץ האודיו ריק"}), 400

    # Derive extension from MIME type for Whisper filename hint
    if "ogg" in mime_type:
        ext = "ogg"
    elif "mp4" in mime_type:
        ext = "mp4"
    else:
        ext = "webm"

    try:
        transcript_resp = openai_client.audio.transcriptions.create(
            model="whisper-1",
            file=(f"recording.{ext}", data, mime_type),
            language="he",
        )
        transcript_text = transcript_resp.text
    except Exception as exc:
        log.error("Whisper error: %s", exc)
        return jsonify({"error": f"שגיאה בתמלול: {exc}"}), 500

    if not transcript_text.strip():
        return jsonify({"error": "לא ניתן לתמלל את ההקלטה — נסה שוב"}), 400

    today = date.today().strftime("%d/%m/%Y")
    duration_line = f"⏱️ משך: {duration}" if duration else ""

    prompt = (
        "אתה עוזר לימודים. סכם את ההרצאה שהוקלטה בעברית בפורמט הזה בדיוק:\n\n"
        f"🎙️ סיכום הרצאה\n"
        f"📅 תאריך: {today}\n"
        f"{duration_line}\n\n"
        "📖 נושאים מרכזיים:\n"
        "1. [נושא]\n"
        "2. [נושא]\n"
        "3. [המשך לפי הצורך]\n\n"
        "💡 רעיונות חשובים:\n"
        "- [רעיון]\n\n"
        "📌 לזכור לבחינה:\n"
        "- [נקודה]\n\n"
        "תמליל ההרצאה:\n" + transcript_text[:8000]
    )

    try:
        summary = _claude(prompt)
    except Exception as exc:
        log.error("Claude error in transcribe: %s", exc)
        return jsonify({"error": f"שגיאה בסיכום: {exc}"}), 500

    return jsonify({"summary": summary, "transcript": transcript_text, "duration": duration})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
