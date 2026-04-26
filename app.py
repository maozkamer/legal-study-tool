import os
import io
import uuid
import json
import logging
from datetime import date, datetime
from flask import Flask, request, render_template, jsonify, send_file, send_from_directory
import PyPDF2
from pptx import Presentation
from docx import Document
from docx.oxml import OxmlElement
import anthropic

logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s")
log = logging.getLogger(__name__)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB

client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))

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

def _claude(prompt: str, max_tokens: int = 2048) -> str:
    response = client.messages.create(
        model="claude-opus-4-5",
        max_tokens=max_tokens,
        messages=[{"role": "user", "content": prompt}],
    )
    return response.content[0].text


# ─────────────────────────────────────────────────────────────
#  Long-transcript chunking helpers
# ─────────────────────────────────────────────────────────────

CHUNK_SIZE    = 40_000   # characters per chunk
MAX_CHUNKS    = 4        # max 4 chunks (~3 hours of recording)
_CHUNK_OVERLAP = 500     # character overlap between consecutive chunks

# Shared case-law detection instructions injected into every full-summary prompt
_CASE_LAW_INSTRUCTIONS = (
    "זיהוי פסקי דין:\n"
    '- חפש כל אזכור של שם תיק בפורמט: X נגד Y, X v Y, ע"א XXXX, בג"ץ XXXX, רע"א XXXX\n'
    "- גם אם המרצה אומר רק שם משפחה כמו 'פרשת כהן' או 'עניין לוי' — זה פסק דין\n"
    "- חלץ: שם התיק, העיקרון המשפטי שהמרצה הסביר, הרלוונטיות לנושא השיעור\n"
    "- אם לא הוזכרו פסקי דין — case_law=[]\n\n"
)

# Full JSON schema string reused in both direct and merge prompts
_JSON_SCHEMA = (
    "{\n"
    '  "subject": "שם הנושא",\n'
    '  "sections": [\n'
    '    {"level": 1, "heading": "כותרת ראשית", "content": "תוכן"},\n'
    '    {"level": 2, "heading": "כותרת משנה",  "content": "תוכן"}\n'
    "  ],\n"
    '  "concepts":  [{"term": "מושג", "definition": "הגדרה", "example": "דוגמה"}],\n'
    '  "case_law":  [{"name": "שם התיק", "principle": "עיקרון", "relevance": "רלוונטיות"}],\n'
    '  "statutes":  [{"law": "שם החוק", "section": "סעיף", "content": "תוכן"}],\n'
    '  "important_moments":  ["רגע חשוב — הוצא מסימון ⭐ בתמלול"],\n'
    '  "related_topics":     "נושאים קשורים",\n'
    '  "instructor_emphasis":["נקודה שהמרצה הדגיש"],\n'
    '  "key_points":         ["נקודה 1","נקודה 2","נקודה 3","נקודה 4","נקודה 5"]\n'
    "}"
)


def _chunk_text(text: str) -> list[str]:
    """Split text into overlapping chunks up to MAX_CHUNKS."""
    chunks, start = [], 0
    while start < len(text) and len(chunks) < MAX_CHUNKS:
        chunks.append(text[start : start + CHUNK_SIZE])
        start += CHUNK_SIZE - _CHUNK_OVERLAP
    return chunks


def _extract_chunk_summary(chunk: str, chunk_num: int, total: int, lesson_name: str) -> str:
    """Ask Claude to extract key information from a single transcript chunk."""
    prompt = (
        f'אתה עוזר לימודים. זהה מהקטע הבא (קטע {chunk_num} מתוך {total}) '
        f'משיעור בשם "{lesson_name}":\n\n'
        "- נושאים שנדונו\n"
        "- פסקי דין שהוזכרו (שם תיק + עיקרון)\n"
        "- חוקים וסעיפים שהוזכרו\n"
        "- הגדרות ומושגים\n"
        "- נקודות שהמרצה הדגיש\n\n"
        f"קטע:\n{chunk}"
    )
    return _claude(prompt, max_tokens=2048)


# ─────────────────────────────────────────────────────────────
#  Routes
# ─────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/manifest.json")
def manifest():
    return send_from_directory("static", "manifest.json")


@app.route("/sw.js")
def service_worker():
    return send_from_directory("static", "sw.js", mimetype="application/javascript")


@app.route("/icon-192.svg")
def icon_192():
    return send_from_directory("static", "icon-192.svg", mimetype="image/svg+xml")


@app.route("/icon-512.svg")
def icon_512():
    return send_from_directory("static", "icon-512.svg", mimetype="image/svg+xml")


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

    doc_type     = detect_type(text)
    summary_type = request.form.get("summary_type", "full")

    if summary_type == "short":
        prompt = (
            "אתה עוזר לימודים משפטי. סכם את המסמך הבא ב-5–7 נקודות עיקריות בלבד. "
            "קצר, ממוקד, ללא פירוט מיותר. השתמש בפורמט:\n\n"
            "⚡ **סיכום מקוצר**\n\n"
            "1. [נקודה ראשונה]\n"
            "2. [נקודה שנייה]\n"
            "...\n\n"
            "📌 **מסקנה עיקרית:** [משפט אחד]\n\n"
            "המסמך:\n" + text[:12000]
        )
    else:
        prompt = PROMPTS[doc_type] + text[:12000]

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


@app.route("/summarize-lecture", methods=["POST"])
def summarize_lecture():
    data         = request.json or {}
    lesson_name  = data.get("lesson_name", "שיעור")
    transcript   = data.get("transcript", "").strip()
    duration     = data.get("duration", "")
    summary_type = data.get("summary_type", "full")

    if not transcript:
        return jsonify({"error": "התמלול ריק — ודא שהמיקרופון פעל"}), 400

    today    = date.today().strftime("%d/%m/%Y")
    now_time = datetime.now().strftime("%H:%M")

    # ── Short summary (bullet points, no chunking) ─────────────
    if summary_type == "short":
        prompt = (
            f'אתה עוזר לימודים משפטי. קיבלת תמליל של שיעור בשם: "{lesson_name}"\n'
            f"תאריך: {today}, משך: {duration}\n\n"
            "סכם ב-5–7 נקודות עיקריות בלבד. קצר, ממוקד, ללא פירוט מיותר.\n"
            "השתמש בפורמט:\n\n"
            f"⚡ **סיכום מקוצר — {lesson_name}**\n\n"
            "1. [נקודה ראשונה]\n"
            "2. [נקודה שנייה]\n"
            "...\n\n"
            "📌 **מסקנה עיקרית:** [משפט אחד]\n\n"
            "תמליל השיעור:\n" + transcript[:CHUNK_SIZE]
        )
        try:
            raw = _claude(prompt, max_tokens=1024)
        except Exception as exc:
            log.error("Claude error: %s", exc)
            return jsonify({"error": f"שגיאה בסיכום: {exc}"}), 500
        return jsonify({"summary": raw, "subject": lesson_name, "structured": None})

    # ── Full summary ───────────────────────────────────────────
    try:
        if len(transcript) > CHUNK_SIZE:
            # ── Multi-chunk path: extract then merge ───────────
            chunks = _chunk_text(transcript)
            log.info(
                "Long transcript: %d chars split into %d chunks",
                len(transcript), len(chunks),
            )
            partials = []
            for i, chunk in enumerate(chunks, 1):
                partials.append(_extract_chunk_summary(chunk, i, len(chunks), lesson_name))

            merged_input = "\n\n---\n\n".join(
                f"סיכום קטע {i + 1}:\n{p}" for i, p in enumerate(partials)
            )
            prompt = (
                f'אתה עוזר לימודים משפטי. קיבלת סיכומי ביניים של שיעור בשם: "{lesson_name}"\n'
                f"תאריך: {today}, שעה: {now_time}, משך: {duration}\n\n"
                "אחד אותם לסיכום אחד מלא ומובנה. "
                "זהה את הנושא המשפטי (דיני עבודה / דיני עונשין / משפט מנהלי / "
                "משפט חוקתי / דיני חוזים / אחר).\n\n"
                + _CASE_LAW_INSTRUCTIONS
                + "אם אין פסקי דין — case_law=[]. אם אין סעיפי חוק — statutes=[].\n"
                "רגעים חשובים מסומנים ב-⭐ — חלץ אותם.\n\n"
                "החזר JSON בלבד — ללא markdown, ללא טקסט לפני או אחרי הסוגריים:\n"
                + _JSON_SCHEMA + "\n\n"
                "סיכומי הביניים:\n" + merged_input
            )
        else:
            # ── Single-chunk path ──────────────────────────────
            prompt = (
                f'אתה עוזר לימודים משפטי. קיבלת תמליל של שיעור בשם: "{lesson_name}"\n'
                f"תאריך: {today}, שעה: {now_time}, משך: {duration}\n\n"
                "זהה את הנושא המשפטי (דיני עבודה / דיני עונשין / משפט מנהלי / "
                "משפט חוקתי / דיני חוזים / אחר).\n\n"
                + _CASE_LAW_INSTRUCTIONS
                + "אם אין פסקי דין — case_law=[]. אם אין סעיפי חוק — statutes=[].\n"
                "רגעים חשובים מסומנים ב-⭐ בתמלול — חלץ אותם.\n\n"
                "החזר JSON בלבד — ללא markdown, ללא טקסט לפני או אחרי הסוגריים:\n"
                + _JSON_SCHEMA + "\n\n"
                "תמליל השיעור:\n" + transcript
            )

        raw        = _claude(prompt, max_tokens=4096)
        js         = raw[raw.find("{") : raw.rfind("}") + 1]
        js         = js.strip().lstrip("```json").lstrip("```").rstrip("```").strip()
        structured = json.loads(js)

    except (json.JSONDecodeError, ValueError):
        log.warning("JSON parse failed — returning plain summary")
        structured = None
    except Exception as exc:
        log.error("Claude error in summarize_lecture: %s", exc)
        return jsonify({"error": f"שגיאה בסיכום: {exc}"}), 500

    if structured is None:
        return jsonify({"summary": raw, "subject": "שיעור", "structured": None})

    subject = structured.get("subject", "שיעור")

    # Build display text
    lines = [
        f"🎓 **נושא:** {subject}",
        f"📅 **תאריך:** {today}   ⏱️ **משך:** {duration}\n",
    ]
    for sec in structured.get("sections", []):
        lines.append(f"\n**{sec.get('heading', '')}**")
        lines.append(sec.get("content", ""))
    if structured.get("concepts"):
        lines.append("\n**💡 מושגים מרכזיים:**")
        for c in structured["concepts"]:
            lines.append(f"• **{c.get('term','')}** — {c.get('definition','')}")
    if structured.get("important_moments"):
        lines.append("\n**⭐ רגעים חשובים:**")
        for m in structured["important_moments"]:
            lines.append(f"• {m}")
    if structured.get("instructor_emphasis"):
        lines.append("\n**📍 נקודות שהמרצה הדגיש:**")
        for e in structured["instructor_emphasis"]:
            lines.append(f"• {e}")
    if structured.get("key_points"):
        lines.append("\n**📌 5 נקודות עיקריות:**")
        for i, kp in enumerate(structured["key_points"][:5], 1):
            lines.append(f"{i}. {kp}")

    return jsonify({
        "summary":    "\n".join(lines),
        "subject":    subject,
        "structured": structured,
    })


@app.route("/export-lecture-docx", methods=["POST"])
def export_lecture_docx():
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn

    data        = request.json or {}
    lesson_name = data.get("lesson_name", "שיעור")
    dt_str      = data.get("date", "")
    duration    = data.get("duration", "")
    subject     = data.get("subject", "")
    structured  = data.get("structured") or {}
    if isinstance(structured, str):
        try:
            structured = json.loads(structured)
        except (json.JSONDecodeError, ValueError):
            structured = {}
    summary_txt = data.get("summary", "")

    doc = Document()

    def _rtl(paragraph):
        pPr = paragraph._p.get_or_add_pPr()
        bidi = OxmlElement("w:bidi")
        pPr.append(bidi)

    def _heading(text, level):
        h = doc.add_heading(text, level=level)
        _rtl(h)
        return h

    def _cell_write(cell, text, bold=False, white_text=False):
        """Write text into a cell with optional bold/white and RTL paragraph."""
        cell.text = text
        p = cell.paragraphs[0]
        for run in p.runs:
            if bold:
                run.bold = True
            if white_text:
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        _rtl(p)

    def _cell_bg(cell, hex_color):
        """Apply a solid background fill to a table cell."""
        tc   = cell._tc
        tcPr = tc.get_or_add_tcPr()
        shd  = OxmlElement("w:shd")
        shd.set(qn("w:val"),   "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"),  hex_color)
        tcPr.append(shd)

    def _tbl_rtl(tbl):
        """Mark the table as right-to-left so column 0 appears on the right."""
        tblPr = tbl._tbl.tblPr
        if tblPr is None:
            tblPr = OxmlElement("w:tblPr")
            tbl._tbl.insert(0, tblPr)
        tblPr.append(OxmlElement("w:bidiVisual"))

    # ── Cover page ──────────────────────────────────────────────
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(lesson_name)
    run.bold = True
    run.font.size = Pt(22)
    _rtl(p)

    for line in filter(None, [
        f"📅 תאריך: {dt_str}",
        f"⏱️ משך: {duration}",
        f"🎓 נושא: {subject}" if subject else "",
    ]):
        mp = doc.add_paragraph(line)
        mp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _rtl(mp)

    doc.add_page_break()

    # ── Body sections ────────────────────────────────────────────
    for sec in structured.get("sections", []):
        _heading(sec.get("heading", ""), level=min(sec.get("level", 1), 2))
        content = sec.get("content", "")
        if content:
            cp = doc.add_paragraph(content)
            _rtl(cp)

    # ── Important moments ────────────────────────────────────────
    moments = structured.get("important_moments", [])
    if moments:
        _heading("⭐ רגעים חשובים", level=1)
        for m in moments:
            mp = doc.add_paragraph()
            run = mp.add_run(m)
            run.bold = True
            run.font.color.rgb = RGBColor(0xB8, 0x86, 0x00)
            _rtl(mp)

    # ── Concepts table ───────────────────────────────────────────
    # RTL order: col 0 (right) = מושג, col 1 (middle) = הגדרה, col 2 (left) = דוגמה
    concepts = structured.get("concepts", [])
    if concepts:
        _heading("💡 טבלת מושגים", level=1)
        tbl = doc.add_table(rows=1, cols=3)
        tbl.style = "Table Grid"
        _tbl_rtl(tbl)
        hdr = tbl.rows[0].cells
        for cell, title in zip(hdr, ["מושג", "הגדרה", "דוגמה"]):
            _cell_write(cell, title, bold=True, white_text=True)
            _cell_bg(cell, "1E3A5F")
        for c in concepts:
            row = tbl.add_row().cells
            _cell_write(row[0], c.get("term", ""),       bold=True)
            _cell_bg(row[0], "D6E4F7")
            _cell_write(row[1], c.get("definition", ""))
            _cell_bg(row[1], "F5F9FF")
            _cell_write(row[2], c.get("example", ""))
            _cell_bg(row[2], "FEF9EC")
        doc.add_paragraph("")

    # ── Case-law table ───────────────────────────────────────────
    # RTL order: col 0 (right) = שם התיק, col 1 = עיקרון, col 2 (left) = רלוונטיות
    case_law = structured.get("case_law", [])
    if case_law:
        _heading("⚖️ טבלת פסיקה", level=1)
        tbl = doc.add_table(rows=1, cols=3)
        tbl.style = "Table Grid"
        _tbl_rtl(tbl)
        hdr = tbl.rows[0].cells
        for cell, title in zip(hdr, ["שם התיק", "עיקרון", "רלוונטיות"]):
            _cell_write(cell, title, bold=True, white_text=True)
            _cell_bg(cell, "1E3A5F")
        for c in case_law:
            row = tbl.add_row().cells
            _cell_write(row[0], c.get("name", ""),      bold=True)
            _cell_bg(row[0], "EAF0FA")
            _cell_write(row[1], c.get("principle", ""))
            _cell_bg(row[1], "F2F6FC")
            _cell_write(row[2], c.get("relevance", ""))
            _cell_bg(row[2], "F8FAFE")
        doc.add_paragraph("")

    # ── Statutes table ───────────────────────────────────────────
    # RTL order: col 0 (right) = שם החוק, col 1 = סעיף, col 2 (left) = תוכן
    statutes = structured.get("statutes", [])
    if statutes:
        _heading("📜 סעיפי חוק", level=1)
        tbl = doc.add_table(rows=1, cols=3)
        tbl.style = "Table Grid"
        _tbl_rtl(tbl)
        hdr = tbl.rows[0].cells
        for cell, title in zip(hdr, ["שם החוק", "סעיף", "תוכן"]):
            _cell_write(cell, title, bold=True, white_text=True)
            _cell_bg(cell, "4A235A")
        for s in statutes:
            row = tbl.add_row().cells
            _cell_write(row[0], s.get("law", ""),     bold=True)
            _cell_bg(row[0], "F3E8FA")
            _cell_write(row[1], s.get("section", ""))
            _cell_bg(row[1], "F9F2FD")
            _cell_write(row[2], s.get("content", ""))
            _cell_bg(row[2], "FDF6FF")
        doc.add_paragraph("")

    # ── Related topics ───────────────────────────────────────────
    related = structured.get("related_topics", "")
    if related:
        _heading("🔗 קשרים לנושאים אחרים", level=1)
        rp = doc.add_paragraph(related)
        _rtl(rp)

    # ── Instructor emphasis ──────────────────────────────────────
    emphasis = structured.get("instructor_emphasis", [])
    if emphasis:
        _heading("📍 נקודות לבדיקה", level=1)
        for e in emphasis:
            ep = doc.add_paragraph(e, style="List Bullet")
            _rtl(ep)

    # ── Key points summary ───────────────────────────────────────
    key_points = structured.get("key_points", [])
    if key_points:
        _heading("📌 5 נקודות עיקריות מהשיעור", level=1)
        for i, kp in enumerate(key_points[:5], 1):
            kpp = doc.add_paragraph(f"{i}. {kp}")
            _rtl(kpp)

    # ── Fallback: plain summary ──────────────────────────────────
    if not structured.get("sections") and summary_txt:
        _heading("סיכום", level=1)
        for line in summary_txt.split("\n"):
            clean = line.strip("*").strip()
            if clean:
                lp = doc.add_paragraph(clean)
                _rtl(lp)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)

    safe = "".join(c for c in lesson_name if c.isalnum() or c in " .-_()") or "שיעור"
    return send_file(
        buf, as_attachment=True,
        download_name=f"{safe}.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@app.route("/concept-map", methods=["POST"])
def concept_map():
    token = (request.json or {}).get("token", "")
    text  = _TEXT_CACHE.get(token)
    if not text:
        return jsonify({"error": "הסשן פג תוקף — אנא העלה את הקובץ מחדש"}), 400

    prompt = (
        "אתה עוזר לימודים משפטי. נתח את המסמך הבא וזהה:\n"
        "- פסקי דין שהוזכרו (type: verdict)\n"
        "- חוקים וסעיפים (type: law)\n"
        "- עקרונות משפטיים (type: principle)\n"
        "- מושגים משפטיים חשובים (type: concept)\n"
        "- קשרים בין האלמנטים הללו\n\n"
        "החזר JSON בלבד — ללא markdown, ללא טקסט לפני או אחרי הסוגריים:\n"
        '{\n'
        '  "nodes": [\n'
        '    {"id":"1","label":"שם קצר (עד 35 תווים)","type":"verdict|law|principle|concept"}\n'
        '  ],\n'
        '  "edges": [\n'
        '    {"from":"1","to":"2","label":"קשר קצר (עד 20 תווים)"}\n'
        '  ]\n'
        '}\n\n'
        "כלול 6–15 צמתים ו-5–15 קשרים. אם אין פסקי דין — השמט. label לכל צומת חובה.\n\n"
        "המסמך:\n" + text[:9000]
    )

    try:
        raw  = _claude(prompt, max_tokens=2048)
        js   = raw[raw.find("{") : raw.rfind("}") + 1]
        data = json.loads(js)
        # Basic validation
        if not data.get("nodes"):
            return jsonify({"error": "לא זוהו אלמנטים משפטיים במסמך"}), 400
    except (json.JSONDecodeError, ValueError) as exc:
        log.error("Concept-map JSON error: %s", exc)
        return jsonify({"error": "שגיאה בניתוח המסמך — נסה שוב"}), 500

    return jsonify(data)


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port, debug=False)
