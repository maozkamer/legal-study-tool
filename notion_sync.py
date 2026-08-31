"""
Optional Notion sync for the lecture-recorder feature.

Every call is wrapped so a Notion problem (missing token, network error,
rate limit, etc.) never breaks the core app flow — summarizing / exporting
a lecture must always work even if Notion is unreachable or unconfigured.

Setup (once, on Render / locally):
  1. Create an integration at https://www.notion.so/my-integrations
     and copy its "Internal Integration Secret".
  2. Share the "מעקב הרצאות - לימודי משפטים" database with that integration
     (··· menu on the database → Connections → add the integration).
  3. Set the env var NOTION_TOKEN to the secret.
     (NOTION_LECTURES_DB_ID already defaults to the right database —
     override it only if you point this at a different database.)

If NOTION_TOKEN is not set, every function below is a no-op.
"""

import io
import logging
import os

import requests

log = logging.getLogger(__name__)

NOTION_TOKEN   = os.environ.get("NOTION_TOKEN", "").strip()
NOTION_DB_ID   = os.environ.get("NOTION_LECTURES_DB_ID", "14648d0b706d495b91ab630ff81398db").strip()
NOTION_VERSION = "2022-06-28"
NOTION_API     = "https://api.notion.com/v1"

_COURSE_KEYWORDS = {
    "פלילי":  ["פלילי", "עונשין", "פלילים"],
    "עבודה":  ["עבודה", "דיני עבודה"],
    "חוזים":  ["חוזים", "חוזה", "דיני חוזים"],
}


def _headers(json_body=True):
    h = {
        "Authorization": f"Bearer {NOTION_TOKEN}",
        "Notion-Version": NOTION_VERSION,
    }
    if json_body:
        h["Content-Type"] = "application/json"
    return h


def _guess_course(subject: str) -> str:
    subject = subject or ""
    for course, keywords in _COURSE_KEYWORDS.items():
        if any(kw in subject for kw in keywords):
            return course
    return "אחר"


def create_lecture_page(lesson_name: str, subject: str, date_iso: str) -> str | None:
    """Create a row for this lecture right after it's summarized.
    Returns the new page id, or None if Notion sync is off/failed."""
    if not NOTION_TOKEN:
        return None

    body = {
        "parent": {"database_id": NOTION_DB_ID},
        "properties": {
            "שם ההרצאה": {"title": [{"text": {"content": lesson_name or "שיעור"}}]},
            "קורס":       {"select": {"name": _guess_course(subject)}},
            "נושאים":     {"rich_text": [{"text": {"content": (subject or "")[:2000]}}]},
            "תאריך":      {"date": {"start": date_iso}},
            "סטטוס":      {"status": {"name": "בתהליך"}},
        },
    }
    try:
        resp = requests.post(f"{NOTION_API}/pages", headers=_headers(), json=body, timeout=15)
        resp.raise_for_status()
        page_id = resp.json().get("id")
        log.info("Notion: created lecture page %s", page_id)
        return page_id
    except Exception as exc:
        log.warning("Notion: failed to create lecture page: %s", exc)
        return None


def finalize_lecture_page(page_id: str | None, docx_bytes: bytes, filename: str) -> bool:
    """Attach the exported Word file to the lecture's Notion row and mark it done.
    Safe to call with page_id=None (does nothing)."""
    if not NOTION_TOKEN or not page_id:
        return False

    try:
        # 1. Start a file upload
        resp = requests.post(
            f"{NOTION_API}/file_uploads",
            headers=_headers(),
            json={"filename": filename},
            timeout=15,
        )
        resp.raise_for_status()
        upload = resp.json()
        upload_id  = upload["id"]
        upload_url = upload.get("upload_url", f"{NOTION_API}/file_uploads/{upload_id}/send")

        # 2. Send the file bytes
        mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        resp = requests.post(
            upload_url,
            headers=_headers(json_body=False),
            files={"file": (filename, io.BytesIO(docx_bytes), mime)},
            timeout=60,
        )
        resp.raise_for_status()

        # 3. Attach the uploaded file to the page + mark status done
        resp = requests.patch(
            f"{NOTION_API}/pages/{page_id}",
            headers=_headers(),
            json={
                "properties": {
                    "קובץ Word": {
                        "files": [
                            {"type": "file_upload", "file_upload": {"id": upload_id}, "name": filename}
                        ]
                    },
                    "סטטוס": {"status": {"name": "בוצע"}},
                }
            },
            timeout=15,
        )
        resp.raise_for_status()
        log.info("Notion: attached %s to page %s", filename, page_id)
        return True
    except Exception as exc:
        log.warning("Notion: failed to finalize lecture page %s: %s", page_id, exc)
        return False
