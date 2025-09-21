# utils/generation_ivq.py
# -----------------------------------------------------------
# Utilities to call OpenAI, build structured questions, and render HTML/DOCX

import io
import re
from typing import List, Dict, Any

from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.shared import qn

try:
    from openai import OpenAI
except Exception:  # pragma: no cover
    OpenAI = None

HTML_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Source+Sans+Pro:wght@400;600;700&display=swap');
:root { --fg:#111; --muted:#444; }
body { font-family: 'Source Sans Pro', system-ui, -apple-system, Segoe UI, Roboto, sans-serif; }
.ivq { max-width: 860px; margin: 0 auto; color: var(--fg); }
.hdr { text-align:center; margin: 8px 0 16px; }
.hdr h1 { font-size: 28px; margin: 0; font-weight:700; }
.meta { text-align:center; font-size:14px; color: var(--muted); margin-bottom: 18px; }
.section h2 { font-size: 18px; margin: 24px 0 8px; font-weight: 700; }
.house ul { margin: 6px 0 0 18px; }
.q { border-top: 1px solid #ddd; padding-top: 14px; margin-top: 14px; }
.q h3 { font-size: 16px; margin: 0 0 6px; font-weight: 700; }
.kv { margin: 6px 0; }
.kv .k { font-weight: 600; }
.notes { border: 1px dashed #bbb; height: 90px; margin: 10px 0 6px; }
.footer { margin-top: 30px; font-size: 12px; color: #777; text-align:center; }
</style>
"""

HOUSEKEEPING_BULLETS = [
    "Ensure confidentiality of candidate responses.",
    "Maintain a professional and respectful tone throughout.",
    "Allocate approximately the scheduled duration for the interview.",
    "Use follow‑up questions to probe deeper into responses.",
    "Adhere to employment law and equality guidelines for the jurisdiction.",
]

SYSTEM_PROMPT = (
    "You are a senior talent acquisition partner and interview architect. "
    "Create sharp, behaviour‑based questions tailored to the role. "
    "Return JSON with an array 'questions', where each item has: 'question', 'intent', 'what_good_looks_like', 'follow_ups' (array). "
    "Questions must be clear, concise, and end with a question mark."
)

USER_PROMPT_TMPL = (
    "Role: {role_title}
"
    "Interview type: {interview_type}
"
    "Duration: {duration}
"
    "Jurisdiction: {jurisdiction}
"
    "Brief / JD highlights:
{brief}

"
    "Create {n} core questions."
)


def _ensure_question_mark(text: str) -> str:
    text = text.strip()
    if not text.endswith("?"):
        text = re.sub(r"[\.;:!]+$", "", text) + "?"
    return text


def call_openai_json(model: str, system: str, user: str) -> Dict[str, Any]:
    client = OpenAI()
    resp = client.chat.completions.create(
        model=model,
        temperature=0.4,
        response_format={"type": "json_object"},
        messages=[
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ],
    )
    import json
    content = resp.choices[0].message.content
    return json.loads(content)


def build_ivq_pack(role_title: str, interview_type: str, duration: str, brief: str, jurisdiction: str,
                   model: str, num_questions: int, include_followups: bool, include_wgll: bool,
                   include_housekeeping: bool) -> Dict[str, Any]:

    user_prompt = USER_PROMPT_TMPL.format(
        role_title=role_title or "Role",
        interview_type=interview_type,
        duration=duration or "60 mins",
        jurisdiction=jurisdiction or "",
        brief=brief or "",
        n=num_questions,
    )

    raw = call_openai_json(model=model, system=SYSTEM_PROMPT, user=user_prompt)

    # Normalise
    questions: List[Dict[str, Any]] = []
    for item in raw.get("questions", [])[:num_questions]:
        q = _ensure_question_mark(str(item.get("question", "")))
        intent = str(item.get("intent", "")).strip()
        wgll = str(item.get("what_good_looks_like", "")).strip()
        followups = [_ensure_question_mark(str(f).strip()) for f in item.get("follow_ups", [])]
        questions.append({
            "question": q,
            "intent": intent,
            "what_good_looks_like": wgll,
            "follow_ups": followups,
        })

    # HTML
    parts = [HTML_CSS, '<div class="ivq">']
    parts.append('<div class="hdr"><h1>{}</h1></div>'.format(f"{role_title} — Competency Pack"))
    parts.append('<div class="meta">Interview type: {} · Duration: {}</div>'.format(interview_type, duration))

    if include_housekeeping:
        parts.append('<div class="section house"><h2>Housekeeping</h2><ul>')
        for b in HOUSEKEEPING_BULLETS:
            parts.append(f"<li>{b}</li>")
        parts.append("</ul></div>")

    parts.append('<div class="section"><h2>Core Questions</h2></div>')
    for idx, it in enumerate(questions, start=1):
        parts.append('<div class="q">')
        parts.append(f"<h3>Question {idx}: {it['question']}</h3>")
        parts.append(f"<div class='kv'><span class='k'>Intent:</span> {it['intent']}</div>")
        if include_wgll and it['what_good_looks_like']:
            parts.append(f"<div class='kv'><span class='k'>What good looks like:</span> {it['what_good_looks_like']}</div>")
        if include_followups and it['follow_ups']:
            fus = " ".join([f"• {fu}" for fu in it['follow_ups']])
            parts.append(f"<div class='kv'><span class='k'>Follow‑ups:</span> {fus}</div>")
        parts.append('<div class="notes"></div>')
        parts.append('</div>')

    parts.append('<div class="footer">Generated by PowerDash HR – IVQ Generator</div>')
    parts.append('</div>')

    html = "
".join(parts)

    return {"html": html, "questions": questions, "meta": {
        "role_title": role_title,
        "interview_type": interview_type,
        "duration": duration,
        "housekeeping": include_housekeeping,
    }}


# ---- DOCX RENDERING ----

def _set_run_font(run, size=11, bold=False):
    run.font.name = "Source Sans Pro"
    run._element.rPr.rFonts.set(qn('w:eastAsia'), "Source Sans Pro")
    run.font.size = Pt(size)
    run.bold = bold


def _add_heading(doc: Document, text: str, level: int = 1):
    p = doc.add_paragraph()
    run = p.add_run(text)
    _set_run_font(run, size=20 if level == 1 else 14, bold=True)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if level == 1 else WD_ALIGN_PARAGRAPH.LEFT


def _add_spacer(doc: Document, pts: int = 6):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(pts)


def _add_dotted_notes(doc: Document, height_lines: int = 6):
    for _ in range(height_lines):
        p = doc.add_paragraph("." * 120)
        for r in p.runs:
            _set_run_font(r, size=10)


def _add_logo_header(doc: Document, logo_bytes: bytes | None):
    if not logo_bytes:
        return
    sec = doc.sections[0]
    header = sec.header
    paragraph = header.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = paragraph.add_run()
    run.add_picture(io.BytesIO(logo_bytes), width=Inches(1.1))


def _add_footer_logo(doc: Document, logo_bytes: bytes | None):
    if not logo_bytes:
        return
    sec = doc.sections[0]
    footer = sec.footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run()
    try:
        r.add_picture(io.BytesIO(logo_bytes), width=Inches(0.7))
    except Exception:
        r.add_text("PowerDash HR")
    _set_run_font(r, size=9)


def render_docx(pack: Dict[str, Any], role_title: str, interview_type: str, duration: str,
                logo_bytes: bytes | None, footer_logo_bytes: bytes | None) -> bytes:
    doc = Document()

    # Margins
    for sec in doc.sections:
        sec.top_margin = Inches(0.6)
        sec.bottom_margin = Inches(0.6)
        sec.left_margin = Inches(0.7)
        sec.right_margin = Inches(0.7)

    _add_logo_header(doc, logo_bytes)

    _add_heading(doc, f"{role_title} — Competency Pack", level=1)
    pmeta = doc.add_paragraph()
    run = pmeta.add_run(f"Interview type: {interview_type} · Duration: {duration}")
    _set_run_font(run, size=10)

    if pack["meta"].get("housekeeping"):
        _add_spacer(doc, 4)
        _add_heading(doc, "Housekeeping", level=2)
        for b in HOUSEKEEPING_BULLETS:
            li = doc.add_paragraph(style=None)
            run = li.add_run(f"• {b}")
            _set_run_font(run, size=11)

    _add_spacer(doc, 6)
    _add_heading(doc, "Core Questions", level=2)

    for idx, it in enumerate(pack["questions"], start=1):
        p = doc.add_paragraph()
        r = p.add_run(f"Question {idx}: {it['question']}")
        _set_run_font(r, size=12, bold=True)

        p = doc.add_paragraph()
        r = p.add_run(f"Intent: {it['intent']}")
        _set_run_font(r, size=11)

        if it.get('what_good_looks_like'):
            p = doc.add_paragraph()
            r = p.add_run(f"What good looks like: {it['what_good_looks_like']}")
            _set_run_font(r, size=11)

        if it.get('follow_ups'):
            p = doc.add_paragraph()
            r = p.add_run("Follow‑ups:")
            _set_run_font(r, size=11, bold=True)
            for fu in it['follow_ups']:
                p = doc.add_paragraph()
                r = p.add_run(f"• {fu}")
                _set_run_font(r, size=11)

        _add_dotted_notes(doc, height_lines=6)
        _add_spacer(doc, 6)

    _add_footer_logo(doc, footer_logo_bytes)

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out.read()



# utils/generation_iqt.py
# (Alias module for apps that still import generation_iqt.py)
# Same contents as generation_ivq.py, kept to avoid import errors.

import io
import re
from typing import List, Dict, Any

from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.shared import qn

try:
    from openai import OpenAI
except Exception:  # pragma: no cover
    OpenAI = None

HTML_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Source+Sans+Pro:wght@400;600;700&display=swap');
:root { --fg:#111; --muted:#444; }
body { font-family: 'Source Sans Pro', system-ui, -apple-system, Segoe UI, Roboto, sans-serif; }
.ivq { max-width: 860px; margin: 0 auto; color: var(--fg); }
.hdr { text-align:center; margin: 8px 0 16px; }
.hdr h1 { font-size: 28px; margin: 0; font-weight:700; }
.meta { text-align:center; font-size:14px; color: var(--muted); margin-bottom: 18px; }
.section h2 { font-size: 18px; margin: 24px 0 8px; font-weight: 700; }
.house ul { margin: 6px 0 0 18px; }
.q { border-top: 1px solid #ddd; padding-top: 14px; margin-top: 14px; }
.q h3 { font-size: 16px; margin: 0 0 6px; font-weight: 700; }
.kv { margin: 6px 0; }
.kv .k { font-weight: 600; }
.notes { border: 1px dashed #bbb; height: 90px; margin: 10px 0 6px; }
.footer { margin-top: 30px; font-size: 12px; color: #777; text-align:center; }
</style>
"""

HOUSEKEEPING_BULLETS = [
    "Ensure confidentiality of candidate responses.",
    "Maintain a professional and respectful tone throughout.",
    "Allocate approximately the scheduled duration for the interview.",
    "Use follow‑up questions to probe deeper into responses.",
    "Adhere to employment law and equality guidelines for the jurisdiction.",
]

SYSTEM_PROMPT = (
    "You are a senior talent acquisition partner and interview architect. "
    "Create sharp, behaviour‑based questions tailored to the role. "
    "Return JSON with an array 'questions', where each item has: 'question', 'intent', 'what_good_looks_like', 'follow_ups' (array). "
    "Questions must be clear, concise, and end with a question mark."
)

USER_PROMPT_TMPL = (
    "Role: {role_title}
"
    "Interview type: {interview_type}
"
    "Duration: {duration}
"
    "Jurisdiction: {jurisdiction}
"
    "Brief / JD highlights:
{brief}

"
    "Create {n} core questions."
)


def _ensure_question_mark(text: str) -> str:
    text = text.strip()
    if not text.endswith("?"):
        text = re.sub(r"[\.;:!]+$", "", text) + "?"
    return text


def call_openai_json(model: str, system: str, user: str) -> Dict[str, Any]:
    client = OpenAI()
    resp = client.chat.completions.create(
        model=model,
        temperature=0.4,
        response_format={"type": "json_object"},
        messages=[
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ],
    )
    import json
    content = resp.choices[0].message.content
    return json.loads(content)


def build_ivq_pack(role_title: str, interview_type: str, duration: str, brief: str, jurisdiction: str,
                   model: str, num_questions: int, include_followups: bool, include_wgll: bool,
                   include_housekeeping: bool) -> Dict[str, Any]:

    user_prompt = USER_PROMPT_TMPL.format(
        role_title=role_title or "Role",
        interview_type=interview_type,
        duration=duration or "60 mins",
        jurisdiction=jurisdiction or "",
        brief=brief or "",
        n=num_questions,
    )

    raw = call_openai_json(model=model, system=SYSTEM_PROMPT, user=user_prompt)

    questions: List[Dict[str, Any]] = []
    for item in raw.get("questions", [])[:num_questions]:
        q = _ensure_question_mark(str(item.get("question", "")))
        intent = str(item.get("intent", "")).strip()
        wgll = str(item.get("what_good_looks_like", "")).strip()
        followups = [_ensure_question_mark(str(f).strip()) for f in item.get("follow_ups", [])]
        questions.append({
            "question": q,
            "intent": intent,
            "what_good_looks_like": wgll,
            "follow_ups": followups,
        })

    parts = [HTML_CSS, '<div class="ivq">']
    parts.append('<div class="hdr"><h1>{}</h1></div>'.format(f"{role_title} — Competency Pack"))
    parts.append('<div class="meta">Interview type: {} · Duration: {}</div>'.format(interview_type, duration))

    if include_housekeeping:
        parts.append('<div class="section house"><h2>Housekeeping</h2><ul>')
        for b in HOUSEKEEPING_BULLETS:
            parts.append(f"<li>{b}</li>")
        parts.append("</ul></div>")

    parts.append('<div class="section"><h2>Core Questions</h2></div>')
    for idx, it in enumerate(questions, start=1):
        parts.append('<div class="q">')
        parts.append(f"<h3>Question {idx}: {it['question']}</h3>")
        parts.append(f"<div class='kv'><span class='k'>Intent:</span> {it['intent']}</div>")
        if include_wgll and it['what_good_looks_like']:
            parts.append(f"<div class='kv'><span class='k'>What good looks like:</span> {it['what_good_looks_like']}</div>")
        if include_followups and it['follow_ups']:
            fus = " ".join([f"• {fu}" for fu in it['follow_ups']])
            parts.append(f"<div class='kv'><span class='k'>Follow‑ups:</span> {fus}</div>")
        parts.append('<div class="notes"></div>')
        parts.append('</div>')

    parts.append('<div class="footer">Generated by PowerDash HR – IVQ Generator</div>')
    parts.append('</div>')

    html = "
".join(parts)

    return {"html": html, "questions": questions, "meta": {
        "role_title": role_title,
        "interview_type": interview_type,
        "duration": duration,
        "housekeeping": include_housekeeping,
    }}


def _set_run_font(run, size=11, bold=False):
    run.font.name = "Source Sans Pro"
    run._element.rPr.rFonts.set(qn('w:eastAsia'), "Source Sans Pro")
    run.font.size = Pt(size)
    run.bold = bold


def _add_heading(doc: Document, text: str, level: int = 1):
    p = doc.add_paragraph()
    run = p.add_run(text)
    _set_run_font(run, size=20 if level == 1 else 14, bold=True)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if level == 1 else WD_ALIGN_PARAGRAPH.LEFT


def _add_spacer(doc: Document, pts: int = 6):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(pts)


def _add_dotted_notes(doc: Document, height_lines: int = 6):
    for _ in range(height_lines):
        p = doc.add_paragraph("." * 120)
        for r in p.runs:
            _set_run_font(r, size=10)


def _add_logo_header(doc: Document, logo_bytes: bytes | None):
    if not logo_bytes:
        return
    sec = doc.sections[0]
    header = sec.header
    paragraph = header.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = paragraph.add_run()
    run.add_picture(io.BytesIO(logo_bytes), width=Inches(1.1))


def _add_footer_logo(doc: Document, logo_bytes: bytes | None):
    if not logo_bytes:
        return
    sec = doc.sections[0]
    footer = sec.footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run()
    try:
        r.add_picture(io.BytesIO(logo_bytes), width=Inches(0.7))
    except Exception:
        r.add_text("PowerDash HR")
    _set_run_font(r, size=9)


def render_docx(pack: Dict[str, Any], role_title: str, interview_type: str, duration: str,
                logo_bytes: bytes | None, footer_logo_bytes: bytes | None) -> bytes:
    doc = Document()

    for sec in doc.sections:
        sec.top_margin = Inches(0.6)
        sec.bottom_margin = Inches(0.6)
        sec.left_margin = Inches(0.7)
        sec.right_margin = Inches(0.7)

    _add_logo_header(doc, logo_bytes)

    _add_heading(doc, f"{role_title} — Competency Pack", level=1)
    pmeta = doc.add_paragraph()
    run = pmeta.add_run(f"Interview type: {interview_type} · Duration: {duration}")
    _set_run_font(run, size=10)

    if pack["meta"].get("housekeeping"):
        _add_spacer(doc, 4)
        _add_heading(doc, "Housekeeping", level=2)
        for b in HOUSEKEEPING_BULLETS:
            li = doc.add_paragraph(style=None)
            run = li.add_run(f"• {b}")
            _set_run_font(run, size=11)

    _add_spacer(doc, 6)
    _add_heading(doc, "Core Questions", level=2)

    for idx, it in enumerate(pack["questions"], start=1):
        p = doc.add_paragraph()
        r = p.add_run(f"Question {idx}: {it['question']}")
        _set_run_font(r, size=12, bold=True)

        p = doc.add_paragraph()
        r = p.add_run(f"Intent: {it['intent']}")
        _set_run_font(r, size=11)

        if it.get('what_good_looks_like'):
            p = doc.add_paragraph()
            r = p.add_run(f"What good looks like: {it['what_good_looks_like']}")
            _set_run_font(r, size=11)

        if it.get('follow_ups'):
            p = doc.add_paragraph()
            r = p.add_run("Follow‑ups:")
            _set_run_font(r, size=11, bold=True)
            for fu in it['follow_ups']:
                p = doc.add_paragraph()
                r = p.add_run(f"• {fu}")
                _set_run_font(r, size=11)

        _add_dotted_notes(doc, height_lines=6)
        _add_spacer(doc, 6)

    _add_footer_logo(doc, footer_logo_bytes)

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out.read()
