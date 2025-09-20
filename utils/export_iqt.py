import io, os, requests
from typing import Dict, List
from docx import Document
from docx.shared import Pt, Inches
from docx.oxml.ns import qn

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader

FONT_NAME = "Source Sans 3"

def _set_document_defaults(doc: Document):
    style = doc.styles['Normal']
    style.font.name = FONT_NAME
    style._element.rPr.rFonts.set(qn('w:eastAsia'), FONT_NAME)
    style.font.size = Pt(11)

def _add_heading(doc: Document, text: str):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(16)

def _add_question_block(doc: Document, q: Dict):
    doc.add_paragraph("")
    p = doc.add_paragraph()
    p.add_run("Q: ").bold = True
    p.add_run(q.get("q", ""))
    if q.get("intent"):
        p2 = doc.add_paragraph(); p2.add_run("Intent: ").italic = True; p2.add_run(q["intent"])
    if q.get("good"):
        p3 = doc.add_paragraph(); p3.add_run("What good looks like: ").italic = True; p3.add_run(q["good"])
    if q.get("followups"):
        p4 = doc.add_paragraph(); p4.add_run("Follow-ups: ").italic = True; p4.add_run(", ".join(q["followups"]))
    for _ in range(4):
        doc.add_paragraph("_______________________________")

def _parse_questions(lines: List[str]) -> List[Dict]:
    qs = []; q = {"q": "", "intent": "", "followups": [], "good": ""}
    def flush():
        if q["q"].strip(): qs.append(q.copy())
    for raw in lines:
        line = raw.strip()
        if not line: continue
        l = line.lower()
        if l.startswith(("question:", "q:")):
            flush(); q = {"q": line.split(":",1)[1].strip(), "intent":"", "followups":[], "good":""}
        elif l.startswith("intent:"):
            q["intent"] = line.split(":",1)[1].strip()
        elif l.startswith(("follow-up:", "follow-ups:")):
            val = line.split(":",1)[1]; q["followups"] = [x.strip("- • ") for x in val.split(";") if x.strip()]
        elif l.startswith(("what good looks like:", "good:", "model answer:", "scoring hint:")):
            q["good"] = line.split(":",1)[1].strip()
        elif line.startswith("-") and not q["q"]:
            q["q"] = line.lstrip("- "); flush(); q = {"q": "", "intent":"", "followups":[], "good":""}
    flush(); return qs

def _add_footer_powerdash(doc: Document, pd_logo_path: str):
    """Set a footer with the PowerDash logo + text for all sections (repeats every page)."""
    for section in doc.sections:
        footer = section.footer
        footer.is_linked_to_previous = False
        p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        run = p.add_run()
        # Try to place logo if available
        try:
            if pd_logo_path and os.path.exists(pd_logo_path):
                run.add_picture(pd_logo_path, width=Inches(0.6))
                p.add_run("  Powered by PowerDash HR").italic = True
            else:
                p.add_run("Powered by PowerDash HR").italic = True
        except Exception:
            p.add_run("Powered by PowerDash HR").italic = True

def pack_to_docx(pack: Dict, tenant_name: str = "", logo_url: str = "", pd_logo_path: str = "assets/powerdash-logo.png") -> bytes:
    from docx.shared import Inches
    doc = Document()
    _set_document_defaults(doc)

    # Header (client logo + meta)
    if logo_url:
        try:
            img = requests.get(logo_url, timeout=6).content
            doc.add_picture(io.BytesIO(img), width=Inches(1.4))
        except Exception:
            pass

    _add_heading(doc, pack["title"])
    sub = doc.add_paragraph()
    sub.add_run(f"Interview type: {pack['inputs'].get('interview_type')} · Duration: {pack['inputs'].get('duration_mins')} mins")
    if tenant_name:
        t = doc.add_paragraph(tenant_name); t.runs[0].font.size = Pt(10)

    # Sections
    for name, lines in pack["sections"].items():
        doc.add_paragraph(""); _add_heading(doc, name)
        qs = _parse_questions(lines)
        if qs:
            for q in qs: _add_question_block(doc, q)
        else:
            doc.add_paragraph("\n".join([ln for ln in lines if ln.strip()]))

    # Footer on every page
    _add_footer_powerdash(doc, pd_logo_path)

    bio = io.BytesIO(); doc.save(bio); bio.seek(0)
    return bio.getvalue()

def pack_to_pdf(pack: Dict, tenant_name: str = "", logo_url: str = "", pd_logo_path: str = "assets/powerdash-logo.png") -> bytes:
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4

    x = 20 * mm; y = height - 20 * mm

    # Client logo (top)
    if logo_url:
        try:
            c.drawImage(ImageReader(logo_url), x, y-15*mm, width=30*mm, height=15*mm,
                        preserveAspectRatio=True, mask='auto')
        except Exception:
            pass

    # Title/meta
    c.setFont("Helvetica-Bold", 14); c.drawString(x, y-18*mm, pack["title"])
    c.setFont("Helvetica", 10)
    c.drawString(x, y-23*mm, f"Interview type: {pack['inputs'].get('interview_type')} · Duration: {pack['inputs'].get('duration_mins')} mins")
    if tenant_name: c.drawString(x, y-28*mm, tenant_name)

    cur_y = y - 35*mm
    def add_notes_lines(n=4):
        nonlocal cur_y
        for _ in range(n):
            c.line(x, cur_y, width - x, cur_y); cur_y -= 6
        cur_y -= 2

    def write_wrapped(text, bold=False, size=None):
        nonlocal cur_y
        font = "Helvetica-Bold" if bold else "Helvetica"
        size = 11 if bold else 10 if size is None else size
        c.setFont(font, size); max_w = width - 2*x
        words, line = text.split(), ""
        for w in words:
            trial = (line + " " + w).strip()
            if c.stringWidth(trial, font, size) > max_w:
                c.drawString(x, cur_y, line); cur_y -= 12; line = w
            else:
                line = trial
        if line: c.drawString(x, cur_y, line); cur_y -= 12

    def draw_footer():
        # PowerDash logo + text centered at bottom
        footer_y = 12*mm
        try:
            if pd_logo_path and os.path.exists(pd_logo_path):
                img = ImageReader(pd_logo_path)
                img_w = 14*mm; img_h = 14*mm
                cx = width/2
                c.drawImage(img, cx - img_w - 12, footer_y-3, width=img_w, height=img_h,
                            preserveAspectRatio=True, mask='auto')
                c.setFont("Helvetica-Oblique", 9)
                c.drawString(cx - 12, footer_y+3, "Powered by PowerDash HR")
            else:
                c.setFont("Helvetica-Oblique", 9)
                c.drawCentredString(width/2, footer_y+3, "Powered by PowerDash HR")
        except Exception:
            c.setFont("Helvetica-Oblique", 9)
            c.drawCentredString(width/2, footer_y+3, "Powered by PowerDash HR")

    # Content with pagination
    for name, lines in pack["sections"].items():
        write_wrapped(name, bold=True)
        # parse questions lightly: bullet or labelled items
        qs = []
        q = {"q": "", "intent": "", "followups": [], "good": ""}
        def flush():
            if q["q"].strip(): qs.append(q.copy())
        for raw in lines:
            line = raw.strip()
            if not line: continue
            l = line.lower()
            if l.startswith(("question:", "q:")):
                flush(); q = {"q": line.split(":",1)[1].strip(), "intent":"", "followups":[], "good":""}
            elif l.startswith("intent:"):
                q["intent"] = line.split(":",1)[1].strip()
            elif l.startswith(("follow-up:", "follow-ups:")):
                val = line.split(":",1)[1]; q["followups"] = [x.strip("- • ") for x in val.split(";") if x.strip()]
            elif l.startswith(("what good looks like:", "good:", "model answer:", "scoring hint:")):
                q["good"] = line.split(":",1)[1].strip()
            elif line.startswith("-") and not q["q"]:
                q["q"] = line.lstrip("- "); flush(); q = {"q": "", "intent":"", "followups":[], "good":""}
        flush()

        if qs:
            for item in qs:
                write_wrapped(f"Q: {item['q']}")
                if item.get('intent'): write_wrapped(f"Intent: {item['intent']}")
                if item.get('good'): write_wrapped(f"What good looks like: {item['good']}")
                if item.get('followups'): write_wrapped("Follow-ups: " + ", ".join(item['followups']))
                add_notes_lines(4)
                if cur_y < 40*mm:
                    draw_footer(); c.showPage()
                    cur_y = A4[1] - 20*mm
        else:
            write_wrapped(" ".join([ln for ln in lines if ln.strip()]))

        cur_y -= 6
        if cur_y < 40*mm:
            draw_footer(); c.showPage()
            cur_y = A4[1] - 20*mm

    draw_footer(); c.save()
    buffer.seek(0); return buffer.getvalue()
