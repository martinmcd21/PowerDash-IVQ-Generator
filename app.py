# app.py — Interview Question Pack Generator (PowerDash HR)

import os
from datetime import datetime
import streamlit as st

# Local utils
from utils.generation_iqt import generate_interview_pack
from utils.export_iqt import pack_to_docx, pack_to_pdf

# Optional imports for JD parsing
try:
    from docx import Document as DocxDocument
except Exception:
    DocxDocument = None

try:
    import pypdf
except Exception:
    pypdf = None


# =====================
# Page config & styles
# =====================
st.set_page_config(
    page_title="Interview Question Pack Generator · PowerDash HR",
    page_icon="🧩",
    layout="wide",
)

PRIMARY_ACCENT = st.session_state.get("primary_accent", "#111827")  # slate-900 default

APP_CSS = f"""
<style>
h2,h3,h4 {{ margin-bottom: .35rem; }}
.section-title {{ margin: 1rem 0 .5rem; font-weight: 700; font-size: 1.15rem; }}
.callout {{ background: #f7f7fb; border: 1px solid #ececf2; border-radius: 8px; padding: .6rem .8rem; }}
.muted {{ color: #6b7280; font-size:.9rem; }}
footer {{ visibility: hidden; }}

.powered {{
  display:inline-flex; align-items:center; gap:.35rem;
  background:#fff7e5; border:1px solid #ffe2a7; color:#111827;
  border-radius:999px; padding:.25rem .6rem; font-size:.85rem;
}}

.stButton>button[kind="primary"] {{
  background:{PRIMARY_ACCENT} !important;
}}

.small {{ font-size:.85rem; color:#6b7280 }}
.preview {{ border:1px solid #e5e7eb; border-radius:10px; padding:1rem; background:#fff; }}
.q-table{{ border:1px solid #e5e7eb; border-radius:10px; padding:.75rem; margin:.5rem 0 1rem; }}
.q-row{{ display:grid; grid-template-columns:150px 1fr; gap:.5rem; margin:.15rem 0; }}
.q-label{{ font-weight:700; color:#111827; }}
</style>
"""
st.markdown(APP_CSS, unsafe_allow_html=True)


# =====================
# Sidebar — Settings
# =====================
st.sidebar.title("⚙️ Settings")
org_name = st.sidebar.text_input("Organisation (optional)", value="")
primary_colour = st.sidebar.color_picker("Primary colour", value=PRIMARY_ACCENT)
st.session_state["primary_accent"] = primary_colour

client_logo_url = st.sidebar.text_input("Client logo URL (optional)", value="")
show_powered = st.sidebar.toggle("Show 'Powered by PowerDash HR'", value=True)

# Model picker (safe options)
MODEL_OPTIONS = [
    ("gpt-4.1-mini", "Fast & cost-efficient (recommended)"),
    ("gpt-4.1",      "Higher reasoning (slower, pricier)"),
    ("gpt-3.5-turbo","Budget (lower quality)"),
]
labels = {k: f"{k} — {v}" for k, v in MODEL_OPTIONS}
selected_model = st.sidebar.selectbox(
    "OpenAI model",
    options=list(labels.keys()),
    format_func=lambda k: labels[k],
    index=0,
)

creativity = st.sidebar.slider("Creativity (temperature)", 0.0, 1.0, 0.30, 0.05)
language = st.sidebar.selectbox("Language", ["English", "French", "Spanish", "German", "Italian"], index=0)
jurisdiction = st.sidebar.selectbox("Jurisdiction", ["UK", "EU", "US", "Global"], index=0)

# Optional Job Description upload
st.sidebar.markdown("---")
jd_file = st.sidebar.file_uploader("Upload job description (optional)", type=["txt", "docx", "pdf"])
jd_raw = None
if jd_file is not None:
    name = jd_file.name.lower()
    try:
        if name.endswith(".txt"):
            jd_raw = jd_file.read().decode("utf-8", errors="ignore")
        elif name.endswith(".docx"):
            if DocxDocument is None:
                st.sidebar.warning("python-docx not available; cannot read DOCX.")
            else:
                doc = DocxDocument(jd_file)
                jd_raw = "\n".join(p.text for p in doc.paragraphs)
        elif name.endswith(".pdf"):
            if pypdf is None:
                st.sidebar.warning("pypdf not available; cannot read PDF.")
            else:
                reader = pypdf.PdfReader(jd_file)
                jd_raw = "\n".join((page.extract_text() or "") for page in reader.pages)
    except Exception as e:
        st.sidebar.warning(f"Could not read file: {e}")


# =====================
# Main — Inputs
# =====================
st.title("Interview Question Pack Generator")
st.caption("Create polished, structured interview packs tailored to a role, with export to PDF/DOCX.")

c1, c2, c3 = st.columns(3)
with c1:
    num_core = st.slider("Core questions", 0, 10, 4)
with c2:
    num_technical = st.slider("Technical questions", 0, 10, 3)
with c3:
    num_competency = st.slider("Competency questions", 0, 10, 5)

c1, c2, c3 = st.columns(3)
with c1:
    include_followups = st.toggle("Include suggested follow-ups", value=True)
with c2:
    include_good = st.toggle("Include 'what good looks like'", value=True)
with c3:
    include_scoring = st.toggle("Include scoring rubric", value=True)

st.markdown("**Target competencies (one per line)**")
competencies_text = st.text_area(
    "",
    value="Problem solving\nStakeholder management\nCommunication\nOwnership",
    height=120,
    placeholder="One competency per line…",
)

st.markdown("**House guidance (optional)**")
house_guidance = st.text_area(
    "",
    value="Use UK English. Keep questions short, open, and behaviour-based. Avoid discriminatory topics.",
    height=90,
    placeholder="e.g., preferred tone, legal reminders, DEI guidance…",
)

cA, cB, cC = st.columns(3)
with cA:
    role_title = st.text_input("Role title", value="Accountant")
with cB:
    level = st.text_input("Seniority level", value="Mid")
with cC:
    department = st.text_input("Department", value="Finance")

cA, cB, cC = st.columns(3)
with cA:
    interview_type = st.selectbox("Interview type", ["Competency", "Technical", "Mixed"], index=0)
with cB:
    duration_mins = st.number_input("Duration (mins)", min_value=15, max_value=180, value=60, step=5)
with cC:
    pass


# =====================
# JD summary (optional)
# =====================
jd_summary = None
if jd_raw:
    try:
        from openai import OpenAI
        api_key = os.getenv("OPENAI_API_KEY") or st.secrets.get("OPENAI_API_KEY")
        if not api_key:
            st.sidebar.warning("No OPENAI_API_KEY found; JD will not be summarized.")
        else:
            client = OpenAI(api_key=api_key)
            resp = client.chat.completions.create(
                model=selected_model,
                temperature=0.2,
                messages=[
                    {"role": "system", "content": "Summarize the job description into crisp bullets of responsibilities, must-have skills, nice-to-haves, stakeholders, and tools. Keep it under 350 words. Return plain text only."},
                    {"role": "user", "content": jd_raw[:15000]},
                ],
            )
            jd_summary = (resp.choices[0].message.content or "").strip()
    except Exception as e:
        st.sidebar.warning(f"JD summary failed; using raw snippet. ({e})")
        jd_summary = jd_raw[:2000]


# =====================
# Generate
# =====================
if st.button("Generate Interview Pack", type="primary"):
    competencies = [c.strip() for c in competencies_text.splitlines() if c.strip()]

    inputs = {
        "role_title": role_title,
        "level": level,
        "department": department,
        "interview_type": interview_type,
        "duration_mins": int(duration_mins),
        "competencies": competencies,
        "num_core": num_core,
        "num_technical": num_technical,
        "num_competency": num_competency,
        "include_followups": bool(include_followups),
        "include_good_looks_like": bool(include_good),
        "include_scoring": bool(include_scoring),
        "house_guidance": house_guidance.strip(),
        "language": language,
        "jurisdiction": jurisdiction,
        "jd_context": jd_summary or (jd_raw[:2000] if jd_raw else None),
        "tenant_name": org_name.strip(),
        "client_logo_url": client_logo_url.strip(),
    }

    try:
        pack = generate_interview_pack(inputs, model=selected_model, temperature=creativity)
        st.session_state["pack"] = pack
        st.success("Interview pack generated.")
    except Exception as e:
        st.error(f"Could not load generator module: {e}")


# =====================
# Preview & Export
# =====================
pack = st.session_state.get("pack")
if pack:
    # Brand bar
    b1, b2 = st.columns([1, 1])
    with b1:
        st.markdown(f"### {pack['title']}")
        st.caption(f"{interview_type} · {duration_mins} mins")
    with b2:
        if show_powered:
            st.markdown(
                '<div class="powered">⚡ Powered by <strong>PowerDash HR</strong></div>',
                unsafe_allow_html=True,
            )

    # HTML preview
    with st.container(border=True):
        st.markdown(pack["html_preview"], unsafe_allow_html=True)

    st.markdown("### Export")

    # DOCX
    try:
        docx_bytes = pack_to_docx(
            pack,
            tenant_name=org_name,
            logo_url=client_logo_url,
            pd_logo_path="assets/powerdash-logo.png",
        )
        st.download_button(
            "Download DOCX",
            data=docx_bytes,
            file_name=f"{pack['slug']}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    except Exception as e:
        st.warning(f"DOCX export failed: {e}")

    # PDF
    try:
        pdf_bytes = pack_to_pdf(
            pack,
            tenant_name=org_name,
            logo_url=client_logo_url,
            pd_logo_path="assets/powerdash-logo.png",
        )
        st.download_button(
            "Download PDF",
            data=pdf_bytes,
            file_name=f"{pack['slug']}.pdf",
            mime="application/pdf",
        )
    except Exception as e:
        st.warning(f"PDF export failed: {e}")

# Footer badge
if show_powered:
    st.markdown(
        '<div style="margin-top:2rem" class="powered">⚡ Powered by <strong>PowerDash HR</strong></div>',
        unsafe_allow_html=True,
    )
