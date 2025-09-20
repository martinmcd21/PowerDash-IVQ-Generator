import os
import streamlit as st
from dotenv import load_dotenv

load_dotenv()
st.set_page_config(page_title="Interview Question Pack Generator", page_icon="❓", layout="wide")

# ---------- Fonts & CSS ----------
GOOGLE_FONTS = "https://fonts.googleapis.com/css2?family=Source+Sans+3:wght@300;400;600;700&display=swap"
st.markdown(f"<link href='{GOOGLE_FONTS}' rel='stylesheet'>", unsafe_allow_html=True)
st.markdown("""
<style>
  html, body, [class*="css"], textarea, input { font-family:'Source Sans 3', -apple-system, BlinkMacSystemFont, Segoe UI, Roboto, Helvetica, Arial, sans-serif !important; }
  .muted { opacity:.9; }
  .section-title { font-weight:700; font-size:1.1rem; margin:1rem 0 .25rem; }
  .powerdash-footer { margin-top:2rem; padding-top:.75rem; border-top:1px solid rgba(0,0,0,.08); text-align:center; opacity:.85; }
  .brand-chip { display:inline-flex; align-items:center; gap:.5rem; padding:.25rem .5rem; border-radius:999px; background:#fffbe6; border:1px solid #ffe58f; }
  .logo-row { display:flex; gap:12px; justify-content:flex-end; align-items:center; }
  .logo-row img { max-height:40px; border-radius:8px; background:white; }

  /* Executive preview blocks */
  .q-table { width:100%; border:1px solid #E5E7EB; border-radius:12px; padding:12px; margin:.5rem 0; }
  .q-row { display:grid; grid-template-columns:140px 1fr; gap:12px; margin:6px 0; }
  .q-label { font-weight:600; color:#374151; }
  .q-notes { margin-top:8px; }
  .q-line { height:12px; border-bottom:1px dashed #D1D5DB; }
  .callout { background:#F9FAFB; border:1px solid #E5E7EB; border-radius:12px; padding:12px; }
</style>
""", unsafe_allow_html=True)

# ---------- Sidebar ----------
with st.sidebar:
  st.header("⚙️ Settings")
  st.caption("Branding and generation options.")
  tenant_name = st.text_input("Organisation (optional)")
  primary_colour = st.color_picker("Primary colour", "#111827")
  logo_url = st.text_input("Client logo URL (optional)")
  show_powerdash = st.toggle("Show 'Powered by PowerDash HR'", value=True)

  default_model = os.getenv("OPENAI_MODEL", "gpt-4.1-mini")
  model = st.text_input("OpenAI model", value=default_model)
  temperature = st.slider("Creativity (temperature)", 0.0, 1.2, 0.3, 0.1)

  language = st.selectbox("Language", ["English", "Français", "Deutsch", "Español", "Italiano"], 0)
  jurisdiction = st.selectbox("Jurisdiction", ["Global", "UK", "USA", "EU", "Canada", "Australia"], 1)

# ---------- Header ----------
left, right = st.columns([0.72, 0.28])
with left:
  title = "Interview Question Pack Generator"
  if tenant_name:
    title += f" — {tenant_name}"
  st.title(title)
  st.caption("Generate a structured, competency-based interview pack with space for notes and a scoring rubric.")

with right:
  # Client logo (URL) + PowerDash logo (local asset or URL from Secrets)
  html = "<div class='logo-row'>"
  if logo_url:
    html += f"<img src='{logo_url}' alt='Client logo'/>"
  pd_asset = "assets/powerdash-logo.png"
  pd_url = st.secrets.get("POWERDASH_LOGO_URL", "")
  if os.path.exists(pd_asset):
    st.image(pd_asset, width=140)  # render local asset cleanly
    html = ""  # prevent double rendering
  elif pd_url:
    html += f"<img src='{pd_url}' alt='PowerDash HR'/>"
  html += "</div>"
  if html != "</div>":
    st.markdown(html, unsafe_allow_html=True)

# ---------- Inputs ----------
st.subheader("Role Details")
col1, col2, col3 = st.columns(3)
with col1:
  role_title = st.text_input("Role title", placeholder="e.g., Senior Data Analyst")
  level = st.selectbox("Seniority", ["Entry", "Associate", "Mid", "Senior", "Lead", "Manager", "Director"], 3)
with col2:
  department = st.text_input("Department/Function", placeholder="Analytics")
  interview_type = st.selectbox("Interview type", ["Screen", "Technical", "Competency", "Panel", "Final"], 2)
with col3:
  duration_mins = st.number_input("Duration (minutes)", 20, 180, 60, 5)

st.subheader("Content Controls")
col4, col5, col6 = st.columns(3)
with col4:
  num_core = st.slider("# Core questions", 2, 10, 4)
  include_followups = st.toggle("Include suggested follow-ups", True)
with col5:
  num_technical = st.slider("# Technical questions", 0, 10, 3)
  include_good_looks_like = st.toggle("Include 'what good looks like'", True)
with col6:
  num_competency = st.slider("# Competency questions", 0, 12, 5)
  include_scoring = st.toggle("Include scoring rubric", True)

competencies = st.text_area(
  "Target competencies (one per line)",
  height=120,
  placeholder="Problem solving\nStakeholder management\nCommunication\nOwnership",
)
house_guidance = st.text_area(
  "House guidance (optional)", height=100,
  placeholder="Use UK English. Keep questions short, open, and behaviour-based. Avoid discriminatory topics.",
)

# ---------- Action ----------
if st.button("Generate Interview Pack", type="primary"):
  # Lazy imports so a module error doesn't blank the app
  try:
    from utils.generation_iqt import generate_interview_pack
  except Exception as e:
    st.error(f"Couldn't load generator module (utils/generation_iqt.py).\n\n**Import error:** {e}")
    st.stop()
  try:
    from utils.export_iqt import pack_to_docx, pack_to_pdf
  except Exception as e:
    st.warning(f"Export module failed to load; generation will still work.\n\n**Import error:** {e}")
    pack_to_docx = pack_to_pdf = None

  with st.spinner("Drafting your pack…"):
    inputs = {
      "role_title": role_title, "level": level, "department": department,
      "interview_type": interview_type, "duration_mins": duration_mins,
      "num_core": num_core, "num_technical": num_technical, "num_competency": num_competency,
      "competencies": [x.strip() for x in (competencies or "").split("\n") if x.strip()],
      "include_followups": include_followups, "include_good_looks_like": include_good_looks_like,
      "include_scoring": include_scoring, "language": language, "jurisdiction": jurisdiction,
      "tenant_name": tenant_name, "logo_url": logo_url, "primary_colour": primary_colour,
      "house_guidance": house_guidance,
    }
    try:
      pack = generate_interview_pack(inputs=inputs, model=os.getenv("OPENAI_MODEL", "gpt-4.1-mini"), temperature=0.3)
    except Exception as e:
      st.error(f"Generation failed: {e}")
      st.stop()

  st.success("Interview pack ready ✨")
  st.markdown(pack["html_preview"], unsafe_allow_html=True)

  # --- Exports (expects PowerDash logo at assets/powerdash-logo.png) ---
  st.subheader("Export")
  pd_logo_path = "assets/powerdash-logo.png"
  if pack_to_docx:
    try:
      docx_bytes = pack_to_docx(pack, tenant_name=tenant_name, logo_url=logo_url, pd_logo_path=pd_logo_path)
      st.download_button("⬇️ Download DOCX", data=docx_bytes, file_name=f"{pack['slug']}.docx",
                         mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    except Exception as e:
      st.warning(f"DOCX export failed: {e}")
  if pack_to_pdf:
    try:
      pdf_bytes = pack_to_pdf(pack, tenant_name=tenant_name, logo_url=logo_url, pd_logo_path=pd_logo_path)
      st.download_button("⬇️ Download PDF", data=pdf_bytes, file_name=f"{pack['slug']}.pdf",
                         mime="application/pdf")
    except Exception as e:
      st.warning(f"PDF export failed: {e}")

# ---------- Footer (web app) ----------
if show_powerdash:
  st.markdown(
    f"""
    <div class='powerdash-footer'>
      <span class='brand-chip' style='--chip-bg:{primary_colour}10'>⚡️ Powered by <strong>PowerDash HR</strong></span>
    </div>
    """,
    unsafe_allow_html=True,
  )
