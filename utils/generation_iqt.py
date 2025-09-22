# utils/generation_iqt.py
import os, json, re, textwrap
from datetime import date
from typing import Dict, List
from functools import lru_cache
from openai import OpenAI

SECTION_ORDER = [
    "Housekeeping",
    "Overview",
    "Core Questions",
    "Competency Questions",
    "Technical Questions",
    "Culture & Values",
    "Closing Questions",
    "Close-down & Next Steps",   # we’ll guarantee this exists
    "Scoring Rubric",
]

@lru_cache(maxsize=1)
def _client() -> OpenAI:
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        try:
            import streamlit as st
            api_key = st.secrets.get("OPENAI_API_KEY")
        except Exception:
            pass
    if not api_key:
        raise RuntimeError("OPENAI_API_KEY not found in env or Streamlit Secrets.")
    return OpenAI(api_key=api_key)

def _slug(s: str) -> str:
    s = re.sub(r"[^a-zA-Z0-9]+", "-", (s or "interview-pack").lower()).strip("-")
    return f"{s}-{date.today().isoformat()}"

def _json_prompt(inputs: Dict) -> str:
    return f"""
You are an executive-search interviewer. Return ONLY valid JSON with this schema:

{{
  "housekeeping": [ "bullet point", ... ],
  "sections": [
    {{
      "name": "Core Questions" | "Competency Questions" | "Technical Questions" | "Culture & Values" | "Closing Questions" | "Overview" | "Close-down & Next Steps" | "Scoring Rubric",
      "questions": [
        {{
          "question": "short behaviour-based question",
          "intent": "why we ask it",
          "followups": ["optional, short prompts"],
          "good": "what good looks like (optional)"
        }}
      ],
      "notes": "optional brief prose for this section",
      "bullets": ["optional bullet list for guidance"]
    }}
  ]
}}

Guidance:
- Language: {inputs.get('language','English')}; Jurisdiction: {inputs.get('jurisdiction','UK')}.
- Role: {inputs.get('role_title')} · Level: {inputs.get('level')} · Dept: {inputs.get('department')}
- Interview type: {inputs.get('interview_type')} · Duration: {inputs.get('duration_mins')} mins
- Competencies: {", ".join(inputs.get('competencies', []))}
- Include approx {inputs.get('num_core')} core, {inputs.get('num_technical')} technical, {inputs.get('num_competency')} competency questions.
- Include follow-ups: {inputs.get('include_followups')}
- Include "what good looks like": {inputs.get('include_good_looks_like')}
- Include scoring rubric section: {inputs.get('include_scoring')}
- House guidance (use if helpful): {inputs.get('house_guidance') or "None"}

Make sure **Housekeeping** includes opener bullets (welcome, agenda, timings, consent, DEI/legal reminder, note-taking),
and include a section **"Close-down & Next Steps"** with bullets covering: thanking the candidate, what happens next, decision timelines, who contacts them, and how feedback is shared.

Style:
- Executive tone, inclusive, lawful; keep questions concise and behaviour-based.
- Return JSON only (no markdown or prose outside the JSON).
"""

def generate_interview_pack(inputs: Dict, model: str = "gpt-4.1-mini", temperature: float = 0.3) -> Dict:
    """Call OpenAI to produce a structured interview pack as strict JSON, then build HTML preview."""
    client = _client()
    prompt = _json_prompt(inputs)

    resp = client.chat.completions.create(
        model=model,
        temperature=temperature,
        response_format={"type": "json_object"},
        messages=[
            {"role": "system", "content": "You produce strictly valid JSON and nothing else."},
            {"role": "user", "content": textwrap.dedent(prompt).strip()},
        ],
    )
    data = json.loads(resp.choices[0].message.content or "{}")

    # Normalize sections and enforce order
    by_name = { (s.get("name") or "").strip(): s for s in data.get("sections", []) if isinstance(s, dict) }
    ordered: List[Dict] = [by_name[n] for n in SECTION_ORDER if n in by_name] + \
                          [s for s in data.get("sections", []) if (s.get("name") or "") not in SECTION_ORDER]

    # Ensure Close-down section exists (create a sensible default if the model omitted it)
    close_exists = any((s.get("name") or "").strip().lower() in {"close-down & next steps", "close down & next steps", "close-down", "next steps"} for s in ordered)
    if not close_exists:
        ordered.append({
            "name": "Close-down & Next Steps",
            "bullets": [
                "Thank the candidate for their time and engagement.",
                "Outline next steps in the process and expected timelines.",
                "Confirm who will contact them and by when.",
                "Explain how feedback will be shared and through which channel.",
                "Invite final questions and ensure the candidate knows how to follow up."
            ],
            "questions": [],
            "notes": "",
        })

    # Clean up questions (ensure trailing '?', strip, etc.)
    for sec in ordered:
        for q in (sec.get("questions") or []):
            qt = (q.get("question") or "").strip()
            if qt and not qt.endswith("?"):
                # add '?' if it looks like a question and doesn't already end with punctuation
                if not qt.endswith((".", "!", "?”", ".”", "!”")):
                    qt += "?"
            q["question"] = qt

    # ---------- Build polished HTML preview ----------
    def qblock(q: Dict) -> str:
        follow = ", ".join((q.get("followups") or [])[:6]) if q.get("followups") else ""
        parts = [
            "<div class='q-table'>",
            f"<div class='q-row'><div class='q-label'>Question</div><div>{(q.get('question') or '').strip()}</div></div>",
        ]
        if q.get("intent"):
            parts.append(f"<div class='q-row'><div class='q-label'>Intent</div><div>{q['intent']}</div></div>")
        if follow:
            parts.append(f"<div class='q-row'><div class='q-label'>Follow-ups</div><div>{follow}</div></div>")
        if q.get("good"):
            parts.append(f"<div class='q-row'><div class='q-label'>What good looks like</div><div>{q['good']}</div></div>")
        # white-space notes (no dots)
        parts.append("<div style='height:72px'></div>")
        parts.append("</div>")
        return "\n".join(parts)

    html_parts: List[str] = [
        f"<h2 style='margin-bottom:0'>{inputs.get('role_title') or 'Interview Pack'}</h2>",
        f"<div class='muted'>{inputs.get('interview_type')} interview · {inputs.get('duration_mins')} mins</div>",
    ]

    hk = data.get("housekeeping") or []
    if hk:
        html_parts.append("<div class='section-title'>Housekeeping</div>")
        bullets = "".join(f"<li>{x}</li>" for x in hk if x)
        html_parts.append(f"<div class='callout'><ul>{bullets}</ul></div>")

    for sec in ordered:
        name = sec.get("name") or "Section"
        html_parts.append(f"<div class='section-title'>{name}</div>")
        # bullets list (for close-down etc.)
        bullets = sec.get("bullets") or []
        if bullets:
            html_parts.append(f"<div class='callout'><ul>{''.join(f'<li>{b}</li>' for b in bullets)}</ul></div>")
        if sec.get("notes"):
            html_parts.append(f"<div class='muted' style='margin:.25rem 0 .5rem'>{sec['notes']}</div>")
        for q in sec.get("questions", []) or []:
            html_parts.append(qblock(q))

    return {
        "title": f"{inputs.get('role_title') or 'Interview'} — {inputs.get('interview_type')} Pack",
        "slug": _slug(inputs.get("role_title")),
        "inputs": inputs,
        "housekeeping": hk,
        "sections": ordered,          # structured JSON (exporters read this)
        "html_preview": "\n".join(html_parts),
    }
