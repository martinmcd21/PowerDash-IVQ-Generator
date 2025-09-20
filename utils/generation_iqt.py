import os
from openai import OpenAI

def generate_interview_pack(inputs, model="gpt-4.1-mini", temperature=0.3):
    """
    Calls OpenAI to generate a structured interview pack based on role inputs.
    Returns a dict with slug, title, html_preview, and sections.
    """

    client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

    # Prompt
    prompt = f"""
    Create a structured interview pack for the following role:

    Role: {inputs.get('role_title')}
    Level: {inputs.get('level')}
    Department: {inputs.get('department')}
    Interview type: {inputs.get('interview_type')}
    Duration: {inputs.get('duration_mins')} minutes
    Competencies: {", ".join(inputs.get('competencies', []))}
    Guidance: {inputs.get('house_guidance', '')}

    Include:
    - {inputs.get('num_core')} core questions
    - {inputs.get('num_technical')} technical questions
    - {inputs.get('num_competency')} competency questions
    - Follow-ups: {inputs.get('include_followups')}
    - What good looks like: {inputs.get('include_good_looks_like')}
    - Scoring rubric: {inputs.get('include_scoring')}

    Respond in a structured, clear format with sections and short questions.
    """

    resp = client.chat.completions.create(
        model=model,
        temperature=temperature,
        messages=[
            {"role": "system", "content": "You are an expert HR interviewer. Produce interview packs in plain text with clear headings and bullet questions."},
            {"role": "user", "content": prompt}
        ],
    )

    text = resp.choices[0].message.content

    # Very simple packaging for now
    sections = {"Interview Questions": text.split("\n")}

    return {
        "slug": inputs.get("role_title", "interview-pack").lower().replace(" ", "-"),
        "title": f"Interview Pack — {inputs.get('role_title', '')}",
        "html_preview": f"<pre>{text}</pre>",
        "sections": sections,
        "inputs": inputs,
    }
