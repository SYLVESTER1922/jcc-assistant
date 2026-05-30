"""
JCC Assistant - Bible Study & Church Programs Chatbot
======================================================
A Gradio app for Jubilee Celebration Center - AFM.
"""

import os
from datetime import date
from typing import List, Tuple, Optional

import gradio as gr
from supabase import create_client, Client
from openai import OpenAI
from docx import Document
from pptx import Presentation


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
SUPABASE_URL = os.environ["SUPABASE_URL"]
SUPABASE_KEY = os.environ["SUPABASE_SERVICE_KEY"]
OPENAI_API_KEY = os.environ["OPENAI_API_KEY"]
ADMIN_PASSWORD = os.environ["ADMIN_PASSWORD"]

OPENAI_MODEL = "gpt-4o-mini"

sb: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
oai = OpenAI(api_key=OPENAI_API_KEY)


# ---------------------------------------------------------------------------
# Document parsing
# ---------------------------------------------------------------------------
def parse_docx(file_path: str) -> str:
    doc = Document(file_path)
    paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    return "\n\n".join(paragraphs)


def parse_pptx(file_path: str) -> str:
    prs = Presentation(file_path)
    sections = []
    for i, slide in enumerate(prs.slides, start=1):
        slide_lines = [f"## Slide {i}"]
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    txt = "".join(run.text for run in para.runs).strip()
                    if txt:
                        slide_lines.append(txt)
        if len(slide_lines) > 1:
            sections.append("\n".join(slide_lines))
    return "\n\n".join(sections)


def parse_study_document(file_path: str) -> str:
    lower = file_path.lower()
    if lower.endswith(".docx"):
        return parse_docx(file_path)
    elif lower.endswith(".pptx"):
        return parse_pptx(file_path)
    else:
        raise ValueError("Unsupported file type. Please upload .docx or .pptx.")


# ---------------------------------------------------------------------------
# Bible Study DB helpers
# ---------------------------------------------------------------------------
def list_bible_studies() -> List[Tuple[str, str]]:
    """Return list of (label, id) for the dropdown, newest first."""
    res = sb.table("bible_studies") \
        .select("id, week_of, title, presenter") \
        .order("week_of", desc=True) \
        .order("uploaded_at", desc=True) \
        .execute()
    options = []
    for row in res.data:
        label = f"{row['week_of']} - {row['title']}"
        if row.get("presenter"):
            label += f" ({row['presenter']})"
        options.append((label, row["id"]))
    return options


def get_bible_study(study_id: str) -> Optional[dict]:
    if not study_id:
        return None
    res = sb.table("bible_studies").select("*").eq("id", study_id).single().execute()
    return res.data


def upload_bible_study(file, title, presenter, week_of, password):
    """Upload a Bible study. If a study with the same week_of exists, replace it."""
    if password != ADMIN_PASSWORD:
        return "❌ Incorrect admin password."
    if not file:
        return "❌ Please attach a .docx or .pptx file."
    if not title or not title.strip():
        return "❌ Title is required."
    if not week_of:
        return "❌ Week-of date is required."

    try:
        text = parse_study_document(file.name)
    except Exception as e:
        return f"❌ Could not parse the document: {e}"

    if len(text) < 50:
        return "❌ The document looks empty after parsing."

    # Check for existing study with same week_of -> replace it
    existing = sb.table("bible_studies").select("id").eq("week_of", week_of).execute()
    if existing.data:
        existing_id = existing.data[0]["id"]
        try:
            sb.table("bible_studies").update({
                "title": title.strip(),
                "presenter": (presenter or "").strip() or None,
                "document_text": text,
            }).eq("id", existing_id).execute()
            return (
                f"✅ Replaced existing study for week of {week_of} with **{title}**.\n\n"
                f"Document length: {len(text):,} characters."
            )
        except Exception as e:
            return f"❌ Database update failed: {e}"

    # No existing -> insert new
    try:
        sb.table("bible_studies").insert({
            "week_of": week_of,
            "title": title.strip(),
            "presenter": (presenter or "").strip() or None,
            "document_text": text,
        }).execute()
    except Exception as e:
        return f"❌ Database insert failed: {e}"

    return (
        f"✅ Uploaded **{title}** for the week of {week_of}.\n\n"
        f"Document length: {len(text):,} characters."
    )


def delete_bible_study(study_id, password):
    if password != ADMIN_PASSWORD:
        return "❌ Incorrect admin password.", gr.update()
    if not study_id:
        return "❌ Please select a study to delete.", gr.update()
    try:
        sb.table("bible_studies").delete().eq("id", study_id).execute()
    except Exception as e:
        return f"❌ Delete failed: {e}", gr.update()
    options = list_bible_studies()
    return "✅ Deleted.", gr.update(choices=options, value=options[0][1] if options else None)


# ---------------------------------------------------------------------------
# Church Programs context
# ---------------------------------------------------------------------------
def fetch_programs_context() -> str:
    ministries = sb.table("ministries").select("*").execute().data
    events = sb.table("events").select("*").order("event_date").execute().data
    notes = sb.table("ministry_notes").select("*").execute().data

    lines = ["# JCC 2026 MINISTRY PROGRAMS\n"]
    for m in ministries:
        lines.append(f"\n## {m['name']} Ministry")
        if m.get("lead"):
            lines.append(f"Led by: {m['lead']}")

        m_notes = [n for n in notes if n["ministry_id"] == m["id"]]
        for n in m_notes:
            heading = n["section"].replace("_", " ").title()
            lines.append(f"\n**{heading}:** {n['content']}")

        m_events = [e for e in events if e["ministry_id"] == m["id"]]
        if m_events:
            lines.append("\n**Events:**")
            for e in m_events:
                date_str = e.get("date_label") or e.get("event_date") or "TBD"
                line = f"- {date_str}: {e['title']}"
                if e.get("description"):
                    line += f" - {e['description']}"
                if e.get("format"):
                    line += f" ({e['format']})"
                lines.append(line)
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# System prompts
# ---------------------------------------------------------------------------
BIBLE_STUDY_PROMPT = """You are the JCC Bible Study Assistant for Jubilee Celebration Center - AFM.

You answer questions about a specific Bible study document shared with the group. Your answers MUST come only from the document below - do not add interpretation, outside scripture, or commentary.

RULES:
- Quote directly from the document using quotation marks for the key passage.
- Cite the section heading, number, or slide where the quote appears (e.g., "Section 7: Logos and Rhema" or "Slide 4").
- If the question is not addressed in the document, say:
  "This isn't covered in this week's study. Please bring it up with the group or the presenter."
- Do not speculate or fill gaps with general Bible knowledge.
- Keep answers focused. Concise is better than thorough.
- Maintain a respectful, reverent tone appropriate for Bible study.

---
THIS WEEK'S BIBLE STUDY
Title: {title}
Presenter: {presenter}
Week of: {week_of}

{document_text}
---
"""

PROGRAMS_PROMPT = """You are the JCC Programs Assistant for Jubilee Celebration Center - AFM.

You answer questions about JCC's 2026 ministry programs, events, leads, goals, and activities. Your answers MUST come only from the programs information below.

GUIDANCE:
- Be specific. When asked about events, give the date, ministry, and format.
- When asked about a ministry's vision, mission, or goals, summarize from the notes.
- Some questions are about people (e.g., "Who is Pastor Tabu Bere?"). Look across ALL ministries in the data - leads are listed at each ministry section. If a person appears as a lead of any ministry, mention that.
- For example, if asked "Who is the Pastor?", note that Pastor Tabu Bere leads the Outreach Ministry, even though there is no separate "Pastor" entry.
- If the information truly is not in the data below, say:
  "I don't have that information in the 2026 plans. Please check with the relevant ministry lead."
- Do not invent events, dates, or leads. Stick to the data.
- Keep answers focused and helpful.

---
JCC 2026 PROGRAMS DATA
{programs_context}
---
"""


# ---------------------------------------------------------------------------
# Chat
# ---------------------------------------------------------------------------
def chat(message, history, mode, study_id):
    if not message or not message.strip():
        return "Please type a question."

    if mode == "Bible Study":
        study = get_bible_study(study_id) if study_id else None
        if not study:
            return (
                "Please select a Bible study from the dropdown first. "
                "If the dropdown is empty, click the refresh button or ask an admin to upload one."
            )
        system_prompt = BIBLE_STUDY_PROMPT.format(
            title=study["title"],
            presenter=study.get("presenter") or "Not specified",
            week_of=study["week_of"],
            document_text=study["document_text"],
        )
    else:
        try:
            context = fetch_programs_context()
        except Exception as e:
            return f"Could not load programs data: {e}"
        system_prompt = PROGRAMS_PROMPT.format(programs_context=context)

    messages = [{"role": "system", "content": system_prompt}]
    for h in history:
        messages.append({"role": h["role"], "content": h["content"]})
    messages.append({"role": "user", "content": message})

    try:
        resp = oai.chat.completions.create(
            model=OPENAI_MODEL,
            messages=messages,
            temperature=0.2,
        )
        return resp.choices[0].message.content
    except Exception as e:
        return f"OpenAI request failed: {e}"


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------
def refresh_studies():
    options = list_bible_studies()
    if not options:
        return gr.update(choices=[], value=None)
    return gr.update(choices=options, value=options[0][1])


CUSTOM_CSS = """
.gradio-container {
    font-family: 'Inter', 'Helvetica Neue', system-ui, sans-serif !important;
    max-width: 1100px !important;
    margin: 0 auto !important;
}
#app-header {
    border-bottom: 1px solid #e5e7eb;
    padding-bottom: 12px;
    margin-bottom: 8px;
}
#app-header h1 {
    color: #1B2A4E;
    margin: 0 0 4px 0 !important;
    font-size: 1.8em !important;
    font-weight: 700;
}
#app-header .subtitle {
    color: #6b7280;
    font-size: 0.95em;
    margin: 0;
}
#app-header .gold-bar {
    width: 50px;
    height: 3px;
    background: #C9A55C;
    margin: 8px 0;
}
.tab-nav button {
    font-weight: 500 !important;
}
.tab-nav button.selected {
    color: #1B2A4E !important;
    border-bottom-color: #C9A55C !important;
}
footer {
    display: none !important;
}
"""


theme = gr.themes.Soft(
    primary_hue=gr.themes.colors.slate,
    secondary_hue=gr.themes.colors.amber,
    neutral_hue=gr.themes.colors.slate,
    font=[gr.themes.GoogleFont("Inter"), "system-ui", "sans-serif"],
).set(
    button_primary_background_fill="#1B2A4E",
    button_primary_background_fill_hover="#0F1A35",
    button_primary_text_color="white",
    body_background_fill="#F7F3EC",
    block_background_fill="white",
    block_border_color="#e5e7eb",
)


with gr.Blocks(title="JCC Assistant", theme=theme, css=CUSTOM_CSS) as demo:

    gr.HTML("""
    <div id="app-header">
        <h1>JCC Assistant</h1>
        <div class="gold-bar"></div>
        <p class="subtitle">Jubilee Celebration Center — AFM &nbsp;·&nbsp; Bible Study &amp; Ministry Programs</p>
    </div>
    """)

    with gr.Tabs():
        # ============================================================
        # CHAT TAB
        # ============================================================
        with gr.Tab("Chat"):
            with gr.Row(equal_height=True):
                mode = gr.Radio(
                    choices=["Bible Study", "Church Programs"],
                    value="Bible Study",
                    label="Mode",
                    container=True,
                    scale=2,
                )
                study_dropdown = gr.Dropdown(
                    label="Select Bible Study",
                    choices=list_bible_studies(),
                    container=True,
                    scale=4,
                )
                refresh_btn = gr.Button("Refresh", scale=1, size="sm")

            refresh_btn.click(fn=refresh_studies, outputs=study_dropdown)

            gr.ChatInterface(
                fn=chat,
                additional_inputs=[mode, study_dropdown],
                type="messages",
                examples=[
                    ["What does the study say about the main idea?"],
                    ["What scriptures are referenced?"],
                    ["What were the main points?"],
                    ["When is the next Couples Ministry event?"],
                    ["What's the vision for Praise & Worship?"],
                    ["What fundraising activities are planned for 2026?"],
                ],
            )

        # ============================================================
        # ADMIN TAB
        # ============================================================
        with gr.Tab("Admin"):
            gr.Markdown(
                "### Upload a Bible Study\n"
                "Accepts `.docx` or `.pptx` files. "
                "If a study already exists for the same week-of date, it will be **replaced**."
            )

            with gr.Row():
                with gr.Column(scale=2):
                    up_title = gr.Textbox(
                        label="Title",
                        placeholder="e.g. The Foundation of the Word",
                    )
                    up_presenter = gr.Textbox(
                        label="Presenter",
                        placeholder="e.g. Elder, Group 3",
                    )
                    up_week = gr.Textbox(
                        label="Week of (YYYY-MM-DD)",
                        value=str(date.today()),
                    )
                with gr.Column(scale=1):
                    up_password = gr.Textbox(
                        label="Admin Password",
                        type="password",
                    )
                    up_file = gr.File(
                        label="Document (.docx / .pptx)",
                        file_types=[".docx", ".pptx"],
                    )

            up_button = gr.Button("Upload Study", variant="primary")
            up_status = gr.Markdown()

            up_button.click(
                fn=upload_bible_study,
                inputs=[up_file, up_title, up_presenter, up_week, up_password],
                outputs=up_status,
            )

            gr.Markdown("---\n### Delete a Bible Study")
            with gr.Row():
                del_dropdown = gr.Dropdown(
                    label="Select study to delete",
                    choices=list_bible_studies(),
                    scale=4,
                )
                del_refresh = gr.Button("Refresh list", scale=1, size="sm")
                del_password = gr.Textbox(
                    label="Admin Password",
                    type="password",
                    scale=2,
                )
            del_button = gr.Button("Delete Selected", variant="stop")
            del_status = gr.Markdown()

            del_refresh.click(fn=refresh_studies, outputs=del_dropdown)
            del_button.click(
                fn=delete_bible_study,
                inputs=[del_dropdown, del_password],
                outputs=[del_status, del_dropdown],
            )

    gr.Markdown(
        "<div style='text-align:center; color:#9ca3af; font-size:0.85em; padding:12px;'>"
        "JCC Assistant — Prototype · The bot only answers from loaded study and programs data."
        "</div>"
    )


if __name__ == "__main__":
    demo.launch()