"""
JCC Assistant — Bible Study & Church Programs Chatbot
======================================================
A Gradio app for the Jubilee Celebration Center — AFM.

Two modes:
  1. Bible Study — answers questions about a selected week's study,
     quoting only from the elder's notes.
  2. Church Programs — answers questions about JCC ministry events,
     leads, goals, and activities for 2026.

Stack: Gradio + Supabase Postgres + OpenAI (GPT-4o-mini)
Deploy: Hugging Face Spaces (Gradio SDK)

Required Space secrets:
  SUPABASE_URL              e.g. https://xxxx.supabase.co
  SUPABASE_SERVICE_KEY      service_role key (Settings → API)
  OPENAI_API_KEY            sk-...
  ADMIN_PASSWORD            shared password for uploading new studies
"""

import os
import io
import json
from datetime import date, datetime
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
# Bible Study helpers
# ---------------------------------------------------------------------------
def list_bible_studies() -> List[Tuple[str, str]]:
    """Return list of (label, id) for the dropdown, newest first."""
    res = sb.table("bible_studies") \
        .select("id, week_of, title, presenter") \
        .order("week_of", desc=True) \
        .execute()
    options = []
    for row in res.data:
        label = f"{row['week_of']} — {row['title']}"
        if row.get("presenter"):
            label += f" ({row['presenter']})"
        options.append((label, row["id"]))
    return options


def get_bible_study(study_id: str) -> Optional[dict]:
    """Fetch a single Bible study row."""
    if not study_id:
        return None
    res = sb.table("bible_studies").select("*").eq("id", study_id).single().execute()
    return res.data


def parse_docx(file_path: str) -> str:
    """Extract plain text from a .docx file, preserving paragraph breaks."""
    doc = Document(file_path)
    paragraphs = []
    for p in doc.paragraphs:
        txt = p.text.strip()
        if txt:
            paragraphs.append(txt)
    return "\n\n".join(paragraphs)


def parse_pptx(file_path: str) -> str:
    """Extract plain text from a .pptx file, slide by slide."""
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
            elif shape.shape_type == 19 and shape.has_table:  # table
                for row in shape.table.rows:
                    row_text = " | ".join(cell.text.strip() for cell in row.cells)
                    if row_text.strip():
                        slide_lines.append(row_text)
        if len(slide_lines) > 1:
            sections.append("\n".join(slide_lines))
    return "\n\n".join(sections)


def parse_study_document(file_path: str) -> str:
    """Route to the right parser based on file extension."""
    lower = file_path.lower()
    if lower.endswith(".docx"):
        return parse_docx(file_path)
    elif lower.endswith(".pptx"):
        return parse_pptx(file_path)
    else:
        raise ValueError(
            f"Unsupported file type: {file_path}. "
            "Please upload a .docx or .pptx file."
        )


def upload_bible_study(
    file, title: str, presenter: str, week_of: str, password: str
) -> str:
    """Admin: parse a .docx and insert a new Bible study row."""
    if password != ADMIN_PASSWORD:
        return "❌ Incorrect admin password."
    if not file:
        return "❌ Please attach a .docx or .pptx file."
    if not title.strip():
        return "❌ Title is required."
    if not week_of:
        return "❌ Week-of date is required."

    try:
        text = parse_study_document(file.name)
    except Exception as e:
        return f"❌ Could not parse the document: {e}"

    if len(text) < 50:
        return "❌ The document looks empty after parsing."

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


# ---------------------------------------------------------------------------
# Church Programs helpers
# ---------------------------------------------------------------------------
def fetch_programs_context() -> str:
    """Build the full church-programs context string for the LLM."""
    ministries = sb.table("ministries").select("*").execute().data
    events = sb.table("events").select("*").order("event_date").execute().data
    notes = sb.table("ministry_notes").select("*").execute().data

    # Build a per-ministry lookup
    by_id = {m["id"]: m for m in ministries}

    lines = ["# JCC 2026 MINISTRY PROGRAMS\n"]

    for m in ministries:
        lines.append(f"\n## {m['name']} Ministry")
        if m.get("lead"):
            lines.append(f"Led by: {m['lead']}")

        # Notes for this ministry
        m_notes = [n for n in notes if n["ministry_id"] == m["id"]]
        for n in m_notes:
            lines.append(f"\n**{n['section'].replace('_', ' ').title()}:** {n['content']}")

        # Events for this ministry
        m_events = [e for e in events if e["ministry_id"] == m["id"]]
        if m_events:
            lines.append("\n**Events:**")
            for e in m_events:
                date_str = e.get("date_label") or e.get("event_date") or "TBD"
                line = f"- {date_str}: {e['title']}"
                if e.get("description"):
                    line += f" — {e['description']}"
                if e.get("format"):
                    line += f" ({e['format']})"
                lines.append(line)

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# System prompts
# ---------------------------------------------------------------------------
BIBLE_STUDY_PROMPT = """You are the JCC Bible Study Assistant for Jubilee Celebration Center — AFM.

You answer questions about a specific Bible study document shared with the group. Your answers MUST come only from the document below — do not add interpretation, outside scripture, or commentary.

RULES:
- Quote directly from the document using quotation marks for the key passage.
- Cite the section heading or number where the quote appears (e.g., "Section 7: Logos and Rhema").
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

PROGRAMS_PROMPT = """You are the JCC Programs Assistant for Jubilee Celebration Center — AFM.

You answer questions about JCC's 2026 ministry programs, events, leads, goals, and activities. Your answers MUST come only from the programs information below.

RULES:
- Be specific. When asked about events, give the date, ministry, and format.
- When asked about a ministry's vision, mission, or goals, summarize from the notes.
- If the question is not covered in the data below, say:
  "I don't have that information in the 2026 plans. Please check with the relevant ministry lead."
- Do not invent events, dates, or leads. Stick to the data.
- Keep answers focused and helpful.

---
JCC 2026 PROGRAMS DATA
{programs_context}
---
"""


# ---------------------------------------------------------------------------
# Chat function
# ---------------------------------------------------------------------------
def chat(message: str, history: list, mode: str, study_id: str):
    """Main chat handler. Routes by mode."""
    if not message or not message.strip():
        return "Please type a question."

    if mode == "📖 Bible Study":
        study = get_bible_study(study_id) if study_id else None
        if not study:
            return (
                "Please select a Bible study from the dropdown first. "
                "If the dropdown is empty, an admin needs to upload a study."
            )
        system_prompt = BIBLE_STUDY_PROMPT.format(
            title=study["title"],
            presenter=study.get("presenter") or "Not specified",
            week_of=study["week_of"],
            document_text=study["document_text"],
        )
    else:  # Church Programs
        try:
            context = fetch_programs_context()
        except Exception as e:
            return f"Could not load programs data: {e}"
        system_prompt = PROGRAMS_PROMPT.format(programs_context=context)

    # Build messages
    messages = [{"role": "system", "content": system_prompt}]
    for h in history:
        messages.append({"role": h["role"], "content": h["content"]})
    messages.append({"role": "user", "content": message})

    try:
        resp = oai.chat.completions.create(
            model=OPENAI_MODEL,
            messages=messages,
            temperature=0.2,  # low — we want faithful retrieval, not creativity
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


with gr.Blocks(
    title="JCC Assistant",
    theme=gr.themes.Soft(primary_hue="amber", neutral_hue="slate"),
) as demo:

    gr.Markdown(
        """
        # JCC Assistant
        ### Jubilee Celebration Center — AFM
        Ask about this week's Bible study or about church programs and events for 2026.
        """
    )

    with gr.Tabs():
        # ============================================================
        # CHAT TAB
        # ============================================================
        with gr.Tab("💬 Chat"):
            with gr.Row():
                mode = gr.Radio(
                    choices=["📖 Bible Study", "📅 Church Programs"],
                    value="📖 Bible Study",
                    label="Mode",
                    scale=1,
                )
                study_dropdown = gr.Dropdown(
                    label="Select a Bible Study",
                    choices=list_bible_studies(),
                    scale=2,
                )
                refresh_btn = gr.Button("🔄", scale=0, size="sm")

            refresh_btn.click(fn=refresh_studies, outputs=study_dropdown)

            chatbot = gr.ChatInterface(
                fn=chat,
                additional_inputs=[mode, study_dropdown],
                type="messages",
                examples=[
                    ["What does the study say about Logos and Rhema?"],
                    ["What are the main points of this teaching?"],
                    ["What scriptures are referenced in section 6?"],
                    ["When is the next Couples Ministry event?"],
                    ["What's the vision for Praise & Worship?"],
                    ["What fundraising activities are planned for 2026?"],
                ],
            )

        # ============================================================
        # ADMIN TAB
        # ============================================================
        with gr.Tab("⚙️ Admin — Upload Bible Study"):
            gr.Markdown(
                "Upload a new Bible study document. Requires the admin password."
            )
            with gr.Row():
                up_title = gr.Textbox(label="Title", placeholder="e.g. The Foundation of the Word")
                up_presenter = gr.Textbox(label="Presenter", placeholder="e.g. Elder John")
            with gr.Row():
                up_week = gr.Textbox(
                    label="Week of (YYYY-MM-DD)",
                    placeholder=str(date.today()),
                    value=str(date.today()),
                )
                up_password = gr.Textbox(label="Admin Password", type="password")
            up_file = gr.File(
                label="Bible Study Document (.docx or .pptx)",
                file_types=[".docx", ".pptx"],
            )
            up_button = gr.Button("Upload", variant="primary")
            up_status = gr.Markdown()

            up_button.click(
                fn=upload_bible_study,
                inputs=[up_file, up_title, up_presenter, up_week, up_password],
                outputs=up_status,
            )

    gr.Markdown(
        """
        ---
        *JCC Assistant — Prototype.* The bot only answers from the loaded
        Bible study or JCC 2026 program data. For anything else, please
        consult your ministry lead or study presenter.
        """
    )


if __name__ == "__main__":
    demo.launch()
