"""
LOBELS BISCUITS - STORES AI ASSISTANT
======================================
Netrisyl Insights.
Design replicated from JCC Assistant (navy + gold, Inter font, hero + sidebar cards).
Stack: Gradio + Supabase Postgres + GPT-4o-mini.
Pattern: GPT tool-calling returns query params, Python runs the SQL.
"""

import os
import json
import base64
from pathlib import Path
from datetime import date

import gradio as gr
from supabase import create_client, Client
from openai import OpenAI


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
SUPABASE_URL   = os.environ["SUPABASE_URL"]
SUPABASE_KEY   = os.environ["SUPABASE_SERVICE_KEY"]
OPENAI_API_KEY = os.environ["OPENAI_API_KEY"]

OPENAI_MODEL = "gpt-4o-mini"

sb: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
oai = OpenAI(api_key=OPENAI_API_KEY)

CLIENT_ID = "lobels"
MONTH_ORDER = ["JAN","FEB","MAR","APR","MAY","JUN",
               "JUL","AUG","SEP","OCT","NOV","DEC"]


# ---------------------------------------------------------------------------
# Logos as base64
# ---------------------------------------------------------------------------
LOGO_PATH = Path(__file__).parent / "logo.png"
if LOGO_PATH.exists():
    with open(LOGO_PATH, "rb") as f:
        LOGO_B64 = base64.b64encode(f.read()).decode("ascii")
    LOGO_DATA_URI = f"data:image/png;base64,{LOGO_B64}"
else:
    LOGO_DATA_URI = ""

NETRISYL_LOGO_PATH = Path(__file__).parent / "netrisyl_logo.png"
if NETRISYL_LOGO_PATH.exists():
    with open(NETRISYL_LOGO_PATH, "rb") as f:
        NETRISYL_B64 = base64.b64encode(f.read()).decode("ascii")
    NETRISYL_DATA_URI = f"data:image/png;base64,{NETRISYL_B64}"
else:
    NETRISYL_DATA_URI = ""


# ---------------------------------------------------------------------------
# Data access (Python runs the queries)
# NOTE: in DATA EXPORT, monthly values (variance, opening, closing, total)
# repeat on every daily row. daily_issues is genuinely per-day so SUM is right;
# variance must take ONE value per material per month (not summed).
# ---------------------------------------------------------------------------
def _num(v):
    try:
        return float(v or 0)
    except Exception:
        return 0.0


def q_material_total(description=None, month=None):
    qry = sb.table("lobels_stores").select(
        "description, daily_issues, month").eq("client_id", CLIENT_ID)
    if description:
        qry = qry.ilike("description", f"%{description}%")
    if month:
        qry = qry.eq("month", month.upper())
    rows = qry.execute().data
    if not rows:
        return {"found": False, "searched_for": description}
    total = sum(_num(r["daily_issues"]) for r in rows)
    return {"found": True, "material": rows[0]["description"],
            "total_issues_kg": round(total, 2),
            "period": month.upper() if month else "all months (Jan-Jun 2026)"}


def q_top_materials(month=None, limit=10):
    qry = sb.table("lobels_stores").select(
        "description, daily_issues, month").eq("client_id", CLIENT_ID)
    if month:
        qry = qry.eq("month", month.upper())
    rows = qry.execute().data
    agg = {}
    for r in rows:
        agg[r["description"]] = agg.get(r["description"], 0) + _num(r["daily_issues"])
    ranked = sorted(agg.items(), key=lambda x: -x[1])[:limit]
    return {"period": month.upper() if month else "all months",
            "top": [{"material": m, "issues_kg": round(v, 2)} for m, v in ranked]}


def q_variances(month=None, flag="LOSS", limit=10):
    qry = sb.table("lobels_stores").select(
        "description, variance, month").eq("client_id", CLIENT_ID)
    if month:
        qry = qry.eq("month", month.upper())
    rows = qry.execute().data
    seen = {}
    for r in rows:
        seen[(r["description"], r["month"])] = _num(r["variance"])
    agg = {}
    for (mat, mth), var in seen.items():
        agg[mat] = agg.get(mat, 0) + var
    if flag == "LOSS":
        ranked = sorted(agg.items(), key=lambda x: x[1])[:limit]
    else:
        ranked = sorted(agg.items(), key=lambda x: -x[1])[:limit]
    return {"period": month.upper() if month else "all months", "type": flag,
            "items": [{"material": m, "variance_kg": round(v, 2)} for m, v in ranked]}


def q_category_breakdown(month=None):
    qry = sb.table("lobels_stores").select(
        "category, daily_issues, month").eq("client_id", CLIENT_ID)
    if month:
        qry = qry.eq("month", month.upper())
    rows = qry.execute().data
    agg = {}
    for r in rows:
        cat = r["category"] or "Other"
        agg[cat] = agg.get(cat, 0) + _num(r["daily_issues"])
    ranked = sorted(agg.items(), key=lambda x: -x[1])
    return {"period": month.upper() if month else "all months",
            "categories": [{"category": c, "issues_kg": round(v, 2)} for c, v in ranked]}


def q_monthly_trend(description):
    rows = sb.table("lobels_stores").select(
        "description, daily_issues, month").eq("client_id", CLIENT_ID
        ).ilike("description", f"%{description}%").execute().data
    if not rows:
        return {"found": False, "searched_for": description}
    agg = {}
    for r in rows:
        agg[r["month"]] = agg.get(r["month"], 0) + _num(r["daily_issues"])
    ordered = [{"month": m, "issues_kg": round(agg.get(m, 0), 2)}
               for m in MONTH_ORDER if m in agg]
    return {"found": True, "material": rows[0]["description"], "trend": ordered}


# ---------------------------------------------------------------------------
# GPT tool definitions
# ---------------------------------------------------------------------------
TOOLS = [
    {"type": "function", "function": {
        "name": "q_material_total",
        "description": "Total kg issued for a specific raw material, optionally for one month. Use a SHORT distinctive part of the name e.g. 'National Foods', 'Sugar', 'Palm Oil', 'Hex Flour' - the search matches partially.",
        "parameters": {"type": "object", "properties": {
            "description": {"type": "string", "description": "Distinctive part of material name. Prefer short fragments e.g. 'National Foods' not 'Flour National Foods Industrial'."},
            "month": {"type": "string", "description": "Three-letter month e.g. JAN. Omit for full year."}
        }, "required": ["description"]}}},
    {"type": "function", "function": {
        "name": "q_top_materials",
        "description": "Top materials ranked by total kg issued.",
        "parameters": {"type": "object", "properties": {
            "month": {"type": "string"}, "limit": {"type": "integer"}}}}},
    {"type": "function", "function": {
        "name": "q_variances",
        "description": "Materials with biggest stock losses or gains.",
        "parameters": {"type": "object", "properties": {
            "month": {"type": "string"},
            "flag": {"type": "string", "enum": ["LOSS", "GAIN"]},
            "limit": {"type": "integer"}}}}},
    {"type": "function", "function": {
        "name": "q_category_breakdown",
        "description": "Total issues grouped by material category.",
        "parameters": {"type": "object", "properties": {
            "month": {"type": "string"}}}}},
    {"type": "function", "function": {
        "name": "q_monthly_trend",
        "description": "Month-by-month consumption trend for one material.",
        "parameters": {"type": "object", "properties": {
            "description": {"type": "string", "description": "Distinctive part of material name."}
        }, "required": ["description"]}}},
]

FUNC_MAP = {
    "q_material_total": q_material_total,
    "q_top_materials": q_top_materials,
    "q_variances": q_variances,
    "q_category_breakdown": q_category_breakdown,
    "q_monthly_trend": q_monthly_trend,
}

TODAY = date.today().strftime("%d %B %Y")
SYSTEM_PROMPT = f"""You are the Lobels Biscuits Stores AI Assistant, built by Netrisyl Insights.
Today's date is {TODAY}.
You answer questions about raw material stores data: consumption, variances, and trends.
You have data for January to June 2026 across 88 raw materials.

Rules:
- ALWAYS use a tool to fetch real figures. NEVER invent numbers.
- When searching for a material, pass a SHORT distinctive fragment of its name
  (e.g. 'National Foods', 'Sugar', 'Palm Oil') so partial matching works.
- Quote figures exactly as returned by the tools, in kilograms (kg).
- Be concise and professional, in plain language a stores manager understands.
- If a question is outside stores data (HR, finance, etc.), politely say it's out of scope.
- Months available: JAN, FEB, MAR, APR, MAY, JUN (2026).
- If no data is found for a material, say so clearly and suggest checking the name.
"""


def chat_answer(message, history):
    messages = [{"role": "system", "content": SYSTEM_PROMPT}]
    for h in (history or []):
        if isinstance(h, dict) and h.get("role") in ("user", "assistant"):
            messages.append({"role": h["role"], "content": h["content"]})
    messages.append({"role": "user", "content": message})

    resp = oai.chat.completions.create(
        model=OPENAI_MODEL, messages=messages,
        tools=TOOLS, tool_choice="auto", temperature=0)
    msg = resp.choices[0].message

    if msg.tool_calls:
        messages.append({"role": "assistant", "content": msg.content or "",
                         "tool_calls": [tc.model_dump() for tc in msg.tool_calls]})
        for tc in msg.tool_calls:
            fn = FUNC_MAP.get(tc.function.name)
            try:
                args = json.loads(tc.function.arguments or "{}")
                result = fn(**args) if fn else {"error": "unknown tool"}
            except Exception as e:
                result = {"error": str(e)}
            messages.append({"role": "tool", "tool_call_id": tc.id,
                             "content": json.dumps(result)})
        final = oai.chat.completions.create(
            model=OPENAI_MODEL, messages=messages, temperature=0)
        return final.choices[0].message.content
    return msg.content or "I couldn't find an answer to that."


SUGGESTED = [
    "Which materials had the highest losses?",
    "How much Flour National Foods did we use in January?",
    "Top 10 materials by consumption",
    "Compare sugar use across months",
    "Which materials gained stock (surplus)?",
    "Break down consumption by category",
]


# ---------------------------------------------------------------------------
# UI  (replicated from JCC: navy + gold, Inter, hero + sidebar cards)
# ---------------------------------------------------------------------------
CUSTOM_CSS = """
.gradio-container {
    font-family: 'Inter', 'Helvetica Neue', system-ui, sans-serif !important;
    max-width: 1400px !important;
    margin: 0 auto !important;
}
#lobels-hero {
    background: linear-gradient(135deg, #1B2A4E 0%, #2C4170 100%);
    border-radius: 16px;
    padding: 28px 32px;
    margin-bottom: 18px;
    color: white;
    display: flex;
    align-items: center;
    gap: 24px;
    box-shadow: 0 8px 24px rgba(27, 42, 78, 0.18);
    position: relative;
    overflow: hidden;
}
#lobels-hero::after {
    content: "";
    position: absolute;
    bottom: 0; left: 0; right: 0;
    height: 4px;
    background: linear-gradient(90deg, #C9A55C 0%, #E4CC8E 50%, #C9A55C 100%);
}
#lobels-hero img.logo {
    width: 88px;
    height: 88px;
    border-radius: 50%;
    background: white;
    padding: 6px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.2);
    flex-shrink: 0;
    object-fit: contain;
}
#lobels-hero .titles h1 {
    font-size: 1.9em !important;
    font-weight: 700 !important;
    margin: 0 0 4px 0 !important;
    color: white !important;
    letter-spacing: -0.5px;
}
#lobels-hero .titles .brand-name {
    font-size: 0.85em;
    color: #C9A55C;
    letter-spacing: 3px;
    font-weight: 600;
    margin-bottom: 6px;
    text-transform: uppercase;
}
#lobels-hero .titles .tagline {
    font-size: 0.95em;
    color: #cbd5e1;
    margin: 0;
}
.sidebar-card {
    background: white;
    border-radius: 12px;
    padding: 16px;
    border: 1px solid #e5e7eb;
    margin-bottom: 12px;
}
.sidebar-card h3 {
    color: #1B2A4E;
    font-size: 0.78em;
    letter-spacing: 1.5px;
    text-transform: uppercase;
    margin: 0 0 12px 0;
    font-weight: 700;
    border-left: 3px solid #C9A55C;
    padding-left: 10px;
}
.suggest-btn button {
    background: white !important;
    border: 1px solid #e5e7eb !important;
    color: #1B2A4E !important;
    text-align: left !important;
    font-weight: 500 !important;
    font-size: 0.88em !important;
    padding: 10px 12px !important;
    line-height: 1.35 !important;
    white-space: normal !important;
    height: auto !important;
    min-height: 40px !important;
    transition: all 0.15s ease;
    width: 100% !important;
    justify-content: flex-start !important;
}
.suggest-btn button:hover {
    background: #1B2A4E !important;
    color: white !important;
    border-color: #1B2A4E !important;
    transform: translateX(2px);
}
#info-panel {
    background: #fefcf7;
    border: 1px solid #e8dfc7;
    border-radius: 12px;
    padding: 18px 20px;
    font-size: 0.92em;
    line-height: 1.55;
    color: #1f2937;
}
#info-panel h3 { color: #1B2A4E; margin-top: 0; }
#netrisyl-footer {
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 6px;
    padding: 22px 12px 14px 12px;
    margin-top: 8px;
    border-top: 1px solid #e8e1cf;
}
#netrisyl-footer .prototype-note {
    color: #9ca3af;
    font-size: 0.8em;
    text-align: center;
    margin: 0;
}
#netrisyl-footer .powered-row {
    display: flex;
    align-items: center;
    gap: 12px;
}
#netrisyl-footer .powered-row .label {
    font-size: 0.85em;
    color: #6b7280;
    letter-spacing: 2px;
    text-transform: uppercase;
    font-weight: 600;
}
#netrisyl-footer .powered-row img {
    height: 64px;
    width: auto;
    display: block;
}
footer { display: none !important; }
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


with gr.Blocks(title="Lobels Stores AI Assistant", theme=theme, css=CUSTOM_CSS) as demo:

    logo_img_html = (
        f'<img class="logo" src="{LOGO_DATA_URI}" alt="Lobels Logo"/>'
        if LOGO_DATA_URI else ""
    )
    gr.HTML(f"""
    <div id="lobels-hero">
        {logo_img_html}
        <div class="titles">
            <div class="brand-name">Lobels Biscuits &amp; Sweets</div>
            <h1>Stores AI Assistant</h1>
            <p class="tagline">Ask about raw material consumption, stock variances and trends for 2026.</p>
        </div>
    </div>
    """)

    with gr.Row():
        # LEFT SIDEBAR
        with gr.Column(scale=1, min_width=240):
            with gr.Group(elem_classes=["sidebar-card"]):
                gr.HTML("<h3>Suggested Questions</h3>")
                suggest_btns = [
                    gr.Button(q, elem_classes=["suggest-btn"])
                    for q in SUGGESTED
                ]

        # MAIN CHAT
        with gr.Column(scale=2):
            chatbot = gr.Chatbot(
                type="messages",
                height=560,
                avatar_images=(None, str(LOGO_PATH) if LOGO_PATH.exists() else None),
                show_label=False,
                show_copy_button=True,
            )
            msg = gr.Textbox(
                placeholder="Ask about materials, variances, trends...",
                show_label=False,
                container=False,
                autofocus=True,
            )

        # RIGHT INFO PANEL
        with gr.Column(scale=2):
            with gr.Group(elem_classes=["sidebar-card"]):
                gr.HTML("<h3>About this assistant</h3>")
                gr.Markdown(
                    "**Data:** 88 raw materials, Jan–Jun 2026.\n\n"
                    "**I can answer:**\n"
                    "- Material consumption (daily & monthly)\n"
                    "- Stock variances — losses & gains\n"
                    "- Top materials & category breakdowns\n"
                    "- Month-by-month trends\n\n"
                    "All figures come directly from the stores register. "
                    "I never invent numbers.",
                    elem_id="info-panel",
                )

    # ---------- chat plumbing ----------
    def respond(message, history):
        if not message or not message.strip():
            return "", history or []
        history = history or []
        reply = chat_answer(message, history)
        history.append({"role": "user", "content": message})
        history.append({"role": "assistant", "content": reply})
        return "", history

    msg.submit(respond, [msg, chatbot], [msg, chatbot])

    for btn, q in zip(suggest_btns, SUGGESTED):
        btn.click(lambda x=q: x, outputs=msg).then(
            respond, [msg, chatbot], [msg, chatbot])

    netrisyl_img_html = (
        f'<img src="{NETRISYL_DATA_URI}" alt="Netrisyl Insights"/>'
        if NETRISYL_DATA_URI else '<span class="label" style="color:#C9A55C;">Netrisyl Insights</span>'
    )
    gr.HTML(f"""
    <div id="netrisyl-footer">
        <p class="prototype-note">Lobels Stores Assistant &mdash; Prototype. The bot only answers from loaded stores data.</p>
        <div class="powered-row">
            <span class="label">Powered by</span>
            {netrisyl_img_html}
        </div>
    </div>
    """)


if __name__ == "__main__":
    demo.launch()
