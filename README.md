---
title: JCC Assistant
emoji: 📖
colorFrom: indigo
colorTo: yellow
sdk: gradio
sdk_version: 4.44.0
app_file: app.py
pinned: false
---

# JCC Assistant

A chatbot for **Jubilee Celebration Center — AFM** that answers two kinds of
questions:

1. **Bible Study** — questions about the week's Bible study notes
2. **Church Programs** — questions about JCC's 2026 ministry events, leads,
   visions, and activities

The bot only answers from loaded JCC data — it does not generate generic
Bible commentary or invent church events.

---

## Setup (one-time)

### 1. Supabase
1. Create a new Supabase project (under a separate org from your other projects).
2. Open the SQL Editor.
3. Paste and run `schema.sql`. This creates the tables and seeds the 2026
   ministry programs.
4. From Project Settings → API, copy:
   - `URL` → this is `SUPABASE_URL`
   - `service_role` key → this is `SUPABASE_SERVICE_KEY`

### 2. OpenAI
- Get an API key from https://platform.openai.com/api-keys
- This is `OPENAI_API_KEY`

### 3. Hugging Face Space
1. Create a new Space (SDK: Gradio, Hardware: CPU basic, free).
2. Upload these files: `app.py`, `requirements.txt`, this `README.md`.
3. In Space Settings → Variables and secrets, add:
   - `SUPABASE_URL`
   - `SUPABASE_SERVICE_KEY`
   - `OPENAI_API_KEY`
   - `ADMIN_PASSWORD` — a shared password for uploading Bible studies

### 4. Seed the first Bible study (locally)
```bash
pip install supabase python-docx
export SUPABASE_URL=https://xxxx.supabase.co
export SUPABASE_SERVICE_KEY=eyJ...
python seed_bible_study.py /path/to/Foundation_of_the_Word.docx
```

After this, the Space should load with the Foundation of the Word study
selectable in the dropdown.

---

## Usage

### For group members
1. Open the Space URL.
2. **Chat tab** → pick mode:
   - 📖 **Bible Study** — select a week from the dropdown, ask questions.
   - 📅 **Church Programs** — ask about events, ministries, dates, leads.
3. Try the examples below the chat box.

### For admins (uploading the next week's study)
1. **Admin tab** → fill in title, presenter, week_of, password.
2. Attach the `.docx` file.
3. Click Upload. It appears in the Bible Study dropdown immediately.

---

## What it can answer (Bible Study)

- "What does the study say about Logos and Rhema?"
- "What scriptures are referenced in section 6?"
- "Summarize the main points of this teaching."
- "What was the elder's point about the Word as seed?"

## What it can answer (Church Programs)

- "When is the next Couples Ministry event?"
- "Who leads the Hospitality Ministry?"
- "What's the vision for Praise & Worship?"
- "What fundraising activities are planned for 2026?"
- "What's happening in May 2026?"

## What it won't do

- Answer generic Bible questions outside the study notes
- Invent church events or dates
- Provide theological interpretation beyond what's in the documents

---

## Architecture

```
User → Gradio (HF Space) → Supabase (Postgres) → GPT-4o-mini → User
                          ↑
              docx upload (admin tab)
```

No RAG, no embeddings — the relevant document(s) are injected as full
context into the system prompt. Works because each Bible study is
small enough to fit comfortably in GPT-4o-mini's context window.

## Tech stack

- **UI:** Gradio (Hugging Face Spaces)
- **Database:** Supabase Postgres
- **LLM:** OpenAI GPT-4o-mini
- **Parsing:** python-docx

## Costs

- Supabase: free tier
- HF Spaces: free tier (CPU)
- OpenAI: pay-per-use (typical: $1–3/month for a small group)
