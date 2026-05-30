---
title: JCC Assistant
emoji: 📖
colorFrom: indigo
colorTo: yellow
sdk: gradio
sdk_version: 5.49.1
python_version: 3.11
app_file: app.py
pinned: false
---

# JCC Assistant

A chatbot for Jubilee Celebration Center — AFM. Answers questions about:

1. The current week's Bible study
2. JCC 2026 ministry programs and events

The bot only answers from loaded JCC data — it does not generate generic Bible commentary or invent church events.

## Tech stack

- Gradio (Hugging Face Spaces)
- Supabase Postgres
- OpenAI GPT-4o-mini
- python-docx

## Required secrets (in Space settings)

- SUPABASE_URL
- SUPABASE_SERVICE_KEY
- OPENAI_API_KEY
- ADMIN_PASSWORD
