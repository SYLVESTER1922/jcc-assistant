-- =====================================================================
-- JCC Assistant -- Database Schema
-- Office: Jubilee Celebration Center (AFM)
-- Stack: Supabase (PostgreSQL)
-- =====================================================================

CREATE EXTENSION IF NOT EXISTS "pgcrypto";

-- =====================================================================
-- TABLES
-- =====================================================================

CREATE TABLE IF NOT EXISTS bible_studies (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  week_of DATE NOT NULL,
  title TEXT NOT NULL,
  presenter TEXT,
  document_text TEXT NOT NULL,
  suggested_questions JSONB DEFAULT '[]'::jsonb,
  uploaded_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
CREATE INDEX IF NOT EXISTS idx_bible_studies_week_of ON bible_studies (week_of DESC);

CREATE TABLE IF NOT EXISTS ministries (
  id SERIAL PRIMARY KEY,
  name TEXT UNIQUE NOT NULL,
  lead TEXT,
  description TEXT
);

CREATE TABLE IF NOT EXISTS events (
  id SERIAL PRIMARY KEY,
  ministry_id INT REFERENCES ministries(id) ON DELETE CASCADE,
  event_date DATE,
  date_label TEXT,
  quarter INT,
  title TEXT NOT NULL,
  description TEXT,
  format TEXT,
  status TEXT DEFAULT 'planned'
);
CREATE INDEX IF NOT EXISTS idx_events_date ON events (event_date);
CREATE INDEX IF NOT EXISTS idx_events_ministry ON events (ministry_id);

CREATE TABLE IF NOT EXISTS ministry_notes (
  id SERIAL PRIMARY KEY,
  ministry_id INT REFERENCES ministries(id) ON DELETE CASCADE,
  section TEXT NOT NULL,
  content TEXT NOT NULL
);