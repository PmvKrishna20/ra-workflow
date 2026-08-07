-- RA Workflow — Supabase schema
-- Run this once in the Supabase SQL Editor (supabase.com → your project → SQL Editor)

CREATE TABLE IF NOT EXISTS companies (
    id SERIAL PRIMARY KEY,
    company_name TEXT NOT NULL,
    position TEXT,
    added_by TEXT,
    date_added TEXT,
    source TEXT
);
CREATE INDEX IF NOT EXISTS idx_companies_name_date ON companies(company_name, date_added);

CREATE TABLE IF NOT EXISTS block_list (
    id SERIAL PRIMARY KEY,
    keyword TEXT NOT NULL UNIQUE
);

CREATE TABLE IF NOT EXISTS clients (
    id SERIAL PRIMARY KEY,
    company_name TEXT NOT NULL UNIQUE
);

CREATE TABLE IF NOT EXISTS avoid_list (
    id SERIAL PRIMARY KEY,
    company_name TEXT NOT NULL UNIQUE,
    reason TEXT
);

CREATE TABLE IF NOT EXISTS eligible_companies (
    id SERIAL PRIMARY KEY,
    company_name TEXT NOT NULL UNIQUE,
    employee_size TEXT,
    industry TEXT
);

CREATE TABLE IF NOT EXISTS not_eligible_companies (
    id SERIAL PRIMARY KEY,
    company_name TEXT NOT NULL UNIQUE,
    employee_size TEXT,
    industry TEXT,
    reason TEXT
);

CREATE TABLE IF NOT EXISTS needs_review_companies (
    id SERIAL PRIMARY KEY,
    company_name TEXT NOT NULL UNIQUE,
    employee_size TEXT,
    industry TEXT,
    reason TEXT
);

CREATE TABLE IF NOT EXISTS settings (
    key TEXT PRIMARY KEY,
    value TEXT
);

CREATE TABLE IF NOT EXISTS ra_assignments (
    id SERIAL PRIMARY KEY,
    batch_id TEXT NOT NULL,
    company_name TEXT NOT NULL,
    job_title TEXT,
    location TEXT,
    job_url TEXT,
    ra_name TEXT NOT NULL,
    assigned_date TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_assignments_batch ON ra_assignments(batch_id);

CREATE TABLE IF NOT EXISTS prospects (
    id SERIAL PRIMARY KEY,
    first_name TEXT,
    email TEXT,
    company_name TEXT NOT NULL,
    company_key TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_prospects_company_key ON prospects(company_key);

CREATE TABLE IF NOT EXISTS bounced_emails (
    id SERIAL PRIMARY KEY,
    email TEXT NOT NULL UNIQUE
);

CREATE TABLE IF NOT EXISTS eod_submissions (
    id SERIAL PRIMARY KEY,
    ra_name TEXT NOT NULL,
    upload_date TEXT NOT NULL,
    company_name TEXT,
    company_linkedin_url TEXT,
    full_name TEXT,
    first_name TEXT,
    poc_location TEXT,
    designation TEXT,
    email TEXT,
    position TEXT,
    location TEXT,
    job_posting_link TEXT,
    industry TEXT
);
CREATE INDEX IF NOT EXISTS idx_eod_ra_date ON eod_submissions(ra_name, upload_date);

CREATE TABLE IF NOT EXISTS emails_sent (
    id SERIAL PRIMARY KEY,
    first_name TEXT,
    email TEXT,
    position TEXT,
    location TEXT,
    company_name TEXT,
    ra_name TEXT,
    period TEXT
);
CREATE INDEX IF NOT EXISTS idx_emails_sent_email_company ON emails_sent(email, company_name);

CREATE TABLE IF NOT EXISTS positive_responses (
    id SERIAL PRIMARY KEY,
    email TEXT,
    name TEXT,
    response_date TEXT,
    ra_name TEXT,
    position TEXT,
    designation TEXT,
    company_name TEXT
);
CREATE INDEX IF NOT EXISTS idx_responses_email_company ON positive_responses(email, company_name);

CREATE TABLE IF NOT EXISTS title_bucket_keywords (
    id SERIAL PRIMARY KEY,
    keyword TEXT NOT NULL UNIQUE
);

CREATE TABLE IF NOT EXISTS ra_assignments (
    id SERIAL PRIMARY KEY,
    batch_id TEXT NOT NULL,
    company_name TEXT NOT NULL,
    job_title TEXT,
    location TEXT,
    job_url TEXT,
    ra_name TEXT NOT NULL,
    assigned_date TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS users (
    id SERIAL PRIMARY KEY,
    username TEXT NOT NULL UNIQUE,
    password_hash TEXT NOT NULL,
    display_name TEXT NOT NULL,
    role TEXT NOT NULL DEFAULT 'RA',
    active INTEGER NOT NULL DEFAULT 1,
    created_at TEXT
);
-- Note: the app also creates this table automatically on first run (init_db()),
-- and seeds the first Manager account from the [bootstrap_admin] secret.
-- Running this file manually is optional — only needed if you want the table
-- to exist before the app's first launch.

-- Seed default title bucket keywords
INSERT INTO title_bucket_keywords (keyword) VALUES
    ('Project Manager'),('Program Manager'),('Product Manager'),('Account Manager'),
    ('Office Manager'),('General Manager'),('Operations Manager'),('Plant Manager'),
    ('Warehouse Manager'),('Sales Manager'),('Manager'),('Director'),('Vice President'),
    ('President'),('Supervisor'),('Superintendent'),('Coordinator'),('Administrator'),
    ('Specialist'),('Analyst'),('Consultant'),('Controller'),('Bookkeeper'),
    ('Accountant'),('Accounting'),('Estimator'),('Scheduler'),('Planner'),('Buyer'),
    ('Recruiter'),('Engineer'),('Engineering'),('Technician'),('Designer'),('Architect'),
    ('Inspector'),('Operator'),('Machinist'),('Welder'),('Electrician'),('Plumber'),
    ('Mechanic'),('Driver'),('Foreman'),('Executive'),('Officer'),
    ('CFO'),('CEO'),('COO'),('CTO'),('VP')
ON CONFLICT (keyword) DO NOTHING;

-- Workpage, persistent sessions, approvals, exports, and audit history.
ALTER TABLE users ADD COLUMN IF NOT EXISTS work_date_preference TEXT;

CREATE TABLE IF NOT EXISTS sessions (
    token_hash TEXT PRIMARY KEY,
    user_id INTEGER NOT NULL,
    created_at TEXT NOT NULL,
    expires_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS positions (
    id SERIAL PRIMARY KEY,
    company_name TEXT NOT NULL,
    job_title TEXT,
    status TEXT NOT NULL DEFAULT 'Open',
    assigned_ra TEXT,
    source TEXT,
    batch_id TEXT,
    location TEXT,
    job_url TEXT,
    created_by TEXT,
    created_at TEXT,
    updated_at TEXT,
    master_company_id INTEGER,
    ready_work_date TEXT,
    ready_at TEXT,
    ready_by TEXT,
    export_batch_id TEXT,
    exported_at TEXT,
    exported_by TEXT
);

CREATE TABLE IF NOT EXISTS position_prospects (
    id SERIAL PRIMARY KEY,
    position_id INTEGER NOT NULL,
    full_name TEXT,
    first_name TEXT,
    email TEXT,
    designation TEXT,
    location TEXT,
    linkedin_url TEXT,
    status TEXT NOT NULL DEFAULT 'Not contacted',
    notes TEXT,
    added_by TEXT,
    created_at TEXT,
    updated_at TEXT
);

ALTER TABLE positions ADD COLUMN IF NOT EXISTS master_company_id INTEGER;
ALTER TABLE positions ADD COLUMN IF NOT EXISTS ready_work_date TEXT;
ALTER TABLE positions ADD COLUMN IF NOT EXISTS ready_at TEXT;
ALTER TABLE positions ADD COLUMN IF NOT EXISTS ready_by TEXT;
ALTER TABLE positions ADD COLUMN IF NOT EXISTS export_batch_id TEXT;
ALTER TABLE positions ADD COLUMN IF NOT EXISTS exported_at TEXT;
ALTER TABLE positions ADD COLUMN IF NOT EXISTS exported_by TEXT;

CREATE INDEX IF NOT EXISTS idx_positions_assigned_ra ON positions(assigned_ra);
CREATE INDEX IF NOT EXISTS idx_positions_ready_work_date ON positions(ready_work_date);
CREATE INDEX IF NOT EXISTS idx_position_prospects_position ON position_prospects(position_id);
DROP INDEX IF EXISTS idx_positions_one_active_per_company;
CREATE UNIQUE INDEX idx_positions_one_active_per_company
    ON positions(company_name) WHERE status <> 'Exported';

CREATE TABLE IF NOT EXISTS not_eligible_requests (
    id SERIAL PRIMARY KEY,
    company_name TEXT NOT NULL,
    reason TEXT NOT NULL,
    requested_by TEXT NOT NULL,
    requested_by_user_id INTEGER,
    requested_at TEXT NOT NULL,
    snapshot_json TEXT NOT NULL,
    status TEXT NOT NULL DEFAULT 'Pending',
    reviewed_by TEXT,
    reviewed_at TEXT,
    decision_note TEXT
);

CREATE TABLE IF NOT EXISTS reassignment_queue (
    id SERIAL PRIMARY KEY,
    company_name TEXT NOT NULL UNIQUE,
    job_title TEXT,
    location TEXT,
    job_url TEXT,
    source_reason TEXT NOT NULL,
    created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS deletion_audit (
    id SERIAL PRIMARY KEY,
    position_id INTEGER,
    company_name TEXT NOT NULL,
    position_snapshot_json TEXT NOT NULL,
    prospects_snapshot_json TEXT NOT NULL,
    export_batch_id TEXT,
    deleted_by TEXT NOT NULL,
    deleted_by_role TEXT NOT NULL,
    deletion_reason TEXT NOT NULL,
    deleted_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS export_batches (
    batch_id TEXT PRIMARY KEY,
    created_by TEXT NOT NULL,
    created_at TEXT NOT NULL,
    position_count INTEGER NOT NULL,
    prospect_count INTEGER NOT NULL
);

CREATE TABLE IF NOT EXISTS export_batch_items (
    id SERIAL PRIMARY KEY,
    batch_id TEXT NOT NULL,
    position_id INTEGER NOT NULL,
    company_name TEXT NOT NULL,
    ra_name TEXT,
    work_date TEXT,
    first_name TEXT NOT NULL,
    email TEXT NOT NULL,
    job_position TEXT NOT NULL,
    job_location TEXT NOT NULL,
    prospect_snapshot_json TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS audit_log (
    id SERIAL PRIMARY KEY,
    action TEXT NOT NULL,
    entity_type TEXT NOT NULL,
    entity_id TEXT,
    actor TEXT NOT NULL,
    details_json TEXT,
    created_at TEXT NOT NULL
);
