from __future__ import annotations

import argparse
import csv
import hashlib
import json
import re
import sqlite3
from collections import Counter
from datetime import datetime, timedelta
from pathlib import Path

import fitz
import pandas as pd
from docx import Document as WordDocument

SCHEMA_SQL = """
PRAGMA journal_mode=WAL;
CREATE TABLE IF NOT EXISTS items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    source_type TEXT NOT NULL,
    source_id TEXT,
    parent_source_id TEXT,
    file_name TEXT,
    path TEXT,
    email_subject TEXT,
    email_sender TEXT,
    email_recipients TEXT,
    email_date TEXT,
    content TEXT,
    summary TEXT,
    keywords TEXT,
    topic TEXT,
    document_type TEXT,
    entity TEXT,
    year INTEGER,
    sensitivity TEXT,
    hash_sha256 TEXT,
    indexed_at TEXT DEFAULT CURRENT_TIMESTAMP,
    last_modified TEXT,
    UNIQUE(source_type, source_id, path, file_name)
);
CREATE INDEX IF NOT EXISTS idx_items_source_type ON items(source_type);
CREATE INDEX IF NOT EXISTS idx_items_topic ON items(topic);
CREATE INDEX IF NOT EXISTS idx_items_doc_type ON items(document_type);
CREATE INDEX IF NOT EXISTS idx_items_entity ON items(entity);
CREATE INDEX IF NOT EXISTS idx_items_year ON items(year);
"""

UPSERT_SQL = """INSERT INTO items (
    source_type, source_id, parent_source_id, file_name, path,
    email_subject, email_sender, email_recipients, email_date,
    content, summary, keywords, topic, document_type, entity,
    year, sensitivity, hash_sha256, last_modified
) VALUES (
    :source_type, :source_id, :parent_source_id, :file_name, :path,
    :email_subject, :email_sender, :email_recipients, :email_date,
    :content, :summary, :keywords, :topic, :document_type, :entity,
    :year, :sensitivity, :hash_sha256, :last_modified
)
ON CONFLICT(source_type, source_id, path, file_name)
DO UPDATE SET
    content=excluded.content,
    summary=excluded.summary,
    keywords=excluded.keywords,
    topic=excluded.topic,
    document_type=excluded.document_type,
    entity=excluded.entity,
    year=excluded.year,
    sensitivity=excluded.sensitivity,
    hash_sha256=excluded.hash_sha256,
    last_modified=excluded.last_modified,
    indexed_at=CURRENT_TIMESTAMP;"""

TOPIC_RULES = {
    "gdpr": ["gdpr", "privacy", "data protection", "persoonsgegevens", "dpo", "purview"],
    "audit": ["audit", "pwc", "internal control", "sox", "assurance"],
    "payroll": ["payroll", "salary", "indexation", "pc200", "pc310", "eco-cheque", "sd worx"],
    "hr": ["hr", "employee", "works council", "outplacement", "benefits", "compensation", "union"],
    "legal": ["agreement", "clause", "contract", "annex", "compliance", "directive", "law"]
}
ENTITY_RULES = {
    "Belgium": ["belgium", "belgique", "belgië", "brussels", "pc200", "pc310"],
    "France": ["france", "français", "paris", "marseille"],
    "Global": ["global", "emea", "worldwide"],
    "Kyndryl": ["kyndryl", "pi-square", "pisquare"]
}
SENSITIVITY_RULES = {
    "highly confidential": ["medical", "salary increase", "disciplinary", "dismissal", "passport", "national register"],
    "confidential": ["confidential", "restricted", "personal data", "cv", "compensation", "bonus"],
    "internal": ["internal", "team only", "draft"]
}
DOC_TYPE_RULES = {
    "contract": ["contract", "agreement", "msa", "sow", "statement of work"],
    "presentation": ["presentation", "slides", ".pptx"],
    "report": ["report", "analysis", "dashboard", "minutes"],
    "memo": ["memo", "note", "briefing"],
    "spreadsheet": ["budget", "calculator", ".xlsx", "forecast"],
    "email": ["from:", "to:", "subject:"]
}
STOPWORDS = {"the", "and", "for", "with", "from", "this", "that", "have", "will", "your", "you", "de", "het", "een", "van", "met", "voor", "dat", "les", "des", "pour", "dans"}
DEFAULT_FOLDERS = {"Inbox": 6, "Sent Items": 5, "Deleted Items": 3, "Drafts": 16}


def load_config(config_path: str) -> dict:
    return json.loads(Path(config_path).read_text(encoding="utf-8"))


def get_conn(db_path: str):
    return sqlite3.connect(db_path)


def sha256_text(text: str) -> str:
    return hashlib.sha256((text or "").encode("utf-8", errors="ignore")).hexdigest()


def _norm(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "").lower()).strip()


def _best(text: str, rules: dict, default: str) -> str:
    t = _norm(text)
    scores = {label: sum(1 for kw in kws if kw in t) for label, kws in rules.items()}
    scores = {k: v for k, v in scores.items() if v > 0}
    return max(scores, key=scores.get) if scores else default


def extract_keywords(text: str, top_n: int = 12):
    tokens = re.findall(r"[a-zA-ZÀ-ÿ0-9\-]{3,}", _norm(text))
    tokens = [t for t in tokens if t not in STOPWORDS and not t.isdigit()]
    return [w for w, _ in Counter(tokens).most_common(top_n)]


def summarize(text: str, max_chars: int = 400) -> str:
    clean = re.sub(r"\s+", " ", (text or "").strip())
    return clean if len(clean) <= max_chars else clean[:max_chars].rsplit(" ", 1)[0] + " ..."


def classify_record(text: str, *, file_name: str = "", email_subject: str = "") -> dict:
    combined = " ".join(p for p in [file_name, email_subject, text] if p)
    year_match = re.search(r"\b(20\d{2})\b", combined)
    return {
        "summary": summarize(text),
        "keywords": ", ".join(extract_keywords(combined)),
        "topic": _best(combined, TOPIC_RULES, "other"),
        "document_type": _best(combined, DOC_TYPE_RULES, "document"),
        "entity": _best(combined, ENTITY_RULES, "unspecified"),
        "year": int(year_match.group(1)) if year_match else None,
        "sensitivity": _best(combined, SENSITIVITY_RULES, "internal"),
    }


def read_pdf(path: Path) -> str:
    chunks = []
    with fitz.open(path) as doc:
        for page in doc:
            chunks.append(page.get_text("text"))
    return "\n".join(chunks)


def read_docx(path: Path) -> str:
    doc = WordDocument(path)
    return "\n".join(p.text for p in doc.paragraphs if p.text)


def read_xlsx(path: Path) -> str:
    out = []
    excel = pd.ExcelFile(path, engine="openpyxl")
    for sheet in excel.sheet_names[:5]:
        df = pd.read_excel(path, sheet_name=sheet, engine="openpyxl")
        out.append(f"[Sheet: {sheet}]")
        out.append(df.astype(str).head(100).to_csv(index=False))
    return "\n".join(out)


def read_csv(path: Path) -> str:
    rows = []
    with open(path, newline="", encoding="utf-8", errors="ignore") as fh:
        reader = csv.reader(fh)
        for i, row in enumerate(reader):
            if i >= 1000:
                break
            rows.append(", ".join(row))
    return "\n".join(rows)


def read_txt(path: Path) -> str:
    return path.read_text(encoding="utf-8", errors="ignore")


def read_pptx_placeholder(path: Path) -> str:
    return f"Presentation file detected: {path.name}. Add python-pptx later if needed."


def extract_file_text(path: Path) -> str:
    ext = path.suffix.lower()
    if ext == ".pdf":
        return read_pdf(path)
    if ext == ".docx":
        return read_docx(path)
    if ext == ".xlsx":
        return read_xlsx(path)
    if ext == ".csv":
        return read_csv(path)
    if ext == ".txt":
        return read_txt(path)
    if ext == ".pptx":
        return read_pptx_placeholder(path)
    raise ValueError(f"Unsupported extension: {ext}")


def attachment_bytes_to_text(filename: str, content_bytes: bytes) -> str:
    suffix = Path(filename).suffix.lower()
    temp_path = Path(f"__temp_attachment{suffix}")
    temp_path.write_bytes(content_bytes)
    try:
        return extract_file_text(temp_path)
    finally:
        temp_path.unlink(missing_ok=True)


def iter_files(root_paths, allowed_exts):
    allowed = {e.lower() for e in allowed_exts}
    for root in root_paths:
        root_path = Path(root)
        if not root_path.exists():
            continue
        for path in root_path.rglob("*"):
            if path.is_file() and path.suffix.lower() in allowed:
                yield path


def init_db(config_path: str):
    cfg = load_config(config_path)
    conn = get_conn(cfg.get("database_path", "knowledge_hub.db"))
    try:
        conn.executescript(SCHEMA_SQL)
        conn.commit()
    finally:
        conn.close()
    print("Database initialized")


def index_files(config_path: str):
    cfg = load_config(config_path)
    conn = get_conn(cfg.get("database_path", "knowledge_hub.db"))
    processed = 0
    for path in iter_files(cfg.get("document_roots", []), cfg.get("file_extensions", [])):
        try:
            text = extract_file_text(path)
        except Exception as exc:
            print(f"[WARN] {path}: {exc}")
            continue
        c = classify_record(text, file_name=path.name)
        rec = {
            "source_type": "file", "source_id": str(path.resolve()), "parent_source_id": None,
            "file_name": path.name, "path": str(path.resolve()),
            "email_subject": None, "email_sender": None, "email_recipients": None, "email_date": None,
            "content": text, "summary": c["summary"], "keywords": c["keywords"], "topic": c["topic"],
            "document_type": c["document_type"], "entity": c["entity"], "year": c["year"],
            "sensitivity": c["sensitivity"], "hash_sha256": sha256_text(text), "last_modified": str(path.stat().st_mtime),
        }
        conn.execute(UPSERT_SQL, rec)
        processed += 1
        if processed % 25 == 0:
            conn.commit(); print(f"Processed {processed} files...")
    conn.commit(); conn.close()
    print(f"Done. Indexed files: {processed}")


def get_namespace():
    import win32com.client  # type: ignore
    outlook = win32com.client.Dispatch("Outlook.Application")
    return outlook.GetNamespace("MAPI")


def get_mailbox_root(namespace, mailbox_name):
    if not mailbox_name:
        return namespace
    for store in namespace.Folders:
        if str(store.Name).lower() == mailbox_name.lower():
            return store
    raise ValueError(f"Mailbox not found: {mailbox_name}")


def get_folder(namespace_or_store, folder_name):
    if hasattr(namespace_or_store, 'GetDefaultFolder') and folder_name in DEFAULT_FOLDERS:
        return namespace_or_store.GetDefaultFolder(DEFAULT_FOLDERS[folder_name])
    for folder in namespace_or_store.Folders:
        if str(folder.Name).lower() == folder_name.lower():
            return folder
    raise ValueError(f"Folder not found: {folder_name}")


def recipients_string(mail_item):
    values = []
    try:
        for r in mail_item.Recipients:
            values.append(str(getattr(r, 'Name', '')))
    except Exception:
        pass
    return '; '.join(v for v in values if v)


def extract_attachments(mail_item, cfg):
    if not cfg.get("include_email_attachments", True):
        return []
    max_mb = cfg.get("max_attachment_size_mb", 10)
    export_folder = cfg.get("attachment_export_folder")
    save_copy = cfg.get("save_attachments_copy", False)
    if save_copy and export_folder:
        Path(export_folder).mkdir(parents=True, exist_ok=True)
    items = []
    count = getattr(mail_item.Attachments, 'Count', 0)
    for i in range(1, count + 1):
        att = mail_item.Attachments.Item(i)
        filename = str(att.FileName)
        save_dir = Path(export_folder) if (save_copy and export_folder) else Path.cwd()
        temp_file = save_dir / filename
        try:
            att.SaveAsFile(str(temp_file))
            size_mb = temp_file.stat().st_size / (1024 * 1024)
            if size_mb > max_mb:
                temp_file.unlink(missing_ok=True)
                continue
            content_bytes = temp_file.read_bytes()
            try:
                text = attachment_bytes_to_text(filename, content_bytes)
            except Exception:
                text = f"Attachment indexed without text extraction: {filename}"
            items.append((filename, str(temp_file) if save_copy else None, text, str(temp_file.stat().st_mtime)))
        finally:
            if temp_file.exists() and not save_copy:
                temp_file.unlink(missing_ok=True)
    return items


def index_outlook(config_path: str):
    cfg = load_config(config_path)
    conn = get_conn(cfg.get("database_path", "knowledge_hub.db"))
    namespace = get_namespace()
    mailbox_root = get_mailbox_root(namespace, cfg.get("outlook_mailbox"))
    folders = cfg.get("outlook_folders", ["Inbox", "Sent Items"])
    cutoff = datetime.now() - timedelta(days=int(cfg.get("lookback_days", 730)))
    max_len = int(cfg.get("max_email_body_chars", 25000))
    min_year = cfg.get("skip_if_older_than_year")
    processed = 0
    attachments_indexed = 0
    for folder_name in folders:
        folder = get_folder(mailbox_root if cfg.get("outlook_mailbox") else namespace, folder_name)
        items = folder.Items
        items.Sort("[ReceivedTime]", True)
        for item in items:
            if getattr(item, "Class", None) != 43:
                continue
            dt = getattr(item, "ReceivedTime", None) or getattr(item, "SentOn", None)
            if dt and hasattr(dt, 'year'):
                dt_naive = dt.replace(tzinfo=None) if getattr(dt, 'tzinfo', None) else dt
                if dt_naive < cutoff:
                    break
                if min_year and dt.year < int(min_year):
                    continue
            subject = str(getattr(item, "Subject", "") or "")
            sender = str(getattr(item, "SenderName", "") or "")
            body = str(getattr(item, "Body", "") or "")[:max_len]
            recipients = recipients_string(item)
            entry_id = str(getattr(item, "EntryID", "") or f"{folder_name}-{processed}")
            c = classify_record(body, email_subject=subject)
            rec = {
                "source_type": "email", "source_id": entry_id, "parent_source_id": None,
                "file_name": None, "path": None, "email_subject": subject, "email_sender": sender,
                "email_recipients": recipients, "email_date": str(dt) if dt else None,
                "content": body, "summary": c["summary"], "keywords": c["keywords"], "topic": c["topic"],
                "document_type": "email", "entity": c["entity"], "year": dt.year if dt else c["year"],
                "sensitivity": c["sensitivity"], "hash_sha256": sha256_text(body), "last_modified": str(dt) if dt else None,
            }
            conn.execute(UPSERT_SQL, rec)
            processed += 1
            for filename, path_saved, text, last_modified in extract_attachments(item, cfg):
                c2 = classify_record(text, file_name=filename, email_subject=subject)
                rec2 = {
                    "source_type": "attachment", "source_id": f"{entry_id}::{filename}", "parent_source_id": entry_id,
                    "file_name": filename, "path": path_saved, "email_subject": subject, "email_sender": sender,
                    "email_recipients": recipients, "email_date": str(dt) if dt else None,
                    "content": text, "summary": c2["summary"], "keywords": c2["keywords"], "topic": c2["topic"],
                    "document_type": c2["document_type"], "entity": c2["entity"], "year": dt.year if dt else c2["year"],
                    "sensitivity": c2["sensitivity"], "hash_sha256": sha256_text(text), "last_modified": last_modified,
                }
                conn.execute(UPSERT_SQL, rec2)
                attachments_indexed += 1
            if processed % 25 == 0:
                conn.commit(); print(f"Processed {processed} emails; {attachments_indexed} attachments...")
    conn.commit(); conn.close()
    print(f"Done. Indexed emails: {processed}; attachments: {attachments_indexed}")


def main():
    parser = argparse.ArgumentParser(description="Personal Knowledge Hub")
    parser.add_argument("command", choices=["init-db", "index-files", "index-outlook"])
    parser.add_argument("--config", default="config.json")
    args = parser.parse_args()
    if args.command == "init-db":
        init_db(args.config)
    elif args.command == "index-files":
        index_files(args.config)
    elif args.command == "index-outlook":
        index_outlook(args.config)


if __name__ == "__main__":
    main()
