# Personal Knowledge Hub Starter (SQLite + Outlook COM)

This starter package indexes **local files**, **Outlook emails**, and **email attachments** into one SQLite database and lets you search them in a simple Streamlit UI.

## Files included
- `knowledge_hub.py` - database creation + file indexing + Outlook indexing
- `app.py` - Streamlit search UI
- `requirements.txt` - Python dependencies
- `config.example.json` - sample configuration
- `run_initial_index.ps1` - helper to do the first full run on Windows
- `README.md` - overview

## Supported file types
- PDF
- DOCX
- XLSX
- TXT
- CSV
- PPTX (placeholder text only in this starter)

## Quick start
1. Copy `config.example.json` to `config.json`
2. Create a virtual environment
3. Install requirements
4. Run `python knowledge_hub.py init-db --config config.json`
5. Run `python knowledge_hub.py index-files --config config.json`
6. Run `python knowledge_hub.py index-outlook --config config.json`
7. Run `streamlit run app.py`

## Notes
- Outlook indexing requires **Windows + Outlook desktop + configured profile**.
- Attachments are indexed as separate records and linked to the email by `parent_source_id`.
- The starter uses simple rule-based classification so you can adapt it easily to HR / legal / audit content.
