from pathlib import Path
import sqlite3

import json
import argparse
import pandas as pd
import streamlit as st

def load_config(config_path: str) -> dict:
    return json.loads(Path(config_path).read_text(encoding="utf-8"))

parser = argparse.ArgumentParser(description="Personal Knowledge Hub")
parser.add_argument("--config", default="config.json")
args = parser.parse_args()
CONFIG = load_config(args.config)

st.set_page_config(page_title="Personal Knowledge Hub", layout="wide")
st.title("Personal Knowledge Hub")
st.caption("Search across indexed files, emails, and attachments.")



def conn():
    # DB_PATH = Path(__file__).resolve().parent / "knowledge_hub.db"    
    DB_PATH = CONFIG.get("database_path", "knowledge_hub.db")      
    return sqlite3.connect(DB_PATH)


def distinct_values(column: str):
    c = conn()
    try:
        q = f"SELECT DISTINCT {column} FROM items WHERE {column} IS NOT NULL AND {column} <> '' ORDER BY {column}"
        return [r[0] for r in c.execute(q).fetchall()]
    finally:
        c.close()


with st.sidebar:
   
    st.header("Filters")
    source_types = st.multiselect("Source type", distinct_values("source_type"))
    topics = st.multiselect("Topic", distinct_values("topic"))
    doc_types = st.multiselect("Document type", distinct_values("document_type"))
    entities = st.multiselect("Entity", distinct_values("entity"))
    sensitivities = st.multiselect("Sensitivity", distinct_values("sensitivity"))
    years = st.multiselect("Year", distinct_values("year"))
    max_results = st.slider("Max results", 10, 200, 50)

 
search = st.text_input("Search terms", placeholder="e.g. audit 2026 budget Belgium")
conditions = []
params = []
if search.strip():
    conditions.append("(coalesce(file_name,'') || ' ' || coalesce(email_subject,'') || ' ' || coalesce(content,'') || ' ' || coalesce(keywords,'')) LIKE ?")
    params.append(f"%{search.strip()}%")
for col, values in [("source_type", source_types), ("topic", topics), ("document_type", doc_types), ("entity", entities), ("sensitivity", sensitivities), ("year", years)]:
    if values:
        placeholders = ','.join('?' for _ in values)
        conditions.append(f"{col} IN ({placeholders})")
        params.extend(values)
where_clause = 'WHERE ' + ' AND '.join(conditions) if conditions else ''
query = f"SELECT id, source_type, file_name, path, email_subject, email_sender, email_date, topic, document_type, entity, year, sensitivity, keywords, summary FROM items {where_clause} ORDER BY COALESCE(email_date, indexed_at) DESC LIMIT ?"
params.append(max_results)
c = conn()
try:
    df = pd.read_sql_query(query, c, params=params)
finally:
    c.close()

st.subheader(f"Results ({len(df)})")
if df.empty:
    st.info("No results yet. Run the indexing scripts first, or broaden your search.")
else:
    for _, row in df.iterrows():
        title = row['file_name'] or row['email_subject'] or f"Item {row['id']}"
        with st.expander(title):
            meta = []
            for label, value in [("Source", row['source_type']), ("Topic", row['topic']), ("Type", row['document_type']), ("Entity", row['entity']), ("Year", row['year']), ("Sensitivity", row['sensitivity']), ("Sender", row['email_sender']), ("Date", row['email_date'])]:
                if pd.notna(value) and str(value) != '':
                    meta.append(f"**{label}:** {value}")
            if meta:
                st.markdown(' · '.join(meta))
            if pd.notna(row['keywords']) and row['keywords']:
                st.markdown(f"**Keywords:** {row['keywords']}")
            if pd.notna(row['summary']) and row['summary']:
                st.write(row['summary'])
            if pd.notna(row['path']) and row['path']:
                st.code(row['path'], language=None)


