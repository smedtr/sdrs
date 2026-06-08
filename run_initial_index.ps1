param([string]$Config = "config.json")
python knowledge_hub.py init-db --config $Config
python knowledge_hub.py index-files --config $Config
python knowledge_hub.py index-outlook --config $Config
streamlit run app.py
