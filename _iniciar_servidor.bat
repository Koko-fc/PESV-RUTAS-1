@echo off
rem Archivo interno: lo usa Iniciar_PESV.vbs para arrancar la app sin mostrar ventana.
cd /d "%~dp0"
call venv\Scripts\activate.bat
streamlit run app.py --server.port 8501 --server.headless true --browser.gatherUsageStats false > pesv_log.txt 2>&1
