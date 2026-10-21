@echo off
title Outlook Downloader Server
echo ========================================================
echo       INICIANDO SERVIDOR OUTLOOK DOWNLOADER API
echo ========================================================
echo.
echo Servidor escuchando en: http://0.0.0.0:8000
echo Documentacion interactiva: http://localhost:8000/docs
echo.
echo Presiona Ctrl + C si deseas detener el servidor.
echo ========================================================
echo.
.\venv\Scripts\uvicorn.exe servidor_api:app --host 0.0.0.0 --port 8000
pause

