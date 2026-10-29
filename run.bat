@echo off
chcp 65001 > nul
title 유튜브 댓글 수집기

echo ========================================================
echo       유튜브 댓글 수집기 웹앱을 실행합니다.
echo ========================================================
echo.

python -m streamlit run app.py

pause

