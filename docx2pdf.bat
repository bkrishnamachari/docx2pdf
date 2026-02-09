@echo off
powershell -ExecutionPolicy Bypass -File "%~dp0docx2pdf.ps1" %*
