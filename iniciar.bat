@echo off
REM Savian - Sistema de Fechamento DRE (Windows)
REM Na primeira vez instala o Python (se faltar), as bibliotecas e cria o atalho na area de trabalho.
setlocal
cd /d "%~dp0"
title Savian DRE

if not exist "%~dp0app.py" goto semarquivos

REM Empresas e chaves ficam em Documentos, para nao se perderem ao atualizar a pasta do sistema.
if exist "dados\savian.db" (
  set "SAVIAN_DADOS=%~dp0dados"
) else (
  set "SAVIAN_DADOS=%USERPROFILE%\Documents\Savian DRE"
)

if exist ".venv\Scripts\streamlit.exe" goto atalho

echo.
echo === Primeira execucao: instalando o sistema (leva alguns minutos) ===
set "PY="
py -3 --version >nul 2>&1 && set "PY=py -3"
if not defined PY python --version >nul 2>&1 && set "PY=python"
if not defined PY if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" set "PY=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
if not defined PY (
  echo Python nao encontrado. Instalando o Python 3.12...
  winget install -e --id Python.Python.3.12 --scope user --silent --accept-package-agreements --accept-source-agreements
  if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" set "PY=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
)
if not defined PY (
  echo.
  echo Nao consegui instalar o Python automaticamente.
  echo Instale pelo site https://www.python.org/downloads/ marcando "Add python.exe to PATH"
  echo e depois abra este arquivo de novo.
  start https://www.python.org/downloads/
  pause
  exit /b 1
)
%PY% -m venv .venv || goto erro
".venv\Scripts\python.exe" -m pip install --upgrade pip >nul
".venv\Scripts\python.exe" -m pip install -r "%~dp0requirements.txt" || goto erro

:atalho
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$d=[Environment]::GetFolderPath('Desktop'); $l=Join-Path $d 'Savian DRE.lnk';" ^
  "if(-not (Test-Path $l)){ $s=(New-Object -ComObject WScript.Shell).CreateShortcut($l);" ^
  "$s.TargetPath='%~dp0iniciar.bat'; $s.WorkingDirectory='%~dp0'; $s.WindowStyle=7;" ^
  "$s.IconLocation='%SystemRoot%\System32\imageres.dll,109'; $s.Save() }" >nul 2>&1

echo.
echo === Savian DRE aberto no navegador. Deixe esta janela aberta enquanto usa o sistema. ===
".venv\Scripts\python.exe" -m streamlit run "%~dp0app.py" --server.headless false --browser.gatherUsageStats false
goto fim

:semarquivos
echo.
echo Os arquivos do sistema nao estao na pasta deste iniciar.bat.
echo Provavelmente ele foi aberto de dentro do ZIP.
echo Clique com o botao direito no ZIP baixado, escolha "Extrair tudo...",
echo abra a pasta extraida e de dois cliques no iniciar.bat de la.
pause
exit /b 1

:erro
echo.
echo Algo deu errado na instalacao. Tire um print desta janela e envie para o suporte.
pause
exit /b 1

:fim
endlocal
