@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"
cls
echo ============================================================
echo    APURACAO PIS / COFINS - INSTITUICAO FINANCEIRA (COSIF)
echo    Regime cumulativo: PIS 0,65%% + COFINS 4,00%% (sem credito)
echo ============================================================
echo.

set "PY="
where python >nul 2>&1 && set "PY=python"
if not defined PY ( where py >nul 2>&1 && set "PY=py" )
if not defined PY (
  echo  [!] Python nao encontrado. Baixe em https://www.python.org/downloads/
  echo      e marque "Add Python to PATH" na instalacao.
  echo.
  pause
  exit /b 1
)

echo  Preparando o ambiente (pode demorar na 1a vez)...
%PY% -m pip install --quiet openpyxl >nul 2>&1
echo  Pronto.
echo.

set "COMP="
set /p "COMP=  Competencia (ex.: 07/2026) ou Enter para pular: "
set "EMP="
set /p "EMP=  Nome da instituicao (ou Enter para pular): "
echo.

if not exist "balancete.xlsx" (
  echo  [!] Nao encontrei "balancete.xlsx" nesta pasta.
  echo      Coloque o balancete COSIF aqui com esse nome.
  echo.
  pause
  exit /b 1
)
if not exist "cadastro.xlsx" (
  echo  [!] Nao encontrei "cadastro.xlsx" nesta pasta.
  echo.
  pause
  exit /b 1
)

echo  Calculando...
echo.
%PY% apuracao_if.py --balancete "balancete.xlsx" --cadastro "cadastro.xlsx" --competencia "%COMP%" --empresa "%EMP%" --saida "apuracao_resultado.xlsx"

echo.
if exist "apuracao_resultado.xlsx" (
  echo ============================================================
  echo    PRONTO! Gerado: apuracao_resultado.xlsx
  echo ============================================================
  start "" "apuracao_resultado.xlsx"
) else (
  echo  [!] Algo deu errado. Leia as mensagens acima.
)
echo.
pause
endlocal
