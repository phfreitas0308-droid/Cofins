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

REM ---- 0) Confere se os arquivos de programa estao nesta pasta ----
if not exist "apuracao_if.py" (
  echo  [!] Nao encontrei "apuracao_if.py" nesta pasta:
  echo        %cd%
  echo.
  echo  Os arquivos nao estao todos juntos. Extraia o .zip por completo
  echo  e rode o APURAR_IF.bat de DENTRO da pasta "Calculadora_PIS_COFINS_IF",
  echo  onde ficam os arquivos .py (apuracao_if.py, regras_if.py, etc.).
  echo  Nao copie apenas o .bat para outra pasta.
  echo.
  pause
  exit /b 1
)

REM ---- 1) Descobre o Python ----
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

REM ---- 2) Garante a biblioteca openpyxl ----
echo  Preparando o ambiente (pode demorar na 1a vez)...
%PY% -m pip install --quiet openpyxl >nul 2>&1
echo  Pronto.
echo.

REM ---- 3) Dados da apuracao ----
set "COMP="
set /p "COMP=  Competencia (ex.: 07/2026) ou Enter para pular: "
set "EMP="
set /p "EMP=  Nome da instituicao (ou Enter para pular): "
echo.

REM ---- 4) Confere os arquivos de entrada ----
if not exist "balancete.xlsx" (
  echo  [!] Nao encontrei "balancete.xlsx" nesta pasta. Coloque o balancete COSIF aqui.
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

REM ---- 5) Apaga resultado antigo (evita falso "PRONTO") ----
if exist "apuracao_resultado.xlsx" del "apuracao_resultado.xlsx" >nul 2>&1
if exist "apuracao_resultado.xlsx" (
  echo  [!] Nao consegui atualizar "apuracao_resultado.xlsx" - ele parece estar
  echo      ABERTO no Excel. Feche a planilha e rode de novo.
  echo.
  pause
  exit /b 1
)

REM ---- 6) Roda a apuracao ----
echo  Calculando...
echo.
%PY% apuracao_if.py --balancete "balancete.xlsx" --cadastro "cadastro.xlsx" --competencia "%COMP%" --empresa "%EMP%" --saida "apuracao_resultado.xlsx"

REM 'errorlevel 1' = o Python terminou com erro
if errorlevel 1 (
  echo.
  echo  [!] A apuracao falhou. Leia as mensagens acima.
  echo.
  pause
  exit /b 1
)

REM ---- 7) Sucesso ----
if exist "apuracao_resultado.xlsx" (
  echo.
  echo ============================================================
  echo    PRONTO! Gerado: apuracao_resultado.xlsx
  echo ============================================================
  start "" "apuracao_resultado.xlsx"
) else (
  echo  [!] Nao gerou o arquivo de resultado. Leia as mensagens acima.
)
echo.
pause
endlocal
