@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"
cls
echo ============================================================
echo    APURACAO DE PIS / COFINS
echo ============================================================
echo.

REM ---- 1) Descobre o Python ----
set "PY="
where python >nul 2>&1 && set "PY=python"
if not defined PY (
  where py >nul 2>&1 && set "PY=py"
)
if not defined PY (
  echo  [!] Python nao encontrado neste computador.
  echo      Baixe em https://www.python.org/downloads/ e, na instalacao,
  echo      marque a opcao "Add Python to PATH". Depois rode este arquivo de novo.
  echo.
  pause
  exit /b 1
)

REM ---- 2) Garante a biblioteca openpyxl (silencioso, so na 1a vez) ----
echo  Preparando o ambiente (pode demorar um pouco na 1a vez)...
%PY% -m pip install --quiet openpyxl >nul 2>&1
echo  Pronto.
echo.

REM ---- 3) Escolha do regime ----
echo  Qual o regime de apuracao?
echo     1^) Nao-cumulativo  ^(PIS 1,65%% / Cofins 7,6%% - com credito^)
echo     2^) Cumulativo      ^(PIS 0,65%% / Cofins 3,0%% - sem credito^)
echo.
set "OP="
set /p "OP=  Digite 1 ou 2 e tecle Enter: "
if "%OP%"=="2" ( set "REG=cumulativo" ) else ( set "REG=nao_cumulativo" )
echo.

REM ---- 4) Competencia (opcional) ----
set "COMP="
set /p "COMP=  Competencia (ex.: 07/2026) ou apenas Enter para pular: "
echo.

REM ---- 5) Confere os arquivos de entrada ----
if not exist "balancete.xlsx" (
  echo  [!] Nao encontrei "balancete.xlsx" nesta pasta.
  echo      Coloque o seu balancete aqui com esse nome e tente de novo.
  echo.
  pause
  exit /b 1
)
if not exist "cadastro_cofins.xlsx" (
  echo  [!] Nao encontrei "cadastro_cofins.xlsx" nesta pasta.
  echo      Coloque o seu cadastro aqui com esse nome e tente de novo.
  echo.
  pause
  exit /b 1
)

REM ---- 6) Roda a apuracao ----
echo  Calculando...
echo.
%PY% apuracao_pis_cofins.py --balancete "balancete.xlsx" --cadastro "cadastro_cofins.xlsx" --regime %REG% --competencia "%COMP%" --saida "apuracao_resultado.xlsx"

echo.
if exist "apuracao_resultado.xlsx" (
  echo ============================================================
  echo    PRONTO! Gerado o arquivo: apuracao_resultado.xlsx
  echo ============================================================
  start "" "apuracao_resultado.xlsx"
) else (
  echo  [!] Algo deu errado. Leia as mensagens acima.
)
echo.
pause
endlocal
