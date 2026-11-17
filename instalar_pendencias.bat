@echo off
TITLE Instalador de Dependencias - MVP Automacao Garantias
COLOR 0A
echo ========================================================
echo   CONFIGURANDO AMBIENTE E INSTALANDO DEPENDENCIAS
echo ========================================================
echo.

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERRO] Python nao encontrado! Instale o Python e marque 'Add Python to PATH'.
    pause
    exit /b
)

echo [OK] Python detectado! Atualizando pip...
python -m pip install --upgrade pip

echo.
echo Instalando pandas, python-pptx, openpyxl e matplotlib (para graficos)...
pip install pandas python-pptx openpyxl matplotlib

echo.
echo ========================================================
echo   TUDO PRONTO! AMBIENTE CONFIGURADO COM SUCESSO.
echo ========================================================
pause