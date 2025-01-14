@echo off
:: Verifica se o Python está instalado
python --version >nul 2>&1
IF ERRORLEVEL 1 (
    echo Python não está instalado. Por favor, instale o Python antes de continuar.
    pause
    exit /b
)

:: Verifica se o pip está instalado
pip --version >nul 2>&1
IF ERRORLEVEL 1 (
    echo Pip não está instalado. Por favor, instale o pip antes de continuar.
    pause
    exit /b
)

:: Instala as dependências do requirements.txt
echo Instalando dependências...
pip install -r requirements.txt

IF ERRORLEVEL 1 (
    echo Ocorreu um erro ao instalar as dependências.
    pause
    exit /b
)

echo Dependências instaladas com sucesso!
pause
