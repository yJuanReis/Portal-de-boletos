@echo off
:: Define o diretório do repositório como o local onde o .bat está
set "REPO_DIR=%~dp0Portal-de-boletos"

:: Verifica se a pasta do repositório já existe
if exist "%REPO_DIR%" (
    echo A pasta do repositório já existe. Atualizando os arquivos...

    :: Navega até o diretório do repositório
    cd /d "%REPO_DIR%"

    :: Atualiza os arquivos para o commit mais recente
    git fetch origin
    git reset --hard origin/main

    :: Verifica se a atualização foi bem-sucedida
    if %errorlevel% equ 0 (
        echo Arquivos atualizados para o commit mais recente com sucesso!
    ) else (
        echo Ocorreu um erro ao atualizar os arquivos.
        pause
        exit /b 1
    )
) else (
    echo A pasta do repositório não existe. Clonando o repositório...

    :: Clona o repositório
    git clone https://github.com/yJuanReis/Portal-de-boletos.git

    :: Verifica se o clone foi bem-sucedido
    if %errorlevel% neq 0 (
        echo Ocorreu um erro ao clonar o repositório.
        pause
        exit /b 1
    )
)

:: Navega até o diretório do repositório (caso tenha sido clonado ou atualizado)
cd /d "%REPO_DIR%"

:: Cria o ambiente virtual Python, caso não exista
if not exist "venv" (
    echo Criando o ambiente virtual Python...
    python -m venv venv

    :: Verifica se a criação foi bem-sucedida
    if %errorlevel% neq 0 (
        echo Ocorreu um erro ao criar o ambiente virtual.
        pause
        exit /b 1
    )
) else (
    echo Ambiente virtual já existe.
)

:: Ativa o ambiente virtual
call venv\Scripts\activate

pip install -r requirements.txt 

echo Processo concluído com sucesso!
exit