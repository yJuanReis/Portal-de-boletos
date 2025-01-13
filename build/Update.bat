@echo off
:: Define o diretório do repositório como o local onde o .bat está
set REPO_DIR=%~dp0

:: Navega até o diretório do repositório
cd /d %REPO_DIR%

:: Exibe mensagem de início
echo Atualizando os arquivos para o commit mais recente...

:: Baixa as alterações mais recentes do repositório remoto
git fetch origin

:: Reseta os arquivos locais para corresponderem ao último commit da branch main
git reset --hard origin/main

:: Verifica se a atualização foi bem-sucedida
if %errorlevel% equ 0 (
    echo Arquivos atualizados para o commit mais recente com sucesso!
) else (
    echo Ocorreu um erro ao atualizar os arquivos.
)

exit
