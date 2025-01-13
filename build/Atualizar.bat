@echo off
REM Script para atualizar o repositório Git automaticamente

REM Navega para o diretório onde o .bat está localizado
cd /d "%~dp0"

REM Atualiza o repositório local com as alterações remotas
git pull origin main

REM Adiciona todas as alterações ao índice
git add .

REM Cria um commit com uma mensagem padrão
git commit -m "Atualização automática"

REM Envia as alterações para o repositório remoto
git push origin main

REM Mensagem de conclusão
echo Repositório atualizado com sucesso!
exit
