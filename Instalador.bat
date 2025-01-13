@echo off
REM Clonar o repositório
git clone https://github.com/yJuanReis/Portal-de-boletos.git

REM Entrar no diretório clonado
cd Portal-de-boletos

REM Criar o ambiente virtual Python
python -m venv venv

REM Ativar o ambiente virtual
call venv\Scripts\activate

REM Instalar as dependências do projeto
pip install -r requirements.txt

REM Instalar as dependências do projeto 2
pip install -r ../Portal-de-boletos/requirements.txt

exit
