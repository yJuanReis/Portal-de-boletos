import re
from tkinter import filedialog
import customtkinter as ctk
import os

NOME_ARQUIVO_CONFIG = "login_imap.txt"

class App:
    def __init__(self):
        # Configuração inicial do CustomTkinter
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        # Caminho absoluto do arquivo na pasta antes da pasta bin
        self.caminho_arquivo = os.path.join(os.path.dirname(os.path.realpath(__file__)), NOME_ARQUIVO_CONFIG)
        self.pasta_downloads = None

        self.app = ctk.CTk()
        self.configurar_janela()
        self.criar_widgets()
        self.app.mainloop()

    def configurar_janela(self):
        largura_janela = 800
        altura_janela = 500
        largura_tela = self.app.winfo_screenwidth()
        altura_tela = self.app.winfo_screenheight()
        pos_x = (largura_tela - largura_janela) // 2
        pos_y = (altura_tela - altura_janela) // 2

        self.app.geometry(f"{largura_janela}x{altura_janela}+{pos_x}+{pos_y}")
        self.app.title("Login - Portal")

    def criar_widgets(self):
        # Título principal
        ctk.CTkLabel(self.app, text="Portal de Boletos", font=("Arial", 24)).pack(pady=10)

        # Credenciais do Portal de Boletos
        ctk.CTkLabel(self.app, text="Credenciais do Portal de Boletos", font=("Arial", 18)).pack(pady=10)

        self.entrada_email_portal = ctk.CTkEntry(self.app, placeholder_text="Usuário", width=500, justify='center')
        self.entrada_email_portal.pack(pady=5)

        self.entrada_senha_portal = ctk.CTkEntry(self.app, placeholder_text="Senha", show="*", width=500, justify='center')
        self.entrada_senha_portal.pack(pady=5)

        # Credenciais de Acesso IMAP
        ctk.CTkLabel(self.app, text="Credenciais de Acesso IMAP", font=("Arial", 18)).pack(pady=10)

        self.entrada_email = ctk.CTkEntry(self.app, placeholder_text="Email", width=500, justify='center')
        self.entrada_email.pack(pady=5)

        self.entrada_senha = ctk.CTkEntry(self.app, placeholder_text="Senha de APP", show="*", width=500, justify='center')
        self.entrada_senha.pack(pady=5)

        # Botão para selecionar pasta de downloads
        self.botao_pasta_downloads = ctk.CTkButton(self.app, text="Selecionar Pasta de Downloads", command=self.selecionar_pasta_downloads)
        self.botao_pasta_downloads.pack(pady=15)

        # Mensagem sobre a pasta selecionada
        self.label_pasta_downloads = ctk.CTkLabel(self.app, text="Nenhuma pasta selecionada para downloads", text_color="red")
        self.label_pasta_downloads.pack(pady=5)

        # Botão salvar
        ctk.CTkButton(self.app, text="Salvar", command=self.salvar_dados).pack(pady=20)

    def selecionar_pasta_downloads(self):
        pasta_selecionada = filedialog.askdirectory()

        if pasta_selecionada:
            self.pasta_downloads = pasta_selecionada
            self.botao_pasta_downloads.configure(fg_color="green")  # Muda a cor do botão para verde ao selecionar a pasta
            self.label_pasta_downloads.configure(text=f"{pasta_selecionada}", text_color="green")
        else:
            self.botao_pasta_downloads.configure(fg_color="red")  # Mantém a cor vermelha se nenhuma pasta for selecionada
            self.label_pasta_downloads.configure(text="Nenhuma pasta selecionada para downloads", text_color="red")





    def validar_email(self, email):
        """Valida o formato do email usando regex."""
        padrao_email = r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$"
        return re.match(padrao_email, email) is not None

    def salvar_dados(self):
        usuario_email_portal = self.entrada_email_portal.get()
        senha_email_portal = self.entrada_senha_portal.get()
        usuario_email = self.entrada_email.get()
        senha_email = self.entrada_senha.get()

        # Resetar bordas e botão antes da validação
        campos = [
            (self.entrada_email_portal, usuario_email_portal),
            (self.entrada_senha_portal, senha_email_portal),
            (self.entrada_email, usuario_email),
            (self.entrada_senha, senha_email),
        ]

        for campo, valor in campos:
            campo.configure(border_color="blue")

        if not self.pasta_downloads:
            self.botao_pasta_downloads.configure(fg_color="red")  # Destaca o botão em vermelho se nenhuma pasta foi selecionada

        # Validar campos e destacar os que estão vazios em vermelho
        campos_invalidos = [campo for campo, valor in campos if not valor]

        if campos_invalidos or not self.pasta_downloads:
            for campo in campos_invalidos:
                campo.configure(border_color="red")
            if not self.pasta_downloads:
                self.label_pasta_downloads.configure(text="Selecione uma pasta antes de salvar!", text_color="red")
            return

        # Validar formato do email no campo IMAP
        if not self.validar_email(usuario_email):
            self.label_pasta_downloads.configure(text="Email inválido! Verifique o formato.", text_color="red")
            self.entrada_email.configure(border_color="red")
            return

        try:
            # Caminho do arquivo na pasta inicial do código
            with open(self.caminho_arquivo, "w") as arquivo:
                arquivo.write(f"{usuario_email}\n")
                arquivo.write(f"{senha_email}\n")
                arquivo.write(f"{usuario_email_portal}\n")
                arquivo.write(f"{senha_email_portal}\n")
                arquivo.write(f"{self.pasta_downloads}\n")

            self.label_pasta_downloads.configure(text="Dados salvos com sucesso!", text_color="green")
            # Fechar a interface após salvar os dados
            self.fechar_aplicacao()

        except Exception as e:
            self.label_pasta_downloads.configure(text=f"Erro ao salvar: {e}", text_color="red")
    def fechar_aplicacao(self):
        # Fechar a aplicação Tkinter
        self.app.destroy()

if __name__ == "__main__":
    App()