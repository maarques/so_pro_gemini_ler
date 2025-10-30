import tkinter as tk
from tkinter import filedialog

class FilesFunctions:
    def limpar_log(self):
        self.log_text.config(state='normal')
        self.log_text.delete('1.0', tk.END)
        self.log_text.config(state='disabled')

    def selecionar_arquivos(self):
        tipos_arquivo = [
            ("Arquivos de Endereço", "*.docx *.xlsx"),
            ("Todos os arquivos", "*.*")
        ]
        arquivos = filedialog.askopenfilenames(title="Selecione um ou mais arquivos", filetypes=tipos_arquivo)
        if arquivos:
            self.arquivos_selecionados = list(arquivos)
            self.lbl_entrada_status.config(text=f"{len(arquivos)} arquivos selecionados.")
            self.log(f"Entrada definida: {len(arquivos)} arquivos.")

    def selecionar_pasta(self):
        pasta = filedialog.askdirectory(title="Selecione uma pasta com os arquivos")
        if pasta:
            self.arquivos_selecionados = [pasta]
            self.lbl_entrada_status.config(text=f"Pasta selecionada: {pasta}")
            self.log(f"Entrada definida: Pasta {pasta}")

    def selecionar_saida(self):
        caminho_saida = filedialog.asksaveasfilename(
            title="Salvar arquivo Word como...",
            initialfile=self.nome_arquivo_saida,
            defaultextension=".docx",
            filetypes=[("Documento Word", "*.docx")]
        )
        if caminho_saida:
            self.caminho_arquivo_saida_completo = caminho_saida
            self.lbl_saida_status.config(text=f"Será salvo como: {caminho_saida}")
            self.log(f"Saída definida: {caminho_saida}")
