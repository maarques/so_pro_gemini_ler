import tkinter as tk
from tkinter import messagebox, ttk
import backend.etiqueta as backend 
import threading
from classes.filesFunctions import FilesFunctions

class AppGeradorEtiquetas:
    def __init__(self, parent_tab):
        self.map_chaves_padrao = {
            'Nome': ['Nome completo / Razão Social', 'Nome'],
            'CPF': ['CPF ou CNPJ', 'CPF/CNPJ', 'CPF'],
            'Endereço': ['Endereço completo', 'Endereço Completo', 'Endereço'],
            'Telefone': ['Telefone de contato', 'Telefone'],
            'Qtd Cartões': ['Qtd de cartões', 'Qtda Cartões', 'qtd cartões', 'qtda de cartões'],
            'IBGE': ['IBGE de atuação', 'IBGE'],
        }

        self.backend = backend
        self.root = parent_tab 
        self.nome_arquivo_saida = "Cebraspe_Etiquetas_Geradas.docx"

        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True)

        frame_entrada = ttk.LabelFrame(main_frame, text="1. Selecionar Arquivos de Entrada", padding=15)
        frame_entrada.pack(fill=tk.X, padx=10, pady=10)
        
        self.lbl_entrada_status = ttk.Label(frame_entrada, text="Nenhuma entrada selecionada.", wraplength=650)
        
        frame_saida = ttk.LabelFrame(main_frame, text="2. Selecionar Local de Saída", padding=15)
        frame_saida.pack(fill=tk.X, padx=10, pady=5)

        self.lbl_saida_status = ttk.Label(frame_saida, text=f"Será salvo como: {self.nome_arquivo_saida}", wraplength=650)

        frame_acao = ttk.Frame(main_frame)
        
        self.btn_gerar = ttk.Button(frame_acao, text="GERAR ETIQUETAS", command=self.iniciar_processamento)
        self.btn_gerar.pack(fill=tk.X, ipady=10)

        frame_log = ttk.LabelFrame(main_frame, text="Mensagens do Processo", padding=10)

        scrollbar = ttk.Scrollbar(frame_log)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.log_text = tk.Text(frame_log, height=200, wrap=tk.WORD, yscrollcommand=scrollbar.set, state='disabled', font=('Courier New', 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.log_text.yview)

        self.functions = FilesFunctions(self.lbl_entrada_status, self.lbl_saida_status, self.log_text)
        self.functions.nome_arquivo_saida = self.nome_arquivo_saida # Garante que o nome do arquivo padrão seja o desta classe

        btn_arquivos = ttk.Button(frame_entrada, text="Selecionar Arquivos (.docx, .xlsx)", command=self.functions.selecionar_arquivos)
        btn_arquivos.pack(fill=tk.X, pady=5)
        
        btn_pasta = ttk.Button(frame_entrada, text="Selecionar Pasta (contém .docx, .xlsx)", command=self.functions.selecionar_pasta)
        btn_pasta.pack(fill=tk.X, pady=5)
        
        self.lbl_entrada_status.pack(fill=tk.X, pady=5) # .pack() chamado agora

        btn_saida = ttk.Button(frame_saida, text="Selecionar Pasta e Nome do Arquivo", command=self.functions.selecionar_saida)
        btn_saida.pack(fill=tk.X, pady=5)
        
        self.lbl_saida_status.pack(fill=tk.X, pady=5) # .pack() chamado agora

        frame_acao.pack(fill=tk.X, padx=10, pady=10)
        
        frame_log.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)



    def log(self, mensagem: str):
        if self.log_text:
            self.log_text.config(state='normal')
            self.log_text.insert(tk.END, f"{mensagem}\n")
            self.log_text.see(tk.END) # Auto-scroll
            self.log_text.config(state='disabled')
        else:
            print(mensagem)

    def iniciar_processamento(self):
        self.functions.limpar_log() # Usa o limpar_log da instância de functions

        if not self.functions.arquivos_selecionados:
            self.log("ERRO: Por favor, selecione arquivos ou uma pasta de entrada.")
            messagebox.showerror("Erro", "Nenhuma entrada selecionada.")
            return

        if not self.functions.caminho_arquivo_saida_completo:
            self.log("ERRO: Por favor, selecione um local para salvar o arquivo de saída.")
            messagebox.showerror("Erro", "Nenhum local de saída selecionado.")
            return

        self.btn_gerar.config(text="PROCESSANDO...", state='disabled')

        threading.Thread(
            target=self.processar_em_thread,
            args=(self.functions.arquivos_selecionados, self.functions.caminho_arquivo_saida_completo),
            daemon=True
        ).start()

    def processar_em_thread(self, entradas, saida):
        try:
            self.log("Iniciando processamento...")
            
            dados = self.backend.processar_entradas(entradas, self.log)
            
            if dados:
                self.backend.gerar_documento_word(dados, saida, self.log)
                self.log(f"Processamento concluído! Arquivo salvo em: {saida}")
            else:
                self.log("Processamento concluído, mas nenhum dado de etiqueta foi encontrado.")

        except Exception as e:
            self.log(f"\nERRO INESPERADO: {e}")
            self.root.after(0, lambda: messagebox.showerror("Erro Inesperado", str(e)))
        finally:
            self.root.after(0, lambda: self.btn_gerar.config(text="GERAR ETIQUETAS", state='normal'))

