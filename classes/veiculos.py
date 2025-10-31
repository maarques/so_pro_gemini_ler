import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import backend.veiculo as backend
from classes.filesFunctions import FilesFunctions

class AppCadastroVeiculo:
    def __init__(self, parent_tab):
        self.frame = ttk.Frame(parent_tab, padding=10)
        self.frame.pack(fill=tk.BOTH, expand=True)

        main_frame = ttk.Frame(self.frame)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Entrada
        frame_entrada = ttk.LabelFrame(main_frame, text="1. Selecionar Planilha de Veículo", padding=15)
        frame_entrada.pack(fill=tk.X, padx=10, pady=10)

        self.lbl_entrada_status = ttk.Label(frame_entrada, text="Nenhuma entrada selecionada.", wraplength=650)

        btn_arquivos = ttk.Button(frame_entrada, text="Selecionar Planilha (.xlsx)", command=self.selecionar_planilha_veiculo)
        btn_arquivos.pack(fill=tk.X, pady=5)
        self.lbl_entrada_status.pack(fill=tk.X, pady=5)

        # Ações (botão) — pack antes do log
        frame_acao = ttk.Frame(main_frame)
        frame_acao.pack(fill=tk.X, padx=10, pady=10)
        self.btn_gerar = ttk.Button(frame_acao, text="REGISTRAR VEÍCULO", command=self.iniciar_automacao)
        self.btn_gerar.pack(fill=tk.X, ipady=10)

        # Log
        frame_log = ttk.LabelFrame(main_frame, text="Mensagens do Processo", padding=10)
        frame_log.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        scrollbar = ttk.Scrollbar(frame_log)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.log_text = tk.Text(frame_log, wrap=tk.WORD, yscrollcommand=scrollbar.set,
                                state='disabled', font=('Courier New', 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.log_text.yview)

        self.functions = FilesFunctions(self.lbl_entrada_status, None, self.log_text)

    def selecionar_planilha_veiculo(self):
        tipos_arquivo = [("Planilha Excel", "*.xlsx"), ("Todos os arquivos", "*.*")]
        arquivo = filedialog.askopenfilename(title="Selecione a planilha de viagem", filetypes=tipos_arquivo)
        if arquivo:
            self.functions.arquivos_selecionados = [arquivo]
            if self.functions.lbl_entrada_status:
                self.functions.lbl_entrada_status.config(text=f"Arquivo selecionado: {arquivo}")
            self.log(f"Entrada definida: {arquivo}")

    def log(self, mensagem: str):
        if self.log_text:
            self.log_text.config(state='normal')
            self.log_text.insert(tk.END, f"{mensagem}\n")
            self.log_text.see(tk.END)
            self.log_text.config(state='disabled')
        else:
            print(mensagem)

    def iniciar_automacao(self):
        self.functions.limpar_log()
        self.log("--- Iniciando verificação para Registro de Veículo ---")

        arquivos = getattr(self.functions, 'arquivos_selecionados', None)
        if not arquivos:
            self.log("ERRO: Nenhum arquivo de planilha selecionado.")
            messagebox.showerror("Erro", "Por favor, selecione uma planilha de entrada.")
            return

        arquivo_para_processar = arquivos[0]
        self.btn_gerar.config(text="PROCESSANDO...", state='disabled')

        threading.Thread(target=self.processar_em_thread, args=(arquivo_para_processar,), daemon=True).start()

    def processar_em_thread(self, arquivo_path):
        try:
            self.log(f"Iniciando processamento para: {arquivo_path}")
            dados_veiculo = backend.extrair_dados_planilha(arquivo_path, self.log)
            if dados_veiculo:
                self.log(f"Dados extraídos com sucesso: {dados_veiculo['nome']}, {dados_veiculo['placa']}")
                self.log("Iniciando automação web (Selenium)...")
                self.log("VERIFIQUE O CONSOLE (TERMINAL) PARA INSTRUÇÕES DE LOGIN.")
                backend.preencher_formulario_web(dados_veiculo, self.log)
                self.log(f"Processamento do arquivo {arquivo_path} concluído.")
            else:
                self.log(f"ERRO: Não foi possível extrair dados válidos do arquivo {arquivo_path}.")
        except Exception as e:
            self.log(f"\nERRO INESPERADO NA THREAD: {e}")
            self.frame.after(0, lambda: messagebox.showerror("Erro Inesperado", str(e)))
        finally:
            self.frame.after(0, lambda: self.btn_gerar.config(text="REGISTRAR VEÍCULO", state='normal'))
