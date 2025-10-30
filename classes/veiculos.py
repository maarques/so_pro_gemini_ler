import tkinter as tk
import backend.veiculo as backend
from tkinter import ttk
from classes.filesFunctions import FilesFunctions

class AppCadastroVeiculo:
    def __init__(self, parent_tab):
        self.functions = FilesFunctions()
        self.frame = ttk.Frame(parent_tab, padding=10)
        self.frame.pack(fill=tk.BOTH, expand=True)

        main_frame = ttk.Frame(self.frame)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- 1. Seção de Entrada ---
        frame_entrada = ttk.LabelFrame(main_frame, text="1. Selecionar Arquivos de Entrada", padding=15)
        frame_entrada.pack(fill=tk.X, padx=10, pady=10)

        btn_arquivos = ttk.Button(frame_entrada, text="Selecionar Arquivos (.docx, .xlsx)", command=self.functions.selecionar_arquivos)
        btn_arquivos.pack(fill=tk.X, pady=5)

        btn_pasta = ttk.Button(frame_entrada, text="Selecionar Pasta (contém .docx, .xlsx)", command=self.functions.selecionar_pasta)
        btn_pasta.pack(fill=tk.X, pady=5)
        
        self.lbl_entrada_status = ttk.Label(frame_entrada, text="Nenhuma entrada selecionada.", wraplength=650)
        self.lbl_entrada_status.pack(fill=tk.X, pady=5)

        # --- 2. Seção de Ação ---
        frame_acao = ttk.Frame(main_frame)
        frame_acao.pack(fill=tk.X, padx=10, pady=10)
        
        self.btn_gerar = ttk.Button(frame_acao, text="REGISTRAR VEÍCULO")
        self.btn_gerar.pack(fill=tk.X, ipady=10)

        # --- 3. Seção de Log (Mensagens) ---
        frame_log = ttk.LabelFrame(main_frame, text="Mensagens do Processo", padding=10)
        frame_log.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        scrollbar = ttk.Scrollbar(frame_log)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.log_text = tk.Text(frame_log, height=200, wrap=tk.WORD, yscrollcommand=scrollbar.set, state='normal', font=('Courier New', 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.log_text.yview)
