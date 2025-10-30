import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import etiqueta_backend as backend # Importa o backend
import threading

class AppGeradorEtiquetas:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerador de Etiquetas Cebraspe")
        self.root.geometry("700x550")
        
        self.arquivos_selecionados = []
        self.pasta_saida = ""
        self.nome_arquivo_saida = "Cebraspe_Etiquetas_Geradas.docx"

        # --- Estilo ---
        style = ttk.Style()
        style.configure('TButton', font=('Helvetica', 10, 'bold'), padding=10)
        style.configure('TLabel', font=('Helvetica', 10), padding=5)
        style.configure('TFrame', padding=10)
        style.configure('Header.TLabel', font=('Helvetica', 12, 'bold'))

        # --- Container Principal ---
        main_frame = ttk.Frame(root)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- 1. Seção de Entrada ---
        frame_entrada = ttk.LabelFrame(main_frame, text="1. Selecionar Arquivos de Entrada", padding=15)
        frame_entrada.pack(fill=tk.X, padx=10, pady=10)

        btn_arquivos = ttk.Button(frame_entrada, text="Selecionar Arquivos (.docx, .xlsx)", command=self.selecionar_arquivos)
        btn_arquivos.pack(fill=tk.X, pady=5)
        
        btn_pasta = ttk.Button(frame_entrada, text="Selecionar Pasta (contém .docx, .xlsx)", command=self.selecionar_pasta)
        btn_pasta.pack(fill=tk.X, pady=5)
        
        self.lbl_entrada_status = ttk.Label(frame_entrada, text="Nenhuma entrada selecionada.", wraplength=650)
        self.lbl_entrada_status.pack(fill=tk.X, pady=5)

        # --- 2. Seção de Saída ---
        frame_saida = ttk.LabelFrame(main_frame, text="2. Selecionar Local de Saída", padding=15)
        frame_saida.pack(fill=tk.X, padx=10, pady=5)

        btn_saida = ttk.Button(frame_saida, text="Selecionar Pasta e Nome do Arquivo", command=self.selecionar_saida)
        btn_saida.pack(fill=tk.X, pady=5)

        self.lbl_saida_status = ttk.Label(frame_saida, text=f"Será salvo como: {self.nome_arquivo_saida}", wraplength=650)
        self.lbl_saida_status.pack(fill=tk.X, pady=5)

        # --- 3. Seção de Ação ---
        frame_acao = ttk.Frame(main_frame)
        frame_acao.pack(fill=tk.X, padx=10, pady=10)
        
        self.btn_gerar = ttk.Button(frame_acao, text="GERAR ETIQUETAS", command=self.iniciar_processamento)
        self.btn_gerar.pack(fill=tk.X, ipady=10)

        # --- 4. Seção de Log (Mensagens) ---
        frame_log = ttk.LabelFrame(main_frame, text="Mensagens do Processo", padding=10)
        frame_log.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        scrollbar = ttk.Scrollbar(frame_log)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.log_text = tk.Text(frame_log, height=50, wrap=tk.WORD, yscrollcommand=scrollbar.set, state='disabled', font=('Courier New', 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.log_text.yview)

    def log(self, message):
        """ Adiciona uma mensagem ao widget de log na thread principal. """
        def _log():
            self.log_text.config(state='normal')
            self.log_text.insert(tk.END, message + "\n")
            self.log_text.see(tk.END) # Auto-scroll
            self.log_text.config(state='disabled')
            self.log_text.update_idletasks()
        
        # Garante que a atualização da GUI seja feita na thread principal
        self.root.after(0, _log)

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
            self.arquivos_selecionados = [pasta] # A lógica de processamento tratará como pasta
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
            
    def iniciar_processamento(self):
        self.limpar_log()

        if not self.arquivos_selecionados:
            self.log("ERRO: Por favor, selecione arquivos ou uma pasta de entrada.")
            messagebox.showerror("Erro", "Nenhuma entrada selecionada.")
            return

        if not hasattr(self, 'caminho_arquivo_saida_completo') or not self.caminho_arquivo_saida_completo:
            self.log("ERRO: Por favor, selecione um local para salvar o arquivo de saída.")
            messagebox.showerror("Erro", "Nenhum local de saída selecionado.")
            return

        # Desabilita o botão para evitar cliques duplos
        self.btn_gerar.config(text="PROCESSANDO...", state='disabled')

        # Inicia o processamento em uma thread separada para não travar a GUI
        threading.Thread(
            target=self.processar_em_thread,
            args=(self.arquivos_selecionados, self.caminho_arquivo_saida_completo),
            daemon=True
        ).start()

    def processar_em_thread(self, entradas, saida):
        try:
            self.log("Iniciando processamento...")
            
            # 1. Processar entradas (backend)
            # A função self.log é passada como o 'logger'
            dados = backend.processar_entradas(entradas, self.log)
            
            # 2. Gerar documento (backend)
            if dados:
                backend.gerar_documento_word(dados, saida, self.log)
            else:
                self.log("Processamento concluído, mas nenhum dado de etiqueta foi encontrado.")

        except Exception as e:
            self.log(f"\nERRO INESPERADO: {e}")
            # Mostra o erro em um popup também
            self.root.after(0, lambda: messagebox.showerror("Erro Inesperado", str(e)))
        finally:
            # Reabilita o botão na thread principal
            self.root.after(0, lambda: self.btn_gerar.config(text="GERAR ETIQUETAS", state='normal'))


if __name__ == "__main__":
    # Verifica se as dependências estão instaladas
    root = tk.Tk()
    app = AppGeradorEtiquetas(root)
    root.mainloop()
