import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import subprocess
import threading
import sys
import os
import queue
from pathlib import Path


class AutomacaoGUI:
    # Tipos de peça com nome da aba e label
    TIPOS_PECA = [
        ('Pilares', 'Pasta dos Pilares (DWG):'),
        ('Vigas', 'Pasta das Vigas (DWG):'),
        ('Lajes', 'Pasta das Lajes (DWG):'),
    ]

    def __init__(self, root):
        self.root = root
        self.root.title("Automação de Engenharia - Extração de Dados")
        self.root.geometry("750x600")

        # Fila para thread-safe GUI updates
        self.queue = queue.Queue()
        self.check_queue()

        # Variáveis compartilhadas
        self.arquivo_planilha = tk.StringVar()

        # Variáveis por aba (pasta de desenhos de cada tipo)
        self.pastas = {}

        # Layout principal
        main_frame = tk.Frame(root, padx=20, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- Abas de tipos de peça ---
        tk.Label(main_frame, text="1. Pastas dos Desenhos (DWG):",
                 font=("Arial", 10, "bold")).pack(anchor="w")

        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.X, pady=(5, 15))

        for nome_aba, label_texto in self.TIPOS_PECA:
            self._criar_aba(nome_aba, label_texto)

        # --- Seção: Planilha Excel (compartilhada) ---
        tk.Label(main_frame, text="2. Planilha de Controle (.xlsm):",
                 font=("Arial", 10, "bold")).pack(anchor="w")

        frame_excel = tk.Frame(main_frame)
        frame_excel.pack(fill=tk.X, pady=(5, 15))

        self.entry_excel = tk.Entry(frame_excel, textvariable=self.arquivo_planilha, width=50)
        self.entry_excel.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        btn_excel = tk.Button(frame_excel, text="Selecionar Planilha",
                              command=self.selecionar_planilha)
        btn_excel.pack(side=tk.RIGHT)

        # --- Botão Executar ---
        self.btn_executar = tk.Button(main_frame, text="EXECUTAR AUTOMAÇÃO",
                                      command=self.iniciar_automacao,
                                      font=("Arial", 12, "bold"),
                                      bg="#4CAF50", fg="white",
                                      height=2)
        self.btn_executar.pack(fill=tk.X, pady=(0, 15))

        # --- Log de Execução ---
        tk.Label(main_frame, text="Log de Execução:", font=("Arial", 9)).pack(anchor="w")
        self.log_text = scrolledtext.ScrolledText(main_frame, height=15, state='disabled',
                                                   font=("Consolas", 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # Preencher valores padrão
        self._preencher_defaults()

    def _criar_aba(self, nome_aba, label_texto):
        """Cria uma aba no notebook com campo de seleção de pasta."""
        frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(frame, text=nome_aba)

        var_pasta = tk.StringVar()
        self.pastas[nome_aba] = var_pasta

        tk.Label(frame, text=label_texto, font=("Arial", 9)).pack(anchor="w")

        frame_input = tk.Frame(frame)
        frame_input.pack(fill=tk.X, pady=(5, 0))

        entry = tk.Entry(frame_input, textvariable=var_pasta, width=50)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        btn = tk.Button(frame_input, text="Selecionar Pasta",
                        command=lambda v=var_pasta, n=nome_aba: self._selecionar_pasta_tipo(v, n))
        btn.pack(side=tk.RIGHT)

    def _selecionar_pasta_tipo(self, var, nome_tipo):
        pasta = filedialog.askdirectory(title=f"Selecione a pasta com os desenhos de {nome_tipo}")
        if pasta:
            var.set(pasta)

    def _preencher_defaults(self):
        """Preenche caminhos padrão se existirem no diretório atual."""
        cwd = os.getcwd()

        # Planilha padrão
        planilha_padrao = os.path.join(cwd, "CONTROLE DE PEÇAS - CROMA.xlsm")
        if os.path.exists(planilha_padrao):
            self.arquivo_planilha.set(planilha_padrao)

    def selecionar_planilha(self):
        arquivo = filedialog.askopenfilename(
            title="Selecione a planilha Excel",
            filetypes=[("Arquivos Excel com Macro", "*.xlsm"), ("Arquivos Excel", "*.xlsx")]
        )
        if arquivo:
            self.arquivo_planilha.set(arquivo)

    def log(self, mensagem):
        self.queue.put(mensagem)

    def check_queue(self):
        while not self.queue.empty():
            mensagem = self.queue.get()
            self.log_text.config(state='normal')
            self.log_text.insert(tk.END, mensagem + "\n")
            self.log_text.see(tk.END)
            self.log_text.config(state='disabled')
        self.root.after(100, self.check_queue)

    def iniciar_automacao(self):
        excel = self.arquivo_planilha.get()

        if not excel or not os.path.exists(excel):
            messagebox.showerror("Erro", "Selecione um arquivo de planilha válido.")
            return

        # Coletar pastas preenchidas
        pastas_para_processar = []
        for nome_aba, var_pasta in self.pastas.items():
            pasta = var_pasta.get().strip()
            if pasta and os.path.exists(pasta):
                pastas_para_processar.append((nome_aba, pasta))

        if not pastas_para_processar:
            messagebox.showerror("Erro",
                                 "Preencha pelo menos uma pasta de desenhos válida.")
            return

        # Desabilitar interface
        self.btn_executar.config(state='disabled', text="Executando... (Aguarde)")
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')

        tipos_str = ", ".join(nome for nome, _ in pastas_para_processar)
        self.log(f"Iniciando automação para: {tipos_str}")
        self.log(f"Planilha: {excel}")
        self.log("-" * 50)

        thread = threading.Thread(target=self.executar_processo,
                                  args=(pastas_para_processar, excel))
        thread.start()

    def executar_processo(self, pastas_para_processar, excel):
        python_exe = sys.executable
        script_dir = os.path.dirname(os.path.abspath(__file__))
        converter_script = os.path.join(script_dir, "converter_dwg_dxf.py")
        extrair_script = os.path.join(script_dir, "extrair_dados_dxf.py")

        creation_flags = 0
        if os.name == 'nt':
            creation_flags = subprocess.CREATE_NO_WINDOW

        total = len(pastas_para_processar)
        sucesso_total = True

        for idx, (nome_tipo, pasta) in enumerate(pastas_para_processar, 1):
            self.log(f"\n{'='*50}")
            self.log(f">>> [{idx}/{total}] Processando {nome_tipo}: {pasta}")
            self.log(f"{'='*50}")

            try:
                # Etapa 1: Converter DWG -> DXF
                self.log(f"\n  [ETAPA 1/2] Convertendo DWG para DXF ({nome_tipo})...")
                cmd_conv = [python_exe, converter_script, pasta]

                process = subprocess.Popen(cmd_conv, stdout=subprocess.PIPE,
                                           stderr=subprocess.PIPE, text=True,
                                           creationflags=creation_flags)

                while True:
                    line = process.stdout.readline()
                    if not line and process.poll() is not None:
                        break
                    if line:
                        self.log("  " + line.strip())

                if process.returncode != 0:
                    self.log(f"  ERRO NA CONVERSÃO ({nome_tipo})!")
                    stderr = process.stderr.read()
                    self.log(stderr)
                    sucesso_total = False
                    continue

                self.log(f"  Conversão concluída ({nome_tipo}).")

                # Etapa 2: Extrair Dados e Atualizar Planilha
                self.log(f"\n  [ETAPA 2/2] Extraindo dados e atualizando planilha ({nome_tipo})...")
                cmd_ext = [python_exe, extrair_script, pasta, excel]

                process = subprocess.Popen(cmd_ext, stdout=subprocess.PIPE,
                                           stderr=subprocess.PIPE, text=True,
                                           creationflags=creation_flags)

                while True:
                    line = process.stdout.readline()
                    if not line and process.poll() is not None:
                        break
                    if line:
                        self.log("  " + line.strip())

                if process.returncode != 0:
                    self.log(f"  ERRO NA EXTRAÇÃO ({nome_tipo})!")
                    stderr = process.stderr.read()
                    self.log(stderr)
                    sucesso_total = False
                    continue

                self.log(f"  {nome_tipo} processados com sucesso.")

            except Exception as e:
                self.log(f"\n  ERRO CRÍTICO ({nome_tipo}): {str(e)}")
                sucesso_total = False

        self.log(f"\n{'='*50}")
        if sucesso_total:
            self.log("AUTOMAÇÃO CONCLUÍDA COM SUCESSO!")
        else:
            self.log("AUTOMAÇÃO CONCLUÍDA COM ERROS. Verifique o log acima.")
        self.finalizar_thread_safe(sucesso_total)

    def finalizar_thread_safe(self, sucesso):
        self.root.after(0, lambda: self.finalizar(sucesso))

    def finalizar(self, sucesso):
        self.btn_executar.config(state='normal', text="EXECUTAR AUTOMAÇÃO")
        if sucesso:
            messagebox.showinfo("Sucesso", "Processo concluído com sucesso!\nVerifique a planilha.")
        else:
            messagebox.showerror("Erro", "Ocorreram erros durante a execução.\nVerifique o log.")


if __name__ == "__main__":
    root = tk.Tk()
    app = AutomacaoGUI(root)
    root.mainloop()
