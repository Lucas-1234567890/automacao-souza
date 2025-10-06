import warnings
from time import sleep, time
import pandas as pd
import pyautogui
import os
import pyperclip
from openpyxl import load_workbook
import threading
from datetime import datetime
from ttkbootstrap import Window, Style
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox, scrolledtext
from ttkbootstrap import Frame, Label, Button, Progressbar, Notebook, Text
import tkinter as tk
from tkinter import ttk
import json

class SistemaAutomacaoSIC:
    def __init__(self):
        self.app = Window(themename="superhero")
        self.arquivo_selecionado = None
        self.automacao_ativa = False
        self.dados_preview = None
        self.log_operacoes = []
        
        self.setup_window()
        self.create_widgets()
        self.carregar_configuracoes()
        
    def setup_window(self):
        self.app.title("🤖 Sistema de Automação SIC - Auto Elétrica Souza")
        self.app.geometry("1000x700")
        self.app.minsize(950, 650)
        
        # Centralizar janela
        largura_janela = 1000
        altura_janela = 700
        largura_tela = self.app.winfo_screenwidth()
        altura_tela = self.app.winfo_screenheight()
        x_pos = (largura_tela // 2) - (largura_janela // 2)
        y_pos = (altura_tela // 2) - (altura_janela // 2)
        self.app.geometry(f"{largura_janela}x{altura_janela}+{x_pos}+{y_pos}")
        
        # Configurar grid principal
        self.app.columnconfigure(0, weight=1)
        self.app.rowconfigure(1, weight=1)
        
    def create_widgets(self):
        # Header
        self.create_header()
        
        # Notebook (abas)
        self.notebook = Notebook(self.app, bootstyle="info")
        self.notebook.grid(row=1, column=0, sticky="nsew", padx=20, pady=10)
        
        # Aba 1: Configuração e Execução
        self.create_main_tab()
        
        # Aba 2: Preview dos Dados
        self.create_preview_tab()
        
        # Aba 3: Logs e Histórico
        self.create_logs_tab()
        
        # Aba 4: Configurações
        self.create_config_tab()
        
        # Status bar
        self.create_status_bar()
        
    def create_header(self):
        header_frame = Frame(self.app)
        header_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=20)
        
        # Logo/Título
        title_frame = Frame(header_frame)
        title_frame.pack()
        
        Label(title_frame, text="🤖", font=("Segoe UI", 30)).pack(side="left", padx=(0, 10))
        
        title_info = Frame(title_frame)
        title_info.pack(side="left")
        
        Label(title_info, text="Sistema de Automação SIC", 
              font=("Segoe UI", 18, "bold")).pack(anchor="w")
        Label(title_info, text="Auto Elétrica Souza - Cadastro Automatizado", 
              font=("Segoe UI", 10), bootstyle="info").pack(anchor="w")
        
    def create_main_tab(self):
        main_frame = Frame(self.notebook)
        self.notebook.add(main_frame, text="🚀 Execução Principal")
        
        # Grid configuration
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Left panel - Seleção de arquivo
        left_panel = Frame(main_frame)
        left_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=10)
        
        # Card: Seleção de Planilha
        card_arquivo = Frame(left_panel, bootstyle="info", padding=20)
        card_arquivo.pack(fill="x", pady=(0, 20))
        
        Label(card_arquivo, text="📊 Seleção da Planilha", 
              font=("Segoe UI", 14, "bold")).pack(anchor="w")
        
        Label(card_arquivo, text="Selecione a planilha Excel com os dados para automação", 
              font=("Segoe UI", 9), bootstyle="info").pack(anchor="w", pady=(0, 15))
        
        Button(card_arquivo, text="📂 Selecionar Arquivo Excel", 
               command=self.selecionar_arquivo, bootstyle="primary",
               width=25).pack()
        
        self.lbl_arquivo = Label(card_arquivo, text="Nenhum arquivo selecionado", 
                                font=("Segoe UI", 9), bootstyle="secondary")
        self.lbl_arquivo.pack(pady=(10, 0))
        
        # Card: Informações do arquivo
        self.card_info = Frame(left_panel, bootstyle="light", padding=20)
        self.card_info.pack(fill="x", pady=(0, 20))
        
        Label(self.card_info, text="📋 Informações do Arquivo", 
              font=("Segoe UI", 14, "bold")).pack(anchor="w")
        
        self.info_text = Text(self.card_info, height=8, font=("Consolas", 9))
        self.info_text.pack(fill="both", expand=True, pady=(10, 0))
        self.info_text.insert("1.0", "Selecione um arquivo para ver as informações...")
        self.info_text.config(state="disabled")
        
        # Right panel - Controles de execução
        right_panel = Frame(main_frame)
        right_panel.grid(row=0, column=1, sticky="nsew", padx=(10, 0), pady=10)
        
        # Card: Execução
        card_exec = Frame(right_panel, bootstyle="success", padding=20)
        card_exec.pack(fill="x", pady=(0, 20))
        
        Label(card_exec, text="▶️ Controles de Execução", 
              font=("Segoe UI", 14, "bold")).pack(anchor="w")
        
        Label(card_exec, text="Execute a automação após verificar os dados", 
              font=("Segoe UI", 9), bootstyle="info").pack(anchor="w", pady=(0, 15))
        
        self.btn_iniciar = Button(card_exec, text="🚀 Iniciar Automação", 
                                 state="disabled", bootstyle="success-outline",
                                 command=self.iniciar_automacao_thread, width=25)
        self.btn_iniciar.pack(pady=(0, 10))
        
        self.btn_parar = Button(card_exec, text="⏹️ Parar Automação", 
                               state="disabled", bootstyle="danger-outline",
                               command=self.parar_automacao, width=25)
        self.btn_parar.pack()
        
        # Progress frame
        progress_frame = Frame(card_exec)
        progress_frame.pack(fill="x", pady=(15, 0))
        
        Label(progress_frame, text="Progresso:", font=("Segoe UI", 9, "bold")).pack(anchor="w")
        
        self.progress = Progressbar(progress_frame, mode='determinate', bootstyle="success")
        self.progress.pack(fill="x", pady=(5, 0))
        
        self.lbl_progress = Label(progress_frame, text="Aguardando início...", 
                                 font=("Segoe UI", 8), bootstyle="secondary")
        self.lbl_progress.pack(anchor="w", pady=(5, 0))
        
        # Card: Estatísticas em tempo real
        card_stats = Frame(right_panel, bootstyle="warning", padding=20)
        card_stats.pack(fill="both", expand=True)
        
        Label(card_stats, text="📊 Estatísticas da Execução", 
              font=("Segoe UI", 14, "bold")).pack(anchor="w")
        
        stats_grid = Frame(card_stats)
        stats_grid.pack(fill="x", pady=(10, 0))
        
        # Configurar grid 2x2
        for i in range(2):
            stats_grid.columnconfigure(i, weight=1)
            
        self.stats = {}
        self.create_stat_item(stats_grid, "Processados", "0", 0, 0)
        self.create_stat_item(stats_grid, "Sucessos", "0", 0, 1)
        self.create_stat_item(stats_grid, "Erros", "0", 1, 0)
        self.create_stat_item(stats_grid, "Tempo", "00:00", 1, 1)
        
    def create_stat_item(self, parent, label, value, row, col):
        frame = Frame(parent, bootstyle="light", padding=10)
        frame.grid(row=row, column=col, sticky="ew", padx=5, pady=5)
        
        Label(frame, text=label, font=("Segoe UI", 9)).pack()
        stat_label = Label(frame, text=value, font=("Segoe UI", 16, "bold"), 
                          bootstyle="primary")
        stat_label.pack()
        
        self.stats[label.lower()] = stat_label
        
    def create_preview_tab(self):
        preview_frame = Frame(self.notebook)
        self.notebook.add(preview_frame, text="👁️ Preview dos Dados")
        
        # Top controls
        controls_frame = Frame(preview_frame)
        controls_frame.pack(fill="x", padx=20, pady=20)
        
        Label(controls_frame, text="📋 Preview dos Dados da Planilha", 
              font=("Segoe UI", 16, "bold")).pack(side="left")
        
        Button(controls_frame, text="🔄 Atualizar Preview", 
               command=self.atualizar_preview, bootstyle="info-outline").pack(side="right")
        
        # Treeview para mostrar dados
        tree_frame = Frame(preview_frame)
        tree_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)
        
        # Treeview
        columns = ('Gerador', 'Data', 'Quantidade', 'ID Interno', 'Status')
        self.tree_preview = ttk.Treeview(tree_frame, columns=columns, show='headings')
        
        for col in columns:
            self.tree_preview.heading(col, text=col)
            self.tree_preview.column(col, width=120)
        
        # Scrollbars
        v_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree_preview.yview)
        h_scroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree_preview.xview)
        self.tree_preview.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        
        self.tree_preview.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")
        
    def create_logs_tab(self):
        logs_frame = Frame(self.notebook)
        self.notebook.add(logs_frame, text="📝 Logs e Histórico")
        
        # Top controls
        controls_frame = Frame(logs_frame)
        controls_frame.pack(fill="x", padx=20, pady=20)
        
        Label(controls_frame, text="📝 Log de Operações", 
              font=("Segoe UI", 16, "bold")).pack(side="left")
        
        Button(controls_frame, text="🗑️ Limpar Logs", 
               command=self.limpar_logs, bootstyle="danger-outline").pack(side="right", padx=(0, 10))
        
        Button(controls_frame, text="💾 Salvar Logs", 
               command=self.salvar_logs, bootstyle="info-outline").pack(side="right")
        
        # Log text area
        self.log_text = scrolledtext.ScrolledText(logs_frame, font=("Consolas", 9), 
                                                 state="disabled", height=25)
        self.log_text.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        
    def create_config_tab(self):
        config_frame = Frame(self.notebook)
        self.notebook.add(config_frame, text="⚙️ Configurações")
        
        # Configurações do sistema
        Label(config_frame, text="⚙️ Configurações do Sistema", 
              font=("Segoe UI", 16, "bold")).pack(padx=20, pady=20, anchor="w")
        
        # Card: Caminhos
        card_paths = Frame(config_frame, bootstyle="info", padding=20)
        card_paths.pack(fill="x", padx=20, pady=(0, 20))
        
        Label(card_paths, text="📁 Caminhos do Sistema", 
              font=("Segoe UI", 12, "bold")).pack(anchor="w")
        
        # SIC Path
        sic_frame = Frame(card_paths)
        sic_frame.pack(fill="x", pady=(10, 5))
        
        Label(sic_frame, text="Caminho do SIC:", font=("Segoe UI", 9)).pack(anchor="w")
        self.sic_path_var = tk.StringVar(value=r"C:\Users\Lucas\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\sic.lnk")
        sic_entry_frame = Frame(sic_frame)
        sic_entry_frame.pack(fill="x", pady=(2, 0))
        
        tk.Entry(sic_entry_frame, textvariable=self.sic_path_var, 
                font=("Consolas", 9)).pack(side="left", fill="x", expand=True, padx=(0, 5))
        Button(sic_entry_frame, text="📂", command=self.selecionar_sic, 
               bootstyle="outline", width=3).pack(side="right")
        
        # Logs Path
        logs_frame = Frame(card_paths)
        logs_frame.pack(fill="x", pady=5)
        
        Label(logs_frame, text="Pasta de Logs:", font=("Segoe UI", 9)).pack(anchor="w")
        self.logs_path_var = tk.StringVar(value=r"C:\Users\Lucas\OneDrive\Trabalho\Planilhas de excel\logs_automacao")
        logs_entry_frame = Frame(logs_frame)
        logs_entry_frame.pack(fill="x", pady=(2, 0))
        
        tk.Entry(logs_entry_frame, textvariable=self.logs_path_var, 
                font=("Consolas", 9)).pack(side="left", fill="x", expand=True, padx=(0, 5))
        Button(logs_entry_frame, text="📂", command=self.selecionar_logs, 
               bootstyle="outline", width=3).pack(side="right")
        
        # Card: Configurações de Automação
        card_auto = Frame(config_frame, bootstyle="warning", padding=20)
        card_auto.pack(fill="x", padx=20, pady=(0, 20))
        
        Label(card_auto, text="🤖 Configurações de Automação", 
              font=("Segoe UI", 12, "bold")).pack(anchor="w")
        
        # Timeout
        timeout_frame = Frame(card_auto)
        timeout_frame.pack(fill="x", pady=(10, 5))
        
        Label(timeout_frame, text="Timeout para encontrar imagens (segundos):", 
              font=("Segoe UI", 9)).pack(side="left")
        self.timeout_var = tk.StringVar(value="20")
        tk.Entry(timeout_frame, textvariable=self.timeout_var, width=10, 
                font=("Consolas", 9)).pack(side="right")
        
        # Confidence
        conf_frame = Frame(card_auto)
        conf_frame.pack(fill="x", pady=5)
        
        Label(conf_frame, text="Confiança para reconhecimento (0.1-1.0):", 
              font=("Segoe UI", 9)).pack(side="left")
        self.confidence_var = tk.StringVar(value="0.8")
        tk.Entry(conf_frame, textvariable=self.confidence_var, width=10, 
                font=("Consolas", 9)).pack(side="right")
        
        # Botão salvar configurações
        Button(config_frame, text="💾 Salvar Configurações", 
               command=self.salvar_configuracoes, bootstyle="success", 
               width=25).pack(pady=20)
        
    def create_status_bar(self):
        status_frame = Frame(self.app, bootstyle="dark")
        status_frame.grid(row=2, column=0, sticky="ew")
        
        self.status_label = Label(status_frame, text="✅ Sistema pronto para uso", 
                                 font=("Segoe UI", 9))
        self.status_label.pack(side="left", padx=10, pady=5)
        
        # Info adicional
        info_label = Label(status_frame, text=f"Versão 2.0 | {datetime.now().strftime('%d/%m/%Y')}", 
                          font=("Segoe UI", 8), bootstyle="secondary")
        info_label.pack(side="right", padx=10, pady=5)
    
    # ---------- Métodos de funcionalidade ----------
    
    def adicionar_log(self, mensagem, tipo="INFO"):
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {tipo}: {mensagem}\n"
        
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, log_entry)
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")
        
        self.log_operacoes.append({"timestamp": timestamp, "tipo": tipo, "mensagem": mensagem})
        
    def selecionar_arquivo(self):
        caminho = filedialog.askopenfilename(
            title="Selecione a planilha Excel",
            filetypes=[("Planilhas Excel", "*.xls*")]

        )
        
        if caminho:
            self.arquivo_selecionado = caminho
            nome_arquivo = os.path.basename(caminho)
            self.lbl_arquivo.config(text=f"📄 {nome_arquivo}")
            
            # Habilitar botão de execução
            self.btn_iniciar.config(state="normal")
            
            # Carregar informações do arquivo
            self.carregar_info_arquivo(caminho)
            
            # Atualizar preview
            self.atualizar_preview()
            
            self.adicionar_log(f"Arquivo selecionado: {nome_arquivo}")
            self.status_label.config(text=f"📄 Arquivo carregado: {nome_arquivo}")
    
    def carregar_info_arquivo(self, caminho):
        try:
            # Ler dados básicos
            df = pd.read_excel(caminho, sheet_name="Cadastro de materiais", skiprows=3, usecols="F:L")
            
            info = f"""📊 INFORMAÇÕES DO ARQUIVO
            
📁 Caminho: {caminho}
📋 Planilha: Cadastro de materiais
📏 Total de registros: {len(df)}
🔢 Colunas encontradas: {', '.join(df.columns.tolist())}

📈 ESTATÍSTICAS:
• Geradores únicos: {df['Gerador'].nunique() if 'Gerador' in df.columns else 'N/A'}
• Datas únicas: {df['Data'].nunique() if 'Data' in df.columns else 'N/A'}
• Quantidade total: {df['Quantidade'].sum() if 'Quantidade' in df.columns else 'N/A'}

⚠️  VERIFICAÇÕES:
• Dados nulos: {'✅ Sem dados nulos' if not df.isnull().any().any() else '❌ Existem dados nulos'}
• Formato de data: {'✅ Formato válido' if pd.api.types.is_datetime64_any_dtype(pd.to_datetime(df['Data'], errors='coerce')) else '⚠️ Verificar formato'}
            """
            
            self.info_text.config(state="normal")
            self.info_text.delete("1.0", tk.END)
            self.info_text.insert("1.0", info)
            self.info_text.config(state="disabled")
            
        except Exception as e:
            self.info_text.config(state="normal")
            self.info_text.delete("1.0", tk.END)
            self.info_text.insert("1.0", f"❌ Erro ao carregar informações:\n{str(e)}")
            self.info_text.config(state="disabled")
    
    def atualizar_preview(self):
        if not self.arquivo_selecionado:
            return
            
        try:
            # Limpar treeview
            for item in self.tree_preview.get_children():
                self.tree_preview.delete(item)
            
            # Carregar dados
            df = pd.read_excel(self.arquivo_selecionado, sheet_name="Cadastro de materiais", 
                              skiprows=3, usecols="F:L")
            df["Data"] = pd.to_datetime(df["Data"]).dt.strftime("%d/%m/%Y")
            df["Status"] = "Pendente"
            
            self.dados_preview = df
            
            # Inserir no treeview (limitado a 1000 registros para performance)
            for idx, row in df.head(1000).iterrows():
                values = (
                    str(row.get('Gerador', '')),
                    str(row.get('Data', '')),
                    str(row.get('Quantidade', '')),
                    str(row.get('ID Interno', '')),
                    str(row.get('Status', ''))
                )
                self.tree_preview.insert('', 'end', values=values)
            
            self.adicionar_log(f"Preview atualizado: {len(df)} registros carregados")
            
        except Exception as e:
            self.adicionar_log(f"Erro ao atualizar preview: {str(e)}", "ERRO")
    
    def iniciar_automacao_thread(self):
        if not self.arquivo_selecionado:
            messagebox.showwarning("Aviso", "Selecione um arquivo Excel primeiro!")
            return
            
        # Executar em thread separada
        thread = threading.Thread(target=self.iniciar_automacao)
        thread.daemon = True
        thread.start()
    
    def iniciar_automacao(self):
        try:
            self.automacao_ativa = True
            self.btn_iniciar.config(state="disabled")
            self.btn_parar.config(state="normal")
            
            # Reset estatísticas
            self.stats['processados'].config(text="0")
            self.stats['sucessos'].config(text="0")
            self.stats['erros'].config(text="0")
            self.stats['tempo'].config(text="00:00")
            
            inicio_tempo = time()
            
            self.adicionar_log("🚀 Iniciando automação...")
            self.status_label.config(text="🚀 Executando automação...")
            
            # Carregar dados
            if self.arquivo_selecionado.endswith('.xlsm'):
                engine = 'openpyxl'  # Para arquivos com macros
            else:
                engine = None  # Deixa o pandas escolher automaticamente
                
            tabela = pd.read_excel(self.arquivo_selecionado, sheet_name="Cadastro de materiais", 
                                  skiprows=3, usecols="F:L", engine=engine)
            tabela["Data"] = pd.to_datetime(tabela["Data"]).dt.strftime("%d/%m/%Y")
            tabela["Status"] = "Nao"
            
            total_registros = len(tabela)
            self.progress.config(maximum=total_registros)
            
            # Inicializar PyAutoGUI
            pyautogui.FAILSAFE = True
            
            # Abrir SIC
            self.adicionar_log("📂 Abrindo sistema SIC...")
            os.startfile(self.sic_path_var.get())
            sleep(5)
            
            # Login
            self.adicionar_log("🔐 Realizando login...")
            pyautogui.write("123456")
            sleep(1)
            pyautogui.press("tab")
            sleep(0.5)
            pyautogui.press("enter")
            sleep(0.5)
            pyautogui.press('enter')
            sleep(1)
            pyautogui.hotkey("ctrl", "e")
            sleep(0.5)
            
            # Navegar no sistema
            self.adicionar_log("🧭 Navegando no sistema...")
            pyautogui.click(pyautogui.center(self.encontrar_imagem(os.path.join('imagens', 'atualizacao.png'))))
            sleep(0.5)
            pyautogui.click(pyautogui.center(self.encontrar_imagem(os.path.join('imagens', 'saida.png'))))
            sleep(3)
            pyautogui.click(self.esquerda(self.encontrar_imagem(os.path.join('imagens', 'outros.png'))))
            sleep(0.5)
            pyautogui.click(pyautogui.center(self.encontrar_imagem(os.path.join('imagens', 'max.png'))))
            sleep(1)
            
            # Processar dados
            grupos = tabela.groupby(["Gerador", "Data"])
            processados = 0
            sucessos = 0
            erros = 0
            
            for (gerador, data), grupo in grupos:
                if not self.automacao_ativa:
                    break
                    
                try:
                    self.adicionar_log(f"📊 Processando gerador {gerador} - Data {data}")
                    
                    # Lógica de automação (mantida do código original)
                    sleep(1)
                    pyautogui.click(self.direita(self.encontrar_imagem(os.path.join('imagens', 'gerador.png'))))
                    sleep(0.8)
                    pyautogui.write(str(gerador))
                    sleep(0.8)
                    pyautogui.press('enter')
                    sleep(1.5)
                    pyautogui.click(pyautogui.center(self.encontrar_imagem(os.path.join('imagens', 'souza.png'))))
                    sleep(0.8)
                    pyautogui.doubleClick(79, 60)
                    sleep(1)
                    self.escrever_texto(str(data))
                    sleep(1.5)
                    
                    pyautogui.click(pyautogui.center(self.encontrar_imagem(os.path.join('imagens', 'quantidade.png'))))
                    sleep(0.5)
                    
                    for idx, linha in grupo.iterrows():
                        if not self.automacao_ativa:
                            break
                            
                        try:
                            self.escrever_texto(str(int(linha["Quantidade"])))

                            sleep(0.3)
                            pyautogui.press('tab')
                            sleep(0.3)
                            pyautogui.write(str(int(linha["ID Interno"])).zfill(6))
                            sleep(0.3)
                            pyautogui.press('tab')
                            sleep(0.3)
                            pyautogui.press('tab')
                            sleep(0.3)
                            pyautogui.press('tab')
                            sleep(1)
                            
                            imagem = pyautogui.locateCenterOnScreen(os.path.join('imagens', 'sim.png'), confidence=0.9)
                            if imagem:
                                pyautogui.click(imagem.x, imagem.y)
                            sleep(0.3)
                            
                            tabela.at[idx, "Status"] = "Sim"
                            sucessos += 1
                            
                        except Exception as e:
                            self.adicionar_log(f"❌ Erro ao processar linha: {str(e)}", "ERRO")
                            tabela.at[idx, "Status"] = "Nao"
                            erros += 1
                        
                        processados += 1
                        
                        # Atualizar interface
                        self.app.after(0, self.atualizar_estatisticas, processados, sucessos, erros, inicio_tempo)
                        self.app.after(0, lambda: self.progress.config(value=processados))
                    
                    # Gravar
                    sleep(1)
                    pyautogui.click(pyautogui.center(self.encontrar_imagem(os.path.join('imagens', 'gravar.png'))))
                    sleep(1.5)
                    imagem_2 = pyautogui.locateOnScreen(os.path.join('imagens', 'sim_nao.png'), confidence=0.9)
                    sleep(0.3)
                    if imagem_2:
                        sleep(1)
                        pyautogui.click(self.esquerda(imagem_2))
                    sleep(1)
                    sleep(0.8)
                    
                except Exception as e:
                    self.adicionar_log(f"❌ Erro no grupo ({gerador}, {data}): {str(e)}", "ERRO")
                    erros += 1
            
            # Finalizar automação
            self.finalizar_automacao(tabela, processados, sucessos, erros, inicio_tempo)
            
        except Exception as e:
            self.adicionar_log(f"❌ Erro crítico na automação: {str(e)}", "ERRO")
            messagebox.showerror("Erro", f"Ocorreu um erro crítico:\n{str(e)}")
        finally:
            self.automacao_ativa = False
            self.btn_iniciar.config(state="normal")
            self.btn_parar.config(state="disabled")
    
    def finalizar_automacao(self, tabela, processados, sucessos, erros, inicio_tempo):
        # Salvar log
        pasta_logs = self.logs_path_var.get()
        os.makedirs(pasta_logs, exist_ok=True)
        
        data_hora_execucao = datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
        saida_path = os.path.join(pasta_logs, f"log_{data_hora_execucao}.xlsx")
        
        tabela["Data Registro"] = pd.Timestamp.now()
        tabela.to_excel(saida_path, index=False)
        
        tempo_total = time() - inicio_tempo
        tempo_str = f"{int(tempo_total//60):02d}:{int(tempo_total%60):02d}"
        
        self.adicionar_log("✅ Automação finalizada com sucesso!")
        self.adicionar_log(f"📊 Estatísticas finais: {processados} processados, {sucessos} sucessos, {erros} erros")
        self.adicionar_log(f"💾 Log salvo em: {saida_path}")
        self.adicionar_log(f"⏱️ Tempo total: {tempo_str}")
        
        self.status_label.config(text="✅ Automação concluída com sucesso!")
        
        messagebox.showinfo("✅ Concluído", 
                           f"Automação finalizada!\n\n"
                           f"📊 Processados: {processados}\n"
                           f"✅ Sucessos: {sucessos}\n"
                           f"❌ Erros: {erros}\n"
                           f"⏱️ Tempo: {tempo_str}\n\n"
                           f"💾 Log salvo em:\n{saida_path}")
    
    def atualizar_estatisticas(self, processados, sucessos, erros, inicio_tempo):
        self.stats['processados'].config(text=str(processados))
        self.stats['sucessos'].config(text=str(sucessos))
        self.stats['erros'].config(text=str(erros))
        
        tempo_decorrido = time() - inicio_tempo
        tempo_str = f"{int(tempo_decorrido//60):02d}:{int(tempo_decorrido%60):02d}"
        self.stats['tempo'].config(text=tempo_str)
        
        # Atualizar label de progresso
        self.lbl_progress.config(text=f"Processados: {processados} | Sucessos: {sucessos} | Erros: {erros}")
    
    def parar_automacao(self):
        if messagebox.askyesno("⏹️ Parar Automação", 
                              "Deseja realmente parar a automação?\n\nEsta ação não pode ser desfeita."):
            self.automacao_ativa = False
            self.adicionar_log("⏹️ Automação interrompida pelo usuário", "AVISO")
            self.status_label.config(text="⏹️ Automação interrompida")
    
    def limpar_logs(self):
        if messagebox.askyesno("🗑️ Limpar Logs", "Deseja limpar todos os logs?"):
            self.log_text.config(state="normal")
            self.log_text.delete("1.0", tk.END)
            self.log_text.config(state="disabled")
            self.log_operacoes.clear()
            self.adicionar_log("🗑️ Logs limpos")
    
    def salvar_logs(self):
        if not self.log_operacoes:
            messagebox.showinfo("Info", "Não há logs para salvar.")
            return
            
        filename = filedialog.asksaveasfilename(
            title="Salvar logs",
            defaultextension=".txt",
            filetypes=[("Arquivos de texto", "*.txt"), ("Todos os arquivos", "*.*")]
        )
        
        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(f"=== LOG DE AUTOMAÇÃO SIC ===\n")
                    f.write(f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
                    f.write(f"Arquivo processado: {self.arquivo_selecionado or 'N/A'}\n")
                    f.write("=" * 50 + "\n\n")
                    
                    for log in self.log_operacoes:
                        f.write(f"[{log['timestamp']}] {log['tipo']}: {log['mensagem']}\n")
                
                messagebox.showinfo("✅ Sucesso", f"Logs salvos em:\n{filename}")
                self.adicionar_log(f"💾 Logs salvos em: {filename}")
                
            except Exception as e:
                messagebox.showerror("❌ Erro", f"Erro ao salvar logs:\n{str(e)}")
    
    def selecionar_sic(self):
        filename = filedialog.askopenfilename(
            title="Selecionar executável do SIC",
            filetypes=[("Atalhos", "*.lnk"), ("Executáveis", "*.exe"), ("Todos os arquivos", "*.*")]
        )
        if filename:
            self.sic_path_var.set(filename)
    
    def selecionar_logs(self):
        folder = filedialog.askdirectory(title="Selecionar pasta de logs")
        if folder:
            self.logs_path_var.set(folder)
    
    def salvar_configuracoes(self):
        config = {
            "sic_path": self.sic_path_var.get(),
            "logs_path": self.logs_path_var.get(),
            "timeout": self.timeout_var.get(),
            "confidence": self.confidence_var.get()
        }
        
        try:
            with open("config_automacao.json", "w", encoding="utf-8") as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
            
            messagebox.showinfo("✅ Sucesso", "Configurações salvas com sucesso!")
            self.adicionar_log("💾 Configurações salvas")
            
        except Exception as e:
            messagebox.showerror("❌ Erro", f"Erro ao salvar configurações:\n{str(e)}")
    
    def carregar_configuracoes(self):
        try:
            if os.path.exists("config_automacao.json"):
                with open("config_automacao.json", "r", encoding="utf-8") as f:
                    config = json.load(f)
                
                self.sic_path_var.set(config.get("sic_path", self.sic_path_var.get()))
                self.logs_path_var.set(config.get("logs_path", self.logs_path_var.get()))
                self.timeout_var.set(config.get("timeout", "20"))
                self.confidence_var.set(config.get("confidence", "0.8"))
                
                self.adicionar_log("📂 Configurações carregadas")
        except Exception as e:
            self.adicionar_log(f"⚠️ Erro ao carregar configurações: {str(e)}", "AVISO")
    
    # ---------- Funções auxiliares (mantidas do código original) ----------
    
    def encontrar_imagem(self, imagem):
        timeout = int(self.timeout_var.get())
        inicio = time()
        encontrou = None
        while True:
            try:
                confidence = float(self.confidence_var.get())
                encontrou = pyautogui.locateOnScreen(imagem, grayscale=True, confidence=confidence)
                if encontrou:
                    break
            except Exception:
                pass
            if time() - inicio > timeout:
                self.adicionar_log(f'⏱️ Timeout atingido para imagem: {imagem}', "AVISO")
                break
            sleep(1)
        return encontrou
    
    def direita(self, posicoes_imagem):
        return posicoes_imagem[0] + posicoes_imagem[2], posicoes_imagem[1] + posicoes_imagem[3] / 2
    
    def esquerda(self, posicao_imagem, deslocamento=5):
        return posicao_imagem[0] + deslocamento, posicao_imagem[1] + posicao_imagem[3] / 2
    
    def escrever_texto(self, texto):
        pyperclip.copy(texto)
        pyautogui.hotkey('ctrl', 'v')
    
    def executar(self):
        self.adicionar_log("🚀 Sistema de Automação SIC iniciado")
        self.app.mainloop()

# ---------- Execução da aplicação ----------
if __name__ == "__main__":
    app = SistemaAutomacaoSIC()
    app.executar()