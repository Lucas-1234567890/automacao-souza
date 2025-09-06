import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from datetime import datetime
import urllib
import os
from time import sleep
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import threading

class RelatorioGeradoresApp:
    def __init__(self, root):
        self.root = root
        self.mensagens = []
        self.numero_fixo = "71999740292"
        
        self.setup_window()
        self.create_widgets()
        
    def setup_window(self):
        self.root.title("📊 Relatório de Geradores - Sistema Automatizado")
        self.root.geometry("800x700")
        self.root.minsize(750, 650)
        
        # Configurar estilo
        style = ttk.Style()
        style.theme_use('clam')
        
        # Cores personalizadas
        style.configure('Title.TLabel', font=('Segoe UI', 16, 'bold'), foreground='#2c3e50')
        style.configure('Heading.TLabel', font=('Segoe UI', 10, 'bold'), foreground='#34495e')
        style.configure('Info.TLabel', font=('Segoe UI', 9), foreground='#7f8c8d')
        
        # Configurar grid
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)
        
    def create_widgets(self):
        # Header
        header_frame = ttk.Frame(self.root)
        header_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 10))
        
        title_label = ttk.Label(header_frame, text="📊 Sistema de Relatórios de Geradores", 
                               style='Title.TLabel')
        title_label.pack()
        
        subtitle_label = ttk.Label(header_frame, 
                                  text="Controle e envio automático via WhatsApp", 
                                  style='Info.TLabel')
        subtitle_label.pack(pady=(0, 10))
        
        # Main container
        main_container = ttk.Frame(self.root)
        main_container.grid(row=1, column=0, sticky="nsew", padx=20, pady=10)
        main_container.columnconfigure(0, weight=1)
        main_container.columnconfigure(1, weight=1)
        main_container.rowconfigure(0, weight=1)
        
        # Left Panel - Formulário
        self.create_form_panel(main_container)
        
        # Right Panel - Preview e Controles
        self.create_preview_panel(main_container)
        
        # Status bar
        self.create_status_bar()
        
    def create_form_panel(self, parent):
        form_frame = ttk.LabelFrame(parent, text="📝 Dados do Gerador", padding=15)
        form_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        
        # Grid configuration
        form_frame.columnconfigure(1, weight=1)
        
        # Campos do formulário
        fields = [
            ("Nome do Gerador *", "entry_nome", True),
            ("Destino", "entry_destino", False),
            ("Data (dd/mm/aaaa)", "entry_data", False),
            ("Motorista", "entry_motorista", False),
            ("Status", "entry_status", False),
            ("Regime de Trabalho", "entry_regime", False),
        ]
        
        self.entries = {}
        
        for i, (label, var_name, required) in enumerate(fields):
            # Label
            color = '#e74c3c' if required else '#2c3e50'
            lbl = ttk.Label(form_frame, text=label, foreground=color, font=('Segoe UI', 9, 'bold'))
            lbl.grid(row=i, column=0, sticky="w", pady=(5, 2))
            
            # Entry
            entry = ttk.Entry(form_frame, font=('Segoe UI', 10))
            entry.grid(row=i, column=1, sticky="ew", pady=(5, 2), padx=(10, 0))
            self.entries[var_name] = entry
            
            # Validação em tempo real para campos obrigatórios
            if required:
                entry.bind('<KeyRelease>', self.validate_required_fields)
        
        # Campo de observação (maior)
        ttk.Label(form_frame, text="Observação", font=('Segoe UI', 9, 'bold')).grid(row=len(fields), column=0, sticky="nw", pady=(15, 2))
        
        self.text_observacao = scrolledtext.ScrolledText(form_frame, height=4, font=('Segoe UI', 9))
        self.text_observacao.grid(row=len(fields), column=1, sticky="ew", pady=(15, 2), padx=(10, 0))
        
        # Botões de ação
        buttons_frame = ttk.Frame(form_frame)
        buttons_frame.grid(row=len(fields)+1, column=0, columnspan=2, pady=(20, 0))
        
        self.btn_adicionar = ttk.Button(buttons_frame, text="➕ Adicionar Gerador", 
                                       command=self.adicionar_gerador, style='Accent.TButton')
        self.btn_adicionar.pack(side=tk.LEFT, padx=(0, 10))
        
        self.btn_limpar = ttk.Button(buttons_frame, text="🧹 Limpar Campos", 
                                    command=self.limpar_campos)
        self.btn_limpar.pack(side=tk.LEFT)
        
        # Data atual automática
        self.entries['entry_data'].insert(0, datetime.now().strftime("%d/%m/%Y"))
        
    def create_preview_panel(self, parent):
        preview_frame = ttk.LabelFrame(parent, text="👁️ Preview dos Relatórios", padding=15)
        preview_frame.grid(row=0, column=1, sticky="nsew")
        preview_frame.rowconfigure(0, weight=1)
        preview_frame.columnconfigure(0, weight=1)
        
        # Lista de relatórios
        self.lista_frame = ttk.Frame(preview_frame)
        self.lista_frame.grid(row=0, column=0, sticky="nsew")
        self.lista_frame.rowconfigure(0, weight=1)
        self.lista_frame.columnconfigure(0, weight=1)
        
        # Treeview para mostrar os geradores
        columns = ('nome', 'destino', 'data', 'status')
        self.tree = ttk.Treeview(self.lista_frame, columns=columns, show='headings', height=12)
        
        # Configurar colunas
        self.tree.heading('nome', text='Nome')
        self.tree.heading('destino', text='Destino')
        self.tree.heading('data', text='Data')
        self.tree.heading('status', text='Status')
        
        self.tree.column('nome', width=120)
        self.tree.column('destino', width=100)
        self.tree.column('data', width=80)
        self.tree.column('status', width=80)
        
        # Scrollbar para a treeview
        scrollbar = ttk.Scrollbar(self.lista_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        
        # Preview da mensagem
        ttk.Label(preview_frame, text="📱 Preview da Mensagem:", 
                 font=('Segoe UI', 10, 'bold')).grid(row=1, column=0, sticky="w", pady=(15, 5))
        
        self.preview_text = scrolledtext.ScrolledText(preview_frame, height=8, 
                                                     font=('Segoe UI', 9), 
                                                     state='disabled',
                                                     background='#f8f9fa')
        self.preview_text.grid(row=2, column=0, sticky="ew", pady=(0, 15))
        
        # Controles de envio
        control_frame = ttk.Frame(preview_frame)
        control_frame.grid(row=3, column=0, sticky="ew")
        
        # Info do número
        ttk.Label(control_frame, text=f"📞 Enviando para: {self.numero_fixo}", 
                 style='Info.TLabel').pack(anchor="w")
        
        button_frame = ttk.Frame(control_frame)
        button_frame.pack(fill="x", pady=(10, 0))
        
        self.btn_enviar = ttk.Button(button_frame, text="📤 Enviar pelo WhatsApp", 
                                    command=self.enviar_mensagens_thread, 
                                    style='Accent.TButton')
        self.btn_enviar.pack(side=tk.LEFT, padx=(0, 10))
        
        self.btn_remover = ttk.Button(button_frame, text="🗑️ Remover Selecionado", 
                                     command=self.remover_gerador)
        self.btn_remover.pack(side=tk.LEFT)
        
        # Bind para seleção na treeview
        self.tree.bind('<<TreeviewSelect>>', self.on_select)
        
    def create_status_bar(self):
        status_frame = ttk.Frame(self.root)
        status_frame.grid(row=2, column=0, sticky="ew", padx=20, pady=(0, 20))
        
        self.status_label = ttk.Label(status_frame, text="✅ Pronto para uso", 
                                     style='Info.TLabel')
        self.status_label.pack(side=tk.LEFT)
        
        self.progress = ttk.Progressbar(status_frame, mode='indeterminate')
        self.progress.pack(side=tk.RIGHT, padx=(10, 0))
        
    def saudacao(self):
        hora = datetime.now().hour
        if hora < 12:
            return "☀️ Bom dia"
        elif hora < 18:
            return "🌤️ Boa tarde"
        else:
            return "🌃 Boa noite"
    
    def validate_required_fields(self, event=None):
        nome = self.entries['entry_nome'].get().strip()
        self.btn_adicionar.config(state='normal' if nome else 'disabled')
        
    def adicionar_gerador(self):
        nome = self.entries['entry_nome'].get().strip()
        destino = self.entries['entry_destino'].get().strip()
        data = self.entries['entry_data'].get().strip()
        motorista = self.entries['entry_motorista'].get().strip()
        status = self.entries['entry_status'].get().strip()
        regime = self.entries['entry_regime'].get().strip()
        observacao = self.text_observacao.get("1.0", tk.END).strip()
        
        if not nome:
            messagebox.showerror("❌ Erro", "O campo 'Nome' é obrigatório.")
            return
        
        # Criar mensagem
        msg = f"{self.saudacao()}, Dona Socorro!\n\n*Relatório do gerador:* *{nome}*\n"
        if destino:
            msg += f"📍 *Destino:* {destino}\n"
        if data:
            msg += f"📅 *Data:* {data}\n"
        if motorista:
            msg += f"🚚 *Motorista:* {motorista}\n"
        if status:
            msg += f"🔄 *Status:* {status}\n"
        if regime:
            msg += f"🕝 *Regime:* {regime}\n"
        if observacao:
            msg += f"⚠️ *Observação:* {observacao}"
        
        # Adicionar à lista
        gerador_data = {
            'nome': nome,
            'destino': destino or '-',
            'data': data or '-',
            'motorista': motorista or '-',
            'status': status or '-',
            'regime': regime or '-',
            'observacao': observacao or '-',
            'mensagem': msg
        }
        
        self.mensagens.append(gerador_data)
        
        # Adicionar à treeview
        self.tree.insert('', 'end', values=(nome, destino or '-', data or '-', status or '-'))
        
        # Limpar campos
        self.limpar_campos()
        
        # Atualizar status
        self.status_label.config(text=f"✅ {len(self.mensagens)} gerador(es) adicionado(s)")
        
        messagebox.showinfo("✅ Sucesso", f"Gerador '{nome}' adicionado ao relatório.")
        
    def limpar_campos(self):
        for entry in self.entries.values():
            entry.delete(0, tk.END)
        self.text_observacao.delete("1.0", tk.END)
        
        # Restaurar data atual
        self.entries['entry_data'].insert(0, datetime.now().strftime("%d/%m/%Y"))
        
        # Limpar preview
        self.preview_text.config(state='normal')
        self.preview_text.delete("1.0", tk.END)
        self.preview_text.config(state='disabled')
        
    def on_select(self, event):
        selection = self.tree.selection()
        if selection:
            item = self.tree.item(selection[0])
            index = self.tree.index(selection[0])
            
            if index < len(self.mensagens):
                mensagem = self.mensagens[index]['mensagem']
                
                # Mostrar preview
                self.preview_text.config(state='normal')
                self.preview_text.delete("1.0", tk.END)
                self.preview_text.insert("1.0", mensagem)
                self.preview_text.config(state='disabled')
                
    def remover_gerador(self):
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("⚠️ Aviso", "Selecione um gerador para remover.")
            return
            
        if messagebox.askyesno("🗑️ Confirmar", "Deseja remover o gerador selecionado?"):
            index = self.tree.index(selection[0])
            self.tree.delete(selection[0])
            
            if index < len(self.mensagens):
                removed = self.mensagens.pop(index)
                self.status_label.config(text=f"🗑️ Gerador '{removed['nome']}' removido")
                
                # Limpar preview se estava selecionado
                self.preview_text.config(state='normal')
                self.preview_text.delete("1.0", tk.END)
                self.preview_text.config(state='disabled')
    
    def enviar_mensagens_thread(self):
        # Executar em thread separada para não travar a interface
        thread = threading.Thread(target=self.enviar_mensagens)
        thread.daemon = True
        thread.start()
        
    def enviar_mensagens(self):
        if not self.mensagens:
            messagebox.showwarning("⚠️ Aviso", "Nenhum gerador adicionado.")
            return
        
        # Atualizar interface
        self.btn_enviar.config(state='disabled')
        self.progress.start(10)
        self.status_label.config(text="📤 Enviando mensagens...")
        
        try:
            servico = Service(ChromeDriverManager().install())
            options = webdriver.ChromeOptions()
            options.add_argument(r"user-data-dir=C:\Users\Lucas\AppData\Local\Temp\Profile Selenium")
            options.add_argument("--disable-notifications")
            navegador = webdriver.Chrome(service=servico, options=options)
            
            navegador.get("https://web.whatsapp.com")
            
            # Usar messagebox com thread-safe update
            self.root.after(0, lambda: messagebox.showinfo("🔐 Login", 
                           "Faça login no WhatsApp Web se ainda não estiver logado."))
            
            while len(navegador.find_elements(By.ID, 'side')) < 1:
                sleep(1)
            sleep(2)
            
            total = len(self.mensagens)
            for i, gerador in enumerate(self.mensagens, 1):
                mensagem = gerador['mensagem']
                texto = urllib.parse.quote(mensagem)
                link = f"https://web.whatsapp.com/send?phone={self.numero_fixo}&text={texto}"
                
                # Atualizar status
                self.root.after(0, lambda i=i, total=total: 
                               self.status_label.config(text=f"📤 Enviando {i}/{total}..."))
                
                navegador.get(link)
                while len(navegador.find_elements(By.ID, 'side')) < 1:
                    sleep(1)
                sleep(2)
                
                try:
                    btn = navegador.find_element(By.XPATH, 
                                               '//*[@id="main"]/footer/div[1]/div/span/div/div[2]/div/div[4]/button/span')
                    btn.click()
                    sleep(5)
                except:
                    self.root.after(0, lambda: messagebox.showerror("❌ Erro", 
                                    "Falha ao clicar no botão de envio. Verifique se o número é válido."))
                    navegador.quit()
                    return
            
            navegador.quit()
            
            # Sucesso
            self.root.after(0, lambda: messagebox.showinfo("✅ Sucesso", 
                           "Todas as mensagens foram enviadas com sucesso!"))
            
            # Limpar lista
            self.mensagens.clear()
            for item in self.tree.get_children():
                self.tree.delete(item)
                
            self.root.after(0, lambda: self.status_label.config(text="✅ Envio concluído!"))
            
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("❌ Erro", f"Erro durante o envio: {str(e)}"))
        finally:
            # Restaurar interface
            self.root.after(0, lambda: self.progress.stop())
            self.root.after(0, lambda: self.btn_enviar.config(state='normal'))

def main():
    root = tk.Tk()
    app = RelatorioGeradoresApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()