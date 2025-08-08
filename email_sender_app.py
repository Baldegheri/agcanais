import os
import re
import json
import time
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext, simpledialog
import win32com.client as win32
from datetime import datetime
from utils import validar_email, converter_texto_para_html, anexos_validos
from log_manager import extrair_log_excel_function
import threading

class EmailSenderApp:
    """Aplicação principal para enviar e-mails."""
    def __init__(self, root, login_email):
        self.root = root
        try:
            self.root.iconbitmap(r"C:\Users\bruno\OneDrive\Área de Trabalho\AGMS\logo.ico")
        except Exception as e:
            print("Erro ao carregar o ícone:", e)
        self.root.title("Enviar E-mails")
        self.root.geometry("1200x700")
        
        self.login_email = login_email
        self.tabela = None
        self.anexos_por_destinatario = {}  # índice -> lista de anexos
        self.log_por_destinatario = {}     # índice -> log (data/hora)
        self.outlook = win32.Dispatch('outlook.application')
        
        self._setup_ui()

    def _setup_ui(self):
        self._configure_styles()
        self._create_sidebar()
        self._create_content()

    def _configure_styles(self):
        self.style = ttk.Style()
        self.style.theme_use("clam")
        self.style.configure("TButton", font=("Helvetica", 10, "bold"), foreground="white",
                             background="#3498db", padding=5)
        self.style.map("TButton", background=[("active", "#2980b9")])
        self.style.configure("Treeview.Heading", font=("Helvetica", 11, "bold"), background="#bdc3c7")
        self.style.configure("TLabel", font=("Helvetica", 12), background="white", foreground="#2c3e50")

    def _create_sidebar(self):
        self.sidebar_frame = tk.Frame(self.root, width=250, bg="#2c3e50")
        self.sidebar_frame.grid(row=0, column=0, sticky="ns")
        self.sidebar_frame.grid_propagate(False)
        label_menu = tk.Label(self.sidebar_frame, text="MENU", bg="#2c3e50", fg="white",
                              font=("Helvetica", 16, "bold"))
        label_menu.pack(pady=30)
        buttons = [
            ("Carregar Tabela", self.carregar_tabela),
            ("Enviar E-mails", self.enviar_todos_emails),
            ("Salvar Modelo", self.salvar_modelo),
            ("Carregar Modelo", self.carregar_modelo),
            ("Assinatura", self.selecionar_gif),
        ]
        for text, command in buttons:
            ttk.Button(self.sidebar_frame, text=text, command=command)\
                .pack(fill="x", padx=15, pady=8)
        
        # Agrupamento dos botões de anexar (um abaixo do outro)
        anexos_frame = tk.LabelFrame(self.sidebar_frame, text="Anexar Documentos", bg="#2c3e50", fg="white",
                                     font=("Helvetica", 12, "bold"))
        anexos_frame.pack(fill="x", padx=15, pady=8)
        ttk.Button(anexos_frame, text="Adicionar Anexos", command=self.adicionar_anexos)\
            .pack(fill="x", padx=15, pady=4)
        ttk.Button(anexos_frame, text="Anexar em Todos", command=self.adicionar_anexos_todos)\
            .pack(fill="x", padx=15, pady=4)
        # Novo botão para remover anexos de um destinatário específico
        ttk.Button(anexos_frame, text="Remover Anexos", command=self.remover_anexos)\
            .pack(fill="x", padx=15, pady=4)
        
        # Botões de formatação HTML
        label_format = tk.Label(self.sidebar_frame, text="Formatação", bg="#2c3e50", fg="white",
                                  font=("Helvetica", 14, "bold"))
        label_format.pack(pady=(40,10))
        format_buttons = [
            ("Negrito", lambda: self.aplicar_formatacao("bold")),
            ("Itálico", lambda: self.aplicar_formatacao("italic")),
            ("Sublinhado", lambda: self.aplicar_formatacao("underline")),
        ]
        for text, command in format_buttons:
            ttk.Button(self.sidebar_frame, text=text, command=command)\
                .pack(fill="x", padx=15, pady=6)
        
        # Botão para inserir hyperlink
        ttk.Button(self.sidebar_frame, text="Hyperlink", command=self.inserir_hyperlink)\
            .pack(fill="x", padx=15, pady=6)
        
        # Botão para extrair log para Excel
        ttk.Button(self.sidebar_frame, text="Extrair Log Excel", command=self.extrair_log_excel)\
            .pack(fill="x", padx=15, pady=6)

    def _create_content(self):
        self.content_frame = tk.Frame(self.root, bg="white")
        self.content_frame.grid(row=0, column=1, sticky="nsew")
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_rowconfigure(0, weight=1)
        header = tk.Label(self.content_frame, text="Envio de E-mails", font=("Helvetica", 18, "bold"),
                          bg="white", fg="#2c3e50")
        header.grid(row=0, column=0, columnspan=3, padx=10, pady=15, sticky="w")
        self.treeview_destinatarios = ttk.Treeview(
            self.content_frame,
            columns=("Nome", "Email", "CC", "Empresa", "AGCPartner", "Equipe", "Anexos"),
            show="headings"
        )
        for col in ("Nome", "Email", "CC", "Empresa", "AGCPartner", "Equipe", "Anexos"):
            self.treeview_destinatarios.heading(col, text=col)
        self.treeview_destinatarios.grid(row=1, column=0, columnspan=3, padx=10, pady=10, sticky="nsew")
        self.content_frame.grid_columnconfigure(0, weight=1)
        self.content_frame.grid_rowconfigure(1, weight=1)
        self.entry_assunto = tk.Entry(self.content_frame, font=("Helvetica", 12))
        self.entry_assunto.grid(row=2, column=0, columnspan=3, padx=10, pady=10, sticky="ew")
        self.text_mensagem = scrolledtext.ScrolledText(self.content_frame, font=("Helvetica", 12),
                                                       wrap="word", height=10)
        self.text_mensagem.grid(row=3, column=0, columnspan=3, padx=10, pady=10, sticky="nsew")
        # Campo para o GIF
        gif_label = tk.Label(self.content_frame, text="GIF da assinatura:", font=("Helvetica", 12),
                             bg="white", fg="#2c3e50")
        gif_label.grid(row=4, column=0, sticky="w", padx=10, pady=(10,0))
        self.entry_gif = tk.Entry(self.content_frame, font=("Helvetica", 12))
        self.entry_gif.grid(row=4, column=1, sticky="ew", padx=10, pady=(10,0))
        btn_selecionar_gif = ttk.Button(self.content_frame, text="Selecionar GIF", command=self.selecionar_gif)
        btn_selecionar_gif.grid(row=4, column=2, padx=10, pady=(10,0))
        # Controles extras para CC
        cc_frame = tk.Frame(self.content_frame, bg="white")
        cc_frame.grid(row=5, column=0, columnspan=3, padx=10, pady=10, sticky="ew")
        self.var_cc_cx = tk.BooleanVar()
        check_cc = tk.Checkbutton(cc_frame, text="Incluir CC: cx@agcapital.com.br", variable=self.var_cc_cx,
                                  font=("Helvetica", 12), bg="white", fg="#2c3e50")
        check_cc.pack(side="left", padx=5)
        label_team = tk.Label(cc_frame, text="Enviar CC para:", font=("Helvetica", 12), bg="white", fg="#2c3e50")
        label_team.pack(side="left", padx=5)
        self.var_cc_team = tk.StringVar()
        self.combo_cc_team = ttk.Combobox(cc_frame, textvariable=self.var_cc_team, state="readonly", font=("Helvetica", 12))
        self.combo_cc_team['values'] = ("", "canais01@agcapital.com.br", "canais03@agcapital.com.br",
                                         "canais04@agcapital.com.br", "canais06@agcapital.com.br")
        self.combo_cc_team.current(0)
        self.combo_cc_team.pack(side="left", padx=5)
        self.progressbar = ttk.Progressbar(self.content_frame, length=400, mode="determinate")
        self.progressbar.grid(row=6, column=0, columnspan=3, padx=10, pady=10, sticky="ew")
        self.log_area = scrolledtext.ScrolledText(self.content_frame, font=("Helvetica", 10),
                                                  state="disabled", height=8)
        self.log_area.grid(row=7, column=0, columnspan=3, padx=10, pady=10, sticky="nsew")

    def inserir_hyperlink(self):
        try:
            start_index = self.text_mensagem.index("sel.first")
            end_index = self.text_mensagem.index("sel.last")
            texto_selecionado = self.text_mensagem.get(start_index, end_index)
            link_url = simpledialog.askstring("Inserir Hyperlink", "Digite a URL:")
            if link_url:
                hyperlink_html = f'<a href="{link_url}" target="_blank">{texto_selecionado}</a>'
                self.text_mensagem.delete(start_index, end_index)
                self.text_mensagem.insert(start_index, hyperlink_html)
        except tk.TclError:
            messagebox.showerror("Erro", "Selecione um texto para adicionar um hyperlink!")

    def aplicar_formatacao(self, tag):
        try:
            start_index = self.text_mensagem.index("sel.first")
            end_index = self.text_mensagem.index("sel.last")
            texto_selecionado = self.text_mensagem.get(start_index, end_index)
            if tag == "bold":
                texto_formatado = f"<strong>{texto_selecionado}</strong>"
            elif tag == "italic":
                texto_formatado = f"<em>{texto_selecionado}</em>"
            elif tag == "underline":
                texto_formatado = f"<u>{texto_selecionado}</u>"
            else:
                texto_formatado = texto_selecionado
            self.text_mensagem.delete(start_index, end_index)
            self.text_mensagem.insert(start_index, texto_formatado)
        except tk.TclError:
            pass

    def adicionar_anexos_todos(self):
        arquivos = filedialog.askopenfilenames(title="Selecione os anexos para todos os destinatários")
        if arquivos:
            for index in self.tabela.index:
                anexos_atual = self.anexos_por_destinatario.get(index, [])
                for arquivo in arquivos:
                    if arquivo not in anexos_atual:
                        anexos_atual.append(arquivo)
                self.anexos_por_destinatario[index] = anexos_atual
                anexos_texto = ", ".join(os.path.basename(a) for a in anexos_atual)
                self.treeview_destinatarios.set(index, "Anexos", anexos_texto)
            self.log("Anexos adicionados para TODOS os destinatários.")

    def remover_anexos(self):
        # Novo método para remover anexos do destinatário selecionado
        selected = self.treeview_destinatarios.focus()
        if not selected:
            messagebox.showerror("Erro", "Selecione um destinatário na tabela!")
            return
        confirm = messagebox.askyesno("Remover Anexos", "Deseja remover todos os anexos deste destinatário?")
        if confirm:
            index = int(selected)
            self.anexos_por_destinatario[index] = []
            self.treeview_destinatarios.set(selected, "Anexos", "")
            destinatario = self.treeview_destinatarios.item(selected)['values'][0]
            self.log(f"Anexos removidos para {destinatario}.")

    def extrair_log_excel(self):
        if self.tabela is None:
            messagebox.showerror("Extrair Log Excel", "Nenhuma tabela foi carregada!")
            return
        extrair_log_excel_function(self.tabela, self.log_por_destinatario)

    def log(self, mensagem: str):
        self.log_area.config(state="normal")
        self.log_area.insert(tk.END, mensagem + "\n")
        self.log_area.see(tk.END)
        self.log_area.config(state="disabled")
        print(mensagem)

    def enviar_email(self, nome, email, cc, empresa, agcpartner, equipe, assunto, corpo, anexos, caminho_gif):
        try:
            mail = self.outlook.CreateItem(0)
            mail.To = ";".join([addr.strip() for addr in email.split(",") if addr.strip()])
            
            # Processa o campo CC: se for NaN ou não for uma string, converte para string vazia
            if pd.isna(cc):
                cc = ""
            else:
                cc = str(cc).strip()
                
            cc_parts = []
            if cc:
                cc_parts.extend([addr.strip() for addr in re.split(r'[;,]', cc) if addr.strip()])
            # Adiciona as opções extras apenas se selecionadas
            if self.var_cc_cx.get():
                cc_parts.append("cx@agcapital.com.br")
            if self.var_cc_team.get().strip():
                cc_parts.append(self.var_cc_team.get().strip())
            cc_final = ";".join(cc_parts)
            if cc_final:
                mail.CC = cc_final
            self.log("CC final: " + cc_final)
            
            mail.Subject = assunto
            mail.HTMLBody = corpo
            mail.BodyFormat = 2  # HTML
            
            for anexo in anexos:
                if os.path.exists(anexo):
                    mail.Attachments.Add(anexo, 1, 1, os.path.basename(anexo))
                    self.log("Anexo adicionado: " + anexo)
                else:
                    self.log("Anexo não encontrado: " + anexo)
                    
            if caminho_gif and os.path.exists(caminho_gif):
                attachment = mail.Attachments.Add(caminho_gif, 1, 1, "gif_assinatura")
                attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "gif_assinatura")
                self.log("GIF adicionado: " + caminho_gif)
                mail.HTMLBody += '<br><img src="cid:gif_assinatura" alt="GIF da assinatura" style="max-width:500px; height:auto;">'
            else:
                self.log("GIF não encontrado: " + str(caminho_gif))
                
            try:
                namespace = self.outlook.GetNamespace("MAPI")
                sent_folder = namespace.GetDefaultFolder(5)  # Pasta "Itens Enviados"
                mail.SaveSentMessageFolder = sent_folder
                self.log("Pasta de enviados configurada.")
            except Exception as e:
                self.log("Erro ao configurar a pasta de enviados: " + str(e))
                
            if not mail.Recipients.ResolveAll():
                unresolved = [r.Address for r in mail.Recipients if not r.Resolved]
                self.log("Endereços não resolvidos: " + ", ".join(unresolved))
                
            mail.Send()
            self.log(f"E-mail enviado para {email}" + (f" com CC: {cc_final}" if cc_final else ""))
            return True
        except Exception as e:
            self.log(f"Erro ao enviar e-mail para {email}. Erro: {e}")
            return False

    def carregar_tabela(self):
        caminho = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")])
        if caminho:
            try:
                if caminho.endswith('.csv'):
                    self.tabela = pd.read_csv(caminho)
                else:
                    self.tabela = pd.read_excel(caminho)
                self.tabela = self.tabela.drop_duplicates(subset=['Email'])
                self.atualizar_tabela_destinatarios()
                messagebox.showinfo("Sucesso", "Tabela carregada com sucesso!")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao carregar a tabela: {e}")

    def atualizar_tabela_destinatarios(self):
        for item in self.treeview_destinatarios.get_children():
            self.treeview_destinatarios.delete(item)
        for index, row in self.tabela.iterrows():
            anexos = self.anexos_por_destinatario.get(index, [])
            anexos_texto = ", ".join(os.path.basename(a) for a in anexos) if anexos else ""
            self.treeview_destinatarios.insert("", "end", iid=index, values=(
                row.get('Nome', ''),
                row.get('Email', ''),
                row.get('CC', ''),
                row.get('Empresa', ''),
                row.get('AGCPartner', ''),
                row.get('Equipe', ''),
                anexos_texto
            ))

    def adicionar_anexos(self):
        selected = self.treeview_destinatarios.focus()
        if not selected:
            messagebox.showerror("Erro", "Selecione um destinatário na tabela!")
            return
        arquivos = filedialog.askopenfilenames(title="Selecione os anexos")
        if arquivos:
            index = int(selected)
            anexos_atual = self.anexos_por_destinatario.get(index, [])
            for arquivo in arquivos:
                if arquivo not in anexos_atual:
                    anexos_atual.append(arquivo)
            self.anexos_por_destinatario[index] = anexos_atual
            anexos_texto = ", ".join(os.path.basename(a) for a in anexos_atual)
            self.treeview_destinatarios.set(selected, "Anexos", anexos_texto)
            self.log(f"Anexos adicionados para {self.treeview_destinatarios.item(selected)['values'][0]}.")

    def selecionar_gif(self):
        caminho_gif = filedialog.askopenfilename(filetypes=[("GIF files", "*.gif")])
        if caminho_gif:
            self.entry_gif.delete(0, tk.END)
            self.entry_gif.insert(0, caminho_gif)

    def salvar_modelo(self):
        modelo = {
            "assunto": self.entry_assunto.get(),
            "corpo": self.text_mensagem.get("1.0", tk.END)
        }
        caminho = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
        if caminho:
            try:
                with open(caminho, "w", encoding="utf-8") as arquivo:
                    json.dump(modelo, arquivo, ensure_ascii=False, indent=4)
                messagebox.showinfo("Sucesso", "Modelo salvo com sucesso!")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao salvar modelo: {e}")

    def carregar_modelo(self):
        caminho = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if caminho:
            try:
                with open(caminho, "r", encoding="utf-8") as arquivo:
                    modelo = json.load(arquivo)
                self.entry_assunto.delete(0, tk.END)
                self.entry_assunto.insert(0, modelo.get("assunto", ""))
                self.text_mensagem.delete("1.0", tk.END)
                self.text_mensagem.insert("1.0", modelo.get("corpo", ""))
                messagebox.showinfo("Sucesso", "Modelo carregado com sucesso!")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao carregar modelo: {e}")

    def enviar_todos_emails(self):
        if self.tabela is None:
            messagebox.showerror("Erro", "Carregue a tabela primeiro!")
            return

        # Desabilita os widgets da sidebar que suportam a opção "state"
        for widget in self.sidebar_frame.winfo_children():
            try:
                widget.config(state="disabled")
            except tk.TclError:
                pass

        assunto_base = self.entry_assunto.get()
        corpo_html = self.text_mensagem.get("1.0", tk.END)
        caminho_gif = self.entry_gif.get()
        corpo_final = corpo_html + "<br><br><p>Qualquer dúvida, estou à disposição.</p>"
        envios = list(self.tabela.iterrows())
        total_envios = len(envios)
        total_enviados = [0]

        loading_win = tk.Toplevel(self.root)
        loading_win.title("Enviando E-mails...")
        loading_win.geometry("400x150")
        loading_win.resizable(False, False)
        loading_win.transient(self.root)
        loading_win.grab_set()

        lbl_status = tk.Label(loading_win, text=f"Faltam {total_envios} envios", font=("Helvetica", 14))
        lbl_status.pack(pady=10)

        progress = ttk.Progressbar(loading_win, orient="horizontal", length=300, mode="determinate")
        progress.pack(pady=10)
        progress["maximum"] = total_envios
        progress["value"] = 0

        def finish_sending():
            self.log(f"Total de e-mails enviados: {total_enviados[0]}/{total_envios}")
            messagebox.showinfo("Sucesso", f"E-mails enviados: {total_enviados[0]}/{total_envios}")
            for widget in self.sidebar_frame.winfo_children():
                try:
                    widget.config(state="normal")
                except tk.TclError:
                    pass

        def sending_loop(i):
            if i >= total_envios:
                self.root.after(0, finish_sending)
                return

            try:
                index, row = envios[i]
                nome = row.get("Nome", "")
                email = row.get("Email", "")
                if not validar_email(email):
                    self.root.after(0, lambda: self.log(f"E-mail inválido: {email}"))
                    self.log_por_destinatario[index] = f"E-mail inválido - {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"
                else:
                    cc = row.get("CC", "")
                    empresa = row.get("Empresa", "")
                    agcpartner = row.get("AGCPartner", "")
                    equipe = row.get("Equipe", "")
                    assunto = assunto_base.replace("{nome}", str(nome))\
                                          .replace("{empresa}", str(empresa))\
                                          .replace("{agcpartner}", str(agcpartner))\
                                          .replace("{Equipe}", str(equipe))
                    corpo = corpo_final.replace("{nome}", str(nome))\
                                        .replace("{empresa}", str(empresa))\
                                        .replace("{agcpartner}", str(agcpartner))\
                                        .replace("{Equipe}", str(equipe))
                    anexos = self.anexos_por_destinatario.get(index, [])
                    if self.enviar_email(nome, email, cc, empresa, agcpartner, equipe, assunto, corpo, anexos, caminho_gif):
                        msg = f"E-mail enviado com sucesso - {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"
                        self.root.after(0, lambda: self.log(f"E-mail enviado para {email}"))
                        total_enviados[0] += 1
                    else:
                        msg = f"Falha no envio - {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"
                        self.root.after(0, lambda: self.log(f"Falha ao enviar e-mail para {email}"))
                    self.log_por_destinatario[index] = msg
                    self.anexos_por_destinatario[index] = []
                    self.root.after(0, lambda: self.treeview_destinatarios.set(index, "Anexos", ""))
            except Exception as e:
                self.root.after(0, lambda: self.log("Erro inesperado: " + str(e)))
            progress["value"] = i + 1
            restantes = total_envios - (i + 1)
            lbl_status.config(text=f"Faltam {restantes} envios")
            self.root.update_idletasks()
            self.root.after(5000, lambda: sending_loop(i + 1))

        sending_loop(0)
