import os
import re
import json
import time
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import win32com.client as win32
from datetime import datetime
import threading
from utils import validar_email, converter_texto_para_html, anexos_validos
from log_manager import extrair_log_excel_function

class EmailLogic:
    def __init__(self, app):
        self.app = app

    def carregar_tabela(self):
        caminho = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")])
        if caminho:
            try:
                if caminho.endswith(".csv"):
                    self.app.tabela = pd.read_csv(caminho)
                else:
                    self.app.tabela = pd.read_excel(caminho)
                self.app.tabela = self.app.tabela.drop_duplicates(subset=["Email"])
                self.app.content.update_table()
                messagebox.showinfo("Sucesso", "Tabela carregada com sucesso!")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao carregar tabela: {e}")

    def salvar_modelo(self):
        modelo = {
            "assunto": self.app.content.entry_assunto.get(),
            "corpo": self.app.content.text_mensagem.get("1.0", tk.END)
        }
        caminho = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
        if caminho:
            try:
                with open(caminho, "w", encoding="utf-8") as f:
                    json.dump(modelo, f, ensure_ascii=False, indent=4)
                messagebox.showinfo("Sucesso", "Modelo salvo com sucesso!")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao salvar modelo: {e}")

    def carregar_modelo(self):
        caminho = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if caminho:
            try:
                with open(caminho, "r", encoding="utf-8") as f:
                    modelo = json.load(f)
                self.app.content.entry_assunto.delete(0, tk.END)
                self.app.content.entry_assunto.insert(0, modelo.get("assunto", ""))
                self.app.content.text_mensagem.delete("1.0", tk.END)
                self.app.content.text_mensagem.insert("1.0", modelo.get("corpo", ""))
                messagebox.showinfo("Sucesso", "Modelo carregado com sucesso!")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao carregar modelo: {e}")

    def selecionar_gif(self):
        caminho = filedialog.askopenfilename(filetypes=[("GIF files", "*.gif")])
        if caminho:
            self.app.content.entry_gif.delete(0, tk.END)
            self.app.content.entry_gif.insert(0, caminho)

    def inserir_hyperlink(self):
        try:
            start = self.app.content.text_mensagem.index("sel.first")
            end = self.app.content.text_mensagem.index("sel.last")
            txt = self.app.content.text_mensagem.get(start, end)
            url = simpledialog.askstring("Inserir Hyperlink", "Digite a URL:")
            if url:
                html = f'<a href="{url}" target="_blank">{txt}</a>'
                self.app.content.text_mensagem.delete(start, end)
                self.app.content.text_mensagem.insert(start, html)
        except tk.TclError:
            messagebox.showerror("Erro", "Selecione um texto para adicionar um hyperlink!")

    def aplicar_formatacao(self, tag):
        try:
            start = self.app.content.text_mensagem.index("sel.first")
            end = self.app.content.text_mensagem.index("sel.last")
            txt = self.app.content.text_mensagem.get(start, end)
            if tag == "bold":
                html = f"<strong>{txt}</strong>"
            elif tag == "italic":
                html = f"<em>{txt}</em>"
            elif tag == "underline":
                html = f"<u>{txt}</u>"
            else:
                html = txt
            self.app.content.text_mensagem.delete(start, end)
            self.app.content.text_mensagem.insert(start, html)
        except tk.TclError:
            pass

    def adicionar_destinatario_manual(self):
        win = tk.Toplevel(self.app.root)
        win.title("Adicionar Destinatário")
        win.geometry("400x300")
        fields = ["Nome", "Email", "CC", "Empresa", "AGCPartner", "Equipe"]
        entries = {}
        for i, field in enumerate(fields):
            lbl = tk.Label(win, text=f"{field}:", font=("Helvetica", 12))
            lbl.grid(row=i, column=0, padx=10, pady=5, sticky="e")
            ent = tk.Entry(win, font=("Helvetica", 12))
            ent.grid(row=i, column=1, padx=10, pady=5, sticky="w")
            entries[field] = ent

        def salvar():
            data = {field: entries[field].get().strip() for field in fields}
            if not data["Nome"] or not data["Email"]:
                messagebox.showerror("Erro", "Nome e Email são obrigatórios!")
                return
            new_row = pd.DataFrame([data])
            if self.app.tabela is None:
                self.app.tabela = new_row
            else:
                self.app.tabela = pd.concat([self.app.tabela, new_row], ignore_index=True)
            self.app.content.update_table()
            win.destroy()

        btn = tk.Button(win, text="Adicionar", font=("Helvetica", 12),
                        bg="#3498DB", fg="white", command=salvar)
        btn.grid(row=len(fields), column=0, columnspan=2, pady=10)

    def adicionar_anexos(self):
        selected = self.app.content.treeview.focus()
        if not selected:
            messagebox.showerror("Erro", "Selecione um destinatário na tabela!")
            return
        arquivos = filedialog.askopenfilenames(title="Selecione os anexos")
        if arquivos:
            index = int(selected)
            anexos = self.app.anexos_por_destinatario.get(index, [])
            for arq in arquivos:
                if arq not in anexos:
                    anexos.append(arq)
            self.app.anexos_por_destinatario[index] = anexos
            txt = ", ".join(os.path.basename(a) for a in anexos)
            self.app.content.treeview.set(selected, "Anexos", txt)
            self.log(f"Anexos adicionados para {self.app.content.treeview.item(selected)['values'][0]}.")

    def adicionar_anexos_todos(self):
        arquivos = filedialog.askopenfilenames(title="Selecione os anexos para todos os destinatários")
        if arquivos:
            for index in self.app.tabela.index:
                anexos = self.app.anexos_por_destinatario.get(index, [])
                for arq in arquivos:
                    if arq not in anexos:
                        anexos.append(arq)
                self.app.anexos_por_destinatario[index] = anexos
                txt = ", ".join(os.path.basename(a) for a in anexos)
                self.app.content.treeview.set(index, "Anexos", txt)
            self.log("Anexos adicionados para TODOS os destinatários.")

    def enviar_email(self, nome, email, cc, empresa, agcpartner, equipe, assunto, corpo, anexos, caminho_gif):
        try:
            mail = self.app.outlook.CreateItem(0)
            mail.To = ";".join([a.strip() for a in email.split(",") if a.strip()])
            cc_parts = [a.strip() for a in re.split(r'[;,]', cc) if a.strip()]
            if self.app.login_email:
                cc_parts.append(self.app.login_email.strip())
            if self.app.content.var_cc_cx.get():
                cc_parts.append("cx@agcapital.com.br")
            if self.app.content.var_cc_team.get().strip():
                cc_parts.append(self.app.content.var_cc_team.get().strip())
            cc_final = ";".join(cc_parts)
            if cc_final:
                mail.CC = cc_final
            self.log("CC final: " + cc_final)
            mail.Subject = assunto
            mail.HTMLBody = corpo
            mail.BodyFormat = 2
            for anexo in anexos:
                if os.path.exists(anexo):
                    mail.Attachments.Add(anexo, 1, 1, os.path.basename(anexo))
                    self.log("Anexo adicionado: " + anexo)
                else:
                    self.log("Anexo não encontrado: " + anexo)
            if caminho_gif and os.path.exists(caminho_gif):
                att = mail.Attachments.Add(caminho_gif, 1, 1, "gif_assinatura")
                att.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "gif_assinatura")
                self.log("GIF adicionado: " + caminho_gif)
                mail.HTMLBody += '<br><img src="cid:gif_assinatura" alt="GIF da assinatura" style="max-width:500px; height:auto;">'
            else:
                self.log("GIF não encontrado: " + str(caminho_gif))
            try:
                namespace = self.app.outlook.GetNamespace("MAPI")
                sent_folder = namespace.GetDefaultFolder(5)
                mail.SaveSentMessageFolder = sent_folder
                self.log("Pasta de enviados configurada.")
            except Exception as e:
                self.log("Erro ao configurar a pasta de enviados: " + str(e))
            if not mail.Recipients.ResolveAll():
                unresolved = [r.Address for r in mail.Recipients if not r.Resolved]
                self.log("Endereços não resolvidos: " + ", ".join(unresolved))
            mail.Send()
            self.log(f"E-mail enviado para {email} com CC: {cc_final}.")
            return True
        except Exception as e:
            self.log(f"Erro ao enviar e-mail para {email}. Erro: {e}")
            return False

    def enviar_todos_emails(self):
        if self.app.tabela is None:
            messagebox.showerror("Erro", "Carregue a tabela primeiro!")
            return

        for widget in self.app.sidebar.frame.winfo_children():
            widget.config(state="disabled")

        assunto_base = self.app.content.entry_assunto.get()
        corpo_html = self.app.content.text_mensagem.get("1.0", tk.END)
        caminho_gif = self.app.content.entry_gif.get()
        corpo_final = corpo_html + "<br><br><p>Qualquer dúvida, estou à disposição.</p>"
        envios = list(self.app.tabela.iterrows())
        total_envios = len(envios)
        total_enviados = 0

        loading_win = tk.Toplevel(self.app.root)
        loading_win.title("Enviando E-mails...")
        loading_win.geometry("400x150")
        loading_win.resizable(False, False)
        loading_win.transient(self.app.root)
        loading_win.grab_set()

        lbl_status = tk.Label(loading_win, text=f"Faltam {total_envios} envios", font=("Helvetica", 14))
        lbl_status.pack(pady=10)

        progress = ttk.Progressbar(loading_win, orient="horizontal", length=300, mode="determinate")
        progress.pack(pady=10)
        progress["maximum"] = total_envios
        progress["value"] = 0

        def sending_loop(i):
            nonlocal total_envios, total_enviados
            if i >= total_envios:
                loading_win.destroy()
                self.log(f"Total de e-mails enviados: {total_enviados}/{total_envios}")
                messagebox.showinfo("Sucesso", f"E-mails enviados: {total_enviados}/{total_envios}")
                for widget in self.app.sidebar.frame.winfo_children():
                    widget.config(state="normal")
                return

            index, row = envios[i]
            nome = row.get("Nome", "")
            email = row.get("Email", "")
            if not validar_email(email):
                self.log(f"E-mail inválido: {email}")
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
                anexos = self.app.anexos_por_destinatario.get(index, [])
                if self.enviar_email(nome, email, cc, empresa, agcpartner, equipe, assunto, corpo, anexos, caminho_gif):
                    msg = f"E-mail enviado com sucesso - {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"
                    self.log(f"E-mail enviado para {email}")
                    total_enviados += 1
                else:
                    msg = f"Falha no envio - {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"
                    self.log(f"Falha ao enviar e-mail para {email}")
                self.log_por_destinatario[index] = msg
                self.app.anexos_por_destinatario[index] = []
                self.app.content.treeview.set(index, "Anexos", "")
            progress["value"] = i + 1
            restantes = total_envios - (i + 1)
            lbl_status.config(text=f"Faltam {restantes} envios")
            self.app.root.update_idletasks()
            self.app.root.after(5000, lambda: sending_loop(i + 1))

        threading.Thread(target=lambda: sending_loop(0)).start()

    def salvar_modelo(self):
        modelo = {
            "assunto": self.app.content.entry_assunto.get(),
            "corpo": self.app.content.text_mensagem.get("1.0", tk.END)
        }
        caminho = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
        if caminho:
            try:
                with open(caminho, "w", encoding="utf-8") as f:
                    json.dump(modelo, f, ensure_ascii=False, indent=4)
                messagebox.showinfo("Sucesso", "Modelo salvo com sucesso!")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao salvar modelo: {e}")

    def log(self, msg: str):
        self.app.content.log_area.config(state="normal")
        self.app.content.log_area.insert(tk.END, msg + "\n")
        self.app.content.log_area.see(tk.END)
        self.app.content.log_area.config(state="disabled")
        print(msg)
