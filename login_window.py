import tkinter as tk
from tkinter import messagebox

class LoginWindow:
    """Janela de login para capturar o e-mail do usuário."""
    def __init__(self, master):
        self.master = master
        self.master.title("Login")
        self.master.geometry("350x200")
        self.master.configure(bg="#ecf0f1")
        self.email = ""
        self._create_widgets()

    def _create_widgets(self):
        frame = tk.Frame(self.master, bg="#ecf0f1")
        frame.pack(expand=True)
        label = tk.Label(frame, text="Digite seu e-mail:", font=("Helvetica", 14), bg="#ecf0f1")
        label.pack(pady=(20, 10))
        self.entry = tk.Entry(frame, width=30, font=("Helvetica", 12))
        self.entry.pack(pady=5)
        btn = tk.Button(frame, text="Login", font=("Helvetica", 12), bg="#3498DB", fg="white",
                        activebackground="#2980B9", command=self._login)
        btn.pack(pady=20)

    def _login(self):
        self.email = self.entry.get().strip()
        if not self.email:
            messagebox.showerror("Erro", "Por favor, digite seu e-mail!")
        else:
            self.master.destroy()
