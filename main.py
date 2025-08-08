import tkinter as tk
from login_window import LoginWindow
from email_sender_app import EmailSenderApp

def main():
    # Tela de login
    login_root = tk.Tk()
    login_app = LoginWindow(login_root)
    login_root.mainloop()

    # Se o e-mail foi informado, inicia a aplicação principal
    if login_app.email:
        root = tk.Tk()
        app = EmailSenderApp(root, login_app.email)
        root.mainloop()
    else:
        print("Nenhum e-mail informado. Encerrando.")

if __name__ == '__main__':
    main()
