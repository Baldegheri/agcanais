import pandas as pd
from tkinter import filedialog, messagebox

def extrair_log_excel_function(tabela, log_por_destinatario):
    """
    Recebe um DataFrame e um dicionário (índice -> log) e adiciona uma coluna
    "LOG (ENVIADO - DATA E HORA)". Em seguida, permite salvar o Excel.
    """
    df_export = tabela.copy()
    df_export["LOG (ENVIADO - DATA E HORA)"] = df_export.index.map(
        lambda idx: log_por_destinatario.get(idx, "")
    )
    caminho = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
    )
    if caminho:
        try:
            df_export.to_excel(caminho, index=False)
            messagebox.showinfo("Extrair Log Excel", "Log exportado com sucesso!")
        except Exception as e:
            messagebox.showerror("Extrair Log Excel", f"Erro ao exportar log: {e}")
