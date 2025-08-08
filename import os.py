import os
import win32com.client

# Caminho da pasta com os arquivos .docx
pasta = r"C:\Users\bruno.bernardes\Desktop\Modelos - Copia (2)"

# Inicializa o Word via COM
word = win32com.client.Dispatch("Word.Application")
word.Visible = False  # Não mostra a janela do Word

# Percorre todos os arquivos na pasta
for nome_arquivo in os.listdir(pasta):
    if nome_arquivo.lower().endswith(".docx"):
        caminho_arquivo = os.path.join(pasta, nome_arquivo)
        print(f"🔄 Ativando controle de alterações em: {caminho_arquivo}")

        try:
            doc = word.Documents.Open(caminho_arquivo)

            # Ativa o controle de alterações
            doc.TrackRevisions = True

            # Salva e fecha
            doc.Save()
            doc.Close()

            print(f"✅ Controle de alterações ativado: {nome_arquivo}")

        except Exception as e:
            print(f"❌ Erro em '{nome_arquivo}': {e}")

# Encerra o Word
word.Quit()
