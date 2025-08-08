import re
import os

def validar_email(email: str) -> bool:
    """Valida se o e-mail possui um formato básico."""
    regex = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(regex, email) is not None

def converter_texto_para_html(texto: str) -> str:
    """Converte quebras de linha em <br> e retorna HTML básico."""
    html_body = texto.replace("\n", "<br>")
    return f"<html><body>{html_body}</body></html>"


def anexos_validos(anexo: str) -> bool:
    """Retorna True se o arquivo existir."""
    return os.path.exists(anexo)
