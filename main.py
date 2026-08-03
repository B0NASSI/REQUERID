import tkinter as tk

import ttkbootstrap as ttk

import tema
from caminhos import pasta_recursos
from interface import ALTURA_JANELA, LARGURA_JANELA, Janela

NOME_APP = "REQUERID - Requerimento de GERID"
DESCRICAO_APP = "Preenche automaticamente o requerimento administrativo e o cumprimento de exigência"


def _centralizar(root, largura: int, altura: int) -> None:
    root.update_idletasks()
    x = (root.winfo_screenwidth() - largura) // 2
    y = max((root.winfo_screenheight() - altura) // 2 - 30, 0)
    root.geometry(f"{largura}x{altura}+{x}+{y}")


if __name__ == "__main__":
    root = ttk.Window(themename="litera", iconphoto=None)
    root.title(NOME_APP)
    root.resizable(False, False)
    tema.aplicar(root)
    try:
        root.iconbitmap(str(pasta_recursos() / "GERID LOGO.ico"))
    except tk.TclError:
        pass

    janela = Janela(root, NOME_APP, DESCRICAO_APP)  # mantém referência viva (banner/logo dependem disso)
    _centralizar(root, LARGURA_JANELA, ALTURA_JANELA)
    root.mainloop()
