import ctypes
import ctypes.wintypes
import tkinter as tk

import ttkbootstrap as ttk

import tema
from caminhos import pasta_recursos
from interface import ALTURA_JANELA, LARGURA_JANELA, Janela

NOME_APP = "REQUERID - Requerimento de GERID"
DESCRICAO_APP = "Preenche automaticamente o requerimento administrativo e o cumprimento de exigência"

SPI_GETWORKAREA = 0x0030


def _area_util_tela() -> tuple[int, int]:
    """Tamanho da área útil da tela primária, sem a barra de tarefas do
    Windows - pra janela nunca nascer com o rodapé escondido atrás dela em
    telas menores. Se a chamada ao Windows falhar por algum motivo, cai para
    o tamanho cheio da tela (comportamento anterior)."""
    try:
        rect = ctypes.wintypes.RECT()
        ctypes.windll.user32.SystemParametersInfoW(SPI_GETWORKAREA, 0, ctypes.byref(rect), 0)
        largura, altura = rect.right - rect.left, rect.bottom - rect.top
        if largura > 0 and altura > 0:
            return largura, altura
    except Exception:
        pass
    return None


def _centralizar(root, largura: int, altura: int) -> None:
    root.update_idletasks()
    largura_util, altura_util = _area_util_tela() or (root.winfo_screenwidth(), root.winfo_screenheight())
    # nunca maior que a área útil (deixa uma margem pequena), pra caber
    # inteira mesmo em telas menores ou com barra de tarefas grande
    largura = min(largura, largura_util - 20)
    altura = min(altura, altura_util - 20)
    x = (root.winfo_screenwidth() - largura) // 2
    y = max((altura_util - altura) // 2 - 40, 0)
    root.geometry(f"{largura}x{altura}+{x}+{y}")


if __name__ == "__main__":
    root = ttk.Window(themename="litera", iconphoto=None)
    root.title(NOME_APP)
    root.resizable(True, True)
    # a largura mínima é definida pela linha do NB na aba de Exigência (campo
    # + botões "Buscar dados salvos" e "Carregar requerimento" lado a lado)
    root.minsize(860, 820)
    tema.aplicar(root)
    try:
        root.iconbitmap(str(pasta_recursos() / "GERID LOGO.ico"))
    except tk.TclError:
        pass

    janela = Janela(root, NOME_APP, DESCRICAO_APP)  # mantém referência viva (banner/logo dependem disso)
    _centralizar(root, LARGURA_JANELA, ALTURA_JANELA)
    root.mainloop()
