import ctypes
import ctypes.wintypes
import logging
import threading
import tkinter as tk
from tkinter import messagebox

import ttkbootstrap as ttk

import log_setup
import tema
from caminhos import pasta_executavel, pasta_recursos
from interface import ALTURA_JANELA, LARGURA_JANELA, Janela

NOME_APP = "REQUERID - Requerimento de GERID"
DESCRICAO_APP = "Preenche automaticamente o requerimento administrativo e o cumprimento de exigência"

SPI_GETWORKAREA = 0x0030

# interface.py já chama isso na importação acima (mesmo arquivo de log,
# logs/requerid.log); chamar de novo aqui é barato e idempotente - deixa
# main.py independente dessa ordem de import por acidente.
log_setup.configurar_logging("requerid.log", pasta_executavel())
logger = logging.getLogger(__name__)


def _limpar_zips_temporarios() -> None:
    """Apaga zips de atualização (tmp*.zip) esquecidos pelo launcher em
    versões antigas (antes da correção que passou a apagá-los sozinho).
    Fica aqui e não no launcher.py porque o auto-update nunca sobrescreve
    o launcher.py de quem já tem o app instalado - só REQUERID.exe e
    _internal/, que é onde este código roda."""
    for zip_path in pasta_executavel().glob("tmp*.zip"):
        try:
            zip_path.unlink()
            logger.info("Zip temporário de atualização removido: %s", zip_path.name)
        except OSError as exc:
            logger.info("Não foi possível remover zip temporário %s: %r", zip_path.name, exc)


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


def _avisar_se_desatualizado(root) -> None:
    """Rede de segurança para quando o REQUERID.exe é aberto direto, sem
    passar pelo launcher (ex.: atalho fixado errado na barra de tarefas -
    "Pin to taskbar" a partir da janela já aberta em vez do atalho da área
    de trabalho, que fixa o app em vez do launcher). O launcher.py é quem
    normalmente baixa e aplica a atualização; aqui só avisamos, sem baixar
    nem travar o uso - e falha em silêncio se não conseguir checar (sem
    internet, GitHub fora etc.), do mesmo jeito que o launcher já faz."""
    def _checar():
        try:
            from launcher import get_latest_release, is_newer, read_local_version
            release = get_latest_release()
            if release is None or not is_newer(release["tag_name"], read_local_version()):
                return
        except Exception as exc:
            logger.info("Checagem de versão (fora do launcher) não pôde ser concluída: %r", exc)
            return
        root.after(0, lambda: messagebox.showinfo(
            "Nova versão disponível",
            "Há uma versão mais nova do REQUERID disponível.\n\n"
            "Feche o programa e abra pelo atalho da área de trabalho (REQUERID) "
            "para atualizar automaticamente.",
        ))

    threading.Thread(target=_checar, daemon=True).start()


def _tratar_erro_callback(root):
    """Rede de segurança: qualquer exceção não tratada dentro de um callback
    do Tkinter (clique de botão, digitação num campo etc.) passa por aqui em
    vez de só imprimir no console (invisível no .exe empacotado, sem
    console) e travar/fechar o programa."""
    def _tratar(exc_type, exc_value, exc_tb):
        logger.error("Erro não tratado numa ação da interface", exc_info=(exc_type, exc_value, exc_tb))
        try:
            messagebox.showerror(
                "Erro inesperado",
                f"Ocorreu um erro inesperado:\n\n{exc_value}\n\n"
                f"Detalhes salvos em:\n{pasta_executavel() / 'logs' / 'requerid.log'}",
                parent=root,
            )
        except Exception:
            pass  # não deixa um erro ao mostrar o erro derrubar o programa
    return _tratar


if __name__ == "__main__":
    logger.info("REQUERID iniciado.")
    _limpar_zips_temporarios()
    root = ttk.Window(themename="litera", iconphoto=None)
    root.title(NOME_APP)
    root.resizable(True, True)
    # a largura mínima é definida pela linha do NB na aba de Exigência (campo
    # + botões "Buscar dados salvos" e "Carregar requerimento" lado a lado)
    root.minsize(860, 820)
    root.report_callback_exception = _tratar_erro_callback(root)
    tema.aplicar(root)
    try:
        root.iconbitmap(str(pasta_recursos() / "GERID LOGO.ico"))
    except tk.TclError:
        pass

    try:
        janela = Janela(root, NOME_APP, DESCRICAO_APP)  # mantém referência viva (banner/logo dependem disso)
    except Exception:
        logger.exception("Falha ao iniciar o aplicativo")
        raise
    _centralizar(root, LARGURA_JANELA, ALTURA_JANELA)
    _avisar_se_desatualizado(root)
    root.mainloop()
    logger.info("REQUERID encerrado.")
