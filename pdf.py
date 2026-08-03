# -*- coding: utf-8 -*-
"""
Conversão de .docx para .pdf. Não existe biblioteca Python pura que renderize
.docx em PDF preservando 100% a formatação offline; a única forma confiável
é automação COM do Microsoft Word instalado na máquina (o mesmo motor usado
em "Salvar como PDF" no Word). Por isso esse recurso só funciona em máquinas
com Word instalado - se não houver, ConversorPDFIndisponivel é lançada com
uma mensagem clara, e o .docx (já gerado antes) continua válido normalmente.

Mantém uma única instância do Word aberta durante todo um lote (`ConversorPDF`
usado como context manager), em vez de abrir/fechar o Word a cada arquivo.
"""

from pathlib import Path

WD_FORMAT_PDF = 17


class ConversorPDFIndisponivel(Exception):
    """Word não está instalado ou não foi possível automatizá-lo via COM."""


class ConversorPDF:
    def __init__(self):
        try:
            import win32com.client
        except ImportError as exc:
            raise ConversorPDFIndisponivel(
                "Geração de PDF requer o Microsoft Word instalado nesta máquina "
                "(pacote pywin32 não encontrado)."
            ) from exc

        try:
            self._app = win32com.client.DispatchEx("Word.Application")
            self._app.Visible = False
        except Exception as exc:
            raise ConversorPDFIndisponivel(
                "Não foi possível abrir o Microsoft Word automaticamente. Verifique "
                "se ele está instalado nesta máquina."
            ) from exc

    def converter(self, origem_docx: Path, destino_pdf: Path) -> Path:
        doc = self._app.Documents.Open(str(origem_docx))
        try:
            doc.ExportAsFixedFormat(str(destino_pdf), WD_FORMAT_PDF)
        finally:
            doc.Close(False)
        return destino_pdf

    def fechar(self) -> None:
        try:
            self._app.Quit()
        except Exception:
            pass

    def __enter__(self) -> "ConversorPDF":
        return self

    def __exit__(self, *_exc) -> None:
        self.fechar()
