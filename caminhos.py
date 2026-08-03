# -*- coding: utf-8 -*-
"""
Resolução de caminhos que funciona tanto rodando via `python main.py`/`python
gerar.py` quanto a partir do .exe empacotado pelo PyInstaller.

Recursos somente leitura (ex.: modelos/template.docx, GERID LOGO.ico) ficam
dentro do executável e são extraídos para uma pasta temporária em tempo de
execução (sys._MEIPASS); arquivos gerados (output/) precisam ficar ao lado
do .exe, não dentro dessa pasta temporária, para não desaparecerem ao fechar
o programa.
"""

import sys
from pathlib import Path


def pasta_recursos() -> Path:
    return Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))


def pasta_executavel() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


TEMPLATE_PADRAO = pasta_recursos() / "modelos" / "template.docx"
PLANILHA_MODELO = pasta_recursos() / "dados" / "SOLICITAÇÕES GERID.xlsx"
SAIDA_PADRAO = pasta_executavel() / "output"
PLANILHA_LOTE_PADRAO = pasta_executavel() / "dados" / "SOLICITAÇÕES GERID.xlsx"
