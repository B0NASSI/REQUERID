# -*- coding: utf-8 -*-
"""
Tema visual (ttkbootstrap) do Gerador de Requerimento GERID - mesma
identidade visual (navy/laranja RS Rodriguez & Sousa) usada nos outros
aplicativos do escritório (ANEXT, RevisorFAP).
"""

from ttkbootstrap.style import ThemeDefinition

NOME_TEMA = "requerid"

COR_PRIMARIA = "#2D315F"        # azul escuro (banner, ações de destaque)
COR_SECUNDARIA = "#D3782A"      # laranja (ação principal)
COR_FUNDO = "#F4F5F9"
COR_TEXTO = "#22243F"
COR_SUCESSO = "#1F7A4D"
COR_ERRO = "#C00000"

CORES = {
    "primary": COR_PRIMARIA,
    "secondary": COR_SECUNDARIA,
    "success": COR_SUCESSO,
    "info": "#3B82C4",
    "warning": "#E0A526",
    "danger": COR_ERRO,
    "light": COR_FUNDO,
    "dark": COR_TEXTO,
    "bg": COR_FUNDO,
    "fg": COR_TEXTO,
    "selectbg": COR_PRIMARIA,
    "selectfg": "#FFFFFF",
    "border": "#C9CCE0",
    "inputfg": COR_TEXTO,
    "inputbg": "#FFFFFF",
    "active": "#3B4070",
}

TEMA = ThemeDefinition(name=NOME_TEMA, colors=CORES, themetype="light")


def aplicar(root) -> "ttkbootstrap.Style":
    estilo = root.style
    estilo.register_theme(TEMA)
    estilo.theme_use(NOME_TEMA)
    estilo.configure("TNotebook.Tab", font=("Segoe UI", 11, "bold"), padding=(18, 10))
    return estilo
