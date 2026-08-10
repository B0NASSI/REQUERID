# -*- coding: utf-8 -*-
"""Log em arquivo, compartilhado pelo REQUERID.exe e pelo launcher.

Os dois são compilados com console=False (sem janela de terminal), então
sem isso um erro não tratado hoje simplesmente desaparece — nem quem usa
nem quem dá suporte vê rastro nenhum dele.
"""
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path

_configurados = set()


def configurar_logging(nome_arquivo: str, base_dir: Path) -> None:
    """Liga um log rotativo em base_dir/logs/nome_arquivo. Idempotente por
    nome de arquivo, para poder ser chamado mais de uma vez sem duplicar
    handlers (ex.: em testes manuais)."""
    if nome_arquivo in _configurados:
        return

    pasta_logs = base_dir / "logs"
    pasta_logs.mkdir(parents=True, exist_ok=True)

    handler = RotatingFileHandler(
        pasta_logs / nome_arquivo, maxBytes=1_000_000, backupCount=3, encoding="utf-8"
    )
    handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s"))

    raiz = logging.getLogger()
    raiz.setLevel(logging.INFO)
    raiz.addHandler(handler)
    _configurados.add(nome_arquivo)
