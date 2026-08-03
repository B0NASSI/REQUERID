# -*- coding: utf-8 -*-
"""
CLI para gerar requerimentos GERID a partir de modelos/template.docx.

Uso:
    python gerar.py --lote dados/exemplo.xlsx
    python gerar.py --lote dados/exemplo.xlsx --pdf
    python gerar.py --individual
    python gerar.py --individual --empresa "Empresa LTDA" --cnpj 00.000.000/0001-00 \\
        --segurado "Fulano de Tal" --sexo M --cpf 000.000.000-00 --nit 00000000000 \\
        --especie B91 --nb 6251234567 --pdf
"""

import argparse
import sys
from pathlib import Path

from caminhos import SAIDA_PADRAO, TEMPLATE_PADRAO
from core import CAMPOS_OBRIGATORIOS, ErroValidacao, calcular_destino_pdf, gerar_arquivo, resolver_item
from leitura import ler_planilha
from pdf import ConversorPDF, ConversorPDFIndisponivel


def rodar_lote(
    planilha: Path,
    template: Path,
    saida: Path,
    data_override: str | None,
    cnpj_override: str | None,
    gerar_pdf: bool = False,
) -> int:
    if not template.exists():
        print(f"ERRO: template não encontrado em {template}. Rode preparar_template.py antes.", file=sys.stderr)
        return 1

    try:
        linhas = ler_planilha(planilha)
    except ErroValidacao as exc:
        print(f"ERRO: {exc}", file=sys.stderr)
        return 1

    conversor = None
    if gerar_pdf:
        try:
            conversor = ConversorPDF()
        except ConversorPDFIndisponivel as exc:
            print(f"AVISO: {exc}", file=sys.stderr)

    gerados = []
    erros = []
    nomes_usados = {}

    try:
        for posicao, (numero, linha) in enumerate(enumerate(linhas, start=2), start=1):
            item = resolver_item(linha, posicao)
            try:
                destino = gerar_arquivo(linha, template, saida, data_override, cnpj_override, item)
            except ErroValidacao as exc:
                erros.append((numero, str(exc)))
                continue

            if destino.name in nomes_usados:
                erros.append((
                    numero,
                    f"Nome de arquivo '{destino.name}' duplicado (mesma empresa + nb da linha "
                    f"{nomes_usados[destino.name]}). Arquivo desta linha não foi sobrescrito "
                    f"incorretamente, mas verifique os dados.",
                ))
                continue
            if conversor is not None:
                destino_pdf = calcular_destino_pdf(linha, saida, data_override, cnpj_override, item)
                conversor.converter(destino, destino_pdf)
            nomes_usados[destino.name] = numero
            gerados.append((numero, destino))
    finally:
        if conversor is not None:
            conversor.fechar()

    print(f"\n{len(gerados)} requerimento(s) gerado(s) em {saida}:")
    for numero, destino in gerados:
        print(f"  linha {numero}: {destino.name}")

    if erros:
        print(f"\n{len(erros)} linha(s) com erro (NÃO geradas):", file=sys.stderr)
        for numero, msg in erros:
            print(f"  linha {numero}: {msg}", file=sys.stderr)
        return 1

    return 0


def pedir(rotulo: str, valor_cli: str | None, padrao: str | None = None) -> str:
    if valor_cli is not None:
        return valor_cli
    sufixo = f" [{padrao}]" if padrao else ""
    resposta = input(f"{rotulo}{sufixo}: ").strip()
    if not resposta and padrao is not None:
        return padrao
    return resposta


def rodar_individual(args, template: Path, saida: Path) -> int:
    if not template.exists():
        print(f"ERRO: template não encontrado em {template}. Rode preparar_template.py antes.", file=sys.stderr)
        return 1

    linha = {
        "empresa": pedir("Empresa (razão social)", args.empresa),
        "segurado": pedir("Segurado (nome completo)", args.segurado),
        "sexo": pedir("Sexo do segurado (M/F, Enter = detectar pelo nome)", args.sexo, padrao=""),
        "cpf": pedir("CPF do segurado", args.cpf),
        "nit": pedir("NIT do segurado", args.nit),
        "especie": pedir("Espécie (ex.: B91 ou apenas 1)", args.especie),
        "nb": pedir("Número do benefício (NB)", args.nb),
        "data": pedir("Data do requerimento (DD/MM/AAAA, Enter = hoje)", args.data, padrao=""),
    }
    cnpj_override = pedir("CNPJ da empresa", args.cnpj)

    try:
        destino = gerar_arquivo(linha, template, saida, cnpj_override=cnpj_override, item=1)
    except ErroValidacao as exc:
        print(f"ERRO: {exc}", file=sys.stderr)
        return 1

    print(f"\nRequerimento gerado em: {destino}")

    if args.pdf:
        destino_pdf = calcular_destino_pdf(linha, saida, cnpj_override=cnpj_override, item=1)
        try:
            with ConversorPDF() as conversor:
                conversor.converter(destino, destino_pdf)
            print(f"PDF gerado em: {destino_pdf}")
        except ConversorPDFIndisponivel as exc:
            print(f"AVISO: {exc}", file=sys.stderr)

    return 0


def main():
    parser = argparse.ArgumentParser(description="Gera requerimentos administrativos GERID (.docx).")
    modo = parser.add_mutually_exclusive_group(required=True)
    modo.add_argument("--lote", type=Path, metavar="PLANILHA", help="Planilha .xlsx ou .csv com uma linha por requerimento")
    modo.add_argument("--individual", action="store_true", help="Gera um único requerimento via prompts/flags")

    parser.add_argument("--template", type=Path, default=TEMPLATE_PADRAO, help="Caminho de modelos/template.docx")
    parser.add_argument("--saida", type=Path, default=SAIDA_PADRAO, help="Pasta de saída dos .docx gerados")
    parser.add_argument("--data", type=str, default=None, help="Sobrescreve a data do requerimento (DD/MM/AAAA)")
    parser.add_argument(
        "--cnpj", type=str, default=None,
        help="CNPJ da empresa (lote: aplicado a todas as linhas; individual: do requerimento)",
    )
    parser.add_argument(
        "--pdf", action="store_true",
        help="Gera também um .pdf de cada requerimento (requer Microsoft Word instalado)",
    )

    for campo in CAMPOS_OBRIGATORIOS:
        parser.add_argument(f"--{campo}", type=str, default=None, help=argparse.SUPPRESS)
    parser.add_argument("--sexo", type=str, default=None, help=argparse.SUPPRESS)

    args = parser.parse_args()

    if args.lote:
        codigo = rodar_lote(args.lote, args.template, args.saida, args.data, args.cnpj, args.pdf)
    else:
        codigo = rodar_individual(args, args.template, args.saida)

    sys.exit(codigo)


if __name__ == "__main__":
    main()
