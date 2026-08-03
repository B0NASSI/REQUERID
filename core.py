# -*- coding: utf-8 -*-
"""
Lógica compartilhada entre o modo lote e o modo individual: validação dos
dados de entrada, derivação dos campos de concordância de gênero e da data,
montagem do contexto para o docxtpl, sanitização do nome do arquivo e
renderização do .docx final.
"""

import re
from datetime import date, datetime
from pathlib import Path

from docxtpl import DocxTemplate
from jinja2 import Environment, StrictUndefined
from jinja2.exceptions import UndefinedError

from nomes_genero import inferir_genero

CAMPOS_OBRIGATORIOS = [
    "empresa",
    "segurado",
    "cpf",
    "nit",
    "especie",
    "nb",
]

# Nome do benefício é determinado pela espécie (sempre a mesma descrição para
# cada código), por isso não é mais um campo de entrada — é derivado aqui.
ESPECIE_NOMES = {
    "B91": "auxílio por incapacidade temporária por acidente de trabalho",
    "B92": "aposentadoria por incapacidade permanente por acidente de trabalho",
    "B93": "pensão por morte por acidente de trabalho",
    "B94": "auxílio-acidente por acidente de trabalho",
    "B31": "auxílio por incapacidade temporária previdenciário",
    "B32": "aposentadoria por incapacidade permanente previdenciário",
    "B36": "auxílio-acidente previdenciário",
    "B41": "aposentadoria por idade",
    "B42": "aposentadoria por tempo de contribuição",
    "B46": "aposentadoria especial",
}

MESES_PT = {
    1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril", 5: "maio", 6: "junho",
    7: "julho", 8: "agosto", 9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro",
}


class ErroValidacao(Exception):
    """Erro de dados de entrada (campo ausente, vazio ou em formato inválido)."""


def data_por_extenso(d: date) -> str:
    return f"{d.day} de {MESES_PT[d.month]} de {d.year}"


def resolver_data(valor: str | None) -> str:
    """Aceita data em branco (usa hoje), "DD/MM/AAAA", ou um texto já por
    extenso (ex.: "15 de maio de 2026"), que é usado como está."""
    if not valor or not str(valor).strip():
        return data_por_extenso(date.today())
    valor = str(valor).strip()
    try:
        d = datetime.strptime(valor, "%d/%m/%Y").date()
        return data_por_extenso(d)
    except ValueError:
        return valor


def normalizar_especie(valor: str, campo: str = "especie") -> str:
    if valor is None:
        raise ErroValidacao(f"Campo '{campo}' está vazio.")
    texto = _val_str(valor).strip()
    if not texto:
        raise ErroValidacao(f"Campo '{campo}' está vazio.")

    maiusculo = texto.upper()
    if re.match(r'^B\d{2}$', maiusculo):
        # aceita "B91", "b31", etc. → código completo em maiúsculo
        codigo = maiusculo
    elif re.match(r'^\d{2}$', texto):
        # aceita "91", "31" → "B91", "B31"
        codigo = f"B{texto}"
    elif re.match(r'^[1-4]$', texto):
        # atalho legado: "1"-"4" → "B91"-"B94"
        codigo = f"B9{texto}"
    else:
        codigo = maiusculo

    if codigo not in ESPECIE_NOMES:
        raise ErroValidacao(
            f"Campo '{campo}': espécie não reconhecida (valor recebido: {valor!r}). "
            f"Use uma das espécies suportadas: {', '.join(sorted(ESPECIE_NOMES))}."
        )
    return codigo


def beneficio_por_especie(especie: str) -> str:
    return ESPECIE_NOMES[especie]


def normalizar_sexo(valor: str, campo: str = "sexo") -> str:
    if valor is None:
        raise ErroValidacao(f"Campo '{campo}' está vazio.")
    texto = str(valor).strip().upper()
    if texto not in ("M", "F"):
        raise ErroValidacao(f"Campo '{campo}' deve ser 'M' ou 'F'. Valor recebido: {valor!r}")
    return texto


def resolver_sexo(valor: str | None, segurado: str, campo: str = "sexo") -> str:
    """Usa o valor explícito de `sexo` se houver; senão tenta inferir pelo
    nome do segurado. Lança ErroValidacao se o nome for ambíguo/desconhecido
    e nenhum valor explícito tiver sido informado."""
    if valor and str(valor).strip():
        return normalizar_sexo(valor, campo)

    resultado = inferir_genero(segurado)
    if resultado in ("M", "M_PROVAVEL"):
        return "M"
    if resultado in ("F", "F_PROVAVEL"):
        return "F"
    raise ErroValidacao(
        f"Não foi possível identificar automaticamente o sexo pelo nome '{segurado}' "
        f"(nome ambíguo ou desconhecido). Informe a coluna '{campo}' (M ou F) para esta linha."
    )


def _val_str(valor) -> str:
    """Converte float inteiro (ex.: 91.0, 1234567890.0) para string sem '.0'.
    Células numéricas do Excel chegam como float via openpyxl."""
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    return str(valor)


def exigir_texto(valor, campo: str) -> str:
    if valor is None:
        raise ErroValidacao(f"Campo '{campo}' está vazio.")
    texto = _val_str(valor).strip()
    if not texto:
        raise ErroValidacao(f"Campo '{campo}' está vazio.")
    return texto


def formatar_nit(valor: str) -> str:
    """Se vier como 11 dígitos puros (ex.: planilhas "SOLICITAÇÕES GERID",
    que trazem o NIT sem formatação), formata como NIT/PIS padrão
    (XXX.XXXXX.XX-X). Valores já formatados ou fora do padrão são mantidos
    como estão."""
    digitos = re.sub(r"\D", "", _val_str(valor))
    if len(digitos) == 11:
        return f"{digitos[0:3]}.{digitos[3:8]}.{digitos[8:10]}-{digitos[10:]}"
    return _val_str(valor).strip()


def formatar_cpf(valor: str) -> str:
    """Mesma ideia de formatar_nit, mas para CPF (XXX.XXX.XXX-XX)."""
    digitos = re.sub(r"\D", "", _val_str(valor))
    if len(digitos) == 11:
        return f"{digitos[0:3]}.{digitos[3:6]}.{digitos[6:9]}-{digitos[9:]}"
    return _val_str(valor).strip()


def formatar_cnpj(valor: str) -> str:
    """Mesma ideia de formatar_nit/formatar_cpf, mas para CNPJ (XX.XXX.XXX/XXXX-XX)."""
    digitos = re.sub(r"\D", "", _val_str(valor))
    if len(digitos) == 14:
        return f"{digitos[0:2]}.{digitos[2:5]}.{digitos[5:8]}/{digitos[8:12]}-{digitos[12:]}"
    return _val_str(valor).strip()


def resolver_cnpj(valor_linha, cnpj_override: str | None, campo: str = "cnpj") -> str:
    """O CNPJ é o mesmo para todas as linhas de um lote (mesma empresa), por
    isso aceita um valor "global" (`cnpj_override`, ex.: campo da aba de
    lote ou flag --cnpj) além da coluna por linha. O override, se houver,
    tem prioridade."""
    bruto = cnpj_override if cnpj_override else valor_linha
    return formatar_cnpj(exigir_texto(bruto, campo))


def montar_contexto(linha: dict, data_override: str | None = None, cnpj_override: str | None = None) -> dict:
    """Valida `linha` (dict com as chaves de CAMPOS_OBRIGATORIOS, mais
    opcionalmente "sexo", "data" e "cnpj") e devolve o contexto completo
    para o docxtpl. Lança ErroValidacao com mensagem clara em caso de campo
    ausente/vazio ou em formato inválido.
    """
    faltando = [c for c in CAMPOS_OBRIGATORIOS if c not in linha]
    if faltando:
        raise ErroValidacao(f"Coluna(s) ausente(s) na planilha: {', '.join(faltando)}")

    empresa = exigir_texto(linha["empresa"], "empresa")
    cnpj = resolver_cnpj(linha.get("cnpj"), cnpj_override, "cnpj")
    segurado = exigir_texto(linha["segurado"], "segurado")
    sexo = resolver_sexo(linha.get("sexo"), segurado, "sexo")
    cpf = formatar_cpf(exigir_texto(linha["cpf"], "cpf"))
    nit = formatar_nit(exigir_texto(linha["nit"], "nit"))
    especie = normalizar_especie(linha["especie"], "especie")
    beneficio = beneficio_por_especie(especie)
    nb = exigir_texto(linha["nb"], "nb")

    data_linha = linha.get("data")
    data_extenso = resolver_data(data_override if data_override else data_linha)

    return {
        "empresa": empresa,
        "cnpj": cnpj,
        "segurado": segurado,
        "sexo": sexo,
        "tratamento_cap": "O segurado" if sexo == "M" else "A segurada",
        "inscrito": "inscrito" if sexo == "M" else "inscrita",
        "empregado": "empregado" if sexo == "M" else "empregada",
        "cpf": cpf,
        "nit": nit,
        "beneficio": beneficio,
        "especie": especie,
        "nb": nb,
        "data_extenso": data_extenso,
    }


def resolver_item(linha: dict, posicao_sequencial: int):
    """Número do item usado no nome do arquivo: a coluna 'item' da linha
    (quando existir, ex.: planilha "SOLICITAÇÕES GERID"), senão a posição
    sequencial da linha no lote (1, 2, 3...)."""
    valor = linha.get("item")
    if valor is None or str(valor).strip() == "":
        return posicao_sequencial
    if isinstance(valor, float) and valor.is_integer():
        return int(valor)
    return valor


def sanitizar_nome_arquivo_windows(texto: str) -> str:
    """Remove só os caracteres realmente inválidos em nome de arquivo no
    Windows (\\ / : * ? " < > |); mantém acentos, espaços e hífens, já que
    o padrão de nome adotado (ex.: "1. Requerimento -  Segurada - ...")
    depende deles."""
    sem_invalidos = re.sub(r'[\\/:*?"<>|]', "", str(texto))
    limpo = sem_invalidos.strip().rstrip(".")
    return limpo or "SEM_NOME"


def _ultimos_digitos(valor: str, quantidade: int = 3) -> str:
    digitos = re.sub(r"\D", "", str(valor))
    return digitos[-quantidade:] if digitos else ""


def _primeiro_nome_empresa(empresa: str) -> str:
    palavras = str(empresa).split()
    if not palavras:
        return ""
    primeiro = palavras[0]
    if len(primeiro) < 3 and len(palavras) > 1:
        return f"{palavras[0]} {palavras[1]}"
    return primeiro


def _dois_primeiros_nomes(nome: str) -> str:
    palavras = str(nome).split()
    if len(palavras) >= 3 and len(palavras[1]) < 4:
        return " ".join(palavras[:3])
    return " ".join(palavras[:2])


def _nome_base_documento(contexto: dict, empresa_arquivo_override: str | None = None) -> str:
    tratamento = "Segurado" if contexto["sexo"] == "M" else "Segurada"
    nb_curto = _ultimos_digitos(contexto["nb"])
    if empresa_arquivo_override and empresa_arquivo_override.strip():
        empresa_curta = sanitizar_nome_arquivo_windows(empresa_arquivo_override.strip()).upper()
    else:
        empresa_curta = _primeiro_nome_empresa(contexto["empresa"]).upper()
    segurado = _dois_primeiros_nomes(contexto["segurado"]).upper()
    return f"Requerimento -  {tratamento} - {segurado} - NB - {nb_curto} - Empresa - {empresa_curta}"


def nome_arquivo_docx(contexto: dict, item, empresa_arquivo_override: str | None = None) -> str:
    base = sanitizar_nome_arquivo_windows(f"{item}. {_nome_base_documento(contexto, empresa_arquivo_override)}")
    return f"{base}.docx"


def nome_arquivo_pdf(contexto: dict, item, empresa_arquivo_override: str | None = None) -> str:
    base = sanitizar_nome_arquivo_windows(f"{item}.1 {_nome_base_documento(contexto, empresa_arquivo_override)}")
    return f"{base}.pdf"


def renderizar(contexto: dict, template_path: Path, destino: Path) -> None:
    doc = DocxTemplate(str(template_path))
    env = Environment(undefined=StrictUndefined)
    try:
        doc.render(contexto, jinja_env=env)
    except UndefinedError as exc:
        raise ErroValidacao(f"Variável do template não preenchida: {exc}") from exc
    destino.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(destino))


def calcular_destino(
    linha: dict,
    pasta_saida: Path,
    data_override: str | None = None,
    cnpj_override: str | None = None,
    item=1,
    empresa_arquivo_override: str | None = None,
) -> Path:
    contexto = montar_contexto(linha, data_override, cnpj_override)
    return pasta_saida / nome_arquivo_docx(contexto, item, empresa_arquivo_override)


def calcular_destino_pdf(
    linha: dict,
    pasta_saida: Path,
    data_override: str | None = None,
    cnpj_override: str | None = None,
    item=1,
    empresa_arquivo_override: str | None = None,
) -> Path:
    contexto = montar_contexto(linha, data_override, cnpj_override)
    return pasta_saida / nome_arquivo_pdf(contexto, item, empresa_arquivo_override)


def gerar_arquivo(
    linha: dict,
    template_path: Path,
    pasta_saida: Path,
    data_override: str | None = None,
    cnpj_override: str | None = None,
    item=1,
    empresa_arquivo_override: str | None = None,
) -> Path:
    contexto = montar_contexto(linha, data_override, cnpj_override)
    destino = pasta_saida / nome_arquivo_docx(contexto, item, empresa_arquivo_override)
    renderizar(contexto, template_path, destino)
    return destino
