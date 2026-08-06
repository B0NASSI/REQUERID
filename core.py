# -*- coding: utf-8 -*-
"""
Lógica compartilhada entre o modo lote e o modo individual: validação dos
dados de entrada, derivação dos campos de concordância de gênero e da data,
montagem do contexto para o docxtpl, sanitização do nome do arquivo e
renderização do .docx final.
"""

import json
import re
from datetime import date, datetime
from pathlib import Path

from docx import Document
from docxtpl import DocxTemplate
from jinja2 import Environment, StrictUndefined
from jinja2.exceptions import UndefinedError
from pypdf import PdfReader

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


def _nome_base_documento(
    contexto: dict, empresa_arquivo_override: str | None = None, prefixo: str = "Requerimento"
) -> str:
    tratamento = "Segurado" if contexto["sexo"] == "M" else "Segurada"
    nb_curto = _ultimos_digitos(contexto["nb"])
    if empresa_arquivo_override and empresa_arquivo_override.strip():
        empresa_curta = sanitizar_nome_arquivo_windows(empresa_arquivo_override.strip()).upper()
    else:
        empresa_curta = _primeiro_nome_empresa(contexto["empresa"]).upper()
    segurado = _dois_primeiros_nomes(contexto["segurado"]).upper()
    return f"{prefixo} -  {tratamento} - {segurado} - NB - {nb_curto} - Empresa - {empresa_curta}"


def nome_arquivo_docx(
    contexto: dict, item, empresa_arquivo_override: str | None = None, prefixo: str = "Requerimento",
    sufixo_item: str = ".",
) -> str:
    base = sanitizar_nome_arquivo_windows(
        f"{item}{sufixo_item} {_nome_base_documento(contexto, empresa_arquivo_override, prefixo)}"
    )
    return f"{base}.docx"


def nome_arquivo_pdf(
    contexto: dict, item, empresa_arquivo_override: str | None = None, prefixo: str = "Requerimento",
    sufixo_item: str = ".1",
) -> str:
    base = sanitizar_nome_arquivo_windows(
        f"{item}{sufixo_item} {_nome_base_documento(contexto, empresa_arquivo_override, prefixo)}"
    )
    return f"{base}.pdf"


# Campos de identificação do benefício persistidos por NB (um .json por
# benefício em PASTA_DADOS_BENEFICIOS_PADRAO), para a aba de Cumprimento de
# Exigência reaproveitar sem redigitar. Não inclui campos derivados da data
# do requerimento (data_extenso, tratamento_cap, inscrito, empregado), que
# são recalculados a cada geração a partir de sexo/nb/etc.
CAMPOS_DADOS_BENEFICIO = ["empresa", "cnpj", "segurado", "sexo", "cpf", "nit", "especie", "nb"]


def _chave_nb(nb: str) -> str:
    """Nome de arquivo estável para um NB: só os dígitos (NB é sempre
    numérico). Se por algum motivo não sobrar dígito nenhum, cai para o
    texto sanitizado como está, para não gerar um nome de arquivo vazio."""
    digitos = re.sub(r"\D", "", str(nb))
    return digitos or sanitizar_nome_arquivo_windows(str(nb))


def salvar_dados_beneficio(contexto: dict, pasta_dados: Path, item=None, empresa_arquivo: str | None = None) -> None:
    """`item` é o número do requerimento original (o mesmo usado no nome do
    arquivo, ex.: 10, 5, 21) - salvo junto para a aba de Cumprimento de
    Exigência montar o nome do arquivo (ex.: "10.3 Cumprimento de
    Exigência...") sem precisar que o usuário redigite o número.

    `empresa_arquivo` é o nome abreviado da empresa como consta no nome do
    arquivo do requerimento original (a "Empresa - X" do nome do arquivo) -
    só precisa ser passado quando o usuário digitou um valor manual (o campo
    "nome da empresa no arquivo" é opcional). Se None/vazio nesta chamada,
    preserva o que já estava salvo (não apaga um valor salvo anteriormente
    só porque esta geração em particular não informou um override)."""
    pasta_dados.mkdir(parents=True, exist_ok=True)
    destino = pasta_dados / f"{_chave_nb(contexto['nb'])}.json"

    anterior = {}
    if destino.is_file():
        try:
            anterior = json.loads(destino.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            anterior = {}

    dados = {campo: contexto[campo] for campo in CAMPOS_DADOS_BENEFICIO}
    dados["item"] = item if item is not None else anterior.get("item", 1)
    empresa_arquivo_final = empresa_arquivo or anterior.get("empresa_arquivo")
    if empresa_arquivo_final:
        dados["empresa_arquivo"] = empresa_arquivo_final
    dados["atualizado_em"] = datetime.now().isoformat(timespec="seconds")
    destino.write_text(json.dumps(dados, ensure_ascii=False, indent=2), encoding="utf-8")


def carregar_dados_beneficio(nb: str, pasta_dados: Path) -> dict | None:
    """Devolve os dados salvos para o NB (dict pronto para uso como `linha`
    de montar_contexto, mais as chaves "item" e "empresa_arquivo"), ou None
    se nunca foi gerado nenhum requerimento para esse NB nesta instalação."""
    caminho = pasta_dados / f"{_chave_nb(nb)}.json"
    if not caminho.is_file():
        return None
    try:
        dados = json.loads(caminho.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None
    resultado = {campo: dados.get(campo, "") for campo in CAMPOS_DADOS_BENEFICIO}
    resultado["item"] = dados.get("item") or 1
    resultado["empresa_arquivo"] = dados.get("empresa_arquivo") or ""
    return resultado


# Padrão do corpo do texto do requerimento (mesma redação em template.docx e
# template_exigencia.docx, parágrafos 5 e 35) usado para recuperar os dados
# originais de um .docx/.pdf já gerado quando o NB não está salvo em
# PASTA_DADOS_BENEFICIOS_PADRAO (ex.: gerado antes dessa funcionalidade
# existir, ou em outra máquina que ainda não compartilhou os dados).
#
# Os campos numéricos aceitam espaços internos ([0-9./\- ]) porque o texto
# extraído de PDF (via pypdf) às vezes traz um espaço "fantasma" colado a um
# hífen ou vírgula em textos justificados pelo Word (ex.: "077.563.609 -66")
# - removido depois via _limpar_numero, não faz parte do dado real.
_RE_EMPRESA_CNPJ = re.compile(
    r"para representar a empresa (?P<empresa>.+?), pessoa jurídica de direito privado, "
    r"inscrita no CNPJ sob o nº (?P<cnpj>[0-9./\- ]+),"
)

_RE_DADOS_SEGURADO = re.compile(
    r"(?P<tratamento>O segurado|A segurada) (?P<segurado>.+?), "
    r"(?:inscrito|inscrita) no CPF nº (?P<cpf>[0-9./\- ]+) e sob o NIT nº (?P<nit>[0-9./\- ]+), "
    r"era (?:empregado|empregada) da empresa representada quando lhe foi concedido o "
    r"benefício de .+?, da espécie (?P<especie>B\s?\d\s?\d), nº (?P<nb>[0-9./\- ]+),"
)


def _dados_do_nome_arquivo(nome_arquivo: str) -> tuple:
    """Extrai do nome do arquivo (não do conteúdo) o número do item e o nome
    abreviado da empresa usados no documento original, para o novo arquivo
    seguir a mesma numeração/abreviação. O primeiro token (antes do primeiro
    espaço) é sempre "{item}" + um sufixo de numeração (".", ".1" no pdf do
    requerimento, ".3" na exigência) - removido pelo regex abaixo, que casa
    exatamente o que nome_arquivo_docx/pdf acrescentaram no final do token.
    Ex.: "21.3 Cumprimento de Exigência -  Segurada - ÉRICA - NB - 234 -
    Empresa - UNIÃO.pdf" → item=21, empresa_arquivo="UNIÃO"."""
    stem = Path(nome_arquivo).stem
    primeiro_token, _, resto = stem.partition(" ")
    item_texto = re.sub(r"\.\d*$", "", primeiro_token).strip()
    item = int(item_texto) if item_texto.isdigit() else (item_texto or 1)
    empresa_arquivo = resto.rsplit("Empresa - ", 1)[-1].strip() if "Empresa - " in resto else None
    return item, empresa_arquivo


def _normalizar_espacos(texto: str) -> str:
    """Colapsa qualquer sequência de espaços/quebras de linha em um único
    espaço - necessário porque o pypdf preserva as quebras de linha do
    Word (o texto de um mesmo parágrafo do template vem fatiado em várias
    linhas), enquanto o python-docx já devolve cada parágrafo inteiro.

    Também remove espaço logo antes de vírgula/ponto/ponto-e-vírgula/dois-
    pontos (ex.: "da espécie B36 , nº" → "da espécie B36, nº") - artefato do
    texto justificado que o pypdf às vezes introduz e que, sem essa limpeza,
    quebra os regexes de extração porque eles esperam a pontuação colada
    direto no fim do trecho anterior."""
    texto = re.sub(r"\s+", " ", texto)
    return re.sub(r"\s+([,.;:])", r"\1", texto)


def _limpar_numero(valor: str) -> str:
    return re.sub(r"\s+", "", valor).strip()


def _ler_texto_docx(caminho: Path) -> str:
    doc = Document(str(caminho))
    return "\n".join(p.text for p in doc.paragraphs)


def _ler_texto_pdf(caminho: Path) -> str:
    leitor = PdfReader(str(caminho))
    return "\n".join(pagina.extract_text() or "" for pagina in leitor.pages)


def extrair_dados_requerimento(caminho: Path) -> dict:
    """Lê um .docx ou .pdf de requerimento (ou cumprimento de exigência) já
    gerado por este programa e recupera os dados originais (empresa, cnpj,
    segurado, sexo, cpf, nit, especie, nb) a partir do texto do documento,
    além do número do item e do nome abreviado da empresa a partir do nome
    do arquivo. Usado pela aba de Cumprimento de Exigência quando o NB não
    está salvo em PASTA_DADOS_BENEFICIOS_PADRAO. Lança ErroValidacao se o
    formato não for suportado ou o documento não for reconhecido (editado
    manualmente, PDF escaneado sem texto, ou não gerado por este programa)."""
    if not caminho.is_file():
        raise ErroValidacao(f"Arquivo não encontrado: {caminho}")

    sufixo = caminho.suffix.lower()
    if sufixo == ".docx":
        texto_bruto = _ler_texto_docx(caminho)
    elif sufixo == ".pdf":
        texto_bruto = _ler_texto_pdf(caminho)
    else:
        raise ErroValidacao(f"Formato não suportado ({sufixo or 'sem extensão'}). Selecione um .docx ou .pdf.")

    texto = _normalizar_espacos(texto_bruto)

    match_empresa = _RE_EMPRESA_CNPJ.search(texto)
    match_segurado = _RE_DADOS_SEGURADO.search(texto)
    if not match_empresa or not match_segurado:
        raise ErroValidacao(
            "Não foi possível reconhecer os dados nesse documento (pode ter sido editado "
            "manualmente, ser um PDF escaneado sem texto, ou não ter sido gerado por este "
            "programa). Preencha os campos manualmente."
        )

    item, empresa_arquivo = _dados_do_nome_arquivo(caminho.name)

    return {
        "empresa": match_empresa.group("empresa").strip(),
        "cnpj": _limpar_numero(match_empresa.group("cnpj")),
        "segurado": match_segurado.group("segurado").strip(),
        "sexo": "M" if match_segurado.group("tratamento") == "O segurado" else "F",
        "cpf": _limpar_numero(match_segurado.group("cpf")),
        "nit": _limpar_numero(match_segurado.group("nit")),
        "especie": _limpar_numero(match_segurado.group("especie")),
        "nb": _limpar_numero(match_segurado.group("nb")),
        "item": item,
        "empresa_arquivo": empresa_arquivo,
    }


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
    prefixo_arquivo: str = "Requerimento",
    sufixo_item: str = ".",
) -> Path:
    contexto = montar_contexto(linha, data_override, cnpj_override)
    return pasta_saida / nome_arquivo_docx(contexto, item, empresa_arquivo_override, prefixo_arquivo, sufixo_item)


def calcular_destino_pdf(
    linha: dict,
    pasta_saida: Path,
    data_override: str | None = None,
    cnpj_override: str | None = None,
    item=1,
    empresa_arquivo_override: str | None = None,
    prefixo_arquivo: str = "Requerimento",
    sufixo_item: str = ".1",
) -> Path:
    contexto = montar_contexto(linha, data_override, cnpj_override)
    return pasta_saida / nome_arquivo_pdf(contexto, item, empresa_arquivo_override, prefixo_arquivo, sufixo_item)


def gerar_arquivo(
    linha: dict,
    template_path: Path,
    pasta_saida: Path,
    data_override: str | None = None,
    cnpj_override: str | None = None,
    item=1,
    empresa_arquivo_override: str | None = None,
    prefixo_arquivo: str = "Requerimento",
    pasta_dados_beneficios: Path | None = None,
    sufixo_item: str = ".",
) -> Path:
    contexto = montar_contexto(linha, data_override, cnpj_override)
    destino = pasta_saida / nome_arquivo_docx(contexto, item, empresa_arquivo_override, prefixo_arquivo, sufixo_item)
    renderizar(contexto, template_path, destino)
    if pasta_dados_beneficios is not None:
        salvar_dados_beneficio(contexto, pasta_dados_beneficios, item=item, empresa_arquivo=empresa_arquivo_override)
    return destino
