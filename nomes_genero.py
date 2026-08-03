# -*- coding: utf-8 -*-
"""
Listas de nomes próprios (masculinos, femininos e ambíguos) e função
auxiliar para inferir o gênero do segurado paradigma citado em cada tópico.

As listas não pretendem ser exaustivas: cobrem os nomes brasileiros mais
comuns. Quando o primeiro nome não é encontrado em nenhuma lista, o
software tenta uma heurística por terminação (sufixo) e, na dúvida,
marca o nome como "AMBIGUO" para revisão manual.
"""

import unicodedata

# Nomes explicitamente citados como ambíguos/propensos a confusão no
# enunciado do projeto, mais alguns outros nomes brasileiros que geram
# dúvida frequente quanto ao gênero.
NOMES_AMBIGUOS = {
    "DIORANDE", "ROSILEI", "GILVANE", "ANGELO", "ANGELA",
    "ALEX", "DANIEL", "JOAO", "MARIA", "VALMIR", "VALDIR",
    "ALTAIR", "ALTAMIR", "ERINEU", "EVERTON", "JURACI",
    "JUCELI", "MARLI", "MARLENE", "IVONE", "IVAN",
}

NOMES_FEMININOS = {
    "MARIA", "ANA", "FRANCISCA", "ANTONIA", "ADRIANA", "JULIANA", "MARCIA",
    "FERNANDA", "PATRICIA", "ALINE", "SANDRA", "CAMILA", "AMANDA", "BRUNA",
    "JESSICA", "LETICIA", "JULIA", "LUCIANA", "VANESSA", "MARIANA", "GABRIELA",
    "VALERIA", "CRISTINA", "BEATRIZ", "ROSANA", "ROSANGELA", "SIMONE",
    "CLAUDIA", "RAQUEL", "DANIELA", "TATIANE", "TATIANA", "PRISCILA",
    "ELAINE", "ELIANE", "ROSEMARY", "ROSEMEIRE", "SOLANGE", "SONIA",
    "REGINA", "MONICA", "DENISE", "DEBORA", "ALICE", "ALESSANDRA",
    "CAROLINA", "CAROLINE", "RENATA", "ROBERTA", "VIVIANE", "SILVANA",
    "SILVIA", "LUCIA", "LUCILENE", "LUCIMAR", "LUCINEIDE", "EDNA",
    "EDILEUZA", "EDILENE", "EDLEUZA", "JOANA", "JOSEFA", "JOSIANE",
    "JOSILENE", "JANAINA", "JAQUELINE", "JOSEANE", "KARINA", "KARLA",
    "KATIA", "MARILENE", "MARILZA", "MARISA", "MARTA", "MICHELE",
    "MIRIAM", "NATALIA", "NEUSA", "NEUZA", "NILZA", "NOEMIA", "PAULA",
    "RITA", "ROSA", "ROSARIO", "ROSELI", "ROSELY", "SUELI", "SUELEN",
    "TEREZA", "TEREZINHA", "VERA", "VITORIA", "YASMIN", "ZULEIDE",
    "APARECIDA", "CONCEICAO", "IVONE", "IRACEMA", "IRENE", "LOURDES",
    "MARGARETE", "MARGARIDA", "ODETE", "OLINDA", "ZENAIDE", "GISELE",
    "GISLENE", "GLAUCIA", "HELENA", "HELOISA", "INES", "DALVA",
    "DALILA", "CLEONICE", "CLEUSA", "CLEIDE", "CICERA", "BERENICE",
    "ANGELA", "MARLENE", "MARLI", "JUCELI", "JURACI", "FRANCIELI", "JANETE",
    "ROSANE", "ELIZETE",
    # Expansão (cobertura ampliada para reduzir casos "AMBIGUO" em listas
    # reais de funcionários, que trazem muito mais nomes do que os mais
    # comuns do dia a dia forense).
    "DANIELE", "DANIELLE", "MARCIELE", "MARCIELLE", "TAIZE", "TAYSE", "TAIS",
    "SIRLEI", "SIRLEY", "SARAI", "LEONI", "LARISSA", "ISABELA", "ISABELLA",
    "ISADORA", "MANUELA", "MANOELA", "LAURA", "LAVINIA", "VALENTINA",
    "SOPHIA", "SOFIA", "LIVIA", "RAFAELA", "RAFAELLA", "EMANUELE",
    "EMANUELLE", "MARIANE", "MARIANY", "STEFANI", "STEPHANIE", "ESTEFANI",
    "ELLEN", "EMILY", "EMILI", "EMILLY", "GEOVANA", "GIOVANA", "GIOVANNA",
    "YASMIM", "NICOLE", "NICOLLE", "BRENDA", "ALANA", "ALANE", "MILENA",
    "MILENE", "CAROLINY", "LORENA", "RAYANE", "RAYANA", "RAIANE", "RAIANA",
    "KETLEN", "KETLIN", "JAMILE", "JAMILY", "JAMILLY", "GREICE", "GREICY",
    "GREYCE", "FRANCIELE", "FRANCIELLE", "ROSILENE", "ROSILDA", "ROSIMERI",
    "ROSIMERY", "EDILAINE", "EDILEIA", "EDIVANIA", "MARILUCE", "MARILDA",
    "MARINALVA", "NEUZIA", "VANUSA", "VANDA", "VANIA", "IVANIA", "IVANILDA",
    "MARCILENE", "MARCIANE", "MARCIRENE", "ROSICLER", "ROSICLEIDE",
    "VALQUIRIA", "ELIZANGELA", "ELISANGELA", "GLEICE", "GLEIDE", "GLEIDIANE",
    "GISLAINE", "GIRLENE", "GIRLEIDE", "ROSIMAR", "JOELMA", "JOSELIA",
    "JUCIMARA", "LUCIMARA", "LUZIA", "LUZIMAR", "MARLUCE", "NADIA", "NADIR",
    "NAIR", "NEIDE", "NEIVA", "NEUSIMAR", "ROSILEIA", "SALETE", "SOLEDADE",
    "VANIRA", "VILMA", "ZILDA", "ZILMA", "ZORAIDE", "MARLUCIA", "DULCE",
    "DULCINEIA", "DULCINEA", "ERICA", "EDILAINE", "ELISETE", "ELISABETE",
    "ELIZABETE", "ELZA", "ERONDINA", "ERONICE", "EUNICE", "EVANI", "EVELYN",
    "EVELIN", "GENI", "GISELDA", "GRACIELA", "GRACIELE", "GRACINHA",
    "IARA", "IDALINA", "ILDA", "ILMA", "INAIA", "IRANI", "ISAURA", "ISMENIA",
    "IVETE", "IVANETE", "JACIRA", "JANICE", "JOELI", "JULIANE", "JUSSARA",
    "KEILA", "KELLY", "KELI", "LAIS", "LAIZ", "LENI", "LENICE", "LENIRA",
    "LIDIA", "LIGIA", "LINDALVA", "LUANA", "LUDIMILA", "LUDMILA",
    "MAGDA", "MAGNA", "MARLY", "MARILUZ", "MAYARA", "NATACHA", "NATASHA",
    "NEILA", "NELI", "NELIA", "ODALIA", "ODILA", "ROZANGELA", "ROZANE",
    "ROZINEIDE", "SELMA", "SHIRLEY", "SHIRLEI", "SUZANA", "SUZANE",
    "SUZANI", "TANIA", "TELMA", "VALDETE", "VALDIRENE", "VANILDA",
    "VANILDE", "VERONICA", "VITORIA", "WILMA", "WILMARA", "YARA",
}

NOMES_MASCULINOS = {
    "JOSE", "JOAO", "ANTONIO", "FRANCISCO", "CARLOS", "PAULO", "PEDRO",
    "LUCAS", "LUIZ", "MARCOS", "LUIS", "GABRIEL", "RAFAEL", "DANIEL",
    "MARCELO", "BRUNO", "EDUARDO", "FELIPE", "RAIMUNDO", "RODRIGO",
    "MANOEL", "MANUEL", "FABIO", "ALEXANDRE", "RICARDO", "LEONARDO",
    "ROBERTO", "SEBASTIAO", "GUSTAVO", "ANDRE", "MARCO", "VALDIR",
    "VALMIR", "WALTER", "WAGNER", "WELLINGTON", "WILLIAN", "WILLIAM",
    "VINICIUS", "ADRIANO", "ALESSANDRO", "ALOISIO", "ALTAIR", "ALTAMIR",
    "AMAURI", "ANDERSON", "ANTONIO", "ARIOVALDO", "BENEDITO", "CELSO",
    "CICERO", "CLAUDIO", "CLEBER", "CLOVIS", "DENILSON", "DIEGO",
    "DIRCEU", "DOUGLAS", "EDEMILSON", "EDERSON", "EDILSON", "EDMILSON",
    "EDSON", "ELIAS", "ELIEZER", "ELIO", "EMERSON", "ERINEU", "EVANDRO",
    "EVERTON", "EZEQUIEL", "FABRICIO", "FERNANDO", "FLAVIO", "GERALDO",
    "GILBERTO", "GILMAR", "GILSON", "GIOVANI", "HELIO", "HENRIQUE",
    "HUMBERTO", "IVAN", "IVANILDO", "JAIME", "JAIR", "JEFFERSON",
    "JOAQUIM", "JONATAS", "JONATHAN", "JORGE", "JULIO", "JUNIOR",
    "LEANDRO", "LUCIANO", "MAICON", "MARCIO", "MAURICIO", "MAURO",
    "MILTON", "MOACIR", "NELSON", "NEY", "NILTON", "NIVALDO", "ODAIR",
    "OSMAR", "OSVALDO", "OSCAR", "OTAVIO", "PERICLES", "REINALDO",
    "RENATO", "REGINALDO", "RONALDO", "RONALDO", "ROGERIO", "RUBENS",
    "SALVADOR", "SAMUEL", "SERGIO", "SIDNEI", "SILVIO", "TARCISIO",
    "THIAGO", "TIAGO", "URIEL", "VALTER", "VANDERLEI", "VICENTE",
    "VITOR", "WAGNER", "WASHINGTON", "WESLEY", "ANGELO", "ZECA",
    # Expansão (mesma motivação da lista feminina, acima).
    "DAVID", "DAVI", "DENIS", "DENYS", "ALAN", "ALLAN", "RENAN", "ROQUE",
    "IGOR", "IURI", "YURI", "KEVIN", "BRYAN", "MATHEUS", "MATEUS", "ENZO",
    "NOAH", "NOE", "ARTHUR", "ARTUR", "BERNARDO", "THEO", "TEO", "MIGUEL",
    "ISAAC", "ISAQUE", "EMANUEL", "NATAN", "NATHAN", "CAIO", "CAIQUE",
    "KAIQUE", "LORENZO", "BENJAMIN", "BENICIO", "RYAN", "JOAB", "JONAS",
    "ABEL", "ABRAAO", "ISMAEL", "JEREMIAS", "MOISES", "TOBIAS", "AMOS",
    "LUCCA", "LUCA", "GUILHERME", "OTAVIO", "AUGUSTO", "CESAR", "ANTHONY",
    "ANTONY", "NICOLAS", "NICOLAU", "VICTOR", "HEITOR", "ITALO", "IVO",
    "KAUA", "KAUE", "DIOGO", "RUI", "NUNO", "AFONSO", "GONCALO", "FELIX",
    "JULIANO", "EVERALDO", "ADEMIR", "ADEMAR", "EDMAR", "EDIMAR",
    "EDIVALDO", "GENIVALDO", "GERVASIO", "IVANALDO", "JOELSON", "JOILSON",
    "JOSIMAR", "JOSIVALDO", "LUCIVALDO", "MARCIVALDO", "NEUTON", "NEWTON",
    "ROSEMBERG", "ROSEVALDO", "VALDEMAR", "VALDEVINO", "WANDERLEY",
    "WANDERSON", "ADILSON", "ADEMILSON", "AGENOR", "AGOSTINHO", "AILTON",
    "AIRTON", "ALAOR", "ALCIDES", "ALDO", "ALMIR", "AMILTON", "ANISIO",
    "ANTENOR", "APARECIDO", "ARI", "ARLINDO", "ARMANDO", "ARNALDO",
    "ASSIS", "AURELIO", "BALDUINO", "BASILIO", "BENJAMIM", "CASSIO",
    "CLEITON", "CLEYTON", "CRISTIANO", "DARCI", "DARIO", "DEIVID",
    "DEIVISON", "DEJAIR", "DELVAIR", "DERLI", "DEVANIR", "DIONISIO",
    "DIONIZIO", "DIVINO", "DJALMA", "EDIVAN", "EDVALDO", "EDVAN", "EDIVAR",
    "EGIDIO", "ELENILSON", "ELIVELTON", "ELTON", "ENEAS", "EPITACIO",
    "ERALDO", "ERICO", "EROS", "EVALDO", "EVANILDO", "EXPEDITO", "EZIO",
    "FAUSTO", "FELICIANO", "GENILSON", "GENIVAL", "GERSON", "GETULIO",
    "GIDEAO", "GILVAN", "GIVANILDO", "GUARACI", "HAROLDO", "HEBER",
    "HERMES", "HILARIO", "HILTON", "HOMERO", "HORACIO", "IBRAIM",
    "INACIO", "IRINEU", "ISIDORO", "ITAMAR", "IVERSON", "IVO", "IZAIAS",
    "JACI", "JACINTO", "JADSON", "JANIO", "JEAN", "JOABE", "JOACI",
    "JOELMO", "JOILTON", "JONILSON", "JOSEMAR", "JOSENILDO", "JOSIAS",
    "JUVENAL", "LAERCIO", "LAUDELINO", "LAURO", "LAZARO", "LINDOMAR",
    "LOURIVAL", "LUCIDIO", "LUIDSON", "MAGNO", "MAIKON", "MANASSES",
    "MARLON", "MAXIMIANO", "MAXIMILIANO", "NATALICIO", "NAZARENO",
    "NILO", "NILSON", "NORBERTO", "OLAVO", "OLEGARIO", "OLIMPIO",
    "ORLANDO", "OSEAS", "OSIEL", "OSNI", "OZEIAS", "PASCOAL", "PRISCO",
    "RAUL", "REGIVALDO", "REMI", "ROSALVO", "SAMIR", "SANDRO", "SAULO",
    "SAVIO", "SEVERINO", "TADEU", "TARCIO", "TELMO", "URBANO", "VALERIO",
    "VALNEI", "VANILSON", "VOLNEI", "WAGNO", "WALACE", "WALLACE", "WALMOR",
    "WANDICK", "WESCLEY", "ZACARIAS",
}

# Sufixos que, na ausência de nome cadastrado, ajudam a inferir o gênero.
# "ELE"/"IZE" cobrem nomes como "Daniele"/"Marciele"/"Taize"; os demais já
# existiam.
SUFIXOS_FEMININOS = ("A", "HA", "IA", "EZ", "ETE", "ANE", "ENE", "ONE", "ELE", "IZE", "ICE")
SUFIXOS_MASCULINOS = ("O", "OR", "EL", "IM", "OM", "IR", "ON", "OS", "US", "AR", "ALDO", "ILSON")


def _normalizar(nome: str) -> str:
    """Remove acentos e coloca em caixa alta para comparação nas listas."""
    nome = nome.strip().upper()
    nome = unicodedata.normalize("NFKD", nome)
    nome = "".join(c for c in nome if not unicodedata.combining(c))
    return nome


def primeiro_nome(nome_completo: str) -> str:
    """Extrai o primeiro nome (ignorando títulos como 'Sr.', 'Sra.')."""
    titulos = {"SR", "SRA", "SR.", "SRA.", "DR", "DRA", "DR.", "DRA."}
    partes = [p for p in nome_completo.strip().split() if _normalizar(p) not in titulos]
    if not partes:
        return ""
    return partes[0]


def inferir_genero(nome_completo: str) -> str:
    """
    Tenta inferir o gênero do nome informado.

    Retorna uma das strings:
      'F' / 'M'            — nome consta na lista curada (alta confiança);
      'F_PROVAVEL' / 'M_PROVAVEL' — nome desconhecido, inferido só pela
                              terminação (baixa confiança — usar para
                              alertas mais brandos, não para CRÍTICO/IMPORTANTE);
      'AMBIGUO'             — consta na lista de nomes para revisão manual,
                              ou não foi possível decidir com segurança;
      ''                    — nenhum nome foi informado.
    """
    primeiro = primeiro_nome(nome_completo)
    if not primeiro:
        return ""

    chave = _normalizar(primeiro)

    if chave in NOMES_FEMININOS:
        return "F"
    if chave in NOMES_MASCULINOS:
        return "M"

    # Primeiro nome ambíguo (ou desconhecido): tenta desambiguar pelos
    # demais nomes (ex.: "DIORANDE JOSÉ DA ROSA" — DIORANDE é ambíguo,
    # mas "JOSÉ" desambigua para masculino), ignorando preposições comuns.
    titulos_e_preposicoes = {"SR", "SRA", "SR.", "SRA.", "DR", "DRA", "DR.", "DRA.", "DE", "DA", "DO", "DAS", "DOS", "E"}
    demais_tokens = [
        t for t in nome_completo.strip().split()[1:] if _normalizar(t) not in titulos_e_preposicoes
    ]
    for token in demais_tokens:
        chave_token = _normalizar(token)
        if chave_token in NOMES_FEMININOS:
            return "F"
        if chave_token in NOMES_MASCULINOS:
            return "M"

    if chave in NOMES_AMBIGUOS:
        return "AMBIGUO"

    # Heurística por terminação, usada apenas quando o nome não está
    # cadastrado em nenhuma das listas acima — confiança baixa, por isso
    # retorna um marcador "_PROVAVEL" em vez de 'M'/'F' diretamente.
    if chave.endswith(SUFIXOS_MASCULINOS):
        return "M_PROVAVEL"
    if chave.endswith(SUFIXOS_FEMININOS):
        return "F_PROVAVEL"

    return "AMBIGUO"
