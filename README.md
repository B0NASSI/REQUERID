# Gerador de Requerimento GERID

Preenche automaticamente o requerimento administrativo de FAP (INSS/GERID),
gerando um .docx por segurado/benefício, em lote (planilha) ou individualmente,
preservando 100% a formatação do modelo original. Funciona totalmente offline.

## 1. Instalar dependências

Crie um ambiente virtual dentro da pasta do projeto e instale as dependências
nele (evita conflitar com outros projetos Python da máquina):

```powershell
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
```

A partir daqui, troque `python` por `.venv\Scripts\python.exe` nos comandos
abaixo (ou ative o ambiente com `.venv\Scripts\activate` e use `python` normalmente).

## 2. Interface gráfica (recomendado)

```powershell
.venv\Scripts\python.exe main.py
```

Abre uma janela com duas abas. O link "❓ Manual rápido" no topo abre um guia
rápido (colunas da planilha, espécies de benefício, detecção de sexo, PDF
etc.) sem precisar consultar este README.

- **Requerimento individual**: preencha os campos e clique em "Gerar requerimento".
  O arquivo é salvo em `output/`. "Limpar campos" reseta o formulário (volta aos
  valores padrão) para preencher o próximo requerimento sem fechar o programa.
  Se o .docx (ou .pdf) já existir na pasta de salvamento, pede confirmação antes
  de substituir.
- **Gerar em lote (planilha)**: escolha a planilha (.xlsx/.csv) e, opcionalmente, a
  pasta de salvamento e uma data para sobrescrever todas as linhas. Clique em "Gerar
  todos" — o resultado linha a linha (gerado ou erro) aparece na caixa de texto.
  O campo já vem preenchido com `dados/SOLICITAÇÕES GERID.xlsx` (padrão atual). "Limpar"
  reseta planilha/pasta/CNPJ/data e a caixa de resultado.
  Se houver segurados cujo sexo não pôde ser identificado automaticamente pelo
  nome, uma janela é aberta antes de gerar qualquer arquivo, pedindo para
  confirmar Masculino/Feminino de cada um (cancelar não gera nenhum arquivo do
  lote). Se algum .docx/.pdf do lote já existir na pasta de salvamento, mostra
  a lista (até 10, com contagem do restante) e pede confirmação antes de gerar
  qualquer arquivo. A geração roda em segundo plano com uma barra de progresso
  — a janela não fica congelada durante o lote (PDF via Word é a etapa mais
  lenta).

### Formato de planilha aceito

Além do esquema próprio (colunas `empresa`, `segurado`, `cpf`, `nit`, `especie`,
`nb`, ver seção 4), o lote também lê **diretamente** as planilhas
"SOLICITAÇÕES GERID" usadas no dia a dia com cada cliente, sem precisar
reformatar nada — colunas `Empresa`, `Número do Benefício`, `Espécie`, `NIT do
Empregado` (dígitos puros) `NIT do Empregado` (na verdade o CPF, formatado —
erro de digitação na planilha original; identificado pelo dígito verificador
de cada um) e `Segurados`. Esse formato é detectado automaticamente pelo
cabeçalho; as colunas são localizadas pela ordem relativa entre si, não pela
posição exata, então funciona com ou sem colunas `Item` e/ou `CNPJ` adicionais
(em qualquer posição) e tolera uma linha de título acima do cabeçalho (ex.:
"PEDIDO DE CÓPIA DE PROCESSO - ..."). Veja `dados/SOLICITAÇÕES GERID.xlsx`.
Se não houver coluna `sexo`, é identificado pelo nome (ver observação sobre
detecção automática mais abaixo); CPF/NIT com 11 dígitos puros são formatados
automaticamente.

## 3. Preparar o template (rodar uma única vez)

O modelo original com os trechos grifados em amarelo já está em
`modelos/1. Requerimento (GERID) - Atualizado.docx`. Para gerar o
`modelos/template.docx` (com os tokens Jinja2 e sem o realce amarelo):

```powershell
python preparar_template.py
```

Só é necessário repetir esse passo se o modelo original (.docx) for substituído.
O script valida que cada trecho grifado em amarelo tem exatamente o texto
esperado; se o modelo tiver sido editado de forma incompatível, ele para com
um erro em vez de gerar um template incorreto.

## 4. Gerar requerimentos em lote (linha de comando)

Alternativa ao uso pela interface gráfica. Planilha `.xlsx` ou `.csv` com uma
linha por requerimento. Colunas obrigatórias:

| coluna   | descrição                                                            |
| -------- | ---------------------------------------------------------------------- |
| empresa  | razão social da empresa representada                                  |
| segurado | nome completo do segurado                                             |
| cpf      | CPF do segurado, já formatado como deve aparecer no documento         |
| nit      | NIT do segurado, já formatado como deve aparecer no documento         |
| especie  | `B91`, `B92`, `B93` ou `B94` (aceita também só `91`/`92`/`93`/`94` ou `1`/`2`/`3`/`4`) |
| nb       | número do benefício (NB)                                              |

CNPJ não está nessa lista porque é o mesmo para toda a empresa: vem da coluna
opcional `cnpj` (abaixo) **ou** de um valor único aplicado a todas as linhas
do lote (flag `--cnpj` na linha de comando, ou campo "CNPJ" na aba de lote da
interface gráfica).

O nome do benefício (ex.: "auxílio por incapacidade temporária por acidente
de trabalho", citado no texto como "...lhe foi concedido o benefício de
auxílio por incapacidade...") é determinado automaticamente pela espécie —
não é mais necessário informá-lo:

| espécie | nome do benefício                                                |
| ------- | ------------------------------------------------------------------ |
| B91     | auxílio por incapacidade temporária por acidente de trabalho      |
| B92     | aposentadoria por incapacidade permanente por acidente de trabalho |
| B93     | pensão por morte por acidente de trabalho                         |
| B94     | auxílio-acidente por acidente de trabalho                         |

Colunas opcionais:

| coluna | descrição                                                                              |
| ------ | ----------------------------------------------------------------------------------------- |
| sexo   | `M` ou `F`. Se omitida ou vazia, é **detectada automaticamente pelo primeiro nome do segurado** (lista de nomes brasileiros comuns). Se o nome for ambíguo/desconhecido, a interface gráfica abre uma janela para confirmar o sexo de cada um antes de gerar o lote (na linha de comando, a linha falha pedindo para informar o sexo explicitamente). |
| cnpj   | CNPJ da empresa. Se omitida ou vazia em alguma linha, usa o valor do `--cnpj`/campo "CNPJ" do lote; se nenhum dos dois existir, a linha falha pedindo para informar. Aceita 14 dígitos puros (formata automaticamente) ou já formatado. |
| data   | data do requerimento. Em branco = data atual. Aceita `DD/MM/AAAA` ou texto por extenso    |
| item   | número do item (ex.: na planilha "SOLICITAÇÕES GERID"). Usado para nomear o arquivo; se omitida, usa a posição da linha no lote (1, 2, 3...). |

Veja `dados/exemplo.xlsx` como modelo. Para gerar:

```powershell
python gerar.py --lote dados/exemplo.xlsx
```

### Nome dos arquivos gerados

```text
{item}. Requerimento -  Segurado/Segurada - {NOME DO SEGURADO} - NB - {3 últimos dígitos do NB} - Empresa - {primeiro nome da empresa}.docx
{item}.1 Requerimento -  Segurado/Segurada - {NOME DO SEGURADO} - NB - {3 últimos dígitos do NB} - Empresa - {primeiro nome da empresa}.pdf
```

Exemplo:

```text
1. Requerimento -  Segurada - FRANCINEIDE DA SILVA CORREA OLIVEIRA - NB - 068 - Empresa - AGILE.docx
1.1 Requerimento -  Segurada - FRANCINEIDE DA SILVA CORREA OLIVEIRA - NB - 068 - Empresa - AGILE.pdf
```

- O número do item vem da coluna `item` da planilha (quando existir); senão, é a posição da linha no lote. No requerimento individual é sempre `1`.
- "Empresa" é só o primeiro nome da razão social (ex.: "AGILE SOLUÇÕES..." → "AGILE"); se o primeiro nome tiver menos de 3 caracteres, inclui o segundo também (ex.: "RS RODRIGUEZ...” → "RS RODRIGUEZ").
- O nome do arquivo preserva acentos e espaços (só remove os caracteres realmente inválidos em nome de arquivo no Windows: `\ / : * ? " < > |`).
- **Atenção a caminhos muito longos**: esses nomes são mais longos que antes. Se a pasta de salvamento escolhida for muito profunda (muitas subpastas) e o nome do segurado/empresa for muito longo, o caminho completo pode passar de 260 caracteres e o Windows recusar salvar o arquivo. Nesse caso, escolha uma pasta de salvamento mais simples (ex.: `C:\Requerimentos`).

Se uma coluna obrigatória estiver ausente da planilha, o script para
imediatamente, sem gerar nenhum arquivo. Se uma célula obrigatória estiver
vazia ou em formato inválido em uma linha específica, essa linha é pulada
(com aviso no console indicando a linha e o campo) e o script termina com
código de saída diferente de zero — as demais linhas válidas são geradas
normalmente.

## 5. Gerar um requerimento individual (linha de comando)

Modo interativo (pede cada campo):

```powershell
python gerar.py --individual
```

Ou tudo via linha de comando, sem prompts (lembre de incluir `--data`, senão
esse campo continua pedindo digitação). `--sexo` é opcional — se omitido, é
detectado automaticamente pelo nome:

```powershell
python gerar.py --individual --empresa "Empresa LTDA" --cnpj "00.000.000/0001-00" `
  --segurado "Fulano de Tal" --cpf "000.000.000-00" --nit "0.00000.000-0" `
  --especie B91 --nb "625111111-1" --data "26/06/2026" --pdf
```

## Gerar também em PDF

Tanto a interface gráfica (checkbox em cada aba) quanto a linha de comando
(flag `--pdf`) podem gerar, além do `.docx`, um `.pdf` com o mesmo nome ao
lado dele. Não existe biblioteca Python pura que converta `.docx` para PDF
preservando a formatação offline — a conversão usa automação do **Microsoft
Word**, que precisa estar instalado na máquina. Se o Word não estiver
disponível, o `.docx` é gerado normalmente e um aviso explica que o PDF não
pôde ser criado. No modo lote, uma única instância do Word é reaproveitada
para todos os arquivos (não abre/fecha o Word a cada linha).

## Observação sobre concordância de gênero

As palavras do texto que se referem ao segurado concordam com a coluna
`sexo`: "O segurado/A segurada", "inscrito/inscrita" e "empregado/empregada".

## Estrutura

```text
main.py                 # interface gráfica (abas Individual / Lote)
interface.py            # janelas e widgets da interface gráfica
preparar_template.py    # gera modelos/template.docx (rodar uma vez)
gerar.py                # CLI: --lote planilha.xlsx|csv  |  --individual
core.py                 # validação, contexto e renderização (docxtpl)
leitura.py              # leitura de .xlsx/.csv sem pandas
nomes_genero.py         # inferência de sexo pelo nome do segurado
pdf.py                  # conversão .docx -> .pdf via automação do Word
caminhos.py             # resolução de caminhos (dev e .exe empacotado)
modelos/template.docx   # gerado pelo preparar_template.py
dados/exemplo.xlsx      # planilha de exemplo
output/                 # .docx gerados
```
