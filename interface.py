import logging
import os
import threading
import tkinter as tk
from datetime import date
from pathlib import Path
from tkinter import filedialog, messagebox

import shutil

import pythoncom
import ttkbootstrap as ttk
from PIL import Image, ImageTk
from tkinterdnd2 import DND_FILES, TkinterDnD
from ttkbootstrap.widgets.scrolled import ScrolledText

import log_setup
import tema
import visual
from caminhos import (PASTA_DADOS_BENEFICIOS_PADRAO, PLANILHA_LOTE_PADRAO, PLANILHA_MODELO, SAIDA_PADRAO,
                      TEMPLATE_EXIGENCIA_PADRAO, TEMPLATE_PADRAO, pasta_executavel, pasta_recursos)

NOME_PASTA_NOTAS = "NOTAS DE ATUALIZAÇÃO"

log_setup.configurar_logging("requerid.log", pasta_executavel())
logger = logging.getLogger(__name__)
CAMINHO_LOG = pasta_executavel() / "logs" / "requerid.log"


def registrar_erro(contexto: str, exc: BaseException) -> None:
    """Grava contexto + traceback completo no log rotativo (logs/requerid.log,
    ao lado do .exe) - pra dar pra diagnosticar um erro reportado pelo
    usuário sem precisar reproduzir na hora nem pedir o arquivo problemático
    de novo."""
    logger.error(contexto, exc_info=exc)


def _versao_para_ordenacao(nome_sem_extensao: str):
    return tuple(int(p) if p.isdigit() else 0 for p in nome_sem_extensao.split("."))


def _carregar_notas_atualizacao() -> str:
    pasta = pasta_recursos() / NOME_PASTA_NOTAS
    if not pasta.is_dir():
        return "Nenhuma nota de atualização encontrada."
    arquivos = sorted(pasta.glob("*.txt"), key=lambda p: _versao_para_ordenacao(p.stem), reverse=True)
    blocos = [f"VERSÃO {a.stem}\n\n{a.read_text(encoding='utf-8-sig').strip()}" for a in arquivos]
    return "\n\n\n".join(blocos) if blocos else "Nenhuma nota de atualização encontrada."
from core import (ESPECIE_NOMES, ErroValidacao, calcular_destino, calcular_destino_pdf, carregar_dados_beneficio,
                  extrair_dados_requerimento, gerar_arquivo, resolver_item, salvar_dados_beneficio,
                  _dois_primeiros_nomes, _primeiro_nome_empresa, _ultimos_digitos, sanitizar_nome_arquivo_windows)
from leitura import ler_planilha
from nomes_genero import inferir_genero
from pdf import ConversorPDF, ConversorPDFIndisponivel

LARGURA_JANELA = 860
ALTURA_JANELA = 960
ALTURA_BANNER = 110

RADIOS_ESPECIE = list(ESPECIE_NOMES.keys())

RAZAO_SOCIAL_CAMPOS = [
    ("empresa", "Empresa (razão social)"),
    ("cnpj", "CNPJ da empresa"),
    ("segurado", "Segurado (nome completo)"),
    ("cpf", "CPF do segurado"),
    ("nit", "NIT do segurado"),
    ("nb", "Número do benefício (NB)"),
]


def _carregar_imagem_altura(caminho: Path, altura: int) -> ImageTk.PhotoImage:
    imagem = Image.open(str(caminho)).convert("RGBA")
    proporcao = altura / imagem.height
    imagem = imagem.resize((int(imagem.width * proporcao), altura), Image.LANCZOS)
    return ImageTk.PhotoImage(imagem)


def _confirmar_sobrescricao(parent, arquivos: list[Path]) -> bool:
    """Pergunta antes de sobrescrever arquivos que já existem na pasta de
    salvamento. `arquivos` deve conter só os que de fato já existem."""
    if not arquivos:
        return True

    lista = "\n".join(f"  • {p.name}" for p in arquivos[:10])
    if len(arquivos) > 10:
        lista += f"\n  ... e mais {len(arquivos) - 10}"
    plural = "arquivo" if len(arquivos) == 1 else "arquivos"
    verbo = "já existe e será substituído" if len(arquivos) == 1 else "já existem e serão substituídos"

    return messagebox.askyesno(
        "Arquivos serão substituídos",
        f"{len(arquivos)} {plural} {verbo}:\n\n{lista}\n\nContinuar?",
        icon="warning",
        parent=parent,
    )


class AbaIndividual:
    def __init__(self, parent):
        self.frame = ttk.Frame(parent, padding=(20, 18))
        self.frame.columnconfigure(0, weight=1)

        cartao = ttk.Labelframe(self.frame, text=" Dados do requerimento ", padding=18, bootstyle="secondary")
        cartao.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        cartao.columnconfigure(1, weight=1)
        self.entries = {}

        linha = 0
        for chave, rotulo in RAZAO_SOCIAL_CAMPOS:
            ttk.Label(cartao, text=f"{rotulo}:").grid(row=linha, column=0, sticky="w", pady=5)
            entrada = ttk.Entry(cartao, width=48)
            entrada.grid(row=linha, column=1, sticky="w", pady=5, padx=(8, 0))
            self.entries[chave] = entrada
            linha += 1

            if chave == "cnpj":
                ttk.Label(cartao, text="Nome da empresa no arquivo word e pdf (opcional):").grid(
                    row=linha, column=0, sticky="w", pady=5
                )
                frame_emp_arq = ttk.Frame(cartao)
                frame_emp_arq.grid(row=linha, column=1, sticky="w", pady=5, padx=(8, 0))
                self.entry_empresa_arquivo = ttk.Entry(frame_emp_arq, width=28)
                self.entry_empresa_arquivo.pack(side="left")
                self.entry_empresa_arquivo.bind("<KeyRelease>", self._atualizar_preview)
                self.entry_empresa_arquivo.bind("<FocusOut>", self._atualizar_preview)
                ttk.Label(
                    frame_emp_arq, text="← só para o nome do arquivo",
                    foreground="#888888", font=("Segoe UI", 8)
                ).pack(side="left", padx=(8, 0))
                linha += 1

            entrada.bind("<KeyRelease>", self._atualizar_preview)
            entrada.bind("<FocusOut>", self._atualizar_preview)

            if chave == "segurado":
                entrada.bind("<KeyRelease>", self._detectar_sexo, add="+")
                entrada.bind("<FocusOut>", self._detectar_sexo, add="+")

                ttk.Label(cartao, text="Sexo do segurado:").grid(row=linha, column=0, sticky="w", pady=5)
                frame_sexo = ttk.Frame(cartao)
                frame_sexo.grid(row=linha, column=1, sticky="w", pady=5, padx=(8, 0))
                self.var_sexo = tk.StringVar(value="M")
                ttk.Radiobutton(frame_sexo, text="Masculino", variable=self.var_sexo, value="M").pack(
                    side="left", padx=(0, 14)
                )
                ttk.Radiobutton(frame_sexo, text="Feminino", variable=self.var_sexo, value="F").pack(
                    side="left"
                )
                linha += 1

                self.label_deteccao_sexo = ttk.Label(cartao, foreground="#666666")
                self.label_deteccao_sexo.grid(row=linha, column=0, columnspan=2, sticky="w", pady=(0, 5))
                linha += 1

        ttk.Label(cartao, text="Espécie do benefício:").grid(row=linha, column=0, sticky="w", pady=5)
        frame_especie = ttk.Frame(cartao)
        frame_especie.grid(row=linha, column=1, sticky="w", pady=5, padx=(8, 0))
        self.var_especie = tk.StringVar(value=RADIOS_ESPECIE[0])
        for col, especie in enumerate(RADIOS_ESPECIE[:4]):
            ttk.Radiobutton(frame_especie, text=especie, variable=self.var_especie, value=especie).grid(
                row=0, column=col, sticky="w", padx=(0, 12), pady=(0, 2)
            )
        for col, especie in enumerate(RADIOS_ESPECIE[4:]):
            ttk.Radiobutton(frame_especie, text=especie, variable=self.var_especie, value=especie).grid(
                row=1, column=col, sticky="w", padx=(0, 12), pady=(2, 0)
            )
        linha += 1

        self.label_beneficio = ttk.Label(cartao, foreground="#666666", wraplength=700)
        self.label_beneficio.grid(row=linha, column=0, columnspan=2, sticky="w", pady=(0, 5))
        self.var_especie.trace_add("write", self._atualizar_beneficio)
        self._atualizar_beneficio()
        self.var_sexo.trace_add("write", self._atualizar_preview)
        linha += 1

        ttk.Label(cartao, text="Data do requerimento:").grid(row=linha, column=0, sticky="w", pady=5)
        self.entry_data = ttk.Entry(cartao, width=20)
        self.entry_data.insert(0, date.today().strftime("%d/%m/%Y"))
        self.entry_data.grid(row=linha, column=1, sticky="w", pady=5, padx=(8, 0))
        linha += 1

        ttk.Label(cartao, text="Pasta de salvamento:").grid(row=linha, column=0, sticky="w", pady=5)
        frame_saida = ttk.Frame(cartao)
        frame_saida.grid(row=linha, column=1, sticky="ew", pady=5, padx=(8, 0))
        frame_saida.columnconfigure(0, weight=1)
        self.entry_saida = ttk.Entry(frame_saida)
        self.entry_saida.grid(row=0, column=0, sticky="ew")
        ttk.Button(
            frame_saida, text="Procurar...", command=self._escolher_saida, bootstyle="primary-outline"
        ).grid(row=0, column=1, padx=(8, 0))
        linha += 1

        self.var_pdf = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            cartao,
            text="Gerar também em PDF (requer Microsoft Word instalado)",
            variable=self.var_pdf,
            bootstyle="secondary",
        ).grid(row=linha, column=0, columnspan=2, sticky="w", pady=(4, 0))

        self.label_preview = ttk.Label(self.frame, text="", foreground="#666666", font=("Consolas", 8), wraplength=820)
        self.label_preview.grid(row=1, column=0, sticky="w", pady=(0, 20))

        self.frame_progresso = ttk.Frame(self.frame)
        self.barra_progresso = ttk.Progressbar(self.frame_progresso, mode="indeterminate", bootstyle="secondary")
        self.barra_progresso.pack(fill="x", pady=(0, 4))
        self.label_progresso = ttk.Label(self.frame_progresso, text="Gerando...", font=("Segoe UI", 9))
        self.label_progresso.pack(anchor="w")

        acoes = ttk.Frame(self.frame)
        acoes.grid(row=2, column=0, sticky="ew", pady=(2, 0))
        self.botao_abrir_pasta = ttk.Button(
            acoes,
            text="📂  Abrir pasta de salvamento",
            command=self._abrir_pasta_saida,
            bootstyle="primary-outline",
            state="disabled",
        )
        self.botao_abrir_pasta.pack(side="left")
        self._ultimo_pdf = None
        self.botao_abrir_pdf = ttk.Button(
            acoes, text="👁  Abrir PDF gerado", command=self._abrir_pdf, bootstyle="primary-outline",
            state="disabled",
        )
        self.botao_abrir_pdf.pack(side="left", padx=(10, 0))
        self.botao_limpar = ttk.Button(
            acoes, text="🧹  Limpar campos", command=self.limpar, bootstyle="danger-outline"
        )
        self.botao_limpar.pack(side="left", padx=(10, 0))
        self.botao_gerar = ttk.Button(
            acoes, text="📝  Gerar requerimento", command=self.gerar, bootstyle="secondary", width=20
        )
        self.botao_gerar.pack(side="right")

    def _abrir_pdf(self):
        if self._ultimo_pdf and self._ultimo_pdf.is_file():
            os.startfile(self._ultimo_pdf)

    def _detectar_sexo(self, _event=None):
        nome = self.entries["segurado"].get().strip()
        if not nome:
            self.label_deteccao_sexo.configure(text="")
            return

        resultado = inferir_genero(nome)
        if resultado in ("M", "M_PROVAVEL"):
            self.var_sexo.set("M")
            self.label_deteccao_sexo.configure(text="(sexo detectado automaticamente pelo nome)")
        elif resultado in ("F", "F_PROVAVEL"):
            self.var_sexo.set("F")
            self.label_deteccao_sexo.configure(text="(sexo detectado automaticamente pelo nome)")
        else:
            self.label_deteccao_sexo.configure(text="(nome ambíguo — confirme o sexo manualmente)")

    def _atualizar_beneficio(self, *_args):
        nome = ESPECIE_NOMES.get(self.var_especie.get(), "")
        nome_exibicao = nome[:1].upper() + nome[1:] if nome else nome
        self.label_beneficio.configure(text=f"Benefício: {nome_exibicao}")

    def _atualizar_preview(self, *_args):
        empresa = self.entries["empresa"].get().strip()
        segurado = self.entries["segurado"].get().strip()
        nb = self.entries["nb"].get().strip()
        sexo = self.var_sexo.get()
        empresa_arquivo = self.entry_empresa_arquivo.get().strip()

        if not empresa and not segurado and not nb:
            self.label_preview.configure(text="")
            return

        tratamento = "Segurado" if sexo == "M" else "Segurada"
        nome_seg = _dois_primeiros_nomes(segurado).upper() if segurado else "NOME"
        if empresa_arquivo:
            emp_curta = sanitizar_nome_arquivo_windows(empresa_arquivo).upper()
        elif empresa:
            emp_curta = _primeiro_nome_empresa(empresa).upper()
        else:
            emp_curta = "EMPRESA"
        nb_curto = _ultimos_digitos(nb) if nb else "NNN"

        preview = f"1. Requerimento - {tratamento} - {nome_seg} - NB - {nb_curto} - Empresa - {emp_curta}.docx"
        self.label_preview.configure(text=f"Arquivo: {preview}")

    def _escolher_saida(self):
        caminho = filedialog.askdirectory(title="Selecionar pasta de salvamento")
        if caminho:
            self.entry_saida.delete(0, "end")
            self.entry_saida.insert(0, caminho)

    def _abrir_pasta_saida(self):
        saida = self.entry_saida.get().strip()
        if saida and Path(saida).is_dir():
            os.startfile(saida)

    def limpar(self):
        for entrada in self.entries.values():
            entrada.delete(0, "end")
        self.entry_empresa_arquivo.delete(0, "end")
        self.label_preview.configure(text="")
        self.var_sexo.set("M")
        self.label_deteccao_sexo.configure(text="")
        self.var_especie.set(RADIOS_ESPECIE[0])
        self.entry_data.delete(0, "end")
        self.entry_data.insert(0, date.today().strftime("%d/%m/%Y"))
        self.entry_saida.delete(0, "end")
        self.var_pdf.set(True)
        self.botao_abrir_pasta.config(state="disabled")
        self._ultimo_pdf = None
        self.botao_abrir_pdf.config(state="disabled")
        self.entries["empresa"].focus_set()

    def _coletar_linha(self) -> dict:
        linha = {chave: self.entries[chave].get().strip() for chave, _ in RAZAO_SOCIAL_CAMPOS}
        linha["sexo"] = self.var_sexo.get()
        linha["especie"] = self.var_especie.get()
        linha["data"] = self.entry_data.get().strip()
        return linha

    def _iniciar_progresso(self):
        self.botao_gerar.config(state="disabled")
        self.botao_limpar.config(state="disabled")
        self.frame_progresso.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        self.barra_progresso.start(10)

    def _parar_progresso(self):
        self.barra_progresso.stop()
        self.frame_progresso.grid_remove()
        self.botao_gerar.config(state="normal")
        self.botao_limpar.config(state="normal")

    def gerar(self):
        if not TEMPLATE_PADRAO.exists():
            messagebox.showerror(
                "Template não encontrado",
                f"Rode preparar_template.py antes (esperado em {TEMPLATE_PADRAO}).",
            )
            return

        saida_texto = self.entry_saida.get().strip()
        if not saida_texto:
            messagebox.showwarning("Aviso", "Escolha a pasta de salvamento antes de gerar.")
            return
        saida = Path(saida_texto)
        linha = self._coletar_linha()
        item = 1  # requerimento individual: sempre o item 1 (sem planilha)
        empresa_arquivo_override = self.entry_empresa_arquivo.get().strip() or None

        try:
            destino_previsto = calcular_destino(linha, saida, item=item, empresa_arquivo_override=empresa_arquivo_override)
        except ErroValidacao as exc:
            messagebox.showwarning("Aviso", str(exc))
            return

        candidatos = [destino_previsto]
        if self.var_pdf.get():
            candidatos.append(calcular_destino_pdf(linha, saida, item=item, empresa_arquivo_override=empresa_arquivo_override))
        existentes = [p for p in candidatos if p.exists()]
        if not _confirmar_sobrescricao(self.frame.winfo_toplevel(), existentes):
            return

        gerar_pdf = self.var_pdf.get()
        self._iniciar_progresso()
        threading.Thread(
            target=self._gerar_worker, args=(linha, saida, item, empresa_arquivo_override, gerar_pdf), daemon=True
        ).start()

    def _gerar_worker(self, linha, saida, item, empresa_arquivo_override, gerar_pdf: bool) -> None:
        pythoncom.CoInitialize()
        try:
            try:
                destino = gerar_arquivo(
                    linha, TEMPLATE_PADRAO, saida, item=item, empresa_arquivo_override=empresa_arquivo_override,
                    pasta_dados_beneficios=PASTA_DADOS_BENEFICIOS_PADRAO,
                )
            except ErroValidacao as exc:
                self.frame.after(0, self._gerar_falhou, "Aviso", str(exc), messagebox.showwarning)
                return
            except Exception as exc:
                registrar_erro("Gerar - falha inesperada", exc)
                self.frame.after(0, self._gerar_falhou, "Erro ao gerar", str(exc), messagebox.showerror)
                return

            mensagem = f"Requerimento gerado em:\n{destino}"
            destino_pdf_gerado = None
            if gerar_pdf:
                destino_pdf = calcular_destino_pdf(linha, saida, item=item, empresa_arquivo_override=empresa_arquivo_override)
                try:
                    with ConversorPDF() as conversor:
                        conversor.converter(destino, destino_pdf)
                    mensagem += f"\n\nPDF gerado em:\n{destino_pdf}"
                    destino_pdf_gerado = destino_pdf
                except ConversorPDFIndisponivel as exc:
                    mensagem += f"\n\nAviso: PDF não foi gerado ({exc})"

            logger.info(
                "Requerimento gerado: %s (nb=%s, empresa=%s, pdf=%s)",
                destino, linha.get("nb"), linha.get("empresa"), destino_pdf_gerado,
            )
        finally:
            pythoncom.CoUninitialize()

        self.frame.after(0, self._gerar_concluiu, mensagem, destino_pdf_gerado)

    def _gerar_falhou(self, titulo: str, mensagem: str, mostrar) -> None:
        self._parar_progresso()
        mostrar(titulo, mensagem)

    def _gerar_concluiu(self, mensagem: str, destino_pdf_gerado: Path | None) -> None:
        self._parar_progresso()
        self.botao_abrir_pasta.config(state="normal")
        if destino_pdf_gerado is not None:
            self._ultimo_pdf = destino_pdf_gerado
            self.botao_abrir_pdf.config(state="normal")
        messagebox.showinfo("Sucesso", mensagem)


class AbaExigencia:
    """Cumprimento de exigência (ex.: comprovação de representatividade
    pedida pelo INSS numa diligência). Reaproveita os dados do requerimento
    original a partir do NB, salvos automaticamente em PASTA_DADOS_BENEFICIOS_PADRAO
    toda vez que um requerimento (individual ou lote) é gerado; se o NB não
    tiver histórico salvo nesta instalação, os campos ficam em branco para
    preenchimento manual (ou podem ser recuperados a partir do próprio .docx
    do requerimento original, pelo botão "Carregar requerimento").

    O nome do arquivo segue sempre "{item}.3 Cumprimento de Exigência - ...",
    onde {item} é o mesmo número do requerimento original (salvo no JSON, ou
    lido do nome do arquivo carregado) - não reinicia em 1."""

    PREFIXO_ARQUIVO = "Cumprimento de Exigência"
    SUFIXO_ITEM = ".3"

    def __init__(self, parent):
        self.frame = ttk.Frame(parent, padding=(20, 18))
        self.frame.columnconfigure(0, weight=1)
        self._item_carregado = 1

        cartao = ttk.Labelframe(self.frame, text=" Cumprimento de exigência ", padding=18, bootstyle="secondary")
        cartao.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        cartao.columnconfigure(1, weight=1)
        self.entries = {}

        ttk.Label(cartao, text="Número do benefício (NB):").grid(row=0, column=0, sticky="w", pady=5)
        frame_nb = ttk.Frame(cartao)
        frame_nb.grid(row=0, column=1, sticky="w", pady=5, padx=(8, 0))
        self.entry_nb = ttk.Entry(frame_nb, width=16)
        self.entry_nb.pack(side="left")
        ttk.Button(
            frame_nb, text="🔎  Buscar dados salvos", command=self._buscar, bootstyle="primary-outline"
        ).pack(side="left", padx=(8, 0))
        ttk.Button(
            frame_nb, text="📄  Carregar requerimento", command=self._carregar_arquivo,
            bootstyle="primary-outline",
        ).pack(side="left", padx=(8, 0))
        self.entry_nb.bind("<KeyRelease>", self._atualizar_preview)
        self.entry_nb.bind("<FocusOut>", self._atualizar_preview)

        frame_dica = ttk.Frame(cartao)
        frame_dica.grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 8))
        ttk.Label(frame_dica, text="💡").pack(side="left")
        ttk.Label(
            frame_dica, text=" Dica: também dá pra arrastar o arquivo direto pra esta janela.",
            foreground="#5A5A5A", font=("Segoe UI", 9),
        ).pack(side="left")

        self.label_busca = ttk.Label(cartao, foreground="#666666", wraplength=760)
        self.label_busca.grid(row=2, column=0, columnspan=2, sticky="w", pady=(0, 5))

        linha = 3
        for chave, rotulo in RAZAO_SOCIAL_CAMPOS:
            if chave == "nb":
                continue  # NB já foi coletado no topo (é a chave de busca)
            ttk.Label(cartao, text=f"{rotulo}:").grid(row=linha, column=0, sticky="w", pady=5)
            entrada = ttk.Entry(cartao, width=48)
            entrada.grid(row=linha, column=1, sticky="w", pady=5, padx=(8, 0))
            self.entries[chave] = entrada
            linha += 1

            if chave == "cnpj":
                ttk.Label(cartao, text="Nome da empresa no arquivo word e pdf (opcional):").grid(
                    row=linha, column=0, sticky="w", pady=5
                )
                frame_emp_arq = ttk.Frame(cartao)
                frame_emp_arq.grid(row=linha, column=1, sticky="w", pady=5, padx=(8, 0))
                self.entry_empresa_arquivo = ttk.Entry(frame_emp_arq, width=28)
                self.entry_empresa_arquivo.pack(side="left")
                self.entry_empresa_arquivo.bind("<KeyRelease>", self._atualizar_preview)
                self.entry_empresa_arquivo.bind("<FocusOut>", self._atualizar_preview)
                ttk.Label(
                    frame_emp_arq, text="← só para o nome do arquivo",
                    foreground="#888888", font=("Segoe UI", 8)
                ).pack(side="left", padx=(8, 0))
                linha += 1

            entrada.bind("<KeyRelease>", self._atualizar_preview)
            entrada.bind("<FocusOut>", self._atualizar_preview)

            if chave == "segurado":
                entrada.bind("<KeyRelease>", self._detectar_sexo, add="+")
                entrada.bind("<FocusOut>", self._detectar_sexo, add="+")

                ttk.Label(cartao, text="Sexo do segurado:").grid(row=linha, column=0, sticky="w", pady=5)
                frame_sexo = ttk.Frame(cartao)
                frame_sexo.grid(row=linha, column=1, sticky="w", pady=5, padx=(8, 0))
                self.var_sexo = tk.StringVar(value="M")
                ttk.Radiobutton(frame_sexo, text="Masculino", variable=self.var_sexo, value="M").pack(
                    side="left", padx=(0, 14)
                )
                ttk.Radiobutton(frame_sexo, text="Feminino", variable=self.var_sexo, value="F").pack(
                    side="left"
                )
                linha += 1

                self.label_deteccao_sexo = ttk.Label(cartao, foreground="#666666")
                self.label_deteccao_sexo.grid(row=linha, column=0, columnspan=2, sticky="w", pady=(0, 5))
                linha += 1

        ttk.Label(cartao, text="Espécie do benefício:").grid(row=linha, column=0, sticky="w", pady=5)
        frame_especie = ttk.Frame(cartao)
        frame_especie.grid(row=linha, column=1, sticky="w", pady=5, padx=(8, 0))
        self.var_especie = tk.StringVar(value=RADIOS_ESPECIE[0])
        for col, especie in enumerate(RADIOS_ESPECIE[:4]):
            ttk.Radiobutton(frame_especie, text=especie, variable=self.var_especie, value=especie).grid(
                row=0, column=col, sticky="w", padx=(0, 12), pady=(0, 2)
            )
        for col, especie in enumerate(RADIOS_ESPECIE[4:]):
            ttk.Radiobutton(frame_especie, text=especie, variable=self.var_especie, value=especie).grid(
                row=1, column=col, sticky="w", padx=(0, 12), pady=(2, 0)
            )
        linha += 1

        self.label_beneficio = ttk.Label(cartao, foreground="#666666", wraplength=700)
        self.label_beneficio.grid(row=linha, column=0, columnspan=2, sticky="w", pady=(0, 5))
        self.var_especie.trace_add("write", self._atualizar_beneficio)
        self._atualizar_beneficio()
        self.var_sexo.trace_add("write", self._atualizar_preview)
        linha += 1

        ttk.Label(cartao, text="Data do cumprimento:").grid(row=linha, column=0, sticky="w", pady=5)
        self.entry_data = ttk.Entry(cartao, width=20)
        self.entry_data.insert(0, date.today().strftime("%d/%m/%Y"))
        self.entry_data.grid(row=linha, column=1, sticky="w", pady=5, padx=(8, 0))
        linha += 1

        ttk.Label(cartao, text="Pasta de salvamento:").grid(row=linha, column=0, sticky="w", pady=5)
        frame_saida = ttk.Frame(cartao)
        frame_saida.grid(row=linha, column=1, sticky="ew", pady=5, padx=(8, 0))
        frame_saida.columnconfigure(0, weight=1)
        self.entry_saida = ttk.Entry(frame_saida)
        self.entry_saida.grid(row=0, column=0, sticky="ew")
        ttk.Button(
            frame_saida, text="Procurar...", command=self._escolher_saida, bootstyle="primary-outline"
        ).grid(row=0, column=1, padx=(8, 0))
        linha += 1

        self.var_pdf = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            cartao,
            text="Gerar também em PDF (requer Microsoft Word instalado)",
            variable=self.var_pdf,
            bootstyle="secondary",
        ).grid(row=linha, column=0, columnspan=2, sticky="w", pady=(4, 0))

        self.label_preview = ttk.Label(self.frame, text="", foreground="#666666", font=("Consolas", 8), wraplength=820)
        self.label_preview.grid(row=1, column=0, sticky="w", pady=(0, 20))

        self.frame_progresso = ttk.Frame(self.frame)
        self.barra_progresso = ttk.Progressbar(self.frame_progresso, mode="indeterminate", bootstyle="secondary")
        self.barra_progresso.pack(fill="x", pady=(0, 4))
        self.label_progresso = ttk.Label(self.frame_progresso, text="Gerando...", font=("Segoe UI", 9))
        self.label_progresso.pack(anchor="w")

        acoes = ttk.Frame(self.frame)
        acoes.grid(row=2, column=0, sticky="ew", pady=(2, 0))
        self.botao_abrir_pasta = ttk.Button(
            acoes,
            text="📂  Abrir pasta de salvamento",
            command=self._abrir_pasta_saida,
            bootstyle="primary-outline",
            state="disabled",
        )
        self.botao_abrir_pasta.pack(side="left")
        self._ultimo_pdf = None
        self.botao_abrir_pdf = ttk.Button(
            acoes, text="👁  Abrir PDF gerado", command=self._abrir_pdf, bootstyle="primary-outline",
            state="disabled",
        )
        self.botao_abrir_pdf.pack(side="left", padx=(10, 0))
        self.botao_limpar = ttk.Button(
            acoes, text="🧹  Limpar campos", command=self.limpar, bootstyle="danger-outline"
        )
        self.botao_limpar.pack(side="left", padx=(10, 0))
        self.botao_gerar = ttk.Button(
            acoes, text="📝  Gerar cumprimento de exigência", command=self.gerar, bootstyle="secondary"
        )
        self.botao_gerar.pack(side="right")

    def _abrir_pdf(self):
        if self._ultimo_pdf and self._ultimo_pdf.is_file():
            os.startfile(self._ultimo_pdf)

    def _buscar(self):
        nb = self.entry_nb.get().strip()
        if not nb:
            messagebox.showwarning("Aviso", "Informe o NB antes de buscar.")
            return

        dados = carregar_dados_beneficio(nb, PASTA_DADOS_BENEFICIOS_PADRAO)
        if dados is None:
            self.label_busca.configure(
                text=(
                    "Nenhum dado salvo para esse NB nesta instalação — preencha manualmente ou "
                    'use "Carregar requerimento" (ou arraste o arquivo) para recuperar os dados '
                    "de um .docx/.pdf já gerado."
                ),
                foreground="#a15c00",
            )
            return

        for chave in ("empresa", "cnpj", "segurado", "cpf", "nit"):
            self.entries[chave].delete(0, "end")
            self.entries[chave].insert(0, dados.get(chave, ""))
        if dados.get("sexo"):
            self.var_sexo.set(dados["sexo"])
        if dados.get("especie"):
            self.var_especie.set(dados["especie"])
        self.entry_empresa_arquivo.delete(0, "end")
        if dados.get("empresa_arquivo"):
            self.entry_empresa_arquivo.insert(0, dados["empresa_arquivo"])
        self._item_carregado = dados.get("item") or 1
        self.label_deteccao_sexo.configure(text="")
        self.label_busca.configure(text="Dados encontrados e preenchidos a partir do requerimento salvo.", foreground="#1a7a1a")
        self._atualizar_preview()

    def _carregar_arquivo(self):
        caminho_texto = filedialog.askopenfilename(
            title="Selecionar requerimento gerado (.docx ou .pdf)",
            filetypes=[("Requerimento gerado", "*.docx *.pdf"), ("Todos os arquivos", "*.*")],
        )
        if caminho_texto:
            self.carregar_de_caminho(caminho_texto)

    def carregar_de_caminho(self, caminho_texto) -> None:
        """Para NBs sem histórico salvo: lê os dados diretamente de um
        .docx/.pdf de requerimento já gerado (nesta ou noutra máquina) em
        vez de exigir preenchimento manual. Também grava o resultado em
        PASTA_DADOS_BENEFICIOS_PADRAO, para não precisar carregar de novo.
        Usado tanto pelo botão "Carregar requerimento" quanto ao soltar um
        arquivo arrastado na janela (ver o dnd_bind em Janela)."""
        caminho = Path(caminho_texto)
        try:
            dados = extrair_dados_requerimento(caminho)
        except ErroValidacao as exc:
            registrar_erro(f"Carregar requerimento - falha esperada ao ler {caminho}", exc)
            messagebox.showwarning("Não foi possível carregar", str(exc))
            return
        except Exception as exc:
            registrar_erro(f"Carregar requerimento - falha inesperada ao ler {caminho}", exc)
            messagebox.showerror(
                "Erro ao carregar",
                f"Ocorreu um erro inesperado ao ler esse arquivo.\n\nDetalhes salvos em:\n{CAMINHO_LOG}",
            )
            return

        logger.info("Dados carregados de %s (nb=%s)", caminho, dados.get("nb"))

        self.entry_nb.delete(0, "end")
        self.entry_nb.insert(0, dados["nb"])
        for chave in ("empresa", "cnpj", "segurado", "cpf", "nit"):
            self.entries[chave].delete(0, "end")
            self.entries[chave].insert(0, dados.get(chave, ""))
        self.var_sexo.set(dados["sexo"])
        self.var_especie.set(dados["especie"])
        self.entry_empresa_arquivo.delete(0, "end")
        if dados.get("empresa_arquivo"):
            self.entry_empresa_arquivo.insert(0, dados["empresa_arquivo"])
        self._item_carregado = dados.get("item") or 1
        self.label_deteccao_sexo.configure(text="")
        self.label_busca.configure(
            text=f"Dados recuperados de: {caminho.name}", foreground="#1a7a1a"
        )
        self._atualizar_preview()

        try:
            contexto_para_salvar = {campo: dados[campo] for campo in
                                     ("empresa", "cnpj", "segurado", "sexo", "cpf", "nit", "especie", "nb")}
            salvar_dados_beneficio(
                contexto_para_salvar, PASTA_DADOS_BENEFICIOS_PADRAO, item=self._item_carregado,
                empresa_arquivo=dados.get("empresa_arquivo"),
            )
        except OSError:
            pass  # não impede o uso dos dados já carregados na tela

    def _detectar_sexo(self, _event=None):
        nome = self.entries["segurado"].get().strip()
        if not nome:
            self.label_deteccao_sexo.configure(text="")
            return

        resultado = inferir_genero(nome)
        if resultado in ("M", "M_PROVAVEL"):
            self.var_sexo.set("M")
            self.label_deteccao_sexo.configure(text="(sexo detectado automaticamente pelo nome)")
        elif resultado in ("F", "F_PROVAVEL"):
            self.var_sexo.set("F")
            self.label_deteccao_sexo.configure(text="(sexo detectado automaticamente pelo nome)")
        else:
            self.label_deteccao_sexo.configure(text="(nome ambíguo — confirme o sexo manualmente)")

    def _atualizar_beneficio(self, *_args):
        nome = ESPECIE_NOMES.get(self.var_especie.get(), "")
        nome_exibicao = nome[:1].upper() + nome[1:] if nome else nome
        self.label_beneficio.configure(text=f"Benefício: {nome_exibicao}")

    def _atualizar_preview(self, *_args):
        empresa = self.entries["empresa"].get().strip()
        segurado = self.entries["segurado"].get().strip()
        nb = self.entry_nb.get().strip()
        sexo = self.var_sexo.get()
        empresa_arquivo = self.entry_empresa_arquivo.get().strip()

        if not empresa and not segurado and not nb:
            self.label_preview.configure(text="")
            return

        tratamento = "Segurado" if sexo == "M" else "Segurada"
        nome_seg = _dois_primeiros_nomes(segurado).upper() if segurado else "NOME"
        if empresa_arquivo:
            emp_curta = sanitizar_nome_arquivo_windows(empresa_arquivo).upper()
        elif empresa:
            emp_curta = _primeiro_nome_empresa(empresa).upper()
        else:
            emp_curta = "EMPRESA"
        nb_curto = _ultimos_digitos(nb) if nb else "NNN"

        preview = (
            f"{self._item_carregado}{self.SUFIXO_ITEM} {self.PREFIXO_ARQUIVO} - {tratamento} - {nome_seg} "
            f"- NB - {nb_curto} - Empresa - {emp_curta}.docx"
        )
        self.label_preview.configure(text=f"Arquivo: {preview}")

    def _escolher_saida(self):
        caminho = filedialog.askdirectory(title="Selecionar pasta de salvamento")
        if caminho:
            self.entry_saida.delete(0, "end")
            self.entry_saida.insert(0, caminho)

    def _abrir_pasta_saida(self):
        saida = self.entry_saida.get().strip()
        if saida and Path(saida).is_dir():
            os.startfile(saida)

    def limpar(self):
        self.entry_nb.delete(0, "end")
        for entrada in self.entries.values():
            entrada.delete(0, "end")
        self.entry_empresa_arquivo.delete(0, "end")
        self.label_preview.configure(text="")
        self.label_busca.configure(text="")
        self.var_sexo.set("M")
        self.label_deteccao_sexo.configure(text="")
        self.var_especie.set(RADIOS_ESPECIE[0])
        self.entry_data.delete(0, "end")
        self.entry_data.insert(0, date.today().strftime("%d/%m/%Y"))
        self.entry_saida.delete(0, "end")
        self.var_pdf.set(True)
        self._item_carregado = 1
        self.botao_abrir_pasta.config(state="disabled")
        self._ultimo_pdf = None
        self.botao_abrir_pdf.config(state="disabled")
        self.entry_nb.focus_set()

    def _coletar_linha(self) -> dict:
        linha = {chave: self.entries[chave].get().strip() for chave, _ in RAZAO_SOCIAL_CAMPOS if chave != "nb"}
        linha["nb"] = self.entry_nb.get().strip()
        linha["sexo"] = self.var_sexo.get()
        linha["especie"] = self.var_especie.get()
        linha["data"] = self.entry_data.get().strip()
        return linha

    def _iniciar_progresso(self):
        self.botao_gerar.config(state="disabled")
        self.botao_limpar.config(state="disabled")
        self.frame_progresso.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        self.barra_progresso.start(10)

    def _parar_progresso(self):
        self.barra_progresso.stop()
        self.frame_progresso.grid_remove()
        self.botao_gerar.config(state="normal")
        self.botao_limpar.config(state="normal")

    def gerar(self):
        if not TEMPLATE_EXIGENCIA_PADRAO.exists():
            messagebox.showerror(
                "Template não encontrado",
                f"Rode preparar_template_exigencia.py antes (esperado em {TEMPLATE_EXIGENCIA_PADRAO}).",
            )
            return

        saida_texto = self.entry_saida.get().strip()
        if not saida_texto:
            messagebox.showwarning("Aviso", "Escolha a pasta de salvamento antes de gerar.")
            return
        saida = Path(saida_texto)
        linha = self._coletar_linha()
        item = self._item_carregado
        empresa_arquivo_override = self.entry_empresa_arquivo.get().strip() or None

        try:
            destino_previsto = calcular_destino(
                linha, saida, item=item, empresa_arquivo_override=empresa_arquivo_override,
                prefixo_arquivo=self.PREFIXO_ARQUIVO, sufixo_item=self.SUFIXO_ITEM,
            )
        except ErroValidacao as exc:
            messagebox.showwarning("Aviso", str(exc))
            return

        candidatos = [destino_previsto]
        if self.var_pdf.get():
            candidatos.append(calcular_destino_pdf(
                linha, saida, item=item, empresa_arquivo_override=empresa_arquivo_override,
                prefixo_arquivo=self.PREFIXO_ARQUIVO, sufixo_item=self.SUFIXO_ITEM,
            ))
        existentes = [p for p in candidatos if p.exists()]
        if not _confirmar_sobrescricao(self.frame.winfo_toplevel(), existentes):
            return

        gerar_pdf = self.var_pdf.get()
        self._iniciar_progresso()
        threading.Thread(
            target=self._gerar_worker, args=(linha, saida, item, empresa_arquivo_override, gerar_pdf), daemon=True
        ).start()

    def _gerar_worker(self, linha, saida, item, empresa_arquivo_override, gerar_pdf: bool) -> None:
        pythoncom.CoInitialize()
        try:
            try:
                destino = gerar_arquivo(
                    linha, TEMPLATE_EXIGENCIA_PADRAO, saida, item=item, empresa_arquivo_override=empresa_arquivo_override,
                    prefixo_arquivo=self.PREFIXO_ARQUIVO, sufixo_item=self.SUFIXO_ITEM,
                    pasta_dados_beneficios=PASTA_DADOS_BENEFICIOS_PADRAO,
                )
            except ErroValidacao as exc:
                self.frame.after(0, self._gerar_falhou, "Aviso", str(exc), messagebox.showwarning)
                return
            except Exception as exc:
                registrar_erro("Gerar - falha inesperada", exc)
                self.frame.after(0, self._gerar_falhou, "Erro ao gerar", str(exc), messagebox.showerror)
                return

            mensagem = f"Cumprimento de exigência gerado em:\n{destino}"
            destino_pdf_gerado = None
            if gerar_pdf:
                destino_pdf = calcular_destino_pdf(
                    linha, saida, item=item, empresa_arquivo_override=empresa_arquivo_override,
                    prefixo_arquivo=self.PREFIXO_ARQUIVO, sufixo_item=self.SUFIXO_ITEM,
                )
                try:
                    with ConversorPDF() as conversor:
                        conversor.converter(destino, destino_pdf)
                    mensagem += f"\n\nPDF gerado em:\n{destino_pdf}"
                    destino_pdf_gerado = destino_pdf
                except ConversorPDFIndisponivel as exc:
                    mensagem += f"\n\nAviso: PDF não foi gerado ({exc})"

            logger.info(
                "Cumprimento de exigência gerado: %s (nb=%s, empresa=%s, pdf=%s)",
                destino, linha.get("nb"), linha.get("empresa"), destino_pdf_gerado,
            )
        finally:
            pythoncom.CoUninitialize()

        self.frame.after(0, self._gerar_concluiu, mensagem, destino_pdf_gerado)

    def _gerar_falhou(self, titulo: str, mensagem: str, mostrar) -> None:
        self._parar_progresso()
        mostrar(titulo, mensagem)

    def _gerar_concluiu(self, mensagem: str, destino_pdf_gerado: Path | None) -> None:
        self._parar_progresso()
        self.botao_abrir_pasta.config(state="normal")
        if destino_pdf_gerado is not None:
            self._ultimo_pdf = destino_pdf_gerado
            self.botao_abrir_pdf.config(state="normal")
        messagebox.showinfo("Sucesso", mensagem)


class DialogoConfirmarSexo:
    """Janela modal que pede para confirmar o sexo de segurados cujo nome
    não pôde ser identificado automaticamente, antes de gerar o lote.
    Devolve um dict {número da linha: 'M'/'F'}, ou None se cancelado."""

    def __init__(self, master, pendentes: list[tuple[int, dict]]):
        self.resultado = None
        self.vars: dict[int, tk.StringVar] = {}

        self.top = ttk.Toplevel(master)
        self.top.title("Confirme o sexo dos segurados")
        self.top.transient(master)
        self.top.minsize(480, 280)

        largura, altura = 560, 420
        master.update_idletasks()
        x = master.winfo_rootx() + (master.winfo_width() - largura) // 2
        y = master.winfo_rooty() + (master.winfo_height() - altura) // 2
        self.top.geometry(f"{largura}x{altura}+{max(x, 0)}+{max(y, 0)}")
        self.top.grab_set()

        ttk.Label(
            self.top,
            text=(
                "Não foi possível identificar automaticamente o sexo dos segurados "
                "abaixo pelo nome. Selecione antes de gerar os requerimentos:"
            ),
            wraplength=520,
            padding=(16, 16, 16, 8),
        ).pack(anchor="w")

        container = ttk.Frame(self.top, padding=(16, 0, 8, 8))
        container.pack(fill="both", expand=True)

        canvas = tk.Canvas(container, highlightthickness=0, background=tema.COR_FUNDO)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        interno = ttk.Frame(canvas)

        interno.bind("<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=interno, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind(
            "<MouseWheel>", lambda e: canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
        )

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        for numero, linha in pendentes:
            nome = str(linha.get("segurado", "")).strip() or "(sem nome)"
            linha_frame = ttk.Frame(interno)
            linha_frame.pack(fill="x", pady=4, padx=(0, 8))
            ttk.Label(linha_frame, text=f"Linha {numero}: {nome}", width=40, anchor="w").pack(
                side="left"
            )
            var = tk.StringVar(value="")
            self.vars[numero] = var
            ttk.Radiobutton(linha_frame, text="Masculino", variable=var, value="M").pack(
                side="left", padx=(8, 8)
            )
            ttk.Radiobutton(linha_frame, text="Feminino", variable=var, value="F").pack(side="left")

        botoes = ttk.Frame(self.top, padding=(16, 8, 16, 16))
        botoes.pack(fill="x")
        ttk.Button(
            botoes, text="Cancelar (não gera nada)", command=self._cancelar, bootstyle="primary-outline"
        ).pack(side="right", padx=(8, 0))
        ttk.Button(
            botoes, text="Confirmar e gerar", command=self._confirmar, bootstyle="secondary"
        ).pack(side="right")

        self.top.protocol("WM_DELETE_WINDOW", self._cancelar)

    def _confirmar(self):
        faltando = [numero for numero, var in self.vars.items() if not var.get()]
        if faltando:
            messagebox.showwarning(
                "Aviso",
                f"Selecione o sexo para todas as linhas (faltando: "
                f"{', '.join(str(n) for n in faltando)}).",
                parent=self.top,
            )
            return
        self.resultado = {numero: var.get() for numero, var in self.vars.items()}
        self.top.destroy()

    def _cancelar(self):
        self.resultado = None
        self.top.destroy()

    def exibir(self) -> dict | None:
        self.top.wait_window()
        return self.resultado


class AbaLote:
    def __init__(self, parent):
        self.frame = ttk.Frame(parent, padding=(20, 18))
        self.frame.columnconfigure(0, weight=1)
        self._cancelar = threading.Event()
        self._saida_manual = False

        cartao = ttk.Labelframe(self.frame, text=" Arquivos ", padding=18, bootstyle="secondary")
        cartao.grid(row=0, column=0, sticky="ew", pady=(0, 18))
        cartao.columnconfigure(0, weight=1)

        ttk.Label(cartao, text="Planilha (.xlsx ou .csv)", font=("Segoe UI", 9, "bold")).grid(
            row=0, column=0, sticky="w"
        )
        linha1 = ttk.Frame(cartao)
        linha1.grid(row=1, column=0, sticky="ew", pady=(6, 16))
        linha1.columnconfigure(0, weight=1)
        self.entry_planilha = ttk.Entry(linha1)
        self.entry_planilha.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(
            linha1, text="Procurar...", command=self._escolher_planilha, bootstyle="primary-outline"
        ).grid(row=0, column=1)
        linha_download = ttk.Frame(linha1)
        linha_download.grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))
        ttk.Button(
            linha_download, text="📥  Baixar planilha modelo",
            command=self._baixar_planilha_modelo, bootstyle="secondary-outline",
        ).pack(side="left")
        ttk.Label(
            linha_download,
            text="Planilha em branco para preencher e importar no programa",
            foreground="#888888", font=("Segoe UI", 8),
        ).pack(side="left", padx=(6, 0))

        ttk.Label(cartao, text="Pasta de salvamento", font=("Segoe UI", 9, "bold")).grid(
            row=2, column=0, sticky="w"
        )
        linha2 = ttk.Frame(cartao)
        linha2.grid(row=3, column=0, sticky="ew", pady=(6, 16))
        linha2.columnconfigure(0, weight=1)
        self.entry_saida = ttk.Entry(linha2)
        self.entry_saida.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self._saida_manual = False
        ttk.Button(
            linha2, text="Procurar...", command=self._escolher_saida, bootstyle="primary-outline"
        ).grid(row=0, column=1)

        linha3 = ttk.Frame(cartao)
        linha3.grid(row=4, column=0, sticky="ew", pady=(0, 6))
        linha3.columnconfigure((0, 1), weight=1)
        sub_cnpj = ttk.Frame(linha3)
        sub_cnpj.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Label(sub_cnpj, text="CNPJ (opcional, sobrescreve todas as linhas)", font=("Segoe UI", 9, "bold")).pack(
            anchor="w"
        )
        self.entry_cnpj = ttk.Entry(sub_cnpj)
        self.entry_cnpj.pack(fill="x", pady=(6, 0))

        sub_data = ttk.Frame(linha3)
        sub_data.grid(row=0, column=1, sticky="ew")
        ttk.Label(sub_data, text="Data (opcional, sobrescreve todas as linhas)", font=("Segoe UI", 9, "bold")).pack(
            anchor="w"
        )
        self.entry_data = ttk.Entry(sub_data)
        self.entry_data.pack(fill="x", pady=(6, 0))

        sub_empresa_arquivo = ttk.Frame(linha3)
        sub_empresa_arquivo.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        ttk.Label(
            sub_empresa_arquivo,
            text="Nome da empresa no arquivo word e pdf (opcional, sobrescreve todas as linhas)",
            font=("Segoe UI", 9, "bold"),
        ).pack(anchor="w")
        self.entry_empresa_arquivo = ttk.Entry(sub_empresa_arquivo)
        self.entry_empresa_arquivo.pack(fill="x", pady=(6, 0))

        self.var_pdf = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            cartao,
            text="Gerar também em PDF (requer Microsoft Word instalado)",
            variable=self.var_pdf,
            bootstyle="secondary",
        ).grid(row=5, column=0, sticky="w", pady=(10, 0))

        botoes = ttk.Frame(self.frame)
        botoes.grid(row=1, column=0, sticky="ew", pady=(2, 14))
        self.botao_abrir_pasta = ttk.Button(
            botoes,
            text="📂  Abrir pasta de salvamento",
            command=self._abrir_pasta_saida,
            bootstyle="primary-outline",
            state="disabled",
        )
        self.botao_abrir_pasta.pack(side="left")
        self.botao_limpar = ttk.Button(
            botoes, text="🧹  Limpar", command=self.limpar, bootstyle="danger-outline"
        )
        self.botao_limpar.pack(side="left", padx=(10, 0))
        self.botao_gerar = ttk.Button(
            botoes, text="📝  Gerar todos", command=self.gerar, bootstyle="secondary", width=18
        )
        self.botao_gerar.pack(side="right")

        self.frame_progresso = ttk.Frame(self.frame)
        self.barra_progresso = ttk.Progressbar(self.frame_progresso, mode="determinate", bootstyle="secondary")
        self.barra_progresso.pack(fill="x", pady=(0, 4))
        linha_prog = ttk.Frame(self.frame_progresso)
        linha_prog.pack(fill="x")
        self.label_progresso = ttk.Label(linha_prog, text="", font=("Segoe UI", 9))
        self.label_progresso.pack(side="left")
        self.botao_cancelar = ttk.Button(
            linha_prog, text="⛔  Suspender", command=self._cancelar_geracao, bootstyle="danger"
        )
        self.botao_cancelar.pack(side="right")

        moldura = ttk.Labelframe(self.frame, text=" Resultado ", padding=14, bootstyle="secondary")
        moldura.grid(row=3, column=0, sticky="nsew")
        self.frame.rowconfigure(3, weight=1)

        self.texto_resultado = ScrolledText(moldura, autohide=True, bootstyle="secondary")
        self.texto_resultado.pack(fill="both", expand=True)
        self.texto_resultado.text.configure(font=("Segoe UI", 10), padx=8, pady=8, relief="flat")
        self._mostrar_mensagem_inicial()

    def _mostrar_mensagem_inicial(self):
        self.texto_resultado.text.insert("end", 'Selecione a planilha e clique em "Gerar todos".')
        self.texto_resultado.text.configure(state="disabled")

    def _baixar_planilha_modelo(self):
        if not PLANILHA_MODELO.is_file():
            messagebox.showerror(
                "Arquivo não encontrado",
                f"A planilha modelo não foi encontrada em:\n{PLANILHA_MODELO}",
            )
            return
        destino = filedialog.asksaveasfilename(
            title="Salvar planilha modelo",
            initialfile="SOLICITAÇÕES GERID.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Planilha Excel", "*.xlsx")],
        )
        if not destino:
            return
        try:
            shutil.copyfile(PLANILHA_MODELO, destino)
        except OSError as exc:
            messagebox.showerror("Erro ao salvar", str(exc))
            return
        if messagebox.askyesno(
            "Planilha salva",
            f"Planilha modelo salva em:\n{destino}\n\nAbrir agora?",
        ):
            os.startfile(destino)

    def _escolher_planilha(self):
        caminho = filedialog.askopenfilename(
            title="Selecionar planilha",
            filetypes=[("Planilhas", "*.xlsx *.csv"), ("Todos os arquivos", "*.*")],
        )
        if caminho:
            self.entry_planilha.delete(0, "end")
            self.entry_planilha.insert(0, caminho)
            if not self._saida_manual:
                self.entry_saida.delete(0, "end")
                self.entry_saida.insert(0, str(Path(caminho).parent))

    def _escolher_saida(self):
        caminho = filedialog.askdirectory(title="Selecionar pasta de salvamento")
        if caminho:
            self.entry_saida.delete(0, "end")
            self.entry_saida.insert(0, caminho)
            self._saida_manual = True

    def _abrir_pasta_saida(self):
        saida = self.entry_saida.get().strip()
        if saida and Path(saida).is_dir():
            os.startfile(saida)

    def limpar(self):
        self.entry_planilha.delete(0, "end")
        self.entry_saida.delete(0, "end")
        self._saida_manual = False
        self.entry_cnpj.delete(0, "end")
        self.entry_data.delete(0, "end")
        self.entry_empresa_arquivo.delete(0, "end")
        self.var_pdf.set(True)
        self.botao_abrir_pasta.config(state="disabled")
        self.texto_resultado.text.configure(state="normal")
        self.texto_resultado.text.delete("1.0", "end")
        self._mostrar_mensagem_inicial()

    def _detectar_ambiguos(self, linhas: list[dict]) -> list[tuple[int, dict]]:
        """Linhas sem 'sexo' explícito cujo nome o algoritmo não consegue
        decidir com segurança (AMBIGUO ou desconhecido) - candidatas à
        confirmação manual antes de gerar o lote."""
        pendentes = []
        for numero, linha in enumerate(linhas, start=2):  # linha 1 = cabeçalho
            sexo = linha.get("sexo")
            if sexo and str(sexo).strip():
                continue
            nome = str(linha.get("segurado", "")).strip()
            if not nome:
                continue
            resultado = inferir_genero(nome)
            if resultado not in ("M", "M_PROVAVEL", "F", "F_PROVAVEL"):
                pendentes.append((numero, linha))
        return pendentes

    def _escrever_resultado(self, linhas: list[str]):
        self.texto_resultado.text.configure(state="normal")
        self.texto_resultado.text.delete("1.0", "end")
        self.texto_resultado.text.insert("end", "\n".join(linhas))
        self.texto_resultado.text.configure(state="disabled")

    def gerar(self):
        planilha_texto = self.entry_planilha.get().strip()
        if not planilha_texto:
            messagebox.showwarning("Aviso", "Selecione uma planilha.")
            return
        planilha = Path(planilha_texto)

        saida_texto = self.entry_saida.get().strip()
        saida = Path(saida_texto) if saida_texto else SAIDA_PADRAO

        if not TEMPLATE_PADRAO.exists():
            messagebox.showerror(
                "Template não encontrado",
                f"Rode preparar_template.py antes (esperado em {TEMPLATE_PADRAO}).",
            )
            return

        data_override = self.entry_data.get().strip() or None
        cnpj_override = self.entry_cnpj.get().strip() or None
        empresa_arquivo_override = self.entry_empresa_arquivo.get().strip() or None

        try:
            linhas = ler_planilha(planilha)
        except ErroValidacao as exc:
            messagebox.showerror("Erro na planilha", str(exc))
            return

        pendentes = self._detectar_ambiguos(linhas)
        if pendentes:
            escolhas = DialogoConfirmarSexo(self.frame.winfo_toplevel(), pendentes).exibir()
            if escolhas is None:
                return
            for numero, linha in pendentes:
                linha["sexo"] = escolhas[numero]

        gerar_pdf = self.var_pdf.get()

        existentes = []
        for posicao, linha in enumerate(linhas, start=1):
            item = resolver_item(linha, posicao)
            try:
                destino_previsto = calcular_destino(linha, saida, data_override, cnpj_override, item, empresa_arquivo_override)
            except ErroValidacao:
                continue  # erro de verdade aparece na hora de gerar essa linha
            candidatos = [destino_previsto]
            if gerar_pdf:
                candidatos.append(calcular_destino_pdf(linha, saida, data_override, cnpj_override, item, empresa_arquivo_override))
            existentes.extend(p for p in candidatos if p.exists())
        if not _confirmar_sobrescricao(self.frame.winfo_toplevel(), existentes):
            return

        self._cancelar.clear()
        self.botao_gerar.config(state="disabled")
        self.botao_limpar.config(state="disabled")
        self.botao_abrir_pasta.config(state="disabled")
        self.botao_cancelar.config(state="normal")
        self.barra_progresso.configure(value=0, maximum=len(linhas))
        self.label_progresso.configure(text=f"Gerando 0 de {len(linhas)}...")
        self.frame_progresso.grid(row=2, column=0, sticky="ew", pady=(0, 14))
        self.texto_resultado.text.configure(state="normal")
        self.texto_resultado.text.delete("1.0", "end")
        self.texto_resultado.text.configure(state="disabled")

        threading.Thread(
            target=self._gerar_worker,
            args=(linhas, saida, data_override, cnpj_override, empresa_arquivo_override, gerar_pdf),
            daemon=True,
        ).start()

    def _gerar_worker(
        self, linhas: list[dict], saida: Path, data_override, cnpj_override, empresa_arquivo_override, gerar_pdf: bool
    ) -> None:
        """Roda numa thread separada para não congelar a interface durante o
        lote (cada arquivo, e principalmente a conversão para PDF via Word,
        pode levar um tempo perceptível). Atualiza a barra de progresso e
        mostra o resultado de volta na thread principal via `.after(0, ...)`."""
        pythoncom.CoInitialize()
        conversor = None
        aviso_pdf = None
        if gerar_pdf:
            try:
                conversor = ConversorPDF()
            except ConversorPDFIndisponivel as exc:
                aviso_pdf = str(exc)

        total = len(linhas)
        saida_log = []
        gerados = 0
        erros = 0
        try:
            for posicao, (numero, linha) in enumerate(enumerate(linhas, start=2), start=1):
                if self._cancelar.is_set():
                    break
                log_line = ""
                item = resolver_item(linha, posicao)
                try:
                    destino = gerar_arquivo(
                        linha, TEMPLATE_PADRAO, saida, data_override, cnpj_override, item, empresa_arquivo_override,
                        pasta_dados_beneficios=PASTA_DADOS_BENEFICIOS_PADRAO,
                    )
                except ErroValidacao as exc:
                    log_line = f"linha {numero}: ERRO - {exc}"
                    saida_log.append(log_line)
                    erros += 1
                else:
                    if conversor is not None:
                        destino_pdf = calcular_destino_pdf(linha, saida, data_override, cnpj_override, item, empresa_arquivo_override)
                        conversor.converter(destino, destino_pdf)
                    log_line = f"linha {numero}: gerado - {destino.name}"
                    saida_log.append(log_line)
                    gerados += 1
                self.frame.after(0, self._acrescentar_resultado, log_line)
                self.frame.after(0, self._atualizar_progresso, posicao, total, numero)
        finally:
            if conversor is not None:
                conversor.fechar()
            pythoncom.CoUninitialize()

        self.frame.after(
            0, self._gerar_concluiu, saida_log, gerados, erros, conversor is not None, aviso_pdf, saida, self._cancelar.is_set()
        )

    def _atualizar_progresso(self, posicao: int, total: int, numero: int) -> None:
        self.barra_progresso.configure(value=posicao)
        self.label_progresso.configure(text=f"Gerando requerimento {posicao} de {total} (linha {numero})...")

    def _acrescentar_resultado(self, linha: str) -> None:
        self.texto_resultado.text.configure(state="normal")
        self.texto_resultado.text.insert("end", linha + "\n")
        self.texto_resultado.text.see("end")
        self.texto_resultado.text.configure(state="disabled")

    def _cancelar_geracao(self):
        self._cancelar.set()
        self.botao_cancelar.configure(state="disabled", text="⛔  Suspendendo...")

    def _gerar_concluiu(
        self,
        saida_log: list[str],
        gerados: int,
        erros: int,
        gerou_pdf: bool,
        aviso_pdf: str | None,
        saida: Path,
        cancelado: bool = False,
    ) -> None:
        self.frame_progresso.grid_remove()
        self.botao_cancelar.configure(state="normal", text="⛔  Suspender")
        self.botao_gerar.config(state="normal")
        self.botao_limpar.config(state="normal")
        if gerados > 0:
            self.botao_abrir_pasta.config(state="normal")

        if cancelado:
            resumo = f"Geração suspensa. {gerados} requerimento(s) gerado(s) antes da interrupção."
            if erros:
                resumo += f"\n{erros} linha(s) com erro."
            messagebox.showwarning("Suspenso", resumo)
            return

        resumo = f"{gerados} requerimento(s) gerado(s)"
        if gerou_pdf:
            resumo += " (.docx e .pdf)"
        resumo += f" em:\n{saida}"
        if aviso_pdf:
            resumo += f"\n\nAviso: PDF não foi gerado ({aviso_pdf})"

        if erros:
            messagebox.showwarning(
                "Concluído com erros",
                f"{gerados} requerimento(s) gerado(s), {erros} linha(s) com erro. Veja a lista na tela.",
            )
        else:
            messagebox.showinfo("Sucesso", resumo)


TEXTO_MANUAL = """\
1. GERAR EM LOTE (PLANILHA)

Escolha uma planilha .xlsx ou .csv e clique em "Gerar todos". Aceita dois
formatos:
   • Planilha "SOLICITAÇÕES GERID" usada no dia a dia — lida direto, sem
     reformatar.
   • Planilha simples com colunas: empresa, segurado, cpf, nit, especie, nb
     (sexo e cnpj são opcionais).

2. REQUERIMENTO INDIVIDUAL

Preencha os dados e clique em "Gerar requerimento". O sexo é detectado
automaticamente pelo nome (pode corrigir manualmente se vier ambíguo).

3. CUMPRIMENTO DE EXIGÊNCIA

Para quando o INSS abre uma diligência pedindo procuração ou outros
documentos. Digite o NB e clique em "Buscar dados salvos": se já existe
um requerimento gerado para esse benefício (nesta instalação), os campos
são preenchidos automaticamente.

Se o NB não tiver histórico salvo (ex.: requerimento gerado antes desta
versão, ou em outra máquina), clique em "Carregar requerimento" e
selecione o .docx ou .pdf já gerado — os dados são lidos direto do
documento. Também é possível simplesmente arrastar o arquivo (.docx ou
.pdf) até a janela do REQUERID, em qualquer aba: ele troca sozinho para
esta aba e carrega os dados.

O nome do arquivo gerado acompanha o número do requerimento original
(ex.: requerimento "10. Requerimento..." gera "10.3 Cumprimento de
Exigência...").

4. ESPÉCIE DO BENEFÍCIO

   Acidentários:
   B91 – Auxílio por incapacidade temporária por acidente de trabalho
   B92 – Aposentadoria por incapacidade permanente por acidente de trabalho
   B93 – Pensão por morte por acidente de trabalho
   B94 – Auxílio-acidente por acidente de trabalho

   Previdenciários:
   B31 – Auxílio por incapacidade temporária previdenciário
   B32 – Aposentadoria por incapacidade permanente previdenciário
   B36 – Auxílio-acidente previdenciário
   B41 – Aposentadoria por idade
   B42 – Aposentadoria por tempo de contribuição
   B46 – Aposentadoria especial

5. SEXO DO SEGURADO

Detectado automaticamente pelo nome. Se for ambíguo, o programa pede
confirmação antes de gerar (individual: pelos botões; lote: numa janela
única para todos os casos).

6. GERAR TAMBÉM EM PDF

Marque a opção. Requer Word instalado — sem ele, só o .docx é gerado.

DICAS
- A pasta de salvamento é obrigatória: gerar sem escolher uma pasta
  avisa e não salva nada.
- Avisa antes de substituir arquivo já existente.
- "Limpar" reseta os campos sem fechar o programa.
- "Abrir pasta de salvamento" mostra os arquivos gerados; "Abrir PDF
  gerado" abre direto o último PDF gerado nessa aba.
- "Nome da empresa no arquivo word e pdf" (opcional): o sistema
  identifica automaticamente o nome da empresa na maioria dos casos.
  Preencha somente se o resultado automático não for satisfatório.
  Exemplos:
    HONDA AUTOMOVEIS LTDA → auto: HONDA   (ok, não precisa preencher)
    LOJAS COLOMBO S.A.    → auto: LOJAS   (preencher: LOJAS COLOMBO)
    BANCO DO BRASIL S.A.  → auto: BANCO   (preencher: BANCO DO BRASIL)"""


class Janela:
    def __init__(self, root, nome_app: str, descricao_app: str):
        self.root = root
        # root.report_callback_exception (rede de segurança para exceção não
        # tratada num callback do Tkinter) é ligado em main.py, onde `root`
        # é criado.

        self._instalar_menu_contexto(root)

        self._montar_banner(root, nome_app, descricao_app)

        barra_superior = ttk.Frame(root, padding=(14, 6, 14, 0))
        barra_superior.pack(fill="x")
        ttk.Button(
            barra_superior, text="❓ Manual rápido", command=self._abrir_manual, bootstyle="primary-link"
        ).pack(side="right")
        ttk.Button(
            barra_superior, text="📋 Notas de atualização",
            command=self._abrir_notas_atualizacao, bootstyle="primary-link",
        ).pack(side="right", padx=(0, 8))

        rodape = ttk.Frame(root, padding=(14, 11, 10, 8))
        rodape.pack(fill="x", side="bottom")
        caminho_logo = pasta_recursos() / "Logo-RS-completa-colorida.ico"
        if caminho_logo.exists():
            self._imagem_logo = _carregar_imagem_altura(caminho_logo, 24)
            tk.Label(rodape, image=self._imagem_logo, borderwidth=0, background=tema.COR_FUNDO).pack(side="left")
        ttk.Label(rodape, text="versão 3.10", bootstyle="secondary", font=("Segoe UI", 8)).pack(side="right")

        notebook = ttk.Notebook(root)
        notebook.pack(fill="both", expand=True, padx=14, pady=(14, 0))

        aba_lote = AbaLote(notebook)
        notebook.add(aba_lote.frame, text="  Gerar em lote (planilha)  ")

        aba_individual = AbaIndividual(notebook)
        notebook.add(aba_individual.frame, text="  Requerimento individual  ")

        aba_exigencia = AbaExigencia(notebook)
        notebook.add(aba_exigencia.frame, text="  Cumprimento de exigência  ")

        # Arrastar-e-soltar via tkinterdnd2 em vez do pacote windnd: o windnd
        # tem um bug real de estouro de buffer (informa pra DragQueryFileW o
        # tamanho do buffer em BYTES, quando a API espera a contagem de
        # CARACTERES - metade do valor certo) que estoura a pilha e derruba
        # o programa (STATUS_STACK_BUFFER_OVERRUN) para qualquer caminho de
        # arquivo com mais de ~260 caracteres, bem comum em pastas
        # aninhadas. tkinterdnd2 é a biblioteca padrão/madura pra isso.
        TkinterDnD._require(root)
        root.TkdndVersion = TkinterDnD.TkdndVersion
        if TkinterDnD.DnDWrapper not in type(root).__mro__:
            root.__class__ = type(type(root).__name__, (type(root), TkinterDnD.DnDWrapper), {})
        root.drop_target_register(DND_FILES)

        def _ao_soltar_arquivos(event):
            try:
                # event.data vem no formato de lista do Tcl (caminhos com
                # espaço ficam entre chaves) - tk.splitlist trata isso certo.
                arquivos = root.tk.splitlist(event.data)
                # só nos interessa requerimento .docx/.pdf, sempre carregado
                # na aba de Cumprimento de Exigência.
                caminho = next(
                    (a for a in arquivos if Path(a).suffix.lower() in (".docx", ".pdf")), None
                )
                if caminho is None:
                    messagebox.showwarning(
                        "Arquivo não reconhecido", "Solte um requerimento gerado em .docx ou .pdf."
                    )
                    return "refuse"
                notebook.select(aba_exigencia.frame)
                aba_exigencia.carregar_de_caminho(caminho)
                return "copy"
            except Exception as exc:
                registrar_erro(f"Arrastar e soltar - falha inesperada (event.data={event.data!r})", exc)
                messagebox.showerror(
                    "Erro ao carregar",
                    f"Ocorreu um erro inesperado ao processar o arquivo solto.\n\nDetalhes salvos em:\n{CAMINHO_LOG}",
                )
                return "refuse"

        root.dnd_bind("<<Drop>>", _ao_soltar_arquivos)

    def _instalar_menu_contexto(self, root) -> None:
        """Menu de botão direito (Recortar/Copiar/Colar/Selecionar tudo) em
        todos os campos de texto do aplicativo. Instalado por classe de
        widget, vale também para os campos de janelas e diálogos abertos
        depois."""
        menu = tk.Menu(root, tearoff=0)

        def mostrar(event):
            w = event.widget
            try:
                if str(w.cget("state")) == "disabled":
                    return
                w.focus_set()
            except Exception:
                return

            try:
                if isinstance(w, tk.Text):
                    tem_selecao = bool(w.tag_ranges("sel"))
                else:
                    tem_selecao = w.selection_present()
            except Exception:
                tem_selecao = False
            editavel = str(w.cget("state")) == "normal"

            menu.delete(0, "end")
            menu.add_command(
                label="Recortar", accelerator="Ctrl+X",
                state="normal" if tem_selecao and editavel else "disabled",
                command=lambda: w.event_generate("<<Cut>>"),
            )
            menu.add_command(
                label="Copiar", accelerator="Ctrl+C",
                state="normal" if tem_selecao else "disabled",
                command=lambda: w.event_generate("<<Copy>>"),
            )
            menu.add_command(
                label="Colar", accelerator="Ctrl+V",
                state="normal" if editavel else "disabled",
                command=lambda: w.event_generate("<<Paste>>"),
            )
            menu.add_separator()
            menu.add_command(
                label="Selecionar tudo", accelerator="Ctrl+A",
                command=lambda: w.event_generate("<<SelectAll>>"),
            )
            try:
                menu.tk_popup(event.x_root, event.y_root)
            finally:
                menu.grab_release()
            return "break"

        for classe in ("TEntry", "Entry", "Text", "TSpinbox", "TCombobox"):
            root.bind_class(classe, "<Button-3>", mostrar)

    def _montar_banner(self, root, nome_app: str, descricao_app: str) -> None:
        self._nome_app = nome_app
        self._descricao_app = descricao_app
        self._label_banner = tk.Label(root, borderwidth=0, anchor="w")
        self._label_banner.pack(fill="x")
        self._largura_banner_atual = 0
        self._redesenho_banner_pendente = None
        self._atualizar_banner(LARGURA_JANELA)
        # com a janela redimensionável, o banner é redesenhado do zero na
        # nova largura (esticar a imagem existente borraria) - com atraso
        # curto, para não redesenhar a cada pixel durante o arraste da borda
        root.bind("<Configure>", self._ao_redimensionar_janela, add="+")

    def _atualizar_banner(self, largura: int) -> None:
        largura = max(int(largura), 400)
        if largura == self._largura_banner_atual:
            return
        self._largura_banner_atual = largura
        imagem = visual.gerar_banner(
            largura=largura,
            altura=ALTURA_BANNER,
            cor_inicio=tema.COR_PRIMARIA,
            cor_fim="#1A1C3D",
            cor_destaque=tema.COR_SECUNDARIA,
            icone_path=str(pasta_recursos() / "GERID LOGO.ico"),
            titulo=self._nome_app,
            subtitulo=self._descricao_app,
        )
        self._imagem_banner = ImageTk.PhotoImage(imagem)
        self._label_banner.configure(image=self._imagem_banner)

    def _ao_redimensionar_janela(self, event) -> None:
        if event.widget is not self.root or event.width == self._largura_banner_atual:
            return
        if self._redesenho_banner_pendente is not None:
            self.root.after_cancel(self._redesenho_banner_pendente)
        self._redesenho_banner_pendente = self.root.after(
            120, lambda: self._atualizar_banner(self.root.winfo_width())
        )

    def _abrir_notas_atualizacao(self):
        janela = ttk.Toplevel(self.root)
        janela.title("Notas de atualização")
        janela.geometry("600x560")
        janela.resizable(False, False)
        try:
            janela.iconbitmap(str(pasta_recursos() / "GERID LOGO.ico"))
        except tk.TclError:
            pass

        ttk.Label(janela, text="📋 Notas de atualização", font=("Segoe UI", 14, "bold")).pack(
            anchor="w", padx=20, pady=(18, 2)
        )
        ttk.Label(
            janela, text="Histórico de versões do REQUERID", bootstyle="secondary",
        ).pack(anchor="w", padx=20, pady=(0, 12))

        corpo = ScrolledText(janela, autohide=True, bootstyle="secondary")
        corpo.pack(fill="both", expand=True, padx=20, pady=(0, 14))
        corpo.text.insert("end", _carregar_notas_atualizacao())
        corpo.text.configure(font=("Consolas", 9), padx=10, pady=10, state="disabled")
        corpo.text.see("1.0")

        ttk.Button(janela, text="Fechar", command=janela.destroy, bootstyle="secondary").pack(pady=(0, 18))

        janela.update_idletasks()
        x = self.root.winfo_rootx() + (self.root.winfo_width() - janela.winfo_width()) // 2
        y = self.root.winfo_rooty() + (self.root.winfo_height() - janela.winfo_height()) // 2
        janela.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        janela.transient(self.root)

    def _abrir_manual(self):
        janela = ttk.Toplevel(self.root)
        janela.title("Manual rápido")
        janela.geometry("600x560")
        janela.resizable(False, False)
        try:
            janela.iconbitmap(str(pasta_recursos() / "GERID LOGO.ico"))
        except tk.TclError:
            pass

        ttk.Label(janela, text="📖 Manual rápido", font=("Segoe UI", 14, "bold")).pack(
            anchor="w", padx=20, pady=(18, 2)
        )
        ttk.Label(
            janela,
            text="Como usar o gerador de requerimento GERID",
            bootstyle="secondary",
        ).pack(anchor="w", padx=20, pady=(0, 12))

        corpo = ScrolledText(janela, autohide=True, bootstyle="secondary")
        corpo.pack(fill="both", expand=True, padx=20, pady=(0, 14))
        corpo.text.insert("end", TEXTO_MANUAL)
        corpo.text.configure(font=("Consolas", 9), padx=10, pady=10, state="disabled")

        ttk.Button(janela, text="Fechar", command=janela.destroy, bootstyle="secondary").pack(pady=(0, 18))

        janela.update_idletasks()
        x = self.root.winfo_rootx() + (self.root.winfo_width() - janela.winfo_width()) // 2
        y = self.root.winfo_rooty() + (self.root.winfo_height() - janela.winfo_height()) // 2
        janela.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        janela.transient(self.root)
