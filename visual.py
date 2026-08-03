# -*- coding: utf-8 -*-
"""
Geração do banner (gradiente + blob decorativo + ícone + texto) usado no topo do app.
Tudo composto numa única imagem via Pillow porque ttk não suporta gradiente/sombra nativamente.
"""

from PIL import Image, ImageDraw, ImageFilter, ImageFont

FONTE_REGULAR = r"C:\Windows\Fonts\segoeui.ttf"
FONTE_BOLD = r"C:\Windows\Fonts\segoeuib.ttf"


def _hex_para_rgb(cor: str) -> tuple[int, int, int]:
    cor = cor.lstrip("#")
    return tuple(int(cor[i : i + 2], 16) for i in (0, 2, 4))


def _gradiente_diagonal(largura: int, altura: int, cor_inicio: str, cor_fim: str) -> Image.Image:
    r1, g1, b1 = _hex_para_rgb(cor_inicio)
    r2, g2, b2 = _hex_para_rgb(cor_fim)
    ts_x = [x / max(largura - 1, 1) for x in range(largura)]
    ts_y = [y / max(altura - 1, 1) for y in range(altura)]
    data = bytearray(largura * altura * 3)
    idx = 0
    for y in range(altura):
        fy = ts_y[y]
        for x in range(largura):
            t = (ts_x[x] + fy) / 2
            data[idx]     = int(r1 + (r2 - r1) * t)
            data[idx + 1] = int(g1 + (g2 - g1) * t)
            data[idx + 2] = int(b1 + (b2 - b1) * t)
            idx += 3
    return Image.frombytes("RGB", (largura, altura), bytes(data))


def _fonte_ajustada(draw: ImageDraw.ImageDraw, texto: str, caminho_fonte: str, tamanho_inicial: int, largura_maxima: int, tamanho_minimo: int = 10) -> ImageFont.FreeTypeFont:
    tamanho = tamanho_inicial
    while tamanho > tamanho_minimo:
        fonte = ImageFont.truetype(caminho_fonte, tamanho)
        if draw.textlength(texto, font=fonte) <= largura_maxima:
            return fonte
        tamanho -= 1
    return ImageFont.truetype(caminho_fonte, tamanho_minimo)


def gerar_banner(
    largura: int,
    altura: int,
    cor_inicio: str,
    cor_fim: str,
    cor_destaque: str,
    icone_path: str,
    titulo: str,
    subtitulo: str,
) -> Image.Image:
    base = _gradiente_diagonal(largura, altura, cor_inicio, cor_fim).convert("RGBA")

    blob = Image.new("RGBA", (largura, altura), (0, 0, 0, 0))
    raio = altura * 1.3
    cx, cy = largura - altura * 0.2, -altura * 0.55
    ImageDraw.Draw(blob).ellipse(
        [cx - raio, cy - raio, cx + raio, cy + raio], fill=(*_hex_para_rgb(cor_destaque), 45)
    )
    base = Image.alpha_composite(base, blob.filter(ImageFilter.GaussianBlur(22)))

    icone = Image.open(icone_path).convert("RGBA")
    tam_icone = int(altura * 0.46)
    icone = icone.resize((tam_icone, tam_icone), Image.LANCZOS)
    pos_icone = (int(altura * 0.42), (altura - tam_icone) // 2)
    base.alpha_composite(icone, pos_icone)

    draw = ImageDraw.Draw(base)
    x_texto = pos_icone[0] + tam_icone + int(altura * 0.32)
    largura_disponivel = largura - x_texto - int(altura * 0.2)
    fonte_titulo = _fonte_ajustada(draw, titulo, FONTE_BOLD, int(altura * 0.27), largura_disponivel)
    fonte_subtitulo = _fonte_ajustada(draw, subtitulo, FONTE_REGULAR, int(altura * 0.15), largura_disponivel)

    bbox_titulo = draw.textbbox((0, 0), titulo, font=fonte_titulo)
    bbox_sub = draw.textbbox((0, 0), subtitulo, font=fonte_subtitulo)
    altura_titulo = bbox_titulo[3] - bbox_titulo[1]
    altura_sub = bbox_sub[3] - bbox_sub[1]
    espaco = int(altura * 0.08)
    bloco_altura = altura_titulo + espaco + altura_sub
    y_titulo = (altura - bloco_altura) // 2 - bbox_titulo[1]
    y_sub = y_titulo + altura_titulo + espaco - bbox_sub[1] + bbox_titulo[1]

    draw.text((x_texto, y_titulo), titulo, font=fonte_titulo, fill="white")
    draw.text((x_texto, y_sub), subtitulo, font=fonte_subtitulo, fill=(199, 202, 227, 255))

    draw.rectangle([0, altura - 4, largura, altura], fill=_hex_para_rgb(cor_destaque))

    return base.convert("RGB")
