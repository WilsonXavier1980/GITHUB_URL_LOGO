#!/usr/bin/env python3
"""
Substitui a palavra ORIGINAL por CÓPIA na capa da proposta.

A palavra ORIGINAL nao e texto: faz parte da imagem de fundo da capa
(PNG 1312x1856 que ocupa a pagina inteira). Por isso a substituicao e feita
na propria imagem — apaga-se a palavra reconstruindo o fundo e desenha-se
CÓPIA com a mesma fonte (Carlito Bold), a mesma cor (#044438), a mesma altura
de maiusculas e alinhada a direita pela mesma margem.
"""
import numpy as np
import pymupdf
from PIL import Image, ImageDraw, ImageFont

SRC = "Proposta Infinity CMAM - Lotes 1 a 7 (Actualizada).pdf"
OUT = "Proposta Infinity CMAM - Lotes 1 a 7 (COPIA).pdf"
FONT = "/tmp/fonts/Carlito_Bold.ttf"

# posicao da palavra ORIGINAL dentro da imagem de fundo (pixeis)
ROW0, ROW1 = 164, 221          # topo e base das maiusculas
COL0, COL1 = 920, 1249         # extremos esquerdo e direito
COLOR = (4, 68, 56)            # verde escuro institucional
PAD = 14                       # folga para limpar antialiasing


def main():
    doc = pymupdf.open(SRC)
    page = doc[0]
    xref = page.get_images()[0][0]
    info = doc.extract_image(xref)
    img = Image.open(pymupdf.io.BytesIO(info["image"])).convert("RGB")

    a = np.array(img)

    # --- 1. apagar ORIGINAL ------------------------------------------------
    # o fundo nesta zona e um gradiente suave quase uniforme; reconstroi-se
    # copiando, linha a linha, a cor do fundo imediatamente a esquerda da
    # palavra (zona limpa) para toda a largura ocupada pelo texto.
    r0, r1 = ROW0 - PAD, ROW1 + PAD
    c0, c1 = COL0 - PAD, COL1 + PAD
    for r in range(r0, r1):
        fundo = a[r, c0 - 30:c0 - 5].mean(axis=0)
        a[r, c0:c1] = fundo
    img = Image.fromarray(a.astype("uint8"))

    # --- 2. desenhar CÓPIA -------------------------------------------------
    # tamanho escolhido para reproduzir exactamente a altura de maiusculas
    # original (58 px)
    size = 86
    font = ImageFont.truetype(FONT, size)
    texto = "CÓPIA"

    d = ImageDraw.Draw(img)
    bb = d.textbbox((0, 0), texto, font=font)
    # alinhar pela direita (mesma margem do ORIGINAL) e pela linha de base
    # das maiusculas
    obb = font.getbbox("O")
    x = COL1 - (bb[2] - bb[0]) - bb[0]
    y = ROW0 - obb[1]
    d.text((x, y), texto, font=font, fill=COLOR)

    # --- 3. reinserir a imagem na pagina -----------------------------------
    buf = pymupdf.io.BytesIO()
    img.save(buf, format="PNG")

    page.replace_image(xref, stream=buf.getvalue())

    doc.save(OUT, garbage=4, deflate=True)
    print("Guardado:", OUT, "| paginas:", doc.page_count)


if __name__ == "__main__":
    main()
