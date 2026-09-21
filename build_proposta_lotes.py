#!/usr/bin/env python3
"""
Substitui a planilha da proposta (pagina 16 do modelo original) por 7 planilhas
(Lote 1 .. Lote 7), preservando exatamente o layout/design da pagina original:
mesmo tamanho de pagina (A4 landscape 792x612), mesmo logotipo e posicao,
mesmos cabecalhos, mesmas fontes/tamanhos/cores, mesma grelha e espessuras de
linha, mesma faixa de TOTAL GERAL e mesmo bloco de notas.
"""
import openpyxl
import pymupdf

SRC = "Modelo Proposta Infinity para Sempre (4) - Revisado.pdf"
OUT = "Proposta Infinity - Lotes 1 a 7.pdf"
SHEET_PAGE = 15  # 0-based -> pagina 16

# ---------------------------------------------------------------- geometria
# valores extraidos da pagina original (mesma moldura exterior e mesmas alturas)
FRAME = (5.1, 35.1, 787.0, 401.0)
LOGO_BBOX = (4.117, 35.70, 260.883, 141.50)
TOP = 141.5           # topo da grelha
HDR_NUM_Y = 151.5     # linha abaixo da faixa de numeros de coluna
HDR_BOT = 173.0       # fim do cabecalho de colunas
ROW_H = 26.0          # altura de cada linha de item
TOTAL_H = 16.0        # altura da faixa TOTAL GERAL
GRID_W = 0.6
TOTAL_W = 1.2
FRAME_W = 0.9

# colunas: (titulo, largura relativa, alinhamento)
COLS = [
    ("Item",                        18.5, "c"),
    ("Descrição dos Bens",         268.0, "l"),
    ("País de\nOrigem",             41.5, "c"),
    ("Tempo De\nEntrega",           38.4, "c"),
    ("QTD",                         44.0, "c"),
    ("Unidade\nFísica",             30.0, "c"),
    ("Preço Unitário",              58.0, "r"),
    ("Preço Total s/IVA",           67.1, "r"),
    ("IVA 16%",                     67.0, "r"),
    ("Preço Total c/IVA",           75.0, "r"),
    ("Preço Total c/IVA\n(USD)",    74.4, "r"),
]
X0, X1 = FRAME[0], FRAME[2]
_raw = sum(c[1] for c in COLS)
_scale = (X1 - X0) / _raw
XS = [X0]
for _, w, _a in COLS:
    XS.append(XS[-1] + w * _scale)
XS[-1] = X1

TITLE1 = "EQUIPAMENTO DE PROTEÇÃO INDIVIDUAL E MATERIAL MÉDICO-CIRÚRGICO"
TITLE2 = "CONCURSO PÚBLICO Nº 58A001241/CP/03/OE"
TITLE3 = "MATERIAL MÉDICO-CIRÚRGICO DE GRANDE ROTAÇÃO/026"
CAMBIO = 63.89

NOTAS = [
    "Notas (Concurso Público Nº 58A001241/CP/03/OE – Material Médico-Cirúrgico de Grande Rotação/026):",
    "1. Moeda da proposta: Meticais (MZN); IVA de 16% incluído e identificado separadamente.",
    "2. Câmbio de referência: 1 USD = 63,89 MT.",
    "3. Prazo de entrega: 90 dias úteis a partir da confirmação da encomenda.",
    "4. Validade da proposta: mínimo de 90 dias de calendário, com preços inalteráveis.",
    "5. Garantia: conforme caderno de encargos, incluindo assistência técnica.",
]


def mt(v):
    s = f"{v:,.2f}"
    return s.replace(",", "\x00").replace(".", ",").replace("\x00", ".")


# ---------------------------------------------------------------- dados xlsx
def read_lote(n):
    ws = openpyxl.load_workbook(f"CMAM-Lote-{n}-Atualizado-Feito.xlsx").active
    titulo = ""
    for row in ws.iter_rows(min_row=1, max_row=8, values_only=True):
        for c in row:
            if c and "Lote" in str(c):
                titulo = str(c).strip()
    itens = []
    for r in ws.iter_rows(min_row=11, values_only=True):
        if r[0] is None or str(r[0]).startswith("="):
            continue
        try:
            qtd = float(r[4])
            pu = float(r[6])
        except (TypeError, ValueError):
            continue
        desc = " ".join(str(r[1]).split())
        sem = qtd * pu
        iva = sem * 0.16
        com = sem + iva
        itens.append(dict(item=str(r[0]), desc=desc, pais=str(r[2]),
                          tempo=str(r[3]), qtd=qtd, un=str(r[5]), pu=pu,
                          sem=sem, iva=iva, com=com, usd=com / CAMBIO))
    return titulo, itens


# ---------------------------------------------------------------- desenho
def wrap(page, text, font, size, width):
    """quebra o texto em linhas que cabem em `width`"""
    out, line = [], ""
    for word in text.split():
        test = (line + " " + word).strip()
        if pymupdf.get_text_length(test, font, size) <= width or not line:
            line = test
        else:
            out.append(line)
            line = word
    if line:
        out.append(line)
    return out


def put(page, txt, x0, x1, y, size, align, font="helv", pad=4.0):
    w = pymupdf.get_text_length(txt, font, size)
    if align == "l":
        x = x0 + pad
    elif align == "r":
        x = x1 - pad - w
    else:
        x = (x0 + x1) / 2 - w / 2
    page.insert_text((x, y), txt, fontname=font, fontsize=size, color=(0, 0, 0))


def draw_page(doc, logo_png, titulo, itens, lote_n):
    page = doc.new_page(width=792, height=612)

    # logotipo (mesma posicao/escala da pagina original)
    page.insert_image(pymupdf.Rect(*LOGO_BBOX), stream=logo_png)

    # titulos do topo (mesmas fontes/tamanhos/posicoes do original)
    put(page, TITLE1, 0, 792, 17.7, 9.0, "c")
    put(page, TITLE2, 0, 787, 45.6, 8.6, "r", pad=12.4)
    put(page, TITLE3, 0, 787, 57.4, 7.0, "r", pad=14.0)
    put(page, titulo.upper(), 0, 792, 111.0, 10.2, "c")

    n = len(itens)
    y_rows0 = HDR_BOT
    y_rows1 = y_rows0 + n * ROW_H
    y_tot0 = y_rows1
    y_tot1 = y_tot0 + TOTAL_H

    # verticais da grelha
    for x in XS:
        page.draw_line((x, TOP), (x, y_tot1), color=(0, 0, 0), width=GRID_W)
    # horizontais
    for y in [TOP, HDR_NUM_Y, HDR_BOT] + [y_rows0 + i * ROW_H for i in range(1, n + 1)] + [y_tot1]:
        page.draw_line((X0, y), (X1, y), color=(0, 0, 0), width=GRID_W)

    # faixa de numeros de coluna (6pt, junto a esquerda de cada celula)
    for i in range(len(COLS)):
        xn = XS[i] + 7.6
        if XS[i + 1] - XS[i] < 24:
            xn = (XS[i] + XS[i + 1]) / 2 - 1.7
        page.insert_text((xn, 149.2), str(i + 1), fontname="helv",
                         fontsize=6.0, color=(0, 0, 0))

    # cabecalho de colunas (6.9pt, centrado; 1 ou 2 linhas)
    for i, (title, _w, _a) in enumerate(COLS):
        lines = title.split("\n")
        ys = [162.2] if len(lines) == 1 else [158.2, 167.2]
        for ln, yy in zip(lines, ys):
            put(page, ln, XS[i], XS[i + 1], yy, 6.9, "c")

    # linhas de itens
    for k, it in enumerate(itens):
        y0 = y_rows0 + k * ROW_H
        vals = [
            (it["item"], 7.0, "c"),
            (None, 7.0, "l"),
            (it["pais"], 6.4, "c"),
            (it["tempo"], 6.2, "c"),
            (f"{int(it['qtd']):,}".replace(",", "."), 7.0, "c"),
            (it["un"], 6.6, "c"),
            (mt(it["pu"]), 7.0, "r"),
            (mt(it["sem"]), 7.0, "r"),
            (mt(it["iva"]), 7.0, "r"),
            (mt(it["com"]), 7.0, "r"),
            (mt(it["usd"]), 7.0, "r"),
        ]
        # descricao com quebra de linha (mesma metrica do original: 7pt, 9pt leading)
        dlines = wrap(page, it["desc"], "helv", 7.0, XS[2] - XS[1] - 8.0)[:3]
        base = y0 + ROW_H / 2 - (len(dlines) - 1) * 4.5 - 1.0
        for j, ln in enumerate(dlines):
            put(page, ln, XS[1], XS[2], base + j * 9.0 + 3.0, 7.0, "l")

        for i, (txt, sz, al) in enumerate(vals):
            if txt is None:
                continue
            put(page, txt, XS[i], XS[i + 1], y0 + ROW_H / 2 + 2.6, sz, al)

    # faixa TOTAL GERAL (fundo branco + bordas 1.2, igual ao original)
    page.draw_rect(pymupdf.Rect(X0, y_tot0 - 0.4, X1, y_tot1 + 0.4),
                   color=(1, 1, 1), fill=(1, 1, 1), width=1.0)
    page.draw_line((X0, y_tot0), (X1, y_tot0), color=(0, 0, 0), width=TOTAL_W)
    page.draw_line((X0, y_tot1), (X1, y_tot1), color=(0, 0, 0), width=TOTAL_W)
    for x in XS:
        page.draw_line((x, y_tot0), (x, y_tot1), color=(0, 0, 0), width=TOTAL_W)

    ty = y_tot0 + 10.8
    put(page, f"TOTAL GERAL DO LOTE {lote_n}", XS[1], XS[2], ty, 7.0, "l",
        font="hebo", pad=7.3)
    tot = {k: sum(i[k] for i in itens) for k in ("sem", "iva", "com", "usd")}
    for i, key in ((7, "sem"), (8, "iva"), (9, "com"), (10, "usd")):
        put(page, mt(tot[key]), XS[i], XS[i + 1], ty, 7.0, "r", font="hebo")

    # notas (8pt, x=12, mesmo espacamento de 11pt do original)
    y = y_tot1 + 20.0
    notas_top = y
    for i, nt in enumerate(NOTAS):
        page.insert_text((12.0, y), nt, fontname="hebo" if i == 0 else "helv",
                         fontsize=8.0, color=(0, 0, 0))
        y += 11.0 if i else 13.0

    # moldura exterior (fundo ajustado para envolver sempre as notas)
    bottom = max(FRAME[3], y + 6.0)
    page.draw_rect(pymupdf.Rect(FRAME[0], FRAME[1], FRAME[2], bottom),
                   color=(0, 0, 0), width=FRAME_W)
    return tot


def main():
    src = pymupdf.open(SRC)
    logo = src.extract_image(109)["image"]

    out = pymupdf.open()
    # paginas 1..15 do original
    out.insert_pdf(src, from_page=0, to_page=SHEET_PAGE - 1)

    grand = {k: 0.0 for k in ("sem", "iva", "com", "usd")}
    for n in range(1, 8):
        titulo, itens = read_lote(n)
        t = draw_page(out, logo, titulo, itens, n)
        for k in grand:
            grand[k] += t[k]
        print(f"Lote {n}: {len(itens)} itens | total c/IVA {mt(t['com'])} MT")

    # restantes paginas do original (17 -> fim)
    out.insert_pdf(src, from_page=SHEET_PAGE + 1, to_page=src.page_count - 1)
    out.save(OUT, garbage=4, deflate=True)
    print("TOTAL GERAL 7 LOTES c/IVA:", mt(grand["com"]), "MT  /",
          mt(grand["usd"]), "USD")
    print("Paginas:", out.page_count, "->", OUT)


if __name__ == "__main__":
    main()
