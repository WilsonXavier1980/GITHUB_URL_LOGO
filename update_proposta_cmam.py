#!/usr/bin/env python3
"""
Actualiza o Modelo de Proposta com as informacoes do Caderno de Encargos do
Concurso Publico N 58A001241/CP/03/OE (CMAM, IP).

Preserva integralmente o "chrome" de cada pagina (logotipo, cabecalho verde,
rodape institucional, faixas e margens) e reescreve apenas a area de conteudo,
usando as mesmas fontes embebidas no documento (Carlito Regular / Carlito Bold).
"""
import openpyxl
import pymupdf

SRC = "Proposta Infinity - Lotes 1 a 7.pdf"
OUT = "Proposta Infinity CMAM - Lotes 1 a 7 (Actualizada).pdf"

FONT_REG = "/tmp/fonts/Carlito_Regular.ttf"
FONT_BOLD = "/tmp/fonts/Carlito_Bold.ttf"

# area util de conteudo (entre o cabecalho verde e o rodape)
BODY = pymupdf.Rect(71, 118, 541, 690)
TITLE_Y = 95.1      # linha de base dos titulos de seccao
GREEN = "#1b5e20"
GREEN2 = "#2e7d32"

CAMBIO = 63.89

# ---------------------------------------------------------------- dados
ENTIDADE = "Central de Medicamentos e Artigos Médicos, IP (CMAM, IP)"
CONCURSO = "Concurso Público Nº 58A001241/CP/03/OE"
CONCURSO_FULL = ("Concurso Público Nº 58A001241/CP/03/OE — Material "
                 "Médico-Cirúrgicos de Grande Rotação/026")
OBJECTO = ("Contratação de fornecimento de material médico cirúrgico de grande "
           "rotação (Equipamentos de Proteção Individual) para o SNS")
DATA = "22/09/2026"
DATA_EXT = "22 de Setembro de 2026"

LOTE_NOMES = {
    1: "Material de Proteção Individual — Máscaras",
    2: "Material de Proteção Individual — Bata descartável",
    3: "Material de Proteção Individual — Luvas Cirúrgicas",
    4: "Material de Proteção Individual — Luva Cirúrgica de cano alto",
    5: "Material de Proteção Individual — Luvas de Observação",
    6: "Instrumentos — Estetoscópio, Esfigmomanómetro e Termómetro",
    7: "Material — Bracelete, Pêra de borracha e lâmina",
}
# Caderno de Encargos, ponto 8.4 b) e ponto 24.1
EXPERIENCIA = {1: "68.880.000,00", 2: "303.600.000,00", 3: "432.000.000,00",
               4: "37.440.000,00", 5: "108.000.000,00", 6: "56.640.000,00",
               7: "74.520.000,00"}
GARANTIA = {1: "861.000,00", 2: "3.795.000,00", 3: "5.400.000,00",
            4: "468.000,00", 5: "1.350.000,00", 6: "708.000,00",
            7: "931.500,00"}
GARANTIA_EXT = {
    1: "oitocentos e sessenta e um mil Meticais",
    2: "três milhões, setecentos e noventa e cinco mil Meticais",
    3: "cinco milhões e quatrocentos mil Meticais",
    4: "quatrocentos e sessenta e oito mil Meticais",
    5: "um milhão, trezentos e cinquenta mil Meticais",
    6: "setecentos e oito mil Meticais",
    7: "novecentos e trinta e um mil e quinhentos Meticais",
}
# Caderno de Encargos, Parte II.1 - especificacoes tecnicas resumidas por lote
ESPECS = {
    1: [("Máscara Cirúrgica com fita, Tipo 2 (pack 50)",
         "TNT, mínimo tripla camada com elemento filtrante; camada externa e filtro "
         "resistentes à penetração de fluidos; clipe nasal maleável; EFP &gt; 98%; BFE &gt; 95%."),
        ("Máscara Cirúrgica com Elástico (pack 50)",
         "TNT soldado por ultra-sons; camada interna estruturada; Tipo II segundo a "
         "EN 14683:2019; eficiência de filtração bacteriana &gt; 98%."),
        ("Máscara N95/FFP2/KN95 (pack 30)",
         "Respirador semifacial descartável, modelo concha, cinco camadas; feltro de "
         "poliéster electrostático; clip nasal metálico; embalagem individualizada."),
        ("Plainitos descartáveis / pro-pé, azul (pack 100)",
         "100% polipropileno TNT 17×41 cm; acabamento em elástico; impermeabilidade 80%; "
         "atóxico, hipoalérgico, não inflamável; não estéril; 100 peças = 50 pares."),
        ("Barretes descartáveis, azul (pack 100)",
         "Gramatura 40 g/m²; 100% polipropileno TNT; soldadura por ultrassom; 45×52 cm; "
         "não estéril; elástico revestido; tamanho único, uso único."),
        ("Avental hospitalar plástico descartável",
         "Impermeável, 80×125 cm, descartável, de uso único, conforme padrão descritivo do CBS."),
        ("Avental de borracha reutilizável",
         "110×70 cm, composição PVC/poliéster, 300 g/m², reutilizável e impermeável."), ],
    2: [("Bata descartável tamanho M / S",
         "TNT de polipropileno, punho elástico, uso único, hipoalergénica e impermeável."),
        ("Bata cirúrgica reforçada descartável, estéril (M / L / XL)",
         "Tecido SMS ou SMMS de três camadas coladas termicamente à base de polipropileno; "
         "reforço extra em toda a frente e em ambas as mangas; esterilizada por radiação "
         "gama; embalagem individual em papel grau cirúrgico com guardanapo cirúrgico "
         "incluído; VLS para IB (EN ISO 22610) de 0,98 a 95% de confiança."), ],
    3: [("Luva Cirúrgica nº 6.5 / 7.0 / 7.5 / 8.0 / 8.5 (pack 50)",
         "Estéril, antiderrapante, cor amarela, látex de borracha natural, alta "
         "sensibilidade táctil, atóxica, de uso único; formato anatómico esquerda/direita; "
         "lubrificada com pó bio-absorvível não alergénico; microtextura nas pontas dos "
         "dedos; punho com bainha enrolada; embalagem em papel grau cirúrgico com abertura "
         "asséptica, data de fabrico e validade impressas; comprimento 28 a 30 cm."), ],
    4: [("Luva Cirúrgica de cano alto obstétrica nº 7 / 7.5 / 8 (pack 50)",
         "Estéril, látex de borracha natural, de uso único; cano alto de 40 a 42 cm para "
         "procedimentos obstétricos; formato anatómico, lubrificada com pó bio-absorvível; "
         "embalagem individualizada e resistente à humidade."), ],
    5: [("Luva de observação tamanho S / M / L (pack 100)",
         "Não estéril, formato anatómico, 100% látex natural; lubrificação com pó "
         "bio-absorvível de baixo teor de proteínas; textura homogénea, alta sensibilidade "
         "ao tacto, boa elasticidade e resistência à tracção; ambidestra; comprimento "
         "mínimo de 25 cm; descartável."), ],
    6: [("Estetoscópio biauricular",
         "Auscultador de dois lados com diafragmas sintonizáveis para uso adulto e "
         "pediátrico; diafragma de peça única; olivas de selamento suaves; haste em aço "
         "inoxidável maleável."),
        ("Estetoscópio de Pinard (corneta)",
         "Cilindro oco em tronco de cone com base em trompa e disco plano de apoio ao "
         "ouvido; comprimento de 13 a 20 cm."),
        ("Esfigmomanómetro digital / Aneroide",
         "Aneroide manual graduado em intervalos de 2 mmHg, escala de 0 a 300 mmHg; "
         "braçadeira em nylon com fecho de velcro (22 a 28 cm, adulto); pêra em látex; "
         "válvula em metal cromado; manómetro em epóxi com aro de alumínio e visor em "
         "acrílico; aferição inicial do fabricante."),
        ("Termómetro clínico digital",
         "Faixa de 32,0 °C a 42,9 °C; resolução 0,1 °C; precisão ±0,2 °C; visor LCD; "
         "memória da última medição; modos axilar, oral ou rectal; pilha LR41; ponta à "
         "prova de água; desligamento automático e aviso sonoro."),
        ("Termómetro clínico com escala de Gálio",
         "Termómetro clínico sem mercúrio, escala de 35,5 °C a 42 °C, reutilizável."), ],
    7: [("Bracelete de identificação infantil / adulto, azul e rosa (pack 100)",
         "Laminado de PVC/vinil 100% polivinil, toque macio, alta flexibilidade e "
         "propriedades hipoalergénicas; infantil 15 cm × 12 mm, adulto 20 cm × 17 mm; área "
         "de impressão dimensionada para nomes completos ou códigos de barras; pontos de "
         "regulagem sequenciais; embalagens de 100 unidades."),
        ("Pêra de borracha para aspiração de recém-nascidos",
         "Capacidade de 30 a 60 ml, isenta de BPA, atóxica, de fácil higienização."),
        ("Lâmina para enxerto de pele",
         "Aço inoxidável, estéril, de uso único, em embalagem individual."), ],
}


def mt(v):
    return f"{v:,.2f}".replace(",", "\x00").replace(".", ",").replace("\x00", ".")


def read_lote(n):
    ws = openpyxl.load_workbook(f"CMAM-Lote-{n}-Atualizado-Feito.xlsx").active
    itens = []
    for r in ws.iter_rows(min_row=11, values_only=True):
        if r[0] is None or str(r[0]).startswith("="):
            continue
        try:
            q, pu = float(r[4]), float(r[6])
        except (TypeError, ValueError):
            continue
        sem = q * pu
        itens.append(dict(item=str(r[0]), desc=" ".join(str(r[1]).split()),
                          pais=str(r[2]), qtd=q, pu=pu, sem=sem,
                          iva=sem * .16, com=sem * 1.16, usd=sem * 1.16 / CAMBIO))
    return itens


LOTES = {n: read_lote(n) for n in range(1, 8)}
TOT = {n: dict(sem=sum(i["sem"] for i in v), iva=sum(i["iva"] for i in v),
               com=sum(i["com"] for i in v), usd=sum(i["usd"] for i in v))
       for n, v in LOTES.items()}
GT = {k: sum(TOT[n][k] for n in TOT) for k in ("sem", "iva", "com", "usd")}

CSS = f"""
* {{ font-family: carlito; }}
body {{ font-size: 9.2pt; color: #000; line-height: 1.30; }}
h2 {{ font-family: carlitob; font-size: 11pt; color: {GREEN}; margin: 2pt 0 5pt 0; }}
h3 {{ font-family: carlitob; font-size: 9.6pt; color: {GREEN2}; margin: 7pt 0 3pt 0; }}
p  {{ margin: 0 0 5pt 0; text-align: justify; }}
b, strong, th {{ font-family: carlitob; }}
ul {{ margin: 0 0 5pt 0; padding-left: 13pt; }}
li {{ margin-bottom: 2.2pt; text-align: justify; }}
table {{ width: 100%; border-collapse: collapse; margin: 3pt 0 5pt 0;
         table-layout: fixed; word-wrap: break-word; }}
th {{ background: #e8f5e9; color: {GREEN}; font-size: 8.2pt; padding: 3pt 4pt;
      border: 0.6pt solid #9e9e9e; text-align: left; }}
td {{ font-size: 8.2pt; padding: 3pt 4pt; border: 0.6pt solid #9e9e9e;
      vertical-align: top; }}
td.r, th.r {{ text-align: right; }}
td.c, th.c {{ text-align: center; }}
tr.tot td {{ font-family: carlitob; background: #f1f8e9; }}
.cap {{ font-family: carlitob; font-size: 8.4pt; color: {GREEN}; text-align: center;
        margin-top: 3pt; }}
.small {{ font-size: 8.4pt; }}
"""


def new_archive():
    a = pymupdf.Archive()
    a.add(FONT_REG, "carlito.ttf")
    a.add(FONT_BOLD, "carlitob.ttf")
    return a


ARCH = new_archive()


def clear_body(page, top=118.0):
    """apaga o conteudo da area util preservando cabecalho/rodape"""
    page.draw_rect(pymupdf.Rect(BODY.x0 - 12, top, BODY.x1 + 44, BODY.y1 + 2),
                   color=None, fill=(1, 1, 1), overlay=True)


def put_title(page, text):
    page.draw_rect(pymupdf.Rect(60, 88, 585, 130), color=None, fill=(1, 1, 1))
    page.insert_text((71, 108.5), text, fontfile=FONT_BOLD, fontname="CarB",
                     fontsize=14, color=(0, 0, 0))


def fill(page, html, top=118.0, title=None):
    if title is not None:
        put_title(page, title)
    clear_body(page, top)
    r = pymupdf.Rect(BODY.x0, top, BODY.x1, BODY.y1)
    spare, _ = page.insert_htmlbox(r, html, css=CSS, archive=ARCH, scale_low=0.55)
    if spare < 0:
        print("  ! conteudo ajustado nesta pagina")
    return spare


# ---------------------------------------------------------------- paginas
def page_capa(page):
    """apenas substitui os textos da capa, mantendo o fundo/imagem"""
    reps = [
        ((55, 130, 592, 166), "ASSOCIAÇÃO C-SAÚDE",
         [("CENTRAL DE MEDICAMENTOS E ARTIGOS MÉDICOS, IP", 16)]),
    ]
    # tapa os blocos de texto antigos com o mesmo tom de fundo das faixas
    bg = (0.9764706, 0.9764706, 0.96862745)
    page.draw_rect(pymupdf.Rect(55, 130, 592, 166), color=None, fill=bg)
    page.draw_rect(pymupdf.Rect(18, 386, 600, 460), color=None, fill=bg)
    page.draw_rect(pymupdf.Rect(122, 470, 490, 568), color=None, fill=bg)
    page.draw_rect(pymupdf.Rect(300, 262, 458, 310), color=None, fill=bg)
    # remove a data antiga do modelo (faixa inferior)
    page.draw_rect(pymupdf.Rect(330, 586, 476, 620), color=None,
                   fill=(0.9686, 0.9686, 0.9568))

    def ctr(txt, y, size, bold=True, color=(0, 0, 0), x0=0, x1=612):
        ff = FONT_BOLD if bold else FONT_REG
        fn = "CarB" if bold else "CarR"
        w = pymupdf.Font(fontfile=ff).text_length(txt, size)
        page.insert_text(((x0 + x1) / 2 - w / 2, y), txt, fontfile=ff,
                         fontname=fn, fontsize=size, color=color)

    # cliente
    ctr("CENTRAL DE MEDICAMENTOS E ARTIGOS MÉDICOS, IP", 150, 14.5)
    # referencia do concurso — negrito, corpo 26
    ctr("Concurso Público Nº 58A001241/CP/03/OE", 418, 26)
    ctr("Material Médico-Cirúrgicos de Grande Rotação/026", 450, 26)
    # nome do concurso / objecto
    ctr("MATERIAL MÉDICO-CIRÚRGICO", 492, 15.5, x0=122, x1=490)
    ctr("DE GRANDE ROTAÇÃO", 513, 15.5, x0=122, x1=490)
    ctr("EQUIPAMENTOS DE PROTEÇÃO INDIVIDUAL", 534, 12.5, x0=122, x1=490)
    ctr("LOTES 1 A 7", 555, 13.5, x0=122, x1=490)
    # data
    ctr(f"Data {DATA}", 290, 15, color=(0.121, 0.36, 0.18), x0=300, x1=458)


def html_indice(paginas):
    def row(n, t, p):
        return f"<tr><td class='c'>{n}</td><td>{t}</td><td class='c'>{p}</td></tr>"
    idx = "".join([
        row("1.0", "INFORMAÇÕES DO CONCORRENTE", 4),
        row("2.0", "APRESENTAÇÃO DA INFINITY HEALTH, SA", 5),
        row("3.0", "DESCRIÇÃO DOS PARCEIROS E MARCAS", 6),
        row("4.0", "OBJECTO DA PROPOSTA (Material Médico-Cirúrgico de Grande Rotação)", 8),
        row("5.0", "PROPOSTA TÉCNICA E METODOLOGIA", 10),
        row("5.1", "Garantia de Qualidade, Certificações e Conformidade", 10),
        row("5.2", "Especificações Técnicas e Tabela de Conformidade por Lote", "11-12"),
        row("5.3", "Modelo Proposto Por Infinity Health, SA (Lotes 1 a 7)",
            f"13-{paginas['lote_end']}"),
        row("6.0", "PROPOSTA FINANCEIRA", paginas["fin"]),
        row("", "Planilhas de Preços — Lotes 1 a 7",
            f"{paginas['pl0']}-{paginas['pl1']}"),
        row("7.0", "REQUISITOS PARA QUALIFICAÇÃO E ANEXOS", paginas["req"]),
        row("8.0", "DECLARAÇÕES DE GARANTIA PROVISÓRIA (Lotes 1 a 7)",
            f"{paginas['dec0']}-{paginas['dec1']}"),
    ])
    tabs = "".join([
        row("Tabela 1", "Informações do Concorrente", 4),
        row("Tabela 2", "Exemplos de fornecimento de material médico-cirúrgico", 9),
        row("Tabela 3", "Resumo do objecto da proposta", 8),
        row("Tabela 4", "Tabela de conformidade técnica por lote", "11-12"),
        row("Tabela 5", "Modelo Proposto Por Infinity Health, SA (Lotes 1 a 7)",
            f"13-{paginas['lote_end']}"),
        row("Tabela 6", "Resumo financeiro por lote", paginas["fin"]),
        row("Tabela 7", "Planilhas de preços dos bens (Lotes 1 a 7)",
            f"{paginas['pl0']}-{paginas['pl1']}"),
        row("Tabela 8", "Declarações de Garantia Provisória por lote",
            f"{paginas['dec0']}-{paginas['dec1']}"),
        row("Figura 1", "Marcas e representações", 7),
    ])
    return f"""
<h2>ÍNDICE GERAL</h2>
<table><colgroup><col style='width:9%'/><col style='width:76%'/>
<col style='width:15%'/></colgroup>
<tr><th class='c'>Nº</th><th>TÓPICO</th><th class='c'>PÁGINA</th></tr>{idx}</table>
<h2>ÍNDICE DE TABELAS E FIGURAS</h2>
<table><colgroup><col style='width:15%'/><col style='width:70%'/>
<col style='width:15%'/></colgroup>
<tr><th class='c'>REF.</th><th>DESIGNAÇÃO</th><th class='c'>PÁGINA</th></tr>{tabs}</table>
<p class='small'>Proposta referente ao {CONCURSO_FULL}, promovido pela
{ENTIDADE}. Apresentação de propostas: {DATA_EXT}, 10:00H; abertura: {DATA_EXT},
10:30H.</p>
"""


def html_objecto():
    def r(a, b):
        return (f"<tr><td style='background:#f7faf7'><b>{a}</b></td>"
                f"<td>{b}</td></tr>")
    linhas = "".join([
        r("Concurso", CONCURSO_FULL),
        r("Entidade Contratante", ENTIDADE + "<br/>Av. de Moçambique nº 847, EN1 — "
          "Zimpeto, Cidade de Maputo"),
        r("Objecto", OBJECTO),
        r("Âmbito do fornecimento", "7 lotes / 35 itens de Equipamento de Proteção "
          "Individual e material médico-cirúrgico (máscaras, batas, luvas cirúrgicas e "
          "de observação, instrumentos de diagnóstico, braceletes, pêras e lâminas)"),
        r("Modalidade", "Concurso Público — realizado e avaliado por Lote"),
        r("Regime de contratação", "Fornecimento de bens ao abrigo do Decreto nº 79/2022, "
          "de 30 de Dezembro; incoterm DAP"),
        r("Critério de avaliação", "Menor Preço Avaliado, por lote (admitida a adjudicação "
          "por item quando se mostre vantajosa)"),
        r("Prazo de entrega", "Até 120 dias, contados da data da notificação do início de "
          "execução do contrato"),
        r("Local de entrega", "Armazém Central de Medicamentos — Av. de Moçambique nº 847, "
          "EN1, Zimpeto, Cidade de Maputo"),
        r("Validade da proposta", "120 dias contados da data de abertura das propostas"),
        r("Moeda", "Preços cotados em Meticais (MZN), com IVA de 16% identificado "
          "separadamente; contravalor em USD à taxa de 63,89 MT/USD"),
        r("Garantia provisória", "Apresentada por lote, nos montantes fixados no ponto 24.1 "
          "do Caderno de Encargos, com validade de 150 dias"),
    ])
    return f"""
<p>A presente proposta tem por objecto a <b>contratação de fornecimento de material
médico cirúrgico de grande rotação (Equipamentos de Proteção Individual) para o
Serviço Nacional de Saúde</b>, promovida pela {ENTIDADE}, no âmbito do concurso
abaixo resumido. A Infinity Health, SA apresenta proposta para a totalidade dos
<b>sete lotes</b> e para 100% dos itens de cada lote, conforme exigido no ponto 13.1
do Caderno de Encargos.</p>
<table><colgroup><col style='width:26%'/><col style='width:74%'/></colgroup>{linhas}</table>
<div class='cap'>Tabela 3. Resumo do objecto da proposta</div>
<p class='small'>A Infinity Health, SA compromete-se a fornecer a totalidade dos itens
em conformidade com as especificações técnicas constantes da Parte II.1 —
Especificações Técnicas do Caderno de Encargos, com produtos de produção recente cuja
vida útil remanescente no momento da entrega não seja inferior a 85%, apresentando
amostras e catálogos, autorização do fabricante, certificação ISO 13485:2016 e
certificado de conformidade dos produtos.</p>
"""


def html_tecnica():
    return f"""
<p>A presente proposta técnica foi elaborada com base no <b>{CONCURSO_FULL}</b> e no
respectivo Caderno de Encargos (Parte II — Especificações Técnicas), destinando-se ao
fornecimento de Equipamentos de Proteção Individual e material médico-cirúrgico de
grande rotação para o Serviço Nacional de Saúde, em sete lotes.</p>
<h2>5.1. Garantia de Qualidade, Certificações e Conformidade</h2>
<p>A Infinity Health, SA fornecerá material médico-cirúrgico de fabricantes com sistemas
de gestão da qualidade certificados pela norma <b>ISO 13485:2016</b> ou equivalente,
apresentando o respectivo <b>Certificado de Conformidade dos produtos</b> e a
<b>Autorização do Fabricante</b> com indicação dos produtos abrangidos, para todos os
itens cotados, conforme o ponto 16.1.1 do Caderno de Encargos.</p>
<h3>Documentação técnica apresentada (ponto 10.1.1)</h3>
<ul>
<li>Amostras e catálogos de todos os itens, nas embalagens originais e seladas do
fabricante, contendo a descrição do lote e do item, nome do produto, nome do fabricante,
especificações técnicas, pack size e data de validade (Parte II.1.2);</li>
<li>Autorização do Fabricante com indicação expressa dos produtos abrangidos;</li>
<li>Certificado de conformidade do fabricante pela norma ISO 13485:2016 ou equivalente;</li>
<li>Declaração de compromisso de submissão do Dossier de Registo na ANARME, IP dos
Dispositivos Médicos e DIVs no prazo de 90 dias contados da notificação de adjudicação,
nos termos da Resolução nº 19/2025, de 31 de Outubro de 2025;</li>
<li>Certificado de Conformidade dos produtos e Alvará emitido pela ANARME.</li>
</ul>
<h3>Metodologia de fornecimento</h3>
<ul>
<li><b>Conformidade:</b> cada item cumpre integralmente o padrão descritivo e as
especificações técnicas da Parte II.1 do Caderno de Encargos, com indicação do nome do
fabricante e do país de origem na Planilha de Preços (ponto 10.1.2 ii).</li>
<li><b>Produção e validade:</b> produtos de produção recente, com vida útil remanescente
não inferior a <b>85%</b> à data da entrega, comprovada por documentos de embarque
(BL, carta de porte) no acto da entrega.</li>
<li><b>Importação e regime DAP:</b> a Infinity Health, SA assegura o seguro e transporte
até ao destino final, a armazenagem e movimentação interna, o desembaraço aduaneiro e os
demais custos internos e formalidades, cabendo ao comprador os Direitos Aduaneiros e o
IVA na importação (ponto 33.1).</li>
<li><b>Saída antecipada:</b> os documentos de importação serão consignados à CMAM, IP e a
Carta de Pertença e o Pedido de Isenção serão solicitados com 30 dias de antecedência
relativamente à chegada da mercadoria (pontos 33.2 e 33.3).</li>
<li><b>Entrega:</b> até <b>120 dias</b> contados da notificação do início de execução do
contrato, no Armazém Central de Medicamentos, Zimpeto, Maputo.</li>
</ul>
<h3>Enquadramento da proposta</h3>
<ul>
<li>Cotação correspondente a <b>100% dos itens de cada lote</b> (ponto 13.1), admitindo o
reajuste de quantidades em mais ou menos até 25% (ponto 13.2).</li>
<li>Preços em Meticais, com IVA de 16% identificado separadamente e contravalor em USD à
taxa do Banco de Moçambique de <b>63,89 MT/USD</b> (pontos 15.2 e 15.3).</li>
<li>Proposta em língua portuguesa, assinada e carimbada, em dois exemplares
(“ORIGINAL” e “CÓPIA”) e em formato editável em flash drive (ponto 12.1).</li>
<li>Validade da proposta: <b>120 dias</b> contados da data de abertura (ponto 20.1).</li>
<li>Garantia provisória por lote, com validade de 150 dias; garantia definitiva de 5% do
valor do contrato, válida por 24 meses (ponto 24).</li>
</ul>
"""


def html_conf(lotes):
    rows = ""
    for n in lotes:
        it = LOTES[n]
        specs = "".join(f"<li><b>{a}</b> — {b}</li>" for a, b in ESPECS[n])
        rows += f"""
<h3>Lote {n} — {LOTE_NOMES[n]}</h3>
<table><colgroup><col style='width:16%'/><col style='width:59%'/>
<col style='width:25%'/></colgroup>
<tr><th>Itens do lote</th><th>Especificações técnicas exigidas
(Parte II.1 do Caderno de Encargos)</th><th>Conformidade</th></tr>
<tr><td class='c'>Itens {it[0]['item']} a {it[-1]['item']}<br/>({len(it)} itens)</td>
<td><ul style='margin:0'>{specs}</ul></td>
<td><b>CONFORME</b> — cotados 100% dos itens do lote, com amostras, catálogos,
autorização do fabricante, certificação ISO 13485:2016 e certificado de conformidade.
Experiência comprovada exigida no ponto 8.4 b): {EXPERIENCIA[n]} MT.
Garantia provisória: {GARANTIA[n]} MT.</td></tr>
</table>"""
    return rows


def html_modelo(lotes, cont=False):
    out = ""
    for n in lotes:
        it = LOTES[n]
        linhas = "".join(
            f"<tr><td class='c'>{i['item']}</td><td>{i['desc']}</td>"
            f"<td class='c'>{i['pais']}</td>"
            f"<td class='r'>{int(i['qtd']):,}</td>".replace(",", ".") +
            f"<td class='r'>{mt(i['pu'])}</td><td class='r'>{mt(i['com'])}</td>"
            f"<td class='r'>{mt(i['usd'])}</td></tr>"
            for i in it)
        out += f"""
<h3>Lote {n} — {LOTE_NOMES[n]}</h3>
<table><colgroup><col style='width:5%'/><col style='width:32%'/>
<col style='width:9%'/><col style='width:12%'/><col style='width:11%'/>
<col style='width:16%'/><col style='width:15%'/></colgroup>
<tr><th class='c'>Item</th><th>Descrição do bem proposto</th>
<th class='c'>País de<br/>origem</th><th class='r'>Quantidade</th>
<th class='r'>Preço unit.<br/>(MT)</th><th class='r'>Total c/IVA<br/>(MT)</th>
<th class='r'>Total c/IVA<br/>(USD)</th></tr>
{linhas}
<tr class='tot'><td colspan='5'>TOTAL DO LOTE {n} (c/IVA)</td>
<td class='r'>{mt(TOT[n]['com'])}</td>
<td class='r'>{mt(TOT[n]['usd'])}</td></tr>
</table>"""
    cap = ("Tabela 5. Modelo Proposto Por Infinity Health, SA" +
           (" (continuação)" if cont else ""))
    intro = "" if cont else (
        "<p>Apresentam-se em seguida os bens propostos pela Infinity Health, SA para "
        "cada um dos sete lotes do concurso, com a descrição, o país de origem, as "
        "quantidades e os preços. As especificações técnicas completas constam da Parte "
        "II.1 do Caderno de Encargos e os preços detalhados das Planilhas de Preços "
        "apresentadas adiante.</p>")
    return intro + out + f"<div class='cap'>{cap}</div>"


def html_financeira(pl0, pl1):
    alineas = "".join(
        f"<tr><td class='c'>Lote {n}</td><td>{LOTE_NOMES[n]}</td>"
        f"<td class='r'>{mt(TOT[n]['sem'])}</td><td class='r'>{mt(TOT[n]['iva'])}</td>"
        f"<td class='r'>{mt(TOT[n]['com'])}</td>"
        f"<td class='r'>{mt(TOT[n]['usd'])}</td></tr>" for n in range(1, 8))
    return f"""
<p>a) Examinámos os documentos do <b>{CONCURSO_FULL}</b>, promovido pela {ENTIDADE}, e
apresentamos a nossa proposta sem reservas para a contratação de fornecimento de material
médico cirúrgico de grande rotação (Equipamentos de Proteção Individual) para o Serviço
Nacional de Saúde, em conformidade com o respectivo Caderno de Encargos.</p>
<p>b) Os preços da nossa proposta constam das Planilhas de Preços (páginas {pl0} a {pl1}),
sendo os valores totais por lote, com IVA de 16% incluído, os seguintes:</p>
<table><colgroup><col style='width:8%'/><col style='width:27%'/>
<col style='width:17%'/><col style='width:15%'/><col style='width:17%'/>
<col style='width:16%'/></colgroup>
<tr><th class='c'>Lote</th><th>Designação</th>
<th class='r'>Total s/IVA (MT)</th><th class='r'>IVA 16% (MT)</th>
<th class='r'>Total c/IVA (MT)</th><th class='r'>Total c/IVA (USD)</th></tr>
{alineas}
<tr class='tot'><td colspan='2'>TOTAL GERAL DOS SETE LOTES</td>
<td class='r'>{mt(GT['sem'])}</td><td class='r'>{mt(GT['iva'])}</td>
<td class='r'>{mt(GT['com'])}</td><td class='r'>{mt(GT['usd'])}</td></tr>
</table>
<div class='cap'>Tabela 6. Resumo financeiro por lote</div>
<p class='small'>Valor total da proposta: <b>{mt(GT['com'])} MT</b>, IVA de 16% incluído
(valor sem IVA: {mt(GT['sem'])} MT; IVA: {mt(GT['iva'])} MT), equivalente a
{mt(GT['usd'])} USD à taxa de câmbio de 63,89 MT/USD. Os preços correspondem a 100% dos
itens de cada lote e a adjudicação pode ser feita por lote ou por item.</p>
<p>c) Declaramos que não estamos enquadrados em qualquer situação de impedimento prevista
no Regulamento aprovado pelo Decreto nº 79/2022, de 30 de Dezembro, e que apresentamos a
garantia provisória exigida para cada lote a que concorremos, nos termos do ponto 24.1 do
Caderno de Encargos.</p>
<p>d) Esta proposta tem validade de <b>120 (cento e vinte) dias</b> contados da data de
abertura das propostas, com preços inalteráveis, e, juntamente com a vossa aceitação por
escrito constante da notificação de adjudicação, será considerada como contrato
vinculativo entre as partes, até que um contrato formal seja assinado.</p>
<p>e) Compreendemos que não estão obrigados a aceitar a proposta de menor preço ou
qualquer outra proposta que venham a receber.</p>
<p style='margin-top:10pt'>Maputo, {DATA_EXT}</p>
<p style='margin-top:22pt'>.............................................................<br/>
Assinatura do Representante Autorizado</p>
"""


def html_requisitos():
    exp = "".join(f"<li>Lote {n} — {EXPERIENCIA[n]} MT;</li>" for n in range(1, 8))
    gar = "".join(f"<li>Lote {n} — {GARANTIA[n]} MT;</li>" for n in range(1, 8))
    return f"""
<p class='small'>Documentação apresentada em conformidade com o {CONCURSO_FULL},
promovido pela {ENTIDADE}.</p>
<h3>A) Documentos de elegibilidade e qualificação (ponto 8)</h3>
<ul>
<li>Certificado de Inscrição no Cadastro Único de Empreiteiros, Fornecedores de Bens e
Prestadores de Serviços ao Estado, compatível com o objecto da contratação (ponto 8.1);</li>
<li>Alvará emitido pela Autoridade Nacional Reguladora de Medicamentos — ANARME
(ponto 8.4 a);</li>
<li>Contratos de fornecimento de medicamentos e outros produtos de saúde nos últimos três
anos, com execução mínima de 80% antes da data final de apresentação das propostas,
comprovados por guias/notas de entrega e facturas, nos valores (ponto 8.4 b):</li>
</ul>
<ul>{exp}</ul>
<ul>
<li>Declaração nominal dos reais beneficiários efectivos, após adjudicação (ponto 8.8);</li>
<li>Certidão válida de quitação fiscal, declaração do INSS e declaração de inexistência de
pedido de falência ou concordata, previamente à celebração do contrato (ponto 8.9).</li>
</ul>
<h3>B) Proposta técnica (pontos 10 e 11)</h3>
<ul>
<li>Amostras e catálogos, entregues até <b>12 de Outubro de 2026, 15:00H</b>, no Recinto do
Armazém Central de Medicamentos, Zimpeto; as amostras não são devolvíveis;</li>
<li>Amostras nas embalagens originais e seladas do fabricante, com descrição do lote e do
item, nome do produto, nome do fabricante, especificações técnicas, pack size e data de
validade (Parte II.1.2);</li>
<li>Autorização do Fabricante para todos os itens (ponto 16.1.1);</li>
<li>Certificação do fabricante pela norma ISO 13485:2016 ou equivalente;</li>
<li>Declaração de compromisso de submissão do Dossier de Registo na ANARME, IP no prazo de
90 dias (Resolução nº 19/2025, de 31 de Outubro);</li>
<li>Certificado de Conformidade dos produtos e vida útil remanescente mínima de 85%.</li>
</ul>
<h3>C) Garantias (ponto 24)</h3>
<ul>
<li>Garantia provisória por lote, com validade de 150 dias:</li>
</ul>
<ul>{gar}</ul>
<ul>
<li>Garantia definitiva de 5% do valor do contrato, válida por 24 meses;</li>
<li>Admitida a declaração de garantia reconhecida pelo Cartório Notarial, nos termos do
Diploma Ministerial nº 5/2024, de 04 de Janeiro, bem como garantia bancária, depósito ou
transferência, cheque visado, título de dívida pública ou seguro-garantia.</li>
</ul>
<h3>D) Condições gerais da participação</h3>
<ul>
<li><b>Entrega das propostas:</b> {DATA_EXT}, 10:00H, na Secretaria-Geral da CMAM, IP,
Av. de Moçambique nº 847, EN1 — Zimpeto, Cidade de Maputo; <b>abertura:</b> {DATA_EXT},
10:30H, em sessão pública (pontos 18 e 19);</li>
<li><b>Formato:</b> invólucro opaco, fechado e lacrado, com dois exemplares marcados
“ORIGINAL” e “CÓPIA”, acompanhados de um flash drive em formato editável (ponto 12.1);</li>
<li><b>Língua e moeda:</b> proposta em língua portuguesa; preços em Meticais, convertidos
para USD à taxa do Banco de Moçambique vigente dez dias antes da data de apresentação
(pontos 15 e 17);</li>
<li><b>Validade:</b> 120 dias contados da data de abertura das propostas (ponto 20.1);</li>
<li><b>Prazo de entrega:</b> até 120 dias contados da notificação do início de execução do
contrato, em regime DAP (pontos 33.1 e 33.4);</li>
<li><b>Avaliação:</b> por lote, pelo critério do Menor Preço Avaliado, podendo a
adjudicação ser feita por item quando se mostre vantajosa (pontos 26 e 27);</li>
<li><b>Margem de preferência:</b> 20% do valor do contrato sem imposto para bens
produzidos no País, mediante declaração do produtor ou selo “Orgulho Moçambicano. Made in
Mozambique” (ponto 30);</li>
<li><b>Reclamação e recurso:</b> reclamação em 5 dias úteis à Directora-Geral da CMAM, IP;
recurso hierárquico em 3 dias ao Ministro da Saúde, mediante caução de 125.000,00 MT
(pontos 35 e 36).</li>
</ul>
"""


def html_declaracao(n):
    """Modelo de Declaração de Garantia Provisória — III.2.1.1 do Caderno de
    Encargos, nos termos do nº 2 do artigo 105 do Decreto nº 79/2022."""
    return f"""
<p class='small' style='text-align:center'><b>(Apresentada juntamente com a proposta,
em alternativa à Garantia Provisória, nos termos do ponto 24.2 do Caderno de Encargos e
do nº 2 do artigo 105 do Decreto nº 79/2022, de 30 de Dezembro)</b></p>
<p><b>Nº do Concurso:</b> {CONCURSO_FULL}</p>
<p><b>Lote {n} — {LOTE_NOMES[n]}</b></p>
<p><b>Para:</b> {ENTIDADE}<br/>
Av. de Moçambique nº 847, EN1 — Zimpeto, Cidade de Maputo — Moçambique</p>
<p>Nós, <b>INFINITY HEALTH, SA</b>, com sede na Rua de França, nº 273, Bairro da Coop,
Maputo — Moçambique, NUIT 401550216, representados por
_________________________________ [indicar nome, endereço, identificação civil e NUIT],
na qualidade de _________________________ [indicar a função que exerce], signatários
desta proposta, declaramos nos termos do nº 2 do artigo 105 do Decreto nº 79/2022, de 30
de Dezembro, que:</p>
<p>Entendemos que, de acordo com as condições previstas nos Documentos de Concurso, as
propostas devem ser acompanhadas de uma Declaração de Garantia Provisória no montante de
<b>{GARANTIA[n]} MT</b> ({GARANTIA_EXT[n]}), referente ao <b>Lote {n}</b>, conforme
fixado no ponto 24.1 do Caderno de Encargos.</p>
<p>Aceitamos que seremos automaticamente sujeitos ao pagamento de multa de valor igual ao
da Garantia Provisória ou proibidos de contratar com o Estado por período de um (1) ano
e, em caso de reincidência, por período de cinco (5) anos, a partir da data de
notificação pela Unidade Funcional de Supervisão das Aquisições, de acordo com o
preceituado no artigo 284 do Regulamento, aprovado pelo Decreto nº 79/2022, de 30 de
Dezembro, se violarmos as nossas obrigações nas condições da proposta, nos seguintes
casos:</p>
<ul>
<li>(a) Retirarmos ou modificarmos a nossa proposta antes de expirar a data da validade
da mesma especificada na Proposta; ou</li>
<li>(b) Tendo sido notificados da aceitação da nossa proposta pela Entidade Contratante
dentro da validade da Proposta, (i) nos recusarmos a assinar o Contrato; (ii) nos
recusarmos a fornecer a Garantia Definitiva.</li>
</ul>
<p>Entendemos que esta Declaração de Garantia Provisória expirará nos seguintes casos:
(i) não formos notificados como Concorrente vencedor; ou (ii) ao passar trinta (30) dias
após expirar a data da validade da Proposta.</p>
<p>A presente declaração é válida por <b>150 dias</b>, nos termos do ponto 24.3 do
Caderno de Encargos.</p>
<p style='margin-top:14pt'>Maputo, {DATA_EXT}</p>
<p style='margin-top:26pt'>_______________________________________________<br/>
[Assinatura do Representante com poderes suficientes — reconhecida pelo Cartório
Notarial]</p>
<p>_______________________________________________<br/>
[Função/qualidade com que actua no acto — Proprietário, Director, Gerente, etc.]</p>
<p class='small'>Pela INFINITY HEALTH, SA — NUIT 401550216</p>
"""


# ---------------------------------------------------------------- footer
def fix_footer(page, num, total):
    """reescreve o numero de pagina preservando o estilo original"""
    page.draw_rect(pymupdf.Rect(470, 750, 570, 768), color=None, fill=(1, 1, 1))
    txt = f"Página {num} de {total}"
    w = pymupdf.get_text_length(txt, fontname="helv", fontsize=8.5)
    page.insert_text((562.8 - w, 761.0), txt, fontname="helv", fontsize=8.5,
                     color=(0.105, 0.368, 0.125))


def main():
    doc = pymupdf.open(SRC)

    # --- estrutura final -------------------------------------------------
    # 1-12 mantidas (8,10,11,12 actualizadas); 13-16 modelo dos lotes;
    # 17 financeira; 18-24 planilhas; 25 requisitos
    modelo_grupos = [[1, 2], [3, 4], [5, 6], [7]]
    n_modelo = len(modelo_grupos)
    p_lote0 = 13
    p_lote1 = p_lote0 + n_modelo - 1
    p_fin = p_lote1 + 1
    p_pl0 = p_fin + 1
    p_pl1 = p_pl0 + 6
    p_req = p_pl1 + 1
    p_dec0 = p_req + 1
    p_dec1 = p_dec0 + 6
    total = p_dec1

    # paginas 13 e 14 do original servem de molde; precisamos de n_modelo
    # paginas -> duplicar o molde limpo
    molde = 12  # indice 0-based da pagina 13
    extra = n_modelo - 2
    for _ in range(extra):
        doc.fullcopy_page(13, 14)   # duplica o molde da pag.14 (sem imagens proprias)

    # NOTA: nao se usa delete_image porque o logotipo e partilhado por todas as
    # paginas; as imagens de produto ficam cobertas por clear_body().

    # 7 paginas de Declaracao de Garantia Provisoria, clonadas do molde
    # retrato (mesmo cabecalho, rodape e fundo)
    for _ in range(7):
        doc.fullcopy_page(p_req - 1)

    print("Estrutura:", f"modelo {p_lote0}-{p_lote1}, financeira {p_fin},",
          f"planilhas {p_pl0}-{p_pl1}, requisitos {p_req},",
          f"declaracoes {p_dec0}-{p_dec1}, total {total}")

    # --- capa -------------------------------------------------------------
    page_capa(doc[0])

    # --- indice (pag.2) ---------------------------------------------------
    fill(doc[1], html_indice(dict(lote_end=p_lote1, fin=p_fin, pl0=p_pl0,
                                  pl1=p_pl1, req=p_req, dec0=p_dec0,
                                  dec1=p_dec1)), top=100)

    # --- pag.8 objecto ----------------------------------------------------
    fill(doc[7], html_objecto(), top=122, title="4. Objecto da Proposta")

    # --- pag.10 proposta tecnica -----------------------------------------
    fill(doc[9], html_tecnica(), top=124, title="5. Proposta Técnica e Metodologia")

    # --- pag.11 e 12 conformidade ----------------------------------------
    fill(doc[10], html_conf([1, 2, 3, 4]), top=124,
         title="5.2. Especificações Técnicas e Conformidade")
    fill(doc[11], html_conf([5, 6, 7]), top=124,
         title="5.2. Especificações Técnicas e Conformidade (cont.)")

    # --- paginas de modelo dos lotes -------------------------------------
    for k, grupo in enumerate(modelo_grupos):
        t = ("5.3. Modelo Proposto Por Infinity Health, SA" +
             (" (cont.)" if k else ""))
        fill(doc[molde + k], html_modelo(grupo, cont=bool(k)), top=124, title=t)

    # --- proposta financeira ---------------------------------------------
    fill(doc[p_fin - 1], html_financeira(p_pl0, p_pl1), top=122,
         title="6. Proposta Financeira")

    # --- requisitos -------------------------------------------------------
    fill(doc[p_req - 1], html_requisitos(), top=122,
         title="7. Requisitos para Qualificação e Anexos")

    # --- declaracoes de garantia provisoria -------------------------------
    for k in range(7):
        t = ("8. Declarações de Garantia Provisória" if k == 0
             else "8. Declarações de Garantia Provisória (cont.)")
        fill(doc[p_dec0 - 1 + k], html_declaracao(k + 1), top=122, title=t)

    # --- rodapes ----------------------------------------------------------
    for i in range(1, doc.page_count):
        if doc[i].rect.width < doc[i].rect.height:   # so paginas retrato
            fix_footer(doc[i], i + 1, total)

    doc.save(OUT, garbage=4, deflate=True)
    print("Guardado:", OUT, "| paginas:", doc.page_count)
    print("TOTAL GERAL:", mt(GT["com"]), "MT")


if __name__ == "__main__":
    main()
