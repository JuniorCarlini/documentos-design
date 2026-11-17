"""Estilo padrão Saldus para documentos Word (.docx) — texto no padrão ABNT, com as cores da marca.

Normas seguidas no texto: ABNT NBR 14724:2024 (apresentação), NBR 6024 (numeração progressiva
das seções), NBR 10520:2023 (citações) e NBR 6023 (referências).

Uso básico:
    from estilo_word import *
    doc = novo_documento("Relatório de Custos", "Tecnologia · novembro de 2026")
    titulo(doc, "Resumo")                         # vira "1 RESUMO" (numeração automática)
    texto(doc, "O custo mensal caiu **29%** após as mudanças.")
    cartoes(doc, [{"rotulo": "Custo mensal", "valor": "R$ 1.800,00", "destaque": True}, ...])
    tabela(doc, ["Item", "Valor"], [["Claude", 587.88]], formatos={"Valor": "moeda"}, total=True,
           titulo_tabela="Custos por item", fonte="Saldus (2026).")
    citacao(doc, "Texto citado com mais de três linhas...", "(Silva, 2019, p. 12)")
    referencias(doc, ["SILVA, J. **Título da obra**. São Paulo: Editora, 2019."])
    salvar(doc, "relatorio.docx")

Rodar este arquivo direto gera um documento de exemplo:  python estilo_word.py exemplo.docx
Modelo em branco:                                          python estilo_word.py --modelo modelo.docx
Requer: pip install python-docx  (matplotlib só para grafico_barras)
"""
import datetime as _dt
import io
import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Emu, Pt, RGBColor

# ---------------------------------------------------------------- identidade
COR = {
    "primaria": "12424D", "destaque": "FFB745", "branco": "FFFFFF",
    "primaria_clara": "E7EFF0", "zebra": "F5F8F8", "borda": "D5E0E2",
    "texto": "1E2B2E", "texto_suave": "5F7378", "texto_faixa": "B7CDD2",
    "positivo": "2E7D32", "negativo": "C62828", "alerta": "B26A00",
    "entrada": "FFF1D6",
}
CORES_GRAFICO = ["12424D", "FFB745", "3E8A96", "8CC2C9", "D9822B", "A7B4B7"]
FONTE = "Poppins"
FONTE_TITULO = "Sora"
STATUS = {
    "positivo": ["OK", "Concluído", "Pago", "Ativo", "Em dia", "Aprovado", "Novo"],
    "alerta": ["Pendente", "Atenção", "Em andamento", "Em análise", "Em renovação"],
    "negativo": ["Atrasado", "Erro", "Vencido", "Reprovado"],
    "texto_suave": ["Cancelado", "Inativo", "Encerrado", "Arquivado"],
}
NUMERICOS = {"moeda", "inteiro", "decimal", "percentual", "variacao"}
CENTRO = {"data", "mes", "centro"}

ASSETS = Path(__file__).resolve().parent.parent / "assets"
LOGO_CLARA = ASSETS / "logo.png"          # texto creme: sobre a faixa petróleo
LOGO_ESCURA = ASSETS / "logo-escura.png"  # texto petróleo: sobre fundo branco
SIMBOLO_SVG = ASSETS / "simbolo-s-contorno.svg"  # contorno do "S" da marca: marca d'água das páginas
COR_MARCA_DAGUA = "F8EEDC"  # âmbar bem claro (bege): aparece sem atrapalhar a leitura

# ---------------------------------------------------------------- ABNT (NBR 14724:2024)
# A4 com as margens "Normal" do Word (sem a margem de encadernação da ABNT):
# superior/inferior 2,5 cm, esquerda/direita 3 cm -> largura útil 15 cm
MARGEM = {"superior": Cm(2.5), "esquerda": Cm(3), "inferior": Cm(2.5), "direita": Cm(3)}
LARGURA_UTIL = Cm(21) - MARGEM["esquerda"] - MARGEM["direita"]
TAM_CORPO = 12          # texto, títulos de seção, referências
TAM_MENOR = 10          # citação longa, notas, legendas/fontes de ilustrações e tabelas, tabelas
ENTRELINHA = 1.5        # texto
RECUO_PARAGRAFO = Cm(1.25)
RECUO_CITACAO = Cm(4)
LINHA_EM_BRANCO = 18    # pt: uma linha de 1,5 em fonte 12 (separa títulos e citações do texto)


# ---------------------------------------------------------------- formatos BR
def _milhar(valor, casas):
    s = f"{abs(valor):,.{casas}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def fmt(valor, formato):
    """Formata um valor no padrão brasileiro. Texto passa direto."""
    if valor is None:
        return ""
    if isinstance(valor, str):
        return valor
    if formato == "moeda":
        return ("-" if valor < 0 else "") + "R$ " + _milhar(valor, 2)
    if formato == "inteiro":
        return ("-" if valor < 0 else "") + _milhar(valor, 0)
    if formato == "decimal":
        return ("-" if valor < 0 else "") + _milhar(valor, 2)
    if formato == "percentual":
        return ("-" if valor < 0 else "") + _milhar(valor * 100, 1) + "%"
    if formato == "variacao":
        return ("+" if valor > 0 else "-" if valor < 0 else "") + _milhar(valor * 100, 1) + "%"
    if isinstance(valor, (_dt.date, _dt.datetime)):
        if formato == "mes":
            meses = "jan fev mar abr mai jun jul ago set out nov dez".split()
            return f"{meses[valor.month - 1]}/{valor.year}"
        return valor.strftime("%d/%m/%Y")
    return str(valor)


def data_extenso(data):
    meses = ("janeiro fevereiro março abril maio junho julho agosto setembro outubro "
             "novembro dezembro").split()
    return f"{data.day} de {meses[data.month - 1]} de {data.year}"


# ---------------------------------------------------------------- XML helpers
def _rgb(token):
    return RGBColor.from_string(COR.get(token, token))


def _set_fonte(rpr_owner, nome):
    """Fixa a fonte em todos os scripts e remove fontes de tema (senão o Word ignora)."""
    rpr = rpr_owner.get_or_add_rPr() if hasattr(rpr_owner, "get_or_add_rPr") else rpr_owner
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for attr in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
        rfonts.attrib.pop(qn(f"w:{attr}"), None)
    for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
        rfonts.set(qn(f"w:{attr}"), nome)


def _run(par, texto, tamanho=None, cor="texto", negrito=False, italico=False,
         fonte=None, maiusculas=False):
    fonte = fonte or FONTE
    r = par.add_run(texto)
    f = r.font
    f.name = fonte
    _set_fonte(r._r, fonte)
    if tamanho:
        f.size = Pt(tamanho)
    f.color.rgb = _rgb(cor)
    f.bold = negrito
    f.italic = italico
    f.all_caps = maiusculas
    return r


def _ricos(par, texto, tamanho=None, cor="texto", fonte=None, italico=False, cor_negrito="primaria"):
    """Escreve texto aceitando **negrito** e _itálico_ no meio da frase."""
    for i, pedaco in enumerate(re.split(r"\*\*", texto)):
        if not pedaco:
            continue
        negrito = bool(i % 2)
        # _itálico_: só quando o "_" abraça texto de verdade. Campos em branco ("____", "dia ____")
        # e nomes_com_sublinhado ficam intactos.
        for j, sub in enumerate(re.split(r"(?<![\w_])_([^_\s](?:[^_]*[^_\s])?)_(?![\w_])", pedaco)):
            if sub:
                _run(par, sub, tamanho, cor_negrito if negrito else cor, negrito=negrito,
                     fonte=fonte, italico=italico or bool(j % 2))


def _borda_par(par, lado, cor, sz=6, espaco=4):
    ppr = par._p.get_or_add_pPr()
    pbdr = ppr.find(qn("w:pBdr"))
    if pbdr is None:
        pbdr = OxmlElement("w:pBdr")
        ppr.append(pbdr)  # a ordem do schema é acertada em salvar()
    b = OxmlElement(f"w:{lado}")
    b.set(qn("w:val"), "single")
    b.set(qn("w:sz"), str(sz))
    b.set(qn("w:space"), str(espaco))
    b.set(qn("w:color"), COR.get(cor, cor))
    pbdr.append(b)


def _formato_par(par, antes=0, depois=0, linha=1.0, alinhamento=WD_ALIGN_PARAGRAPH.LEFT,
                 recuo_primeira=0, recuo_esq=0):
    """Formatação direta do parágrafo; zera o recuo/justificado herdados do estilo Normal (ABNT)."""
    pf = par.paragraph_format
    pf.space_before = Pt(antes)
    pf.space_after = Pt(depois)
    pf.line_spacing = linha
    pf.first_line_indent = recuo_primeira
    pf.left_indent = recuo_esq
    par.alignment = alinhamento
    return par


def _sombrear(cel, cor):
    tcpr = cel._tc.get_or_add_tcPr()
    for old in tcpr.findall(qn("w:shd")):
        tcpr.remove(old)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), COR.get(cor, cor))
    tcpr.append(shd)


def _bordas_celula(cel, **lados):
    """lados: top/bottom/left/right = (cor, sz) ou None para sem borda."""
    tcpr = cel._tc.get_or_add_tcPr()
    tcb = tcpr.find(qn("w:tcBorders"))
    if tcb is None:
        tcb = OxmlElement("w:tcBorders")
        tcpr.append(tcb)
    for lado in ("top", "left", "bottom", "right"):
        if lado not in lados:
            continue
        for old in tcb.findall(qn(f"w:{lado}")):
            tcb.remove(old)
        el = OxmlElement(f"w:{lado}")
        if lados[lado] is None:
            el.set(qn("w:val"), "nil")
        else:
            cor, sz = lados[lado]
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), str(sz))
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), COR.get(cor, cor))
        tcb.append(el)
    for lado in ("top", "left", "bottom", "right"):  # ordem exigida pelo schema
        for el in tcb.findall(qn(f"w:{lado}")):
            tcb.remove(el)
            tcb.append(el)


def _margens_celula(cel, topo=60, base=60, esq=110, dir=110):
    tcpr = cel._tc.get_or_add_tcPr()
    for old in tcpr.findall(qn("w:tcMar")):
        tcpr.remove(old)
    mar = OxmlElement("w:tcMar")
    for lado, v in (("top", topo), ("left", esq), ("bottom", base), ("right", dir)):
        el = OxmlElement(f"w:{lado}")
        el.set(qn("w:w"), str(int(v)))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tcpr.append(mar)


def _tabela_base(doc, linhas, colunas, larguras_emu):
    """Tabela sem bordas, largura total = largura útil, colunas fixas."""
    t = doc.add_table(rows=linhas, cols=colunas)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    tblpr = t._tbl.tblPr
    el = tblpr.find(qn("w:tblStyle"))
    if el is not None:
        tblpr.remove(el)
    bordas = OxmlElement("w:tblBorders")
    for lado in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{lado}")
        el.set(qn("w:val"), "nil")
        bordas.append(el)
    tblpr.append(bordas)
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblpr.append(layout)
    tblW = tblpr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblpr.append(tblW)
    tblW.set(qn("w:type"), "dxa")
    tblW.set(qn("w:w"), str(int(sum(larguras_emu) / 635)))
    mar = OxmlElement("w:tblCellMar")
    for lado in ("left", "right"):
        el = OxmlElement(f"w:{lado}")
        el.set(qn("w:w"), "0")
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tblpr.append(mar)
    for i, w in enumerate(larguras_emu):
        t.columns[i].width = Emu(int(w))
        for cel in t.columns[i].cells:
            cel.width = Emu(int(w))
    return t


def _par_celula(cel, alinhamento=WD_ALIGN_PARAGRAPH.LEFT, antes=0, depois=0, linha=1.0, novo=False):
    """1º parágrafo da célula (ou um novo, com novo=True), sem recuo e sem justificar."""
    p = cel.add_paragraph() if novo else cel.paragraphs[0]
    return _formato_par(p, antes, depois, linha, alinhamento)


def _repetir_cabecalho(linha):
    linha._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))


def _nao_quebrar(linha):
    linha._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))


def _altura_linha(linha, cm, regra="atLeast"):
    h = OxmlElement("w:trHeight")
    h.set(qn("w:val"), str(int(Cm(cm) / 635)))
    h.set(qn("w:hRule"), regra)
    linha._tr.get_or_add_trPr().append(h)


def _campo(par, instrucao, tamanho, cor):
    """Campo do Word (PAGE, NUMPAGES...)."""
    for tipo in ("begin", "instr", "separate", "res", "end"):
        r = _run(par, "", tamanho, cor)
        if tipo == "instr":
            it = OxmlElement("w:instrText")
            it.set(qn("xml:space"), "preserve")
            it.text = f" {instrucao} "
            r._r.append(it)
        elif tipo == "res":
            r.text = "1"
        else:
            fc = OxmlElement("w:fldChar")
            fc.set(qn("w:fldCharType"), tipo)
            r._r.append(fc)


def _espaco_apos(doc, pts):
    """Parágrafo vazio de altura exata (espaço entre blocos)."""
    p = _formato_par(doc.add_paragraph())
    p.paragraph_format.line_spacing = Pt(pts)
    return p


def _contador(doc, chave):
    cont = doc.__dict__.setdefault("_saldus_contadores", {})
    cont[chave] = cont.get(chave, 0) + 1
    return cont[chave]


# ---------------------------------------------------------------- estilos
def _estilo(doc, nome, tamanho, cor, fonte=None, negrito=False, italico=False, antes=0, depois=0,
            linha=ENTRELINHA, junto=False, alinhamento=WD_ALIGN_PARAGRAPH.LEFT, recuo=0, maiusculas=False):
    fonte = fonte or FONTE
    st = doc.styles[nome]
    f = st.font
    f.name = fonte
    _set_fonte(st.element.get_or_add_rPr(), fonte)
    f.size = Pt(tamanho)
    f.color.rgb = _rgb(cor)
    f.bold = negrito
    f.italic = italico
    f.underline = False
    f.all_caps = maiusculas
    pf = st.paragraph_format
    pf.space_before = Pt(antes)
    pf.space_after = Pt(depois)
    pf.line_spacing = linha
    pf.keep_with_next = junto
    pf.first_line_indent = recuo
    pf.alignment = alinhamento
    return st


def _configurar_estilos(doc):
    # padrão do documento + idioma pt-BR para a revisão ortográfica
    rpr_def = doc.styles.element.find(qn("w:docDefaults")).find(qn("w:rPrDefault")).find(qn("w:rPr"))
    _set_fonte(rpr_def, FONTE)
    lang = rpr_def.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        rpr_def.append(lang)
    for a in ("val", "eastAsia", "bidi"):
        lang.set(qn(f"w:{a}"), "pt-BR")

    J = WD_ALIGN_PARAGRAPH.JUSTIFY
    # Texto ABNT: fonte 12, entrelinha 1,5, recuo de 1,25 cm na 1ª linha, justificado, sem espaço entre parágrafos
    _estilo(doc, "Normal", TAM_CORPO, "texto", alinhamento=J, recuo=RECUO_PARAGRAFO)
    _estilo(doc, "Title", 20, "primaria", FONTE_TITULO, negrito=True, depois=4, linha=1.0)
    _estilo(doc, "Subtitle", TAM_CORPO, "texto_suave", italico=True, depois=LINHA_EM_BRANCO)
    for nome in ("Title", "Subtitle"):  # tira a linha azul e o espaçamento entre letras do modelo padrão
        el = doc.styles[nome].element
        for tag in ("w:pBdr", "w:spacing", "w:kern", "w:color"):
            for ppr_ou_rpr in (el.find(qn("w:pPr")), el.find(qn("w:rPr"))):
                if ppr_ou_rpr is None:
                    continue
                for x in ppr_ou_rpr.findall(qn(tag)):
                    if tag == "w:spacing" and ppr_ou_rpr.tag == qn("w:pPr"):
                        continue  # espaçamento de parágrafo fica
                    if tag == "w:color":
                        continue
                    ppr_ou_rpr.remove(x)
    # Seções (NBR 6024): todas em fonte 12, destaque gradual, separadas do texto por uma linha em branco
    h1 = _estilo(doc, "Heading 1", TAM_CORPO, "primaria", FONTE_TITULO, negrito=True, maiusculas=True,
                 antes=LINHA_EM_BRANCO * 1.5, depois=LINHA_EM_BRANCO, linha=1.0, junto=True)
    _estilo(doc, "Heading 2", TAM_CORPO, "primaria", FONTE_TITULO, negrito=True,
            antes=LINHA_EM_BRANCO, depois=LINHA_EM_BRANCO, linha=1.0, junto=True)
    _estilo(doc, "Heading 3", TAM_CORPO, "texto", None, negrito=True,
            antes=LINHA_EM_BRANCO, depois=LINHA_EM_BRANCO, linha=1.0, junto=True)
    _estilo(doc, "Heading 4", TAM_CORPO, "texto", None, italico=True,
            antes=LINHA_EM_BRANCO, depois=LINHA_EM_BRANCO, linha=1.0, junto=True)
    # Legenda de ilustrações e tabelas (identificação em cima, fonte embaixo): fonte menor, simples
    _estilo(doc, "Caption", TAM_MENOR, "texto", antes=0, depois=4, linha=1.0, junto=True)
    for nome in ("List Bullet", "List Number"):
        _estilo(doc, nome, TAM_CORPO, "texto", alinhamento=J)
    for nome in ("Header", "Footer"):
        st = _estilo(doc, nome, TAM_MENOR, "texto_suave", linha=1.0)
        tabs = st.element.get_or_add_pPr().find(qn("w:tabs"))
        if tabs is not None:
            tabs.getparent().remove(tabs)


ALTURA_RODAPE_FAIXA = Cm(1.5)


def _rodape_faixa(par, com_numero):
    """Rodapé em faixa: petróleo de ponta a ponta no pé da folha (imagem presa à página, atrás do
    texto), logo clara à esquerda alinhada à margem e, opcional, o número da página em branco à direita."""
    from PIL import Image
    buf = io.BytesIO()
    Image.new("RGB", (210, 15), "#" + COR["primaria"]).save(buf, format="PNG")
    buf.seek(0)
    h = _formato_par(par)
    topo_faixa = Cm(29.7) - ALTURA_RODAPE_FAIXA
    _fundo_pagina(h, buf, Cm(21), ALTURA_RODAPE_FAIXA, x=0, y=topo_faixa, nome="Faixa do rodapé", camada=1)
    if LOGO_CLARA.exists():
        # logo também presa à folha: centro exato na altura da faixa, encostada na margem do texto
        from PIL import Image as _Img
        lw, lh = _Img.open(LOGO_CLARA).size
        alt = Cm(0.55)
        larg = Emu(int(alt * lw / lh))
        _fundo_pagina(h, str(LOGO_CLARA), larg, alt, x=MARGEM["esquerda"],
                      y=topo_faixa + (ALTURA_RODAPE_FAIXA - alt) / 2, nome="Logo do rodapé", camada=2,
                      atras=False)
    if com_numero:
        # linha de altura fixa: o centro dos algarismos cai no centro da faixa (footer_distance)
        h.paragraph_format.line_spacing = Pt(13)
        h.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        _campo(h, "PAGE", 11, "branco")
        for r in h.runs[-5:]:
            r.font.bold = True
            _set_fonte(r._r, FONTE_TITULO)


def _cabecalho_rodape(doc, secao, texto_rodape, logo_na_primeira, rodape_na_primeira, modo=True,
                      estilo_rodape="linha", paginacao=True, impressao=False):
    """Cabeçalho: logo à esquerda e número da página no canto superior direito (ABNT), filete âmbar.
    Rodapé: texto curto da empresa (sem paginação, que fica no alto)."""
    # 1ª página sempre diferente: conta, mas não mostra o número (ABNT)
    secao.different_first_page_header_footer = True
    tab = Emu(int(LARGURA_UTIL))

    def cabecalho(par, com_numero):
        h = _formato_par(par)
        if not modo:  # sem cabeçalho: parágrafo vazio e mínimo (a marca d'água, se houver, mora aqui)
            h.paragraph_format.line_spacing = Pt(1)
            return
        h.paragraph_format.tab_stops.add_tab_stop(tab, WD_TAB_ALIGNMENT.RIGHT)
        if modo == "numero":  # só o número da página, no canto superior direito
            h.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            if com_numero:
                _campo(h, "PAGE", TAM_MENOR, "texto_suave")
            return
        if LOGO_ESCURA.exists():
            h.add_run().add_picture(str(LOGO_ESCURA), height=Cm(0.7))
        else:
            _run(h, "SALDUS", TAM_MENOR, "primaria", negrito=True, fonte=FONTE_TITULO)
        if com_numero:
            _run(h, "\t", TAM_MENOR, "texto_suave")
            _campo(h, "PAGE", TAM_MENOR, "texto_suave")
        _borda_par(h, "bottom", "destaque", sz=12, espaco=6)

    cabecalho(secao.header.paragraphs[0], True)
    if logo_na_primeira:  # sem faixa nem capa: a logo aparece no alto da 1ª página, sem número
        cabecalho(secao.first_page_header.paragraphs[0], False)
    else:
        _formato_par(secao.first_page_header.paragraphs[0])

    numero_no_rodape = paginacao and modo not in (True, "numero")  # sem número no alto, desce para o rodapé
    if estilo_rodape == "faixa":
        secao.footer_distance = Cm(0.5)  # número centrado na altura da faixa
        _rodape_faixa(secao.footer.paragraphs[0], numero_no_rodape)
        if rodape_na_primeira:
            _rodape_faixa(secao.first_page_footer.paragraphs[0], False)
        return

    def rodape(par, com_numero):
        _formato_par(par)
        _borda_par(par, "top", "borda", sz=4, espaco=6)
        par.paragraph_format.tab_stops.add_tab_stop(tab, WD_TAB_ALIGNMENT.RIGHT)
        if impressao and LOGO_ESCURA.exists():  # logo pequena, escura: imprime bem e gasta pouca tinta
            r_logo = par.add_run()
            r_logo.add_picture(str(LOGO_ESCURA), height=Cm(0.45))
            pos = OxmlElement("w:position")
            pos.set(qn("w:val"), "-4")
            r_logo._r.get_or_add_rPr().append(pos)
            if texto_rodape and texto_rodape != "Saldus":
                _run(par, "   " + texto_rodape.replace("Saldus · ", ""), 8, "texto_suave")
        else:
            _run(par, texto_rodape or "", 8, "texto_suave")
        if com_numero:
            _run(par, "\t", TAM_MENOR, "texto_suave")
            if impressao:  # "Página 2 de 6": padrão de documento impresso
                _run(par, "Página ", 8, "texto_suave")
                _campo(par, "PAGE", 8, "texto_suave")
                _run(par, " de ", 8, "texto_suave")
                _campo(par, "NUMPAGES", 8, "texto_suave")
            else:
                _campo(par, "PAGE", TAM_MENOR, "texto_suave")

    if texto_rodape or numero_no_rodape:
        rodape(secao.footer.paragraphs[0], numero_no_rodape)
        if rodape_na_primeira:
            rodape(secao.first_page_footer.paragraphs[0], numero_no_rodape and impressao)


# ---------------------------------------------------------------- API pública
def novo_documento(titulo_doc, subtitulo=None, faixa=True, capa=None, rodape="Saldus", autor="Saldus",
                   fonte=None, marca=True, aviso=None, marca_dagua=False, cabecalho=False,
                   estilo_rodape="faixa", paginacao=True, impressao=False):
    """Cria o documento já no padrão: A4, margens ABNT, estilos, cabeçalho e rodapé.

    faixa=True  -> faixa petróleo com título/subtítulo/logo no topo da 1ª página (padrão).
    capa=dict   -> 1ª página como capa (documentos longos ou formais):
                   {"estilo": "bloco"|"inteira", "info": {"Preparado para": ..., "Data": ...}}
                   "bloco"   = bloco petróleo no alto da folha branca, ficha embaixo (padrão);
                   "inteira" = folha inteira em petróleo, de ponta a ponta, com texto branco.
                   {"estilo": "inteira", "foto": "foto.jpg"} = capa inteira com a foto na parte de
                   cima (de ponta a ponta) e o título embaixo.
                   "partes": ("Manual", "Exportar relatório") força a divisão fina/grossa do título.
                   {"imagem": "capa.png"} = capa pronta, feita fora (Canva, Figma...): a imagem ocupa a
                   folha inteira (corta as sobras se não for A4). Com "texto": True, o título e a
                   ficha são escritos por cima da imagem, em branco.
    rodape      -> texto à esquerda do rodapé (ex.: "Saldus · Confidencial"); None = sem rodapé.
    fonte       -> troca Sora/Poppins por outra família (ex.: "Arial" ou "Times New Roman"),
                   para instituições que exigem essas fontes.
    marca=False -> tudo em preto e cinza (para quem exige texto preto); sem faixa petróleo.
    cabecalho   -> False = SEM cabeçalho (padrão Saldus): logo e número da página ficam no rodapé;
                   True = logo + número da página + filete âmbar no alto (só se pedirem);
                   "numero" = só o número da página no canto superior direito (só se pedirem).
    estilo_rodape -> "faixa" (padrão: faixa petróleo de ponta a ponta, logo clara à esquerda, número
                   da página em branco à direita) ou "linha" (texto pequeno com linha fina em cima).
    paginacao   -> False tira o número da página do rodapé (quando ele não está no cabeçalho).
    impressao   -> True para documento que vai ser IMPRESSO (padrão de política, contrato, termo, ata):
                   nada de ponta a ponta nem fundos grandes — título sem bloco colorido (petróleo sobre
                   branco, logo escura, filete âmbar), rodapé com linha fina, logo pequena e
                   "Página X de Y" em todas as páginas, sem marca d'água e sem capa inteira.
    marca_dagua -> True põe o contorno do "S" da marca, bem claro, no fundo de todas as páginas
                   (menos na capa).
    aviso       -> página de aviso logo após a capa: "confidencial", "uso_interno", "minuta" ou um
                   texto próprio (veja pagina_aviso). Sem capa, vira uma caixa de aviso no topo.
    A 1ª página conta, mas não mostra o número (como na ABNT).
    """
    global FONTE, FONTE_TITULO
    if impressao:
        estilo_rodape, marca_dagua = "linha", False
        if capa and (capa.get("estilo") == "inteira" or capa.get("foto") or capa.get("imagem")):
            capa = {"info": capa.get("info", {})}  # capa inteira não imprime bem: vira capa em bloco
    if fonte:
        FONTE = FONTE_TITULO = fonte
    if not marca:
        COR.update(primaria="000000", destaque="000000", texto="000000", texto_suave="404040",
                   primaria_clara="EDEDED", zebra="FFFFFF", borda="BFBFBF", entrada="F5F5F5",
                   alerta="000000")
        faixa = False
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.top_margin, sec.left_margin = MARGEM["superior"], MARGEM["esquerda"]
    sec.bottom_margin, sec.right_margin = MARGEM["inferior"], MARGEM["direita"]
    sec.header_distance, sec.footer_distance = Cm(1.1), Cm(1.2)
    _configurar_estilos(doc)
    titulo_marcado = titulo_doc  # a marcação **grosso** / ==âmbar== só vale no título da capa inteira
    titulo_doc = re.sub(r"\*\*|==", "", titulo_doc)
    doc.core_properties.title = titulo_doc
    doc.core_properties.author = autor
    doc.core_properties.language = "pt-BR"
    _cabecalho_rodape(doc, sec, rodape, logo_na_primeira=not (capa or faixa), rodape_na_primeira=not capa,
                      modo=cabecalho, estilo_rodape=estilo_rodape, paginacao=paginacao, impressao=impressao)
    if marca_dagua and marca:
        _marca_dagua(sec, na_primeira=not capa)
    if capa:
        if capa.get("imagem"):
            _capa_inteira(doc, titulo_doc, subtitulo, capa.get("info", {}), imagem=capa["imagem"],
                          com_texto=capa.get("texto", False))
        elif capa.get("estilo") == "inteira" or capa.get("foto"):
            _capa_inteira(doc, titulo_marcado, subtitulo, capa.get("info", {}), foto=capa.get("foto"),
                          partes=capa.get("partes"))
        else:
            _capa(doc, titulo_doc, subtitulo, capa.get("info", {}))
    elif faixa:
        faixa_titulo(doc, titulo_doc, subtitulo, leve=impressao)
    else:
        doc.add_paragraph(titulo_doc, style="Title")
        if subtitulo:
            doc.add_paragraph(subtitulo, style="Subtitle")
    if aviso:
        if capa:
            pagina_aviso(doc, aviso)
        else:
            titulo_av, texto_av = _texto_aviso(aviso)
            destaque(doc, texto_av, titulo=titulo_av, tipo="atencao")  # sem capa: só a caixa curta
    return doc


AVISOS = {
    "confidencial": {
        "titulo": "Atenção", "subtitulo": "Documento confidencial",
        "blocos": [
            ("p", "Este documento foi desenvolvido exclusivamente para uso interno da Saldus e contém "
                  "informações estratégicas, operacionais e metodológicas que fazem parte do conhecimento "
                  "da empresa."),
            ("p", "O conteúdo aqui apresentado é estritamente confidencial e protegido pelas normas internas "
                  "da empresa e pela legislação sobre proteção de informações, direitos autorais e sigilo "
                  "empresarial."),
            ("h", "Uso restrito"),
            ("p", "Este documento é disponibilizado exclusivamente ao destinatário identificado, apenas para "
                  "consulta, estudo e execução das atividades autorizadas pela Saldus."),
            ("h", "É expressamente proibido:"),
            ("li", ["Compartilhar, repassar, divulgar ou disponibilizar este material, no todo ou em parte, "
                    "a terceiros;",
                    "Reproduzir, copiar, fotografar, gravar, digitalizar ou distribuir o conteúdo por "
                    "qualquer meio, físico ou digital;",
                    "Usar as informações aqui contidas para fins externos, concorrenciais ou não autorizados."]),
            ("h", "Responsabilidade"),
            ("p", "O uso indevido, o compartilhamento não autorizado ou qualquer violação deste aviso pode "
                  "resultar em medidas administrativas, civis e legais, conforme o contrato, as políticas "
                  "internas da Saldus e a legislação aplicável."),
            ("p", "Ao continuar a leitura, você declara estar ciente destas condições e se compromete a "
                  "cumprir as diretrizes de confidencialidade e segurança da informação da Saldus."),
        ]},
    "uso_interno": {
        "titulo": "Atenção", "subtitulo": "Uso interno",
        "blocos": [
            ("p", "Este documento é de uso interno da Saldus e se destina apenas à equipe e às atividades "
                  "da empresa."),
            ("h", "Não é permitido:"),
            ("li", ["Compartilhar este material com clientes, fornecedores ou pessoas de fora da empresa sem "
                    "autorização do responsável;",
                    "Publicar o conteúdo, no todo ou em parte, em sites, redes sociais ou grupos externos."]),
            ("p", "Em caso de dúvida sobre o que pode ser compartilhado, consulte o responsável pelo documento."),
        ]},
    "minuta": {
        "titulo": "Atenção", "subtitulo": "Minuta – versão para revisão",
        "blocos": [
            ("p", "Este documento é uma minuta e ainda pode sofrer alterações. Ele não tem valor final até a "
                  "aprovação e a assinatura das partes."),
            ("h", "Como revisar"),
            ("li", ["Anote as sugestões diretamente no arquivo, com comentários ou controle de alterações;",
                    "Envie a revisão ao responsável pelo documento dentro do prazo combinado."]),
        ]},
}


def _aviso_dados(aviso):
    """Aceita a chave de AVISOS, um texto simples, (titulo, texto) ou um dict no formato de AVISOS."""
    if isinstance(aviso, dict):
        return aviso
    if isinstance(aviso, (tuple, list)):
        return {"titulo": "Atenção", "subtitulo": aviso[0], "blocos": [("p", aviso[1])]}
    if aviso in AVISOS:
        return AVISOS[aviso]
    return {"titulo": "Atenção", "subtitulo": "Aviso", "blocos": [("p", aviso)]}


def _texto_aviso(aviso):
    """Versão curta (título + 1º parágrafo) para a caixa de aviso sem capa."""
    d = _aviso_dados(aviso)
    primeiro = next((b[1] for b in d["blocos"] if b[0] == "p"), "")
    return d["subtitulo"], primeiro


def _icone_alerta():
    """Triângulo de alerta petróleo com o ponto de exclamação vazado em âmbar (PNG)."""
    from PIL import Image, ImageDraw
    w, h = 220, 196
    img = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    # triângulo com cantos arredondados = triângulo menor + contorno grosso com juntas curvas
    r, cor = 16, "#" + COR["primaria"]
    v = [(w / 2, r + 4), (w - r - 4, h - r - 4), (r + 4, h - r - 4)]
    d.polygon(v, fill=cor)
    d.line(v + [v[0], v[1]], fill=cor, width=2 * r, joint="curve")
    for x, y in v:
        d.ellipse([x - r, y - r, x + r, y + r], fill=cor)
    ambar = "#" + COR["destaque"]
    d.rounded_rectangle([w / 2 - 11, 62, w / 2 + 11, 136], radius=10, fill=ambar)
    d.ellipse([w / 2 - 13, 148, w / 2 + 13, 174], fill=ambar)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf


def faixa_aviso(doc, titulo_av="Atenção", subtitulo="Documento confidencial"):
    """Faixa âmbar de alerta: triângulos petróleo nas pontas, ATENÇÃO grande e o tipo de aviso embaixo."""
    lado = Cm(2.2)
    t = _tabela_base(doc, 1, 3, [lado, LARGURA_UTIL - 2 * lado, lado])
    _nao_quebrar(t.rows[0])
    for j, c in enumerate(t.rows[0].cells):
        _sombrear(c, "destaque")
        c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        _margens_celula(c, 110, 110, 120, 120)
        if j != 1:
            _par_celula(c, WD_ALIGN_PARAGRAPH.CENTER).add_run().add_picture(_icone_alerta(), width=Cm(1.05))
    meio = t.rows[0].cells[1]
    _run(_par_celula(meio, WD_ALIGN_PARAGRAPH.CENTER), titulo_av.upper(), 18, "primaria", negrito=True,
         fonte=FONTE_TITULO)
    if subtitulo:
        _run(_par_celula(meio, WD_ALIGN_PARAGRAPH.CENTER, novo=True), subtitulo.upper(), 8.5, "primaria",
             negrito=True, fonte=FONTE_TITULO)
    _espaco_apos(doc, LINHA_EM_BRANCO)
    return t


def pagina_aviso(doc, aviso="confidencial"):
    """Página de aviso logo depois da capa: faixa âmbar ATENÇÃO + texto completo no padrão ABNT, com
    subtítulos e marcadores. aviso: "confidencial", "uso_interno", "minuta", um texto, (subtitulo, texto)
    ou um dict {"titulo", "subtitulo", "blocos": [("p", texto), ("h", subtítulo), ("li", [itens])]}."""
    d = _aviso_dados(aviso)
    faixa_aviso(doc, d.get("titulo", "Atenção"), d.get("subtitulo"))
    for tipo, conteudo in d["blocos"]:
        if tipo == "h":
            p = _formato_par(doc.add_paragraph(), 6, 6, ENTRELINHA)
            p.paragraph_format.keep_with_next = True
            _run(p, conteudo, TAM_CORPO, "primaria", negrito=True)
        elif tipo == "li":
            lista(doc, conteudo)
        else:
            texto(doc, conteudo)
    quebra_pagina(doc)


def faixa_titulo(doc, titulo_doc, subtitulo=None, leve=False):
    """Faixa petróleo no topo (mesma da planilha): título branco Sora, subtítulo, logo clara à
    direita, filete âmbar embaixo. Ocupa exatamente a largura do texto.
    leve=True (impressão): sem fundo — título petróleo, subtítulo cinza, logo escura, filete âmbar fino."""
    larg_logo = Cm(4.2)
    if leve:
        t = _tabela_base(doc, 1, 2, [LARGURA_UTIL - larg_logo, larg_logo])
        esq, dir_ = t.rows[0].cells
        for c in (esq, dir_):
            _bordas_celula(c, bottom=("destaque", 18))
            c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.BOTTOM
        _margens_celula(esq, 0, 140, 0, 120)
        _margens_celula(dir_, 0, 160, 120, 0)
        p = _par_celula(esq)
        _run(p, titulo_doc, 18, "primaria", negrito=True, fonte=FONTE_TITULO)
        if subtitulo:
            _run(_par_celula(esq, antes=3, novo=True), subtitulo, TAM_MENOR, "texto_suave", italico=True)
        pl = _par_celula(dir_, WD_ALIGN_PARAGRAPH.RIGHT)
        if LOGO_ESCURA.exists():
            pl.add_run().add_picture(str(LOGO_ESCURA), width=Cm(3.0))
        _espaco_apos(doc, 12)
        return t
    t = _tabela_base(doc, 1, 2, [LARGURA_UTIL - larg_logo, larg_logo])
    esq, dir_ = t.rows[0].cells
    for c in (esq, dir_):
        _sombrear(c, "primaria")
        _bordas_celula(c, bottom=("destaque", 24))
        c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    _margens_celula(esq, 260, 260, 340, 120)
    _margens_celula(dir_, 260, 260, 120, 340)
    p = _par_celula(esq)
    _run(p, titulo_doc, 18, "branco", negrito=True, fonte=FONTE_TITULO)
    if subtitulo:
        p2 = _par_celula(esq, antes=4, novo=True)
        _run(p2, subtitulo, TAM_MENOR, "texto_faixa", italico=True)
    pl = _par_celula(dir_, WD_ALIGN_PARAGRAPH.RIGHT)
    if LOGO_CLARA.exists():
        pl.add_run().add_picture(str(LOGO_CLARA), width=Cm(3.2))
    _espaco_apos(doc, 12)
    return t


def _capa(doc, titulo_doc, subtitulo, info):
    # bloco petróleo em 2 linhas: logo no alto, título embaixo (altura mínima, nunca corta texto)
    t = _tabela_base(doc, 2, 1, [LARGURA_UTIL])
    topo, base = t.rows[0].cells[0], t.rows[1].cells[0]
    for c in (topo, base):
        _sombrear(c, "primaria")
    _bordas_celula(base, bottom=("destaque", 36))
    _margens_celula(topo, 600, 0, 600, 600)
    _margens_celula(base, 0, 700, 600, 600)
    _altura_linha(t.rows[0], 8.0)
    _altura_linha(t.rows[1], 5.5)
    topo.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    base.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.BOTTOM
    p = _par_celula(topo)
    if LOGO_CLARA.exists():
        p.add_run().add_picture(str(LOGO_CLARA), width=Cm(4.5))
    pt = _par_celula(base, depois=6)
    _run(pt, titulo_doc, 28, "branco", negrito=True, fonte=FONTE_TITULO)
    if subtitulo:
        ps = _par_celula(base, linha=1.15, novo=True)
        _run(ps, subtitulo, 13, "texto_faixa", italico=True)
    _espaco_apos(doc, 30)
    if info:
        ficha(doc, info, colunas=min(len(info), 3))
    quebra_pagina(doc)


def _fundo_pagina(par, imagem_png, largura=Cm(21), altura=Cm(29.7), x=0, y=0, nome="Fundo da capa", camada=0, atras=True):
    """Imagem atrás do texto, posicionada em relação à folha (x, y a partir do canto superior esquerdo).
    Padrão: folha A4 inteira (fundo da capa)."""
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    inline = par.add_run().add_picture(imagem_png, width=largura, height=altura)._inline
    cx, cy = inline.extent.cx, inline.extent.cy
    doc_pr = inline.docPr
    anchor = parse_xml(
        f'<wp:anchor {nsdecls("wp", "a", "pic", "r")} distT="0" distB="0" distL="0" distR="0" '
        f'simplePos="0" relativeHeight="{camada}" behindDoc="{1 if atras else 0}" locked="1" layoutInCell="1" allowOverlap="1">'
        f'<wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="page"><wp:posOffset>{int(x)}</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="page"><wp:posOffset>{int(y)}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
        f'<wp:docPr id="{doc_pr.id}" name="{nome}"/><wp:cNvGraphicFramePr/></wp:anchor>')
    anchor.append(inline.graphic)
    inline.getparent().replace(inline, anchor)


def _imagem_a4(caminho):
    """Abre a imagem da capa e corta as sobras (centralizado) para a proporção A4, sem distorcer."""
    from PIL import Image
    img = Image.open(caminho).convert("RGB")
    w, h = img.size
    alvo = 21 / 29.7
    if abs(w / h - alvo) > 0.01:
        if w / h > alvo:  # larga demais: corta as laterais
            nw = int(h * alvo)
            img = img.crop(((w - nw) // 2, 0, (w - nw) // 2 + nw, h))
        else:             # alta demais: corta em cima e embaixo
            nh = int(w / alvo)
            img = img.crop((0, (h - nh) // 2, w, (h - nh) // 2 + nh))
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=92)
    buf.seek(0)
    return buf


def _svg_contorno_png(svg, cor, escala=1.0, traco=None):
    """Desenha o contorno de um SVG simples (paths com M/L/C/Z absolutos, como o do símbolo "S") em PNG
    transparente, só com a Pillow. O traço fica do lado de dentro da forma, como no arquivo original."""
    from PIL import Image, ImageChops, ImageDraw
    xml = Path(svg).read_text(encoding="utf-8")
    w = float(re.search(r'viewBox="[\d.\s-]*?\s([\d.]+)\s([\d.]+)"', xml).group(1))
    h = float(re.search(r'viewBox="[\d.\s-]*?\s([\d.]+)\s([\d.]+)"', xml).group(2))
    sw = traco or float((re.search(r'stroke-width="([\d.]+)"', xml) or [None, 20])[1])
    W, H = int(w * escala), int(h * escala)
    preench = Image.new("L", (W, H), 0)
    contorno = Image.new("L", (W, H), 0)
    dp, dc = ImageDraw.Draw(preench), ImageDraw.Draw(contorno)
    for d in re.findall(r'<path[^>]*\sd="([^"]+)"', xml):
        tokens = re.findall(r"[MLCZmlcz]|-?\d*\.?\d+(?:e-?\d+)?", d)
        pts, figuras, i, cmd, atual = [], [], 0, None, (0.0, 0.0)
        while i < len(tokens):
            t = tokens[i]
            if t in "MLCZ":
                cmd = t
                i += 1
                if t == "Z":
                    if pts:
                        figuras.append(pts)
                    pts = []
                continue
            if cmd == "M":
                if pts:
                    figuras.append(pts)
                atual = (float(tokens[i]), float(tokens[i + 1]))
                pts, i, cmd = [atual], i + 2, "L"
            elif cmd == "L":
                atual = (float(tokens[i]), float(tokens[i + 1]))
                pts.append(atual)
                i += 2
            elif cmd == "C":
                c = [float(v) for v in tokens[i:i + 6]]
                p0, p1, p2, p3 = atual, (c[0], c[1]), (c[2], c[3]), (c[4], c[5])
                for k in range(1, 13):  # bezier cúbica em 12 segmentos
                    u = k / 12
                    a, b, cc, dd = (1 - u) ** 3, 3 * u * (1 - u) ** 2, 3 * u * u * (1 - u), u ** 3
                    pts.append((a * p0[0] + b * p1[0] + cc * p2[0] + dd * p3[0],
                                a * p0[1] + b * p1[1] + cc * p2[1] + dd * p3[1]))
                atual, i = p3, i + 6
            else:
                i += 1
        if pts:
            figuras.append(pts)
        for f in figuras:
            f = [(x * escala, y * escala) for x, y in f]
            dp.polygon(f, fill=255)
            dc.line(f + [f[0], f[1]], fill=255, width=max(2, int(sw * escala)), joint="curve")
    alfa = ImageChops.multiply(preench, contorno)  # só a metade de dentro do traço (máscara do SVG)
    # sem transparência: cor sobre branco (alguns programas escurecem PNG transparente atrás do texto)
    img = Image.composite(Image.new("RGB", (W, H), "#" + cor), Image.new("RGB", (W, H), "#FFFFFF"), alfa)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf, w / h


def _marca_dagua(secao, na_primeira=True):
    """Contorno grande do "S" atrás do texto, em todas as páginas (fica no cabeçalho, então se repete
    sozinho). Vaza um pouco das bordas da folha, como nos materiais da marca."""
    if not SIMBOLO_SVG.exists():
        return
    largura = Cm(23)
    for cab, usar in ((secao.header, True), (secao.first_page_header, na_primeira)):
        if not usar:
            continue
        png, prop = _svg_contorno_png(SIMBOLO_SVG, COR_MARCA_DAGUA, escala=0.5, traco=18)
        altura = Emu(int(largura / prop))
        _fundo_pagina(cab.paragraphs[0], png, largura, altura, x=(Cm(21) - largura) / 2,
                      y=(Cm(29.7) - altura) / 2, nome="Marca d'água")


FONTE_FINA = "Poppins Light"  # parte fina do título da capa (instale Poppins Light; sem ela, cai na Poppins)


def _partes_titulo(titulo_doc, partes=None):
    """Divide o título da capa em (parte fina, parte grossa). Ex.: "Manual – Exportar relatório" ->
    ("Manual", "Exportar relatório"). Sem separador ( – , - , : ), fica tudo na parte grossa."""
    if partes:
        return partes
    m = re.match(r"^(.{2,40}?)\s*(?:–|—|-|:)\s+(.+)$", titulo_doc)
    return (m.group(1), m.group(2)) if m else (None, titulo_doc)


def _titulo_misto(titulo_doc, partes=None):
    """Monta o título da capa como UMA frase com dois pesos sobre o fundo escuro:
    texto normal = fino (Poppins Light) branco · **texto** (ou ==texto==) = grosso (Sora) âmbar — o destaque.
    Sem marcação: a parte antes de – - : fica fina e o resto em destaque; sem separador, a 1ª palavra
    fica fina e o resto em destaque."""
    if "**" in titulo_doc or "==" in titulo_doc:
        texto = titulo_doc.replace("==", "**")
    else:
        fina, grossa = _partes_titulo(titulo_doc, partes)
        if not fina and " " in grossa:
            fina, grossa = grossa.split(" ", 1)
        texto = f"{fina} **{grossa}**" if fina else f"**{grossa}**"
    pedacos = []
    for parte in re.split(r"(\*\*.+?\*\*)", texto):
        if parte.startswith("**"):
            pedacos.append((parte[2:-2], "destaque", True))
        elif parte:
            pedacos.append((parte, "branco", False))
    return pedacos, "".join(t for t, _, _ in pedacos)


def _capa_inteira(doc, titulo_doc, subtitulo, info, imagem=None, com_texto=True, foto=None, partes=None):
    """Capa de página inteira, folha toda em petróleo com faixa âmbar no pé.
    foto: imagem principal na parte de cima da capa, de ponta a ponta, cortada sem distorcer. A altura
          se ajusta ao conteúdo: a área petróleo fica só do tamanho da logo + título + ficha, e a foto
          ocupa o resto (entre 11 e 20 cm). Um filete âmbar separa a foto do texto.
    Título numa frase só, misturando fino/grosso e branco/âmbar (veja _titulo_misto), mais embaixo.
    imagem: capa pronta feita fora — ocupa a folha toda; com_texto=False não escreve nada por cima."""
    import math
    from PIL import Image, ImageDraw
    # Mede o conteúdo primeiro (alturas fixas em pt), para a foto ocupar todo o resto da folha
    util = (Cm(29.7) - MARGEM["superior"] - MARGEM["inferior"]) / 12700
    larg = LARGURA_UTIL / 12700
    pedacos, limpo = _titulo_misto(titulo_doc, partes)
    l_grossa = max(1, math.ceil(len(limpo) * 34 * 0.58 / larg))
    l_sub = math.ceil(len(subtitulo) * 13 * 0.5 / larg) if subtitulo else 0
    h_logo, h_grossa, h_sub, h_info = 34, 40, 19, (44 if info else 0)
    bloco_titulo = l_grossa * h_grossa + (8 + l_sub * h_sub if subtitulo else 0)
    entre = 40 if info else 0                       # título -> ficha
    grupo = bloco_titulo + entre + h_info
    logo_titulo = 26                                # logo -> título (com foto)
    pad_topo, folga = 30, 34                        # abaixo da foto / acima da margem de baixo
    alt_foto_cm = 0
    if foto:
        conteudo = pad_topo + h_logo + logo_titulo + grupo + folga
        alt_pt = 842 - (MARGEM["inferior"] / 12700) - conteudo
        alt_foto_cm = min(max(alt_pt / 28.35, 11.0), 20.0)  # entre 11 e 20 cm
    if imagem:
        buf = _imagem_a4(imagem)
    else:
        esc = 6 if foto else 4  # px por mm (mais resolução quando há foto)
        w, h = 210 * esc, 297 * esc
        img = Image.new("RGB", (w, h), "#" + COR["primaria"])
        d = ImageDraw.Draw(img)
        d.rectangle([0, h - 8 * esc, w, h], fill="#" + COR["destaque"])  # 8 mm âmbar no pé
        if foto:
            fh = int(alt_foto_cm * 10 * esc)
            f = Image.open(foto).convert("RGB")
            alvo = w / fh
            fw, fhh = f.size
            if fw / fhh > alvo:  # corta as laterais
                nw = int(fhh * alvo)
                f = f.crop(((fw - nw) // 2, 0, (fw - nw) // 2 + nw, fhh))
            else:                # corta em cima e embaixo
                nh = int(fw / alvo)
                f = f.crop((0, (fhh - nh) // 2, fw, (fhh - nh) // 2 + nh))
            img.paste(f.resize((w, fh), Image.LANCZOS), (0, 0))
            d.rectangle([0, fh, w, fh + 2 * esc], fill="#" + COR["destaque"])  # filete âmbar sob a foto
        buf = io.BytesIO()
        img.save(buf, format="JPEG" if foto else "PNG", quality=90)
        buf.seek(0)

    p = _formato_par(doc.add_paragraph())
    p.paragraph_format.line_spacing = Pt(1)
    _fundo_pagina(p, buf)
    if not com_texto:
        p.add_run().add_break(WD_BREAK.PAGE)
        return

    if foto:  # logo logo abaixo da foto, título logo abaixo da logo
        antes_logo = (alt_foto_cm * 28.35 + pad_topo) - MARGEM["superior"] / 12700
        antes_titulo = logo_titulo
    else:     # sem foto: logo no alto e o bloco título + ficha centralizado no espaço que sobra
        antes_logo = 0
        antes_titulo = max(16, (util - h_logo - grupo - 30) * 0.5)
    resto = entre

    def espaco(pts):
        if pts > 0:
            _espaco_apos(doc, pts)

    espaco(antes_logo)
    pl = _formato_par(doc.add_paragraph())
    # "no mínimo" (e não "exata"): com altura exata o Word corta a parte de baixo da imagem da logo
    from docx.enum.text import WD_LINE_SPACING
    pl.paragraph_format.line_spacing = Pt(h_logo)
    pl.paragraph_format.line_spacing_rule = WD_LINE_SPACING.AT_LEAST
    if LOGO_CLARA.exists():
        pl.add_run().add_picture(str(LOGO_CLARA), width=Cm(4.2))
    espaco(antes_titulo)
    pg = _formato_par(doc.add_paragraph())
    pg.paragraph_format.line_spacing = Pt(h_grossa)
    pg.paragraph_format.keep_with_next = True
    for txt, cor, grosso in pedacos:  # mesma frase: fino branco + grosso branco + grosso âmbar
        _run(pg, txt, 34, cor, negrito=grosso, fonte=FONTE_TITULO if grosso else FONTE_FINA)
    if subtitulo:
        ps = _formato_par(doc.add_paragraph(), 8, 0)
        ps.paragraph_format.line_spacing = Pt(h_sub)
        _run(ps, subtitulo, 13, "texto_faixa", italico=True)
    espaco(resto)
    if info:
        n = min(len(info), 3)
        itens = list(info.items())
        linhas = (len(itens) + n - 1) // n
        t = _tabela_base(doc, linhas, n, [LARGURA_UTIL / n] * n)
        for idx, (rot, val) in enumerate(itens):
            c = t.rows[idx // n].cells[idx % n]
            _bordas_celula(c, top=("3E8A96", 4))
            _margens_celula(c, 100, 120, 0, 200)
            _run(_par_celula(c, depois=2), rot.upper(), 8, "texto_faixa", negrito=True)
            _run(_par_celula(c, linha=1.15, novo=True), fmt(val, "data"), 11, "branco")
    # quebra de página num parágrafo de 1 pt: não escorrega para a página seguinte
    pq = _formato_par(doc.add_paragraph())
    pq.paragraph_format.line_spacing = Pt(1)
    pq.add_run().add_break(WD_BREAK.PAGE)


def titulo(doc, texto_, nivel=1, numerar=True):
    """Título de seção (NBR 6024). Numeração progressiva automática: "1 INTRODUÇÃO", "1.1 Objetivo".
    nivel 1 = primária (MAIÚSCULAS, negrito, Sora), 2 = secundária (negrito, Sora),
    3 = terciária (negrito), 4 = quaternária (itálico). Todos em fonte 12, alinhados à esquerda.
    numerar=False -> título sem número e centralizado (Referências, Apêndice, Anexo, Glossário)."""
    if not numerar:
        p = doc.add_heading(texto_, level=1)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        return p
    nums = doc.__dict__.setdefault("_saldus_secoes", [0, 0, 0, 0])
    nums[nivel - 1] += 1
    for i in range(nivel, 4):
        nums[i] = 0
    indicativo = ".".join(str(n) for n in nums[:nivel])
    # Se o texto já vier numerado ("1 Introdução" / "1. Introdução"), respeita o que veio
    if re.match(r"^\d+(\.\d+)*\.?\s", texto_):
        return doc.add_heading(re.sub(r"^(\d+(?:\.\d+)*)\.\s", r"\1 ", texto_), level=nivel)
    return doc.add_heading(f"{indicativo} {texto_}", level=nivel)


def texto(doc, conteudo, estilo=None):
    """Parágrafo do corpo (ABNT: justificado, recuo 1,25 cm, entrelinha 1,5).
    Aceita **negrito**, _itálico_, links [texto](https://...) e marcadores de manual {1}, {2}
    (o mesmo círculo numerado usado nas capturas de tela)."""
    p = doc.add_paragraph(style=estilo)
    _escrever(p, conteudo)
    return p


_TOKENS = re.compile(r"(\{\d+\}|\[[^\]]+\]\([^)\s]+\))")


def _escrever(par, conteudo, tamanho=None, cor="texto"):
    for parte in _TOKENS.split(conteudo):
        if not parte:
            continue
        if re.fullmatch(r"\{\d+\}", parte):
            marcador_texto(par, int(parte[1:-1]))
        elif re.fullmatch(r"\[[^\]]+\]\([^)\s]+\)", parte):
            rot, url = re.match(r"\[([^\]]+)\]\(([^)]+)\)", parte).groups()
            link(par, rot, url, tamanho)
        else:
            _ricos(par, parte, tamanho, cor)


def link(par, rotulo, url, tamanho=None):
    """Link clicável (petróleo, sublinhado) dentro de um parágrafo."""
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    r_id = par.part.relate_to(url, RT.HYPERLINK, is_external=True)
    h = OxmlElement("w:hyperlink")
    h.set(qn("r:id"), r_id)
    r = _run(par, rotulo, tamanho, "primaria")
    r.font.underline = True
    h.append(r._r)  # move o run para dentro do hyperlink
    par._p.append(h)
    return h


def lista(doc, itens, numerada=False):
    """Lista com marcadores (ou numerada) para itens soltos."""
    estilo = "List Number" if numerada else "List Bullet"
    return [texto(doc, i, estilo) for i in itens]


def alineas(doc, itens):
    """Alíneas da NBR 6024: a) b) c)..., começando em minúscula, terminando em ponto e vírgula e a
    última em ponto. A pontuação final é acrescentada se faltar."""
    pars = []
    for i, item in enumerate(itens):
        item = item.rstrip()
        fim = "." if i == len(itens) - 1 else ";"
        if item[-1:] not in ".;:!?":
            item += fim
        p = _formato_par(doc.add_paragraph(), linha=ENTRELINHA, alinhamento=WD_ALIGN_PARAGRAPH.JUSTIFY,
                         recuo_esq=RECUO_PARAGRAFO + Cm(0.6), recuo_primeira=-Cm(0.6))
        _run(p, f"{chr(97 + i)})\t", TAM_CORPO)
        p.paragraph_format.tab_stops.add_tab_stop(RECUO_PARAGRAFO + Cm(0.6))
        _escrever(p, item, TAM_CORPO)
        pars.append(p)
    return pars


def citacao(doc, conteudo, referencia=None):
    """Citação direta longa (mais de 3 linhas, NBR 10520:2023): recuo de 4 cm, fonte 10,
    entrelinha simples, sem aspas, separada do texto por uma linha em branco.
    referencia: "(Silva, 2019, p. 12)" — sobrenome só com a inicial maiúscula (regra de 2023)."""
    p = _formato_par(doc.add_paragraph(), LINHA_EM_BRANCO, LINHA_EM_BRANCO, 1.0,
                     WD_ALIGN_PARAGRAPH.JUSTIFY, recuo_esq=RECUO_CITACAO)
    _ricos(p, conteudo.strip(), TAM_MENOR)
    if referencia:
        _run(p, " " + referencia.strip(), TAM_MENOR)
    return p


def nota(doc, conteudo):
    """Texto menor (fonte 10, simples) para observações curtas."""
    p = _formato_par(doc.add_paragraph(), 0, 6, 1.0)
    _ricos(p, conteudo, TAM_MENOR, cor="texto_suave")
    return p


def _identificacao(doc, designacao, titulo_ilustracao):
    """Linha acima da ilustração/tabela: "Tabela 1 – Título" (NBR 14724)."""
    n = _contador(doc, designacao)
    p = _formato_par(doc.add_paragraph(style="Caption"), 6, 4, 1.0)
    p.paragraph_format.keep_with_next = True
    _run(p, f"{designacao} {n} – ", TAM_MENOR, "primaria", negrito=True, fonte=FONTE_TITULO)
    _ricos(p, titulo_ilustracao, TAM_MENOR)
    return p


def _fonte_ilustracao(doc, fonte):
    """Linha abaixo da ilustração/tabela: "Fonte: ..." (obrigatória na ABNT, mesmo se autoria própria)."""
    p = _formato_par(doc.add_paragraph(), 4, LINHA_EM_BRANCO, 1.0)
    _run(p, "Fonte: ", TAM_MENOR, "texto", negrito=True)
    _ricos(p, fonte, TAM_MENOR)
    return p


def _larguras_auto(cabecalho, linhas, formatos):
    pesos = []
    for j, nome in enumerate(cabecalho):
        maior = max([len(str(nome))] + [len(fmt(l[j], formatos.get(nome))) for l in linhas] or [4])
        pesos.append(min(max(maior, 6), 40))
    total = sum(pesos)
    return [LARGURA_UTIL * p / total for p in pesos]


def tabela(doc, cabecalho, linhas, formatos=None, larguras=None, total=None, coluna_status=None,
           titulo_tabela=None, fonte="Elaborado pela Saldus.", designacao="Tabela", status_extra=None):
    """Tabela no padrão da planilha (cabeçalho petróleo com filete âmbar, zebra, só linhas
    horizontais, total destacado), com identificação ABNT em cima ("Tabela 1 – ...") e fonte embaixo.

    formatos: {"coluna": "moeda"|"inteiro"|"decimal"|"percentual"|"variacao"|"data"|"mes"|"centro"}
    larguras: lista em cm (somam 15) ou None para automático.
    total:    True = soma das colunas moeda/inteiro/decimal; lista = valores prontos.
    designacao: "Tabela" (dados numéricos) ou "Quadro" (informação textual).
    fonte:    texto após "Fonte:"; None para omitir.
    """
    formatos = formatos or {}
    linhas = [list(l) for l in linhas]
    status = {k: list(v) for k, v in STATUS.items()}
    for k, v in (status_extra or {}).items():
        status.setdefault(k, []).extend(v)
    cor_status = {s.lower(): k for k, v in status.items() for s in v}

    if total is True:
        tot = ["Total"] + [""] * (len(cabecalho) - 1)
        for j, nome in enumerate(cabecalho[1:], 1):
            if formatos.get(nome) in ("moeda", "inteiro", "decimal"):
                tot[j] = sum(l[j] for l in linhas if isinstance(l[j], (int, float)))
        total = tot

    if titulo_tabela:
        _identificacao(doc, designacao, titulo_tabela)
    ws = [Cm(x) for x in larguras] if larguras else _larguras_auto(cabecalho, linhas, formatos)
    escala = LARGURA_UTIL / sum(ws)
    ws = [w * escala for w in ws]
    t = _tabela_base(doc, len(linhas) + 1 + (1 if total else 0), len(cabecalho), ws)

    def alinh(nome):
        f = formatos.get(nome)
        if f in NUMERICOS:
            return WD_ALIGN_PARAGRAPH.RIGHT
        if f in CENTRO or nome == coluna_status:
            return WD_ALIGN_PARAGRAPH.CENTER
        return WD_ALIGN_PARAGRAPH.LEFT

    cab = t.rows[0]
    _repetir_cabecalho(cab)
    _nao_quebrar(cab)
    for j, nome in enumerate(cabecalho):
        c = cab.cells[j]
        _sombrear(c, "primaria")
        _bordas_celula(c, bottom=("destaque", 18))
        _margens_celula(c, 90, 90)
        c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        _run(_par_celula(c, alinh(nome)), str(nome), TAM_MENOR, "branco", negrito=True, fonte=FONTE_TITULO)

    for i, valores in enumerate(linhas, 1):
        row = t.rows[i]
        _nao_quebrar(row)
        for j, nome in enumerate(cabecalho):
            c = row.cells[j]
            if i % 2 == 0:
                _sombrear(c, "zebra")
            ultima = i == len(linhas) and not total
            _bordas_celula(c, bottom=("primaria", 8) if ultima else ("borda", 4))
            _margens_celula(c)
            c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            v = valores[j]
            cor, negrito = "texto", False
            if nome == coluna_status and isinstance(v, str) and v.lower() in cor_status:
                cor, negrito = cor_status[v.lower()], True
            elif isinstance(v, (int, float)) and v < 0 and formatos.get(nome) in NUMERICOS:
                cor = "negativo"
            elif formatos.get(nome) == "variacao" and isinstance(v, (int, float)) and v > 0:
                cor = "positivo"
            _run(_par_celula(c, alinh(nome)), fmt(v, formatos.get(nome)), TAM_MENOR, cor, negrito=negrito)

    if total:
        row = t.rows[-1]
        _nao_quebrar(row)
        for j, nome in enumerate(cabecalho):
            c = row.cells[j]
            _sombrear(c, "primaria_clara")
            _bordas_celula(c, top=("primaria", 12))
            _margens_celula(c, 90, 90)
            c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            v = total[j] if j < len(total) else ""
            cor = "negativo" if isinstance(v, (int, float)) and v < 0 else "primaria"
            _run(_par_celula(c, alinh(nome) if j else WD_ALIGN_PARAGRAPH.LEFT),
                 fmt(v, formatos.get(nome)), TAM_MENOR, cor, negrito=True)

    if fonte:
        _fonte_ilustracao(doc, fonte)
    else:
        _espaco_apos(doc, LINHA_EM_BRANCO)
    return t


def cartoes(doc, itens):
    """Faixa de cartões com o resultado principal (como no Resumo da planilha).
    itens: [{"rotulo": "Custo mensal", "valor": "R$ 1.800,00", "detalhe": "6 itens", "destaque": True}]
    Só UM cartão com destaque=True (fundo âmbar). Ocupam 100% da largura, larguras iguais."""
    n = len(itens)
    t = _tabela_base(doc, 1, n, [LARGURA_UTIL / n] * n)
    _nao_quebrar(t.rows[0])
    for j, k in enumerate(itens):
        c = t.rows[0].cells[j]
        _sombrear(c, "destaque" if k.get("destaque") else "primaria_clara")
        # bordas brancas grossas entre os cartões = o espaço entre eles; nada nas pontas
        _bordas_celula(c, left=("branco", 36) if j else None, right=("branco", 36) if j < n - 1 else None)
        _margens_celula(c, 200, 200, 160, 160)
        c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        cor_rot = "primaria" if k.get("destaque") else "texto_suave"
        C = WD_ALIGN_PARAGRAPH.CENTER
        _run(_par_celula(c, C, depois=4), k["rotulo"].upper(), 8, cor_rot, negrito=True)
        _run(_par_celula(c, C, depois=2, novo=True), str(k["valor"]), 18, "primaria", negrito=True,
             fonte=FONTE_TITULO)
        if k.get("detalhe"):
            _run(_par_celula(c, C, linha=1.1, novo=True), k["detalhe"], 9, cor_rot)
    _espaco_apos(doc, LINHA_EM_BRANCO)
    return t


def destaque(doc, conteudo, titulo=None, tipo="info"):
    """Caixa de destaque: fundo claro com barra lateral âmbar.
    tipo="info" (petróleo claro) ou "atencao" (fundo âmbar claro, barra laranja)."""
    t = _tabela_base(doc, 1, 1, [LARGURA_UTIL])
    c = t.rows[0].cells[0]
    _nao_quebrar(t.rows[0])
    fundo, barra = ("entrada", "alerta") if tipo == "atencao" else ("primaria_clara", "destaque")
    _sombrear(c, fundo)
    _bordas_celula(c, left=(barra, 36))
    _margens_celula(c, 160, 160, 260, 220)
    novo = False
    if titulo:
        _run(_par_celula(c, depois=3), titulo, 11, "primaria", negrito=True, fonte=FONTE_TITULO)
        novo = True
    for bloco in conteudo.split("\n"):
        p = _par_celula(c, WD_ALIGN_PARAGRAPH.JUSTIFY, depois=2, linha=1.25, novo=novo)
        novo = True
        _escrever(p, bloco, 11)
    _espaco_apos(doc, LINHA_EM_BRANCO)
    return t


def ficha(doc, campos, colunas=2):
    """Bloco de dados rótulo/valor (cliente, data, versão, responsável...).
    Rótulo em MAIÚSCULAS cinza, valor embaixo. Distribui em `colunas` colunas iguais."""
    itens = list(campos.items())
    linhas = (len(itens) + colunas - 1) // colunas
    t = _tabela_base(doc, linhas, colunas, [LARGURA_UTIL / colunas] * colunas)
    for idx, (rot, val) in enumerate(itens):
        c = t.rows[idx // colunas].cells[idx % colunas]
        _bordas_celula(c, top=("borda", 4))
        _margens_celula(c, 100, 140, 0, 200)
        _run(_par_celula(c, depois=2), rot.upper(), 8, "texto_suave", negrito=True)
        _ricos(_par_celula(c, linha=1.15, novo=True), fmt(val, "data"), 11)
    _espaco_apos(doc, LINHA_EM_BRANCO)
    return t


def assinaturas(doc, pessoas, local_data=None):
    """Bloco de assinatura. pessoas: [("Nome", "Cargo · Empresa"), ...] lado a lado.
    local_data: "Cidade, 16 de novembro de 2026." (use data_extenso())."""
    if local_data:
        p = _formato_par(doc.add_paragraph(), LINHA_EM_BRANCO, 30, ENTRELINHA, WD_ALIGN_PARAGRAPH.RIGHT)
        p.paragraph_format.keep_with_next = True  # data e assinaturas na mesma página
        _run(p, local_data, TAM_CORPO)
    n = len(pessoas)
    t = _tabela_base(doc, 1, n, [LARGURA_UTIL / n] * n)
    _nao_quebrar(t.rows[0])
    for j, (nome, cargo) in enumerate(pessoas):
        c = t.rows[0].cells[j]
        folga = max(int((LARGURA_UTIL / n - Cm(7.5)) / 635 / 2), 200)  # linha de ~7,5 cm
        _margens_celula(c, 0, 0, folga, folga)
        p = _par_celula(c, WD_ALIGN_PARAGRAPH.CENTER, antes=36, depois=2)
        _borda_par(p, "top", "texto", sz=6, espaco=6)
        _run(p, nome, TAM_CORPO, negrito=True)
        _run(_par_celula(c, WD_ALIGN_PARAGRAPH.CENTER, linha=1.1, novo=True), cargo, TAM_MENOR, "texto_suave")
    return t


def referencias(doc, itens, titulo_secao="Referências"):
    """Seção de referências (NBR 6023): título sem número e centralizado, ordem alfabética,
    alinhadas à esquerda, entrelinha simples, separadas por uma linha em branco simples.
    Use **negrito** para o título da obra: "SILVA, J. **Título**. São Paulo: Editora, 2019." """
    titulo(doc, titulo_secao, numerar=False)
    pars = []
    for ref in sorted(itens, key=lambda s: re.sub(r"[^\w]", "", s).lower()):
        p = _formato_par(doc.add_paragraph(), 0, 12, 1.0)
        _ricos(p, ref, TAM_CORPO, cor_negrito="texto")
        pars.append(p)
    return pars


# ---------------------------------------------------------------- manuais: capturas de tela marcadas
def _fonte_pil(tamanho):
    """Fonte negrito para desenhar os números (Sora > Poppins > Arial > DejaVu)."""
    from PIL import ImageFont
    casa = Path.home()
    candidatas = [ASSETS / "fontes" / "Sora-Bold.ttf", casa / "Library/Fonts/Sora[wght].ttf",
                  casa / "Library/Fonts/Poppins-Bold.ttf", "/Library/Fonts/Arial Bold.ttf",
                  "/System/Library/Fonts/Supplemental/Arial Bold.ttf", "C:/Windows/Fonts/arialbd.ttf",
                  "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", "DejaVuSans-Bold.ttf"]
    for c in candidatas:
        try:
            f = ImageFont.truetype(str(c), tamanho)
        except Exception:
            continue
        if "Sora" in str(c):
            try:
                f.set_variation_by_name("Bold")
            except Exception:
                pass
        return f
    return ImageFont.load_default(tamanho)


def _circulo_numero(draw, cx, cy, r, n):
    """Círculo petróleo com contorno branco e número branco (o marcador do manual)."""
    borda = max(2, round(r * 0.16))
    draw.ellipse([cx - r - borda, cy - r - borda, cx + r + borda, cy + r + borda], fill="#FFFFFF")
    draw.ellipse([cx - r, cy - r, cx + r, cy + r], fill="#" + COR["primaria"])
    f = _fonte_pil(int(r * 1.25))
    draw.text((cx, cy), str(n), font=f, fill="#FFFFFF", anchor="mm")


def _seta(draw, x0, y0, x1, y1, largura):
    """Seta de (x0, y0) até a ponta em (x1, y1): contorno branco por baixo para aparecer em
    qualquer fundo, linha e ponta em petróleo."""
    import math
    ang = math.atan2(y1 - y0, x1 - x0)
    tam = largura * 3.2
    base = (x1 - tam * math.cos(ang), y1 - tam * math.sin(ang))
    ponta = [(x1, y1),
             (base[0] + tam * 0.55 * math.sin(ang), base[1] - tam * 0.55 * math.cos(ang)),
             (base[0] - tam * 0.55 * math.sin(ang), base[1] + tam * 0.55 * math.cos(ang))]
    for cor, extra in (("#FFFFFF", largura * 0.9), ("#" + COR["primaria"], 0)):
        draw.line([(x0, y0), base], fill=cor, width=int(largura + extra))
        if extra:
            draw.polygon(ponta, fill=cor, outline=cor, width=int(extra))
        else:
            draw.polygon(ponta, fill=cor)


def marcar_imagem(imagem, marcas, moldura=True):
    """Desenha as marcações de manual numa captura de tela e devolve um PNG (BytesIO).

    marcas: lista de dicts, uma por passo:
        {"n": 1,                      # número do passo (vai no círculo)
         "alvo": (x, y),              # ponto exato para onde a seta aponta
         "caixa": (x1, y1, x2, y2),   # opcional: retângulo âmbar em volta do item clicado
         "lado": "cima",              # onde fica o círculo: cima, baixo, esquerda, direita
         "distancia": None}           # opcional: distância do círculo até o alvo
    Coordenadas em pixels da imagem original, ou em frações (0 a 1) da largura/altura.
    Sem "alvo" e com "caixa", a seta aponta para o meio do lado da caixa virado para o círculo.
    Se o círculo cair fora da captura (ex.: "cima" num item do menu do topo), a imagem ganha uma
    margem branca e o número fica do lado de fora, sem cobrir a tela.
    moldura=True põe uma borda fina cinza em volta da captura (separa do fundo branco da página).
    """
    from PIL import Image, ImageDraw
    img = Image.open(imagem).convert("RGB")
    fator = 2 if img.width < 1400 else 1  # trabalha em alta resolução para o traço sair liso
    if fator > 1:
        img = img.resize((img.width * fator, img.height * fator), Image.LANCZOS)
    W, H = img.size

    def px(v, eixo):
        tot = W if eixo == 0 else H
        return v * tot if 0 <= v <= 1 and isinstance(v, float) else v * fator

    r = max(14 * fator, round(W * 0.016))
    traco = max(3, round(r * 0.22))
    caixas = []
    for m in marcas:
        if m.get("caixa"):
            c = m["caixa"]
            caixas.append((px(c[0], 0), px(c[1], 1), px(c[2], 0), px(c[3], 1)))
    # 1º calcula onde fica cada círculo; se cair fora da captura, a imagem ganha margem branca
    planos = []
    for m in marcas:
        lado = m.get("lado", "cima")
        if m.get("alvo"):
            ax, ay = px(m["alvo"][0], 0), px(m["alvo"][1], 1)
        else:
            x1, y1, x2, y2 = (px(m["caixa"][0], 0), px(m["caixa"][1], 1),
                              px(m["caixa"][2], 0), px(m["caixa"][3], 1))
            ax, ay = {"cima": ((x1 + x2) / 2, y1), "baixo": ((x1 + x2) / 2, y2),
                      "esquerda": (x1, (y1 + y2) / 2), "direita": (x2, (y1 + y2) / 2)}[lado]
        d = m["distancia"] * fator if m.get("distancia") else 3.4 * r
        dx, dy = {"cima": (0, -1), "baixo": (0, 1), "esquerda": (-1, 0), "direita": (1, 0)}[lado]
        planos.append((m["n"], ax, ay, ax + dx * d, ay + dy * d, dx, dy))
    folga = r * 1.4
    esq = max([0] + [folga - cx for _, _, _, cx, _, _, _ in planos])
    topo = max([0] + [folga - cy for _, _, _, _, cy, _, _ in planos])
    dir_ = max([0] + [cx + folga - W for _, _, _, cx, _, _, _ in planos])
    base = max([0] + [cy + folga - H for _, _, _, _, cy, _, _ in planos])
    if esq or topo or dir_ or base:
        nova = Image.new("RGB", (int(W + esq + dir_), int(H + topo + base)), "#FFFFFF")
        nova.paste(img, (int(esq), int(topo)))
        img = nova
    ox, oy = int(esq), int(topo)
    draw = ImageDraw.Draw(img)
    if moldura:  # moldura só em volta da captura original
        draw.rectangle([ox, oy, ox + W - 1, oy + H - 1], outline="#" + COR["borda"], width=max(2, fator * 2))
    for x1, y1, x2, y2 in caixas:
        draw.rounded_rectangle([ox + x1, oy + y1, ox + x2, oy + y2], radius=traco * 2,
                               outline="#" + COR["destaque"], width=traco)
    for n, ax, ay, cx, cy, dx, dy in planos:
        _seta(draw, ox + cx - dx * r, oy + cy - dy * r, ox + ax - dx * traco * 0.6,
              oy + ay - dy * traco * 0.6, traco)
        _circulo_numero(draw, ox + cx, oy + cy, r, n)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf


def marcador_texto(par, n, altura_cm=0.45):
    """Círculo numerado dentro do texto, igual ao da captura: "Clique em {1} Rural"."""
    from PIL import Image, ImageDraw
    lado = 96
    img = Image.new("RGBA", (lado, lado), (255, 255, 255, 0))
    _circulo_numero(ImageDraw.Draw(img), lado / 2, lado / 2, lado * 0.42, n)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    r = par.add_run()
    r.add_picture(buf, height=Cm(altura_cm))
    pos = OxmlElement("w:position")  # desce um pouco para alinhar com a linha do texto
    pos.set(qn("w:val"), "-4")
    r._r.get_or_add_rPr().append(pos)
    par.add_run("\u00a0")  # espaço que não quebra: o número não fica sozinho no fim da linha
    return r


def captura(doc, imagem, marcas=None, titulo_figura=None, fonte="Elaborado pela Saldus.",
            largura_cm=None):
    """Captura de tela de manual: marcações numeradas + moldura fina, largura total do texto.
    Com titulo_figura, ganha a identificação ABNT ("Figura 3 – ...") e a fonte embaixo."""
    png = marcar_imagem(imagem, marcas or [])
    if titulo_figura:
        return figura(doc, png, titulo_figura, fonte, largura_cm=largura_cm)
    p = _formato_par(doc.add_paragraph(), 6, LINHA_EM_BRANCO, 1.0, WD_ALIGN_PARAGRAPH.CENTER)
    p.add_run().add_picture(png, width=Cm(largura_cm) if largura_cm else LARGURA_UTIL)
    return p


def quebra_pagina(doc):
    _formato_par(doc.add_paragraph()).add_run().add_break(WD_BREAK.PAGE)


def figura(doc, imagem, titulo_figura, fonte="Elaborado pela Saldus.", designacao="Figura", largura_cm=None):
    """Imagem com identificação ABNT em cima ("Figura 1 – ...") e fonte embaixo."""
    _identificacao(doc, designacao, titulo_figura)
    p = _formato_par(doc.add_paragraph(), alinhamento=WD_ALIGN_PARAGRAPH.CENTER)
    p.paragraph_format.keep_with_next = True
    p.add_run().add_picture(imagem, width=Cm(largura_cm) if largura_cm else LARGURA_UTIL)
    if fonte:
        _fonte_ilustracao(doc, fonte)
    return p


def grafico_barras(doc, categorias, valores, titulo_grafico, formato="moeda", horizontal=True,
                   fonte="Elaborado pela Saldus.", altura_cm=None):
    """Gráfico de barras (imagem matplotlib) no padrão: sem moldura, sem grade, valor na ponta da
    barra, cor petróleo, largura total; identificação "Gráfico N – ..." em cima e fonte embaixo."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib import font_manager

    nomes = {f.name for f in font_manager.fontManager.ttflist}
    fam = FONTE if FONTE in nomes else "DejaVu Sans"
    n = len(categorias)
    larg_in = LARGURA_UTIL / 914400
    alt_in = (altura_cm / 2.54) if altura_cm else (0.4 + 0.3 * n if horizontal else 2.6)
    fig, ax = plt.subplots(figsize=(larg_in, alt_in), dpi=200)
    cor = "#" + COR["primaria"]
    pos = list(range(n))
    if horizontal:
        barras = ax.barh(pos, valores, color=cor, height=0.62)
        ax.set_yticks(pos, categorias)
        ax.invert_yaxis()
        ax.set_xticks([])
        ax.set_xlim(0, max(valores) * 1.2)
    else:
        barras = ax.bar(pos, valores, color=cor, width=0.62)
        ax.set_xticks(pos, categorias)
        ax.set_yticks([])
        ax.set_ylim(0, max(valores) * 1.15)
    for s in ax.spines.values():
        s.set_visible(False)
    eixo = ax.spines["left"] if horizontal else ax.spines["bottom"]
    eixo.set_visible(True)
    eixo.set_color("#" + COR["borda"])
    ax.tick_params(length=0, labelsize=8, labelcolor="#" + COR["texto"])
    for lbl in ax.get_xticklabels() + ax.get_yticklabels():
        lbl.set_fontfamily(fam)
    ax.bar_label(barras, labels=[fmt(v, formato) for v in valores], padding=4, fontsize=8,
                 color="#" + COR["texto"], fontfamily=fam)
    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", facecolor="white")
    plt.close(fig)
    buf.seek(0)
    return figura(doc, buf, titulo_grafico, fonte, designacao="Gráfico")


_ORDEM = {
    "tblPr": ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize",
              "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd",
              "tblLayout", "tblCellMar", "tblLook"],
    "trPr": ["cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore", "wAfter", "cantSplit",
             "trHeight", "tblHeader", "tblCellSpacing", "jc", "hidden"],
    "pPr": ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
            "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
            "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
            "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc",
            "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle",
            "rPr", "sectPr", "pPrChange"],
    "tcPr": ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap",
             "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark"],
}


def _ordenar_xml(raiz):
    """O Word recusa arquivos com elementos fora da ordem do schema; reordena por segurança."""
    for tag, ordem in _ORDEM.items():
        pos = {qn(f"w:{n}"): i for i, n in enumerate(ordem)}
        for el in raiz.iter(qn(f"w:{tag}")):
            filhos = sorted(el, key=lambda c: pos.get(c.tag, len(ordem)))
            for c in filhos:
                el.remove(c)
                el.append(c)


def salvar(doc, caminho):
    """Salva garantindo XML válido para o Word (corpo, estilos, cabeçalhos e rodapés)."""
    _ordenar_xml(doc.element.body)
    _ordenar_xml(doc.styles.element)
    for sec in doc.sections:
        for parte in (sec.header, sec.footer, sec.first_page_header, sec.first_page_footer):
            if not parte.is_linked_to_previous:
                _ordenar_xml(parte._element)
    doc.save(caminho)
    return caminho


# ---------------------------------------------------------------- modelo e exemplo
def modelo_em_branco(caminho):
    """Modelo para quem escreve direto no Word (faixa, estilos, cabeçalho e rodapé prontos)."""
    doc = novo_documento("Título do documento", "Assunto · data · área responsável")
    titulo(doc, "Primeira seção")
    texto(doc, "Escreva aqui. O texto já está no padrão ABNT: fonte 12, entrelinha 1,5, recuo de 1,25 cm "
               "na primeira linha e alinhamento justificado. Use os estilos **Título 1** (seção primária), "
               "**Título 2** (secundária) e **Normal** da galeria de estilos do Word. Para trocar o título "
               "do documento, clique na faixa petróleo acima e digite.")
    titulo(doc, "Subseção", 2)
    texto(doc, "Para tabelas, copie o exemplo abaixo, com a identificação em cima e a fonte embaixo.")
    tabela(doc, ["Descrição", "Data", "Valor"], [["Item de exemplo", _dt.date(2026, 1, 15), 1000.0],
                                                  ["Outro item", _dt.date(2026, 2, 1), 250.5]],
           formatos={"Data": "data", "Valor": "moeda"}, total=True, titulo_tabela="Título da tabela")
    destaque(doc, "Use para premissas, conclusões ou avisos importantes.", titulo="Caixa de destaque")
    referencias(doc, ["SOBRENOME, Nome. **Título da obra**. Cidade: Editora, 2026."])
    salvar(doc, caminho)
    return caminho


def _print_exemplo():
    """Tela de exemplo (cinza) para o modelo de manual: mostra onde colar o print e como fica marcado."""
    from PIL import Image, ImageDraw
    w, h = 1600, 820
    img = Image.new("RGB", (w, h), "#F2F4F5")
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, w, 90], fill="#" + COR["primaria"])
    f, fg = _fonte_pil(30), _fonte_pil(44)
    d.text((40, 45), "Sistema", font=f, fill="#FFFFFF", anchor="lm")
    for i, nome in enumerate(["Início", "Cadastros", "Relatórios", "Ajuda"]):
        d.text((420 + i * 230, 45), nome, font=f, fill="#FFFFFF", anchor="lm")
    d.rounded_rectangle([1200, 600, 1520, 700], radius=16, fill="#" + COR["primaria"])
    d.text((1360, 650), "Salvar", font=f, fill="#FFFFFF", anchor="mm")
    d.text((w / 2, 330), "Cole aqui o print da tela", font=fg, fill="#8A9A9E", anchor="mm")
    d.text((w / 2, 400), "e marque os cliques com números e setas", font=f, fill="#8A9A9E", anchor="mm")
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf


def modelo(tipo, caminho):
    """Modelos prontos para abrir no Word e preencher.
    tipo: "documento" (formatado: capa inteira, cartões, tabela, gráfico), "documento-simples" (faixa de
    título), "manual" (capa inteira, aviso, marca d'água, rodapé em faixa, print marcado) ou "proposta"."""
    if tipo == "manual":
        doc = novo_documento("Manual de **nome do processo**", "Passo a passo · mês/ano",
                             capa={"estilo": "inteira", "info": {"Área": "Nome da área", "Versão": "1.0",
                                                                 "Data": "dd/mm/aaaa"}},
                             aviso="confidencial", marca_dagua=True, cabecalho=False, estilo_rodape="faixa")
        titulo(doc, "Objetivo")
        texto(doc, "Explique em duas ou três frases para que serve este manual e quem deve usá-lo.")
        titulo(doc, "Antes de começar")
        lista(doc, ["Acesso ao sistema (usuário e senha);", "Permissão de acesso ao módulo;",
                    "Arquivos ou dados necessários."])
        titulo(doc, "Nome da primeira tarefa")
        texto(doc, "Acesse [o sistema](https://exemplo.com.br) e clique em {1} **“Relatórios”** e depois em "
                   "{2} **“Salvar”**. Troque o print abaixo pelo da sua tela.")
        captura(doc, _print_exemplo(), [
            {"n": 1, "caixa": (870, 18, 1060, 72), "lado": "baixo"},
            {"n": 2, "caixa": (1200, 600, 1520, 700), "lado": "esquerda"},
        ], titulo_figura="Descreva a tela", fonte="captura de tela do sistema (ano).")
        destaque(doc, "Use esta caixa para dicas e cuidados do passo {1}.", titulo="Dica")
        titulo(doc, "Dúvidas frequentes")
        texto(doc, "**Pergunta?** Resposta curta e direta.")
    elif tipo == "proposta":
        doc = novo_documento("Proposta Comercial", "Nome do projeto · Nome do cliente",
                             capa={"info": {"Preparado para": "Nome do cliente", "Data": "dd/mm/aaaa",
                                            "Validade": "30 dias"}}, rodape="Saldus · Confidencial")
        for sec, txt in [("Contexto", "Descreva a situação e a necessidade do cliente."),
                         ("Solução", "Explique o que a Saldus vai entregar e como isso resolve o problema.")]:
            titulo(doc, sec)
            texto(doc, txt)
        titulo(doc, "Escopo")
        alineas(doc, ["primeira entrega", "segunda entrega", "terceira entrega"])
        titulo(doc, "Investimento")
        tabela(doc, ["Etapa", "Prazo", "Valor"], [["Etapa 1", "dd/mm/aaaa", 0.0], ["Etapa 2", "dd/mm/aaaa", 0.0]],
               formatos={"Prazo": "centro", "Valor": "moeda"}, total=True, titulo_tabela="Investimento por etapa",
               fonte="Elaborado pela Saldus (valores sem impostos).")
        titulo(doc, "Condições")
        destaque(doc, "Forma de pagamento, prazo total e validade da proposta.", titulo="Condições")
        assinaturas(doc, [("Nome", "Cargo · Saldus"), ("Nome", "Cargo · Cliente")],
                    local_data="Cidade, dd de mês de aaaa.")
    elif tipo == "documento":
        doc = novo_documento("Relatório de **nome do assunto**", "Área · mês/ano",
                             capa={"estilo": "inteira", "info": {"Área": "Nome da área", "Responsável": "Nome",
                                                                 "Data": "dd/mm/aaaa"}})
        titulo(doc, "Resumo")
        texto(doc, "Comece pelo resultado: o que aconteceu, o número mais importante e o que precisa ser "
                   "decidido. Use **negrito** para destacar o essencial.")
        cartoes(doc, [{"rotulo": "Número principal", "valor": "R$ 0,00", "detalhe": "explicação curta",
                       "destaque": True},
                      {"rotulo": "Segundo número", "valor": "0", "detalhe": "explicação curta"},
                      {"rotulo": "Variação", "valor": "+0,0%", "detalhe": "vs. período anterior"}])
        titulo(doc, "Análise")
        texto(doc, "Explique os dados. Cite a tabela antes dela: a Tabela 1 mostra os valores por item.")
        tabela(doc, ["Item", "Responsável", "Valor", "Status"],
               [["Item 1", "Nome", 0.0, "Ativo"], ["Item 2", "Nome", 0.0, "Pendente"],
                ["Item 3", "Nome", 0.0, "Concluído"]],
               formatos={"Valor": "moeda"}, coluna_status="Status", total=True,
               titulo_tabela="Título da tabela", fonte="indique a origem dos dados.")
        grafico_barras(doc, ["Item 1", "Item 2", "Item 3"], [3, 2, 1], "Título do gráfico",
                       formato="inteiro", fonte="indique a origem dos dados.")
        titulo(doc, "Detalhes", 2)
        texto(doc, "Use subseções (1.1, 1.2) para separar assuntos dentro de uma seção.")
        titulo(doc, "Conclusão e próximos passos")
        alineas(doc, ["primeira ação, com responsável e prazo", "segunda ação, com responsável e prazo"])
        destaque(doc, "Resuma aqui a decisão ou o pedido principal deste documento.", titulo="Conclusão")
        referencias(doc, ["SOBRENOME, Nome. **Título da obra**. Cidade: Editora, ano."])
    elif tipo == "formal":
        doc = novo_documento("Política de nome do assunto",
                             "Área responsável · Versão 1.0 · Vigência a partir de ____/____/______ · Confidencial",
                             rodape="Saldus · Confidencial")
        for i, (sec, itens) in enumerate([("Objeto e abrangência", ["Descreva o que esta política regula e a "
                                            "quem se aplica.", "Indique a relação com outros documentos e contratos."]),
                                           ("Definições", ["Termo: significado.", "Outro termo: significado."]),
                                           ("Regras", ["Primeira regra.", "Segunda regra."]),
                                           ("Disposições gerais", ["Como e quando esta política pode ser revista.",
                                                                   "Quem resolve os casos omissos."])], 1):
            titulo(doc, sec)
            for j, txt in enumerate(itens, 1):
                texto(doc, f"**{i}.{j}** {txt}")
        assinaturas(doc, [("SALDUS ______________ LTDA.", "Representante legal")],
                    local_data="____________________________, ____ de ______________ de 20____.")
    elif tipo == "documento-simples":
        return modelo_em_branco(caminho)
    else:
        raise ValueError('tipo de modelo: "documento", "documento-simples", "formal", "manual" ou "proposta"')
    salvar(doc, caminho)
    return caminho


def _exemplo(saida):
    """Exemplo com dados FICTÍCIOS (Empresa Exemplo). Nunca coloque dados reais da empresa aqui: o
    repositório é público."""
    doc = novo_documento("Relatório de Despesas Operacionais", "Empresa Exemplo · janeiro de 2026 · valores em R$")
    titulo(doc, "Resumo")
    texto(doc, "Com a troca do **sistema de atendimento** e o fim de duas assinaturas pouco usadas, a despesa "
               "mensal caiu de R$ 2.400,00 para R$ 1.800,00, uma economia de **25%**.")
    cartoes(doc, [
        {"rotulo": "Despesa mensal", "valor": "R$ 1.800,00", "detalhe": "5 itens ativos", "destaque": True},
        {"rotulo": "Despesa anual", "valor": "R$ 21.600", "detalhe": "projeção 12 meses"},
        {"rotulo": "Economia / mês", "valor": "R$ 600,00", "detalhe": "-25% vs. antes"},
    ])
    titulo(doc, "Itens ativos")
    texto(doc, "A Tabela 1 lista os serviços em uso.")
    tabela(doc, ["Item", "Fornecedor", "Custo mensal", "Status"], [
        ["Sistema de gestão", "Fornecedor A", 900.00, "Ativo"],
        ["Sistema de atendimento", "Fornecedor B", 450.00, "Novo"],
        ["Hospedagem do site", "Fornecedor C", 250.00, "Ativo"],
        ["E-mail corporativo", "Fornecedor C", 120.00, "Ativo"],
        ["Armazenamento em nuvem", "Fornecedor D", 80.00, "Ativo"],
    ], formatos={"Custo mensal": "moeda"}, coluna_status="Status", total=True,
        titulo_tabela="Custo mensal dos itens ativos", fonte="dados fictícios, para exemplo.")
    grafico_barras(doc, ["Gestão", "Atendimento", "Site", "E-mail", "Nuvem"], [900, 450, 250, 120, 80],
                   "Custo mensal por item", fonte="dados fictícios, para exemplo.")
    titulo(doc, "Premissas", 2)
    texto(doc, "Exemplo de citação longa:")
    citacao(doc, "Este é um texto de exemplo para mostrar a citação direta longa, usada quando o trecho "
                 "copiado de outra obra passa de três linhas. Ela fica recuada, em fonte menor, com "
                 "entrelinha simples e sem aspas, e termina com a chamada de autor, ano e página.",
            "(Sobrenome, 2026, p. 10)")
    titulo(doc, "Próximos passos")
    alineas(doc, ["revisar as assinaturas a cada trimestre", "registrar novos serviços na planilha de custos"])
    destaque(doc, "Valores em moeda estrangeira mudam com o câmbio. Revise todo mês.", titulo="Atenção", tipo="atencao")
    assinaturas(doc, [("Nome do responsável", "Cargo · Empresa Exemplo")],
                local_data=f"Cidade, {data_extenso(_dt.date(2026, 1, 31))}.")
    referencias(doc, ["SOBRENOME, Nome. **Título da obra**: subtítulo. Cidade: Editora, 2026."])
    salvar(doc, saida)
    return saida


if __name__ == "__main__":
    if len(sys.argv) > 2 and sys.argv[1] == "--modelo":
        # python estilo_word.py --modelo saida.docx            (documento)
        # python estilo_word.py --modelo manual saida.docx     (manual | proposta | documento | documento-simples)
        if len(sys.argv) > 3:
            print("Gerado:", modelo(sys.argv[2], sys.argv[3]))
        else:
            print("Gerado:", modelo_em_branco(sys.argv[2]))
    else:
        print("Gerado:", _exemplo(sys.argv[1] if len(sys.argv) > 1 else "exemplo_saldus.docx"))
