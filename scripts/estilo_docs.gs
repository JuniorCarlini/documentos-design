/**
 * Estilo padrão Saldus para Google Docs.
 * Instalação: no documento, Extensões → Apps Script → cole este arquivo → salve → recarregue o doc.
 * Aparece o menu "Estilo".
 *
 * Limitações do Apps Script (fazer à mão):
 *  - Logo: Inserir → Imagem, dentro da faixa petróleo (à direita) e no cabeçalho.
 *  - Número de página: Inserir → Números de página (canto superior direito, padrão ABNT).
 *  - Bordas só horizontais em tabelas não existem no Apps Script: as tabelas ficam sem bordas,
 *    separadas pela zebra (fica limpo e no mesmo espírito).
 */

var COR = {
  primaria: '#12424D', destaque: '#FFB745', branco: '#FFFFFF',
  primariaClara: '#E7EFF0', zebra: '#F5F8F8', borda: '#D5E0E2',
  texto: '#1E2B2E', textoSuave: '#5F7378', textoFaixa: '#B7CDD2',
  positivo: '#2E7D32', negativo: '#C62828', alerta: '#B26A00'
};
var FONTE = 'Poppins';
var FONTE_TITULO = 'Sora';
var CM = 28.3465; // pontos por centímetro

var STATUS = {
  positivo: ['ok', 'concluído', 'pago', 'ativo', 'em dia', 'aprovado', 'novo'],
  alerta: ['pendente', 'atenção', 'em andamento', 'em análise', 'em renovação'],
  negativo: ['atrasado', 'erro', 'vencido', 'reprovado'],
  textoSuave: ['cancelado', 'inativo', 'encerrado', 'arquivado']
};

function onOpen() {
  DocumentApp.getUi().createMenu('Estilo')
    .addItem('Aplicar estilo padrão', 'aplicarEstiloPadrao')
    .addItem('Inserir faixa de título', 'inserirFaixaTitulo')
    .addItem('Inserir caixa de destaque', 'inserirDestaque')
    .addToUi();
}

function _atributos(fonte, tamanho, cor, negrito, italico, antes, depois) {
  var a = {};
  a[DocumentApp.Attribute.FONT_FAMILY] = fonte;
  a[DocumentApp.Attribute.FONT_SIZE] = tamanho;
  a[DocumentApp.Attribute.FOREGROUND_COLOR] = cor;
  a[DocumentApp.Attribute.BOLD] = !!negrito;
  a[DocumentApp.Attribute.ITALIC] = !!italico;
  a[DocumentApp.Attribute.UNDERLINE] = false;
  a[DocumentApp.Attribute.SPACING_BEFORE] = antes;
  a[DocumentApp.Attribute.SPACING_AFTER] = depois;
  return a;
}

var H = DocumentApp.ParagraphHeading;
function _estilos() {
  var e = {};
  // ABNT (NBR 14724:2024 / NBR 6024): tudo em 12; títulos separados do texto por uma linha em branco
  e[H.NORMAL] = _atributos(FONTE, 12, COR.texto, false, false, 0, 0);
  e[H.TITLE] = _atributos(FONTE_TITULO, 20, COR.primaria, true, false, 0, 4);
  e[H.SUBTITLE] = _atributos(FONTE, 12, COR.textoSuave, false, true, 0, 18);
  e[H.HEADING1] = _atributos(FONTE_TITULO, 12, COR.primaria, true, false, 27, 18);
  e[H.HEADING2] = _atributos(FONTE_TITULO, 12, COR.primaria, true, false, 18, 18);
  e[H.HEADING3] = _atributos(FONTE, 12, COR.texto, true, false, 18, 18);
  e[H.HEADING4] = _atributos(FONTE, 12, COR.texto, false, true, 18, 18);
  return e;
}

function aplicarEstiloPadrao() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  // A4, margens "Normal" do Word: 2,5 cm em cima/embaixo, 3 cm nas laterais
  body.setPageWidth(21 * CM).setPageHeight(29.7 * CM)
      .setMarginLeft(3 * CM).setMarginRight(3 * CM)
      .setMarginTop(2.5 * CM).setMarginBottom(2.5 * CM);

  var estilos = _estilos();
  for (var h in estilos) {
    try { body.setHeadingAttributes(h, estilos[h]); } catch (e) {}
  }
  // Parágrafos fora de tabelas: aplica o estilo do seu nível (preserva negrito/itálico dentro do texto)
  var pars = body.getParagraphs();
  for (var i = 0; i < pars.length; i++) {
    var p = pars[i];
    if (_dentroDeTabela(p)) continue;
    var a = estilos[p.getHeading()] || estilos[H.NORMAL];
    var t = p.editAsText();
    if (t.getText().length === 0) continue;
    if (p.getHeading() === H.HEADING1) t.setText(t.getText().toUpperCase()); // seção primária em MAIÚSCULAS
    t.setFontFamily(a[DocumentApp.Attribute.FONT_FAMILY])
     .setFontSize(a[DocumentApp.Attribute.FONT_SIZE])
     .setForegroundColor(a[DocumentApp.Attribute.FOREGROUND_COLOR]);
    if (p.getHeading() !== H.NORMAL) t.setBold(a[DocumentApp.Attribute.BOLD]);
    var normal = p.getHeading() === H.NORMAL;
    p.setSpacingBefore(a[DocumentApp.Attribute.SPACING_BEFORE])
     .setSpacingAfter(a[DocumentApp.Attribute.SPACING_AFTER])
     .setLineSpacing(normal ? 1.5 : 1.0);
    if (normal && p.getType() === DocumentApp.ElementType.PARAGRAPH) {
      // texto ABNT: justificado com recuo de 1,25 cm na 1ª linha
      p.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY).setIndentFirstLine(1.25 * CM);
    }
    _negritoEmPetroleo(t);
  }
  var tabelas = body.getTables();
  for (var k = 0; k < tabelas.length; k++) {
    if (!_ehFaixa(tabelas[k])) _estilizarTabela(tabelas[k]);
  }
  _rodape(doc);
}

function _dentroDeTabela(el) {
  var p = el.getParent();
  while (p) {
    if (p.getType() === DocumentApp.ElementType.TABLE_CELL) return true;
    p = p.getParent ? p.getParent() : null;
  }
  return false;
}

function _negritoEmPetroleo(texto) {
  var s = texto.getText();
  for (var i = 0; i < s.length; i++) {
    if (texto.isBold(i)) texto.setForegroundColor(i, i, COR.primaria);
  }
}

function _ehFaixa(tabela) {
  return tabela.getNumRows() === 1 && tabela.getCell(0, 0).getBackgroundColor() === COR.primaria;
}

function _corStatus(valor) {
  var v = String(valor).trim().toLowerCase();
  for (var k in STATUS) if (STATUS[k].indexOf(v) >= 0) return COR[k];
  return null;
}

function _ehNumero(s) {
  return /^[-+]?\s*(R\$\s*)?-?[\d.]+(,\d+)?\s*%?$/.test(String(s).trim());
}

function _estilizarTabela(t) {
  t.setBorderWidth(0);
  var n = t.getNumRows();
  var ultima = t.getRow(n - 1);
  var temTotal = /^total/i.test(ultima.getCell(0).getText().trim());
  for (var r = 0; r < n; r++) {
    var row = t.getRow(r);
    for (var c = 0; c < row.getNumCells(); c++) {
      var cel = row.getCell(c);
      cel.setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(6).setPaddingRight(6);
      cel.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
      var txt = cel.editAsText();
      var conteudo = cel.getText();
      var fundo = COR.branco, cor = COR.texto, negrito = false, fonte = FONTE, tam = 10;
      if (r === 0) {
        fundo = COR.primaria; cor = COR.branco; negrito = true; fonte = FONTE_TITULO; tam = 10;
      } else if (temTotal && r === n - 1) {
        fundo = COR.primariaClara; cor = COR.primaria; negrito = true;
      } else {
        if (r % 2 === 0) fundo = COR.zebra;
        var st = _corStatus(conteudo);
        if (st) { cor = st; negrito = true; }
        else if (/^-|^-R\$/.test(conteudo.trim()) && _ehNumero(conteudo)) cor = COR.negativo;
      }
      cel.setBackgroundColor(fundo);
      if (conteudo.length) {
        txt.setFontFamily(fonte).setFontSize(tam).setForegroundColor(cor).setBold(negrito);
      }
      // números à direita (olha a 2ª linha da coluna para decidir)
      var amostra = n > 1 ? t.getRow(Math.min(1, n - 1)).getCell(Math.min(c, t.getRow(Math.min(1, n - 1)).getNumCells() - 1)).getText() : '';
      var alin = _ehNumero(amostra) ? DocumentApp.HorizontalAlignment.RIGHT
               : (_corStatus(amostra) || /^\d{2}\/\d{2}\/\d{4}$/.test(amostra.trim()))
                 ? DocumentApp.HorizontalAlignment.CENTER : DocumentApp.HorizontalAlignment.LEFT;
      for (var p = 0; p < cel.getNumChildren(); p++) {
        var ch = cel.getChild(p);
        if (ch.getType() === DocumentApp.ElementType.PARAGRAPH) {
          ch.asParagraph().setAlignment(alin).setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1);
        }
      }
    }
  }
}

function inserirFaixaTitulo() {
  var ui = DocumentApp.getUi();
  var r1 = ui.prompt('Faixa de título', 'Título do documento:', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  var r2 = ui.prompt('Faixa de título', 'Subtítulo (assunto · data · área):', ui.ButtonSet.OK_CANCEL);
  var body = DocumentApp.getActiveDocument().getBody();
  var t = body.insertTable(0, [[r1.getResponseText()]]);
  t.setBorderWidth(0);
  var cel = t.getCell(0, 0);
  cel.setBackgroundColor(COR.primaria)
     .setPaddingTop(13).setPaddingBottom(13).setPaddingLeft(17).setPaddingRight(17);
  var p = cel.getChild(0).asParagraph();
  p.editAsText().setFontFamily(FONTE_TITULO).setFontSize(20).setBold(true).setForegroundColor(COR.branco);
  p.setSpacingAfter(4);
  if (r2.getSelectedButton() === ui.Button.OK && r2.getResponseText()) {
    var s = cel.appendParagraph(r2.getResponseText());
    s.editAsText().setFontFamily(FONTE).setFontSize(10).setBold(false).setItalic(true)
     .setForegroundColor(COR.textoFaixa);
    s.setSpacingAfter(0);
  }
  ui.alert('Faixa criada. Insira a logo clara (assets/logo.png) dentro da faixa, à direita.');
}

function inserirDestaque() {
  var doc = DocumentApp.getActiveDocument();
  var cursor = doc.getCursor();
  var body = doc.getBody();
  var idx = body.getNumChildren();
  if (cursor) {
    var el = cursor.getElement();
    while (el.getParent().getType() !== DocumentApp.ElementType.BODY_SECTION) el = el.getParent();
    idx = body.getChildIndex(el) + 1;
  }
  var t = body.insertTable(idx, [['Título da caixa'], ['Texto da caixa de destaque.']]);
  t.setBorderWidth(0);
  for (var r = 0; r < 2; r++) {
    var c = t.getCell(r, 0);
    c.setBackgroundColor(COR.primariaClara).setPaddingLeft(14).setPaddingRight(12)
     .setPaddingTop(r ? 0 : 8).setPaddingBottom(r ? 8 : 0);
    c.editAsText().setFontFamily(r ? FONTE : FONTE_TITULO).setFontSize(r ? 10 : 10.5)
     .setBold(!r).setForegroundColor(r ? COR.texto : COR.primaria);
  }
}

function _rodape(doc) {
  var f = doc.getFooter() || doc.addFooter();
  if (f.getText().trim().length === 0) {
    var p = f.appendParagraph('Saldus');
    p.editAsText().setFontFamily(FONTE).setFontSize(8).setForegroundColor(COR.textoSuave);
  }
}
