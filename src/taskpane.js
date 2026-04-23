/* global Office, Word */

Office.onReady(function (info) {
  if (info.host === Office.HostType.Word) {
    showStatus("DiscereNow pronto.");
  }
});

// ─── Utilitários ────────────────────────────────────────────

function showStatus(msg) {
  const el = document.getElementById("dn-status");
  if (el) el.textContent = msg;
}

function run(fn) {
  return Word.run(fn).catch(function (err) {
    showStatus("Erro: " + err.message);
    console.error(err);
  });
}

// ─── Aplicar estilo de parágrafo ────────────────────────────

function applyStyle(styleName) {
  run(async function (context) {
    const selection = context.document.getSelection();
    selection.paragraphs.load("items");
    await context.sync();

    selection.paragraphs.items.forEach(function (p) {
      p.styleBuiltIn = Word.Style.normal; // reset antes
      p.style = styleName;
    });

    await context.sync();
    showStatus('Estilo "' + styleName + '" aplicado.');
  });
}

// ─── Inserir Acordeão ────────────────────────────────────────

function insertAccordion() {
  run(async function (context) {
    const selection = context.document.getSelection();

    // Marcador de abertura
    const ccOpen = selection.insertContentControl();
    ccOpen.tag = "DN-BLOCK-START";
    ccOpen.title = "accordion";
    ccOpen.appearance = Word.ContentControlAppearance.hidden;

    // Primeiro item do acordeão
    _insertAccordionItem(context, selection);

    // Marcador de fechamento — inserimos após o selection atual
    const range = selection.getRange("End");
    const ccClose = range.insertContentControl();
    ccClose.tag = "DN-BLOCK-END";
    ccClose.title = "accordion";
    ccClose.appearance = Word.ContentControlAppearance.hidden;

    await context.sync();
    showStatus("Acordeão inserido. Preencha o título e o conteúdo.");
  });
}

function _insertAccordionItem(context, range) {
  // Título do item
  const titlePara = range.insertParagraph("Título do item", "End");
  titlePara.style = "DN-Accordion-Titulo";

  // Conteúdo do item
  const contentPara = range.insertParagraph("Texto do item aqui...", "End");
  contentPara.style = "DN-Accordion-Conteudo";
}

function addAccordionItem() {
  run(async function (context) {
    const selection = context.document.getSelection();
    _insertAccordionItem(context, selection);
    await context.sync();
    showStatus("Novo item de acordeão adicionado.");
  });
}

// ─── Inserir Tabs ────────────────────────────────────────────

function insertTabs() {
  run(async function (context) {
    const selection = context.document.getSelection();

    const ccOpen = selection.insertContentControl();
    ccOpen.tag = "DN-BLOCK-START";
    ccOpen.title = "tabs";
    ccOpen.appearance = Word.ContentControlAppearance.hidden;

    _insertTabItem(context, selection);

    const range = selection.getRange("End");
    const ccClose = range.insertContentControl();
    ccClose.tag = "DN-BLOCK-END";
    ccClose.title = "tabs";
    ccClose.appearance = Word.ContentControlAppearance.hidden;

    await context.sync();
    showStatus("Bloco de Abas inserido. Preencha o título e o conteúdo.");
  });
}

function _insertTabItem(context, range) {
  const titlePara = range.insertParagraph("Título da aba", "End");
  titlePara.style = "DN-Tab-Titulo";

  const contentPara = range.insertParagraph("Conteúdo da aba aqui...", "End");
  contentPara.style = "DN-Tab-Conteudo";
}

function addTabItem() {
  run(async function (context) {
    const selection = context.document.getSelection();
    _insertTabItem(context, selection);
    await context.sync();
    showStatus("Nova aba adicionada.");
  });
}

// ─── Inserir Imagem + Texto ──────────────────────────────────

function insertImgText() {
  run(async function (context) {
    const selection = context.document.getSelection();

    const cc = selection.insertContentControl();
    cc.tag = "DN-BLOCK-START";
    cc.title = "imgText";
    cc.appearance = Word.ContentControlAppearance.hidden;

    // Tabela com 2 colunas: imagem | texto
    const table = selection.insertTable(1, 2, "End", [
      ["[Inserir imagem aqui]", "Texto ao lado da imagem..."],
    ]);
    table.styleBuiltIn = Word.Style.tableGrid;

    const range = selection.getRange("End");
    const ccClose = range.insertContentControl();
    ccClose.tag = "DN-BLOCK-END";
    ccClose.title = "imgText";
    ccClose.appearance = Word.ContentControlAppearance.hidden;

    await context.sync();
    showStatus(
      "Bloco Imagem+Texto inserido. Coluna esquerda = imagem, direita = texto.",
    );
  });
}
