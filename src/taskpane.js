/* global Office, Word */

Office.onReady(function (info) {
  if (info.host === Office.HostType.Word) {
    ensureStyles().then(function () {
      showStatus("DiscereNow pronto.");
    });
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

// ─── Criação automática de estilos DN ───────────────────────

async function ensureStyles() {
  return run(async function (context) {
    const stylesNeeded = [
      {
        name: "DN-Capitulo",
        basedOn: "Heading 1",
        fontSize: 22,
        bold: true,
        color: "1e3c72",
      },
      {
        name: "DN-Licao",
        basedOn: "Heading 2",
        fontSize: 16,
        bold: true,
        color: "2a5298",
      },
      {
        name: "DN-Accordion-Titulo",
        basedOn: "Normal",
        fontSize: 13,
        bold: true,
        color: "333333",
      },
      {
        name: "DN-Accordion-Conteudo",
        basedOn: "Normal",
        fontSize: 12,
        bold: false,
        color: "555555",
      },
      {
        name: "DN-Tab-Titulo",
        basedOn: "Normal",
        fontSize: 13,
        bold: true,
        color: "1e3c72",
      },
      {
        name: "DN-Tab-Conteudo",
        basedOn: "Normal",
        fontSize: 12,
        bold: false,
        color: "555555",
      },
    ];

    for (const s of stylesNeeded) {
      try {
        // Tenta carregar o estilo — se não existir, cria
        const existing = context.document.getStyles();
        existing.load("items/name");
        await context.sync();

        const found = existing.items.find((i) => i.name === s.name);
        if (!found) {
          const newStyle = context.document.addStyle(
            s.name,
            Word.StyleType.paragraph,
          );
          newStyle.font.size = s.fontSize;
          newStyle.font.bold = s.bold;
          newStyle.font.color = s.color;
          await context.sync();
        }
      } catch (e) {
        // Se falhar silenciosamente, não bloqueia o add-in
        console.warn("Estilo não criado:", s.name, e);
      }
    }
  });
}

// ─── Aplicar estilo de parágrafo ────────────────────────────

function applyStyle(styleName) {
  run(async function (context) {
    const selection = context.document.getSelection();
    selection.paragraphs.load("items");
    await context.sync();

    selection.paragraphs.items.forEach(function (p) {
      // Usa styleBuiltIn para reset — funciona em qualquer idioma
      p.styleBuiltIn = Word.BuiltInStyleName.normal;
      p.style = styleName;
    });

    await context.sync();
    showStatus('Estilo "' + styleName + '" aplicado.');
  });
}

// ─── Parágrafo simples ───────────────────────────────────────

function applyNormal() {
  run(async function (context) {
    const selection = context.document.getSelection();
    selection.paragraphs.load("items");
    await context.sync();

    selection.paragraphs.items.forEach(function (p) {
      p.styleBuiltIn = Word.BuiltInStyleName.normal;
    });

    await context.sync();
    showStatus("Estilo Normal aplicado.");
  });
}

// ─── Inserir Acordeão ────────────────────────────────────────

function insertAccordion() {
  run(async function (context) {
    const selection = context.document.getSelection();

    const ccOpen = selection.insertContentControl();
    ccOpen.tag = "DN-BLOCK-START";
    ccOpen.title = "accordion";
    ccOpen.appearance = Word.ContentControlAppearance.hidden;

    _insertAccordionItem(context, selection);

    const range = selection.getRange("End");
    const ccClose = range.insertContentControl();
    ccClose.tag = "DN-BLOCK-END";
    ccClose.title = "accordion";
    ccClose.appearance = Word.ContentControlAppearance.hidden;

    await context.sync();
    showStatus("Acordeão inserido.");
  });
}

function _insertAccordionItem(context, range) {
  const titlePara = range.insertParagraph("Título do item", "End");
  titlePara.style = "DN-Accordion-Titulo";

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
    showStatus("Bloco de Abas inserido.");
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

    const table = selection.insertTable(1, 2, "End", [
      ["[Inserir imagem aqui]", "Texto ao lado da imagem..."],
    ]);
    table.styleBuiltIn = Word.BuiltInStyleName.tableGrid;

    const range = selection.getRange("End");
    const ccClose = range.insertContentControl();
    ccClose.tag = "DN-BLOCK-END";
    ccClose.title = "imgText";
    ccClose.appearance = Word.ContentControlAppearance.hidden;

    await context.sync();
    showStatus("Bloco Imagem+Texto inserido.");
  });
}
