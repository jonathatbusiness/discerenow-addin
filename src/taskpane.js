/* global Office, Word */

// ─── Bootstrap ────────────────────────────────────────────────────────

Office.onReady(function (info) {
  if (info.host !== Office.HostType.Word) return;

  ensureStyles().then(function () {
    setStatus("DiscereNow pronto.", "ok");
  });

  // Liga os handlers de clique nas linhas de bloco e nos [+]
  attachUiHandlers();

  // Liga o listener de SelectionChanged pra contextualidade
  attachSelectionListener();
});

// ─── Status bar ───────────────────────────────────────────────────────

function setStatus(msg, kind) {
  const text = document.getElementById("dn-status-text");
  const bar = document.getElementById("dn-status-bar");
  if (text) text.textContent = msg;
  if (bar) {
    bar.classList.remove("is-warning", "is-error", "is-info");
    if (kind === "warning") bar.classList.add("is-warning");
    else if (kind === "error") bar.classList.add("is-error");
    else if (kind === "info") bar.classList.add("is-info");
  }
}

function run(fn) {
  return Word.run(fn).catch(function (err) {
    setStatus("Erro: " + err.message, "error");
    console.error(err);
  });
}

// ─── UI: bind dos cliques ─────────────────────────────────────────────

function attachUiHandlers() {
  document.querySelectorAll(".dn-row").forEach(function (row) {
    row.addEventListener("click", function (ev) {
      // se o clique foi no [+], não dispara a ação principal
      if (ev.target.classList.contains("dn-add")) return;
      const action = row.getAttribute("data-action");
      if (action) handleAction(action);
    });
  });

  document.querySelectorAll(".dn-add").forEach(function (btn) {
    btn.addEventListener("click", function (ev) {
      ev.stopPropagation();
      const add = btn.getAttribute("data-add");
      if (add) handleAddItem(add);
    });
  });
}

// ─── Despachador de ações principais ──────────────────────────────────

function handleAction(action) {
  switch (action) {
    // ─ Estrutura ─
    case "apply-chapter":
      return applyStyle("DN-Capitulo");
    case "apply-lesson":
      return applyStyle("DN-Licao");
    // ─ Texto ─
    case "apply-paragraph":
      return applyNormal();
    case "insert-callout":
      return placeholder("Callout");
    // ─ Mídia ─
    case "insert-imgtext":
      return insertImgText();
    case "insert-video":
      return placeholder("Vídeo");
    // ─ Interação ─
    case "insert-accordion":
      return insertAccordion();
    case "insert-tabs":
      return insertTabs();
    case "insert-cards":
      return placeholder("Cards");
    case "insert-flipcard":
      return placeholder("FlipCard");
    // ─ Avaliação ─
    case "insert-quiz":
      return placeholder("Quiz");
    // ─ Navegação ─
    case "insert-continue":
      return placeholder("Botão Continuar");
  }
}

// ─── Despachador de "+ adicionar item" ────────────────────────────────

function handleAddItem(kind) {
  switch (kind) {
    case "accordion-item":
      return addAccordionItem();
    case "tab-item":
      return addTabItem();
    case "card-item":
      return placeholder("Adicionar card");
    case "flipcard-item":
      return placeholder("Adicionar flipcard");
    case "quiz-item":
      return placeholder("Adicionar item de quiz");
  }
}

function placeholder(label) {
  setStatus(label + ": em breve.", "info");
}

// ─── Listener de seleção (contextualidade) ────────────────────────────

function attachSelectionListener() {
  Word.run(function (context) {
    return context.sync().then(function () {
      try {
        Office.context.document.addHandlerAsync(
          Office.EventType.DocumentSelectionChanged,
          updateContextHighlight,
        );
      } catch (e) {
        console.warn("addHandlerAsync falhou:", e);
      }
    });
  });
  // chamada inicial pra já refletir o estado atual
  updateContextHighlight();
}

function updateContextHighlight() {
  Word.run(function (context) {
    const sel = context.document.getSelection();
    const cc = sel.parentContentControlOrNullObject;
    cc.load("tag, isNullObject");
    return context.sync().then(function () {
      const tag = cc.isNullObject ? null : cc.tag;
      applyContextualState(tag);
    });
  }).catch(function () {
    applyContextualState(null);
  });
}

function applyContextualState(tag) {
  document.querySelectorAll(".dn-row").forEach(function (row) {
    const rowTag = row.getAttribute("data-context-tag");
    if (rowTag && rowTag === tag) row.classList.add("is-contextual");
    else row.classList.remove("is-contextual");
  });

  if (tag) {
    const friendly = friendlyTagName(tag);
    setStatus("Cursor em: " + friendly, "info");
  } else {
    setStatus("DiscereNow pronto.", "ok");
  }
}

function friendlyTagName(tag) {
  switch (tag) {
    case "DN-accordion":
      return "Acordeão";
    case "DN-tabs":
      return "Abas";
    case "DN-imgText":
      return "Imagem + Texto";
    case "DN-cards":
      return "Cards";
    case "DN-flipcard":
      return "FlipCard";
    case "DN-quiz":
      return "Quiz";
    default:
      return tag;
  }
}

// ─── Criação automática de estilos DN ─────────────────────────────────

async function ensureStyles() {
  return run(async function (context) {
    const stylesNeeded = [
      { name: "DN-Capitulo", fontSize: 22, bold: true, color: "1e3c72" },
      { name: "DN-Licao", fontSize: 16, bold: true, color: "2a5298" },
      {
        name: "DN-Accordion-Titulo",
        fontSize: 13,
        bold: true,
        color: "333333",
      },
      {
        name: "DN-Accordion-Conteudo",
        fontSize: 12,
        bold: false,
        color: "555555",
      },
      { name: "DN-Tab-Titulo", fontSize: 13, bold: true, color: "1e3c72" },
      { name: "DN-Tab-Conteudo", fontSize: 12, bold: false, color: "555555" },
    ];

    for (const s of stylesNeeded) {
      try {
        let style;
        try {
          style = context.document.getStyles().getByName(s.name);
          style.load("name");
          await context.sync();
        } catch (e) {
          style = context.document.addStyle(s.name, Word.StyleType.paragraph);
          await context.sync();
        }
        style.font.size = s.fontSize;
        style.font.bold = s.bold;
        style.font.color = s.color;
        await context.sync();
      } catch (e) {
        console.warn("Estilo não criado:", s.name, e);
      }
    }
  });
}

// ─── Aplicar estilo de parágrafo ──────────────────────────────────────

function applyStyle(styleName) {
  run(async function (context) {
    const selection = context.document.getSelection();
    selection.paragraphs.load("items");
    await context.sync();
    selection.paragraphs.items.forEach(function (p) {
      p.style = styleName;
    });
    await context.sync();
    setStatus('Estilo "' + styleName + '" aplicado.', "ok");
  });
}

function applyNormal() {
  run(async function (context) {
    const selection = context.document.getSelection();
    selection.paragraphs.load("items");
    await context.sync();
    selection.paragraphs.items.forEach(function (p) {
      p.style = "Normal";
    });
    await context.sync();
    setStatus("Parágrafo normal aplicado.", "ok");
  });
}

// ─── Acordeão ─────────────────────────────────────────────────────────

function insertAccordion() {
  run(async function (context) {
    const selection = context.document.getSelection();
    const cc = selection.insertContentControl();
    cc.tag = "DN-accordion";
    cc.title = "Acordeão";
    cc.cannotDelete = false;
    cc.cannotEdit = false;

    const at = cc.insertParagraph("Título do item", "Start");
    at.style = "DN-Accordion-Titulo";
    const ac = cc.insertParagraph("Texto do item aqui...", "End");
    ac.style = "DN-Accordion-Conteudo";

    await context.sync();
    setStatus("Acordeão inserido.", "ok");
  });
}

function addAccordionItem() {
  run(async function (context) {
    const selection = context.document.getSelection();
    const cc = selection.parentContentControlOrNullObject;
    cc.load("tag, isNullObject");
    await context.sync();

    if (cc.isNullObject) {
      setStatus(
        "Coloque o cursor dentro de um acordeão antes de adicionar um item.",
        "warning",
      );
      return;
    }

    const t = cc.insertParagraph("Título do item", "End");
    t.style = "DN-Accordion-Titulo";
    const c = cc.insertParagraph("Texto do item aqui...", "End");
    c.style = "DN-Accordion-Conteudo";

    await context.sync();
    setStatus("Novo item de acordeão adicionado.", "ok");
  });
}

// ─── Tabs ─────────────────────────────────────────────────────────────

function insertTabs() {
  run(async function (context) {
    const selection = context.document.getSelection();
    const cc = selection.insertContentControl();
    cc.tag = "DN-tabs";
    cc.title = "Abas";
    cc.cannotDelete = false;
    cc.cannotEdit = false;

    const tt = cc.insertParagraph("Título da aba", "Start");
    tt.style = "DN-Tab-Titulo";
    const tc = cc.insertParagraph("Conteúdo da aba aqui...", "End");
    tc.style = "DN-Tab-Conteudo";

    await context.sync();
    setStatus("Bloco de Abas inserido.", "ok");
  });
}

function addTabItem() {
  run(async function (context) {
    const selection = context.document.getSelection();
    const cc = selection.parentContentControlOrNullObject;
    cc.load("tag, isNullObject");
    await context.sync();

    if (cc.isNullObject) {
      setStatus(
        "Coloque o cursor dentro de um bloco de Abas antes de adicionar uma aba.",
        "warning",
      );
      return;
    }

    const t = cc.insertParagraph("Título da aba", "End");
    t.style = "DN-Tab-Titulo";
    const c = cc.insertParagraph("Conteúdo da aba aqui...", "End");
    c.style = "DN-Tab-Conteudo";

    await context.sync();
    setStatus("Nova aba adicionada.", "ok");
  });
}

// ─── Imagem + Texto ───────────────────────────────────────────────────

function insertImgText() {
  run(async function (context) {
    const selection = context.document.getSelection();
    const cc = selection.insertContentControl();
    cc.tag = "DN-imgText";
    cc.title = "Imagem + Texto";
    cc.cannotDelete = false;
    cc.cannotEdit = false;

    const table = cc.insertTable(1, 2, "End", [
      ["[Inserir imagem aqui]", "Texto ao lado da imagem..."],
    ]);
    table.style = "Table Grid";

    await context.sync();
    setStatus("Bloco Imagem+Texto inserido.", "ok");
  });
}
