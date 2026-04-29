/* global Office, Word */

// ─── Bootstrap ────────────────────────────────────────────────────────

Office.onReady(function (info) {
  if (info.host !== Office.HostType.Word) return;

  ensureStyles().then(function () {
    setStatus("DiscereNow pronto.", "ok");
  });

  attachUiHandlers();
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

// ─── Despachadores ────────────────────────────────────────────────────

function handleAction(action) {
  switch (action) {
    case "apply-chapter":
      return applyStyle("DN-Capitulo");
    case "apply-lesson":
      return applyStyle("DN-Licao");
    case "apply-paragraph":
      return applyNormal();
    case "insert-callout":
      return insertCallout();
    case "insert-imgtext":
      return insertImgText();
    case "insert-video":
      return insertVideo();
    case "insert-accordion":
      return insertAccordion();
    case "insert-tabs":
      return insertTabs();
    case "insert-cards":
      return insertCards();
    case "insert-flipcard":
      return insertFlipCard();
    case "insert-quiz":
      return insertQuiz();
    case "insert-continue":
      return insertContinue();
  }
}

function handleAddItem(kind) {
  switch (kind) {
    case "accordion-item":
      return addAccordionItem();
    case "tab-item":
      return addTabItem();
    case "card-item":
      return addCardItem();
    case "flipcard-item":
      return addFlipCardItem();
    case "quiz-item":
      return addQuizOption();
  }
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
    setStatus("Cursor em: " + friendlyTagName(tag), "info");
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
    case "DN-callout":
      return "Callout";
    case "DN-video":
      return "Vídeo";
    case "DN-continue":
      return "Botão Continuar";
    default:
      return tag;
  }
}

// ─── Criação automática de estilos DN ─────────────────────────────────

async function ensureStyles() {
  return run(async function (context) {
    const stylesNeeded = [
      // Estrutura
      { name: "DN-Capitulo", fontSize: 22, bold: true, color: "1e3c72" },
      { name: "DN-Licao", fontSize: 16, bold: true, color: "2a5298" },
      // Acordeão / Abas
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
      // Callout
      { name: "DN-Callout-Tipo", fontSize: 10, bold: true, color: "888888" },
      { name: "DN-Callout-Titulo", fontSize: 13, bold: true, color: "1e3c72" },
      {
        name: "DN-Callout-Conteudo",
        fontSize: 12,
        bold: false,
        color: "333333",
      },
      // Vídeo
      { name: "DN-Video-Url", fontSize: 11, bold: false, color: "1565c0" },
      { name: "DN-Video-Legenda", fontSize: 11, bold: false, color: "666666" },
      // Cards
      { name: "DN-Card-Titulo", fontSize: 13, bold: true, color: "1e3c72" },
      { name: "DN-Card-Conteudo", fontSize: 12, bold: false, color: "555555" },
      // FlipCard
      {
        name: "DN-Flip-Frente-Titulo",
        fontSize: 13,
        bold: true,
        color: "1e3c72",
      },
      {
        name: "DN-Flip-Frente-Conteudo",
        fontSize: 12,
        bold: false,
        color: "555555",
      },
      {
        name: "DN-Flip-Verso-Titulo",
        fontSize: 13,
        bold: true,
        color: "2a5298",
      },
      {
        name: "DN-Flip-Verso-Conteudo",
        fontSize: 12,
        bold: false,
        color: "555555",
      },
      // Quiz
      { name: "DN-Quiz-Tipo", fontSize: 10, bold: true, color: "888888" },
      { name: "DN-Quiz-Pergunta", fontSize: 13, bold: true, color: "1e3c72" },
      { name: "DN-Quiz-Opcao", fontSize: 12, bold: false, color: "333333" },
      { name: "DN-Quiz-OpcaoCerta", fontSize: 12, bold: true, color: "2e7d32" },
      {
        name: "DN-Quiz-FeedbackOk",
        fontSize: 12,
        bold: false,
        color: "2e7d32",
      },
      {
        name: "DN-Quiz-FeedbackErro",
        fontSize: 12,
        bold: false,
        color: "c62828",
      },
      // Continue
      { name: "DN-Continue-Texto", fontSize: 13, bold: true, color: "1e3c72" },
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

// ─── Helper: garante estar dentro de um CC com tag esperada ───────────

async function getParentCCByTag(context, expectedTag) {
  const sel = context.document.getSelection();
  const cc = sel.parentContentControlOrNullObject;
  cc.load("tag, isNullObject");
  await context.sync();
  if (cc.isNullObject || cc.tag !== expectedTag) return null;
  return cc;
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

    const table = cc.insertTable(1, 2, "End", [
      ["[N]", "Texto do item aqui..."],
    ]);
    table.style = "Table Grid";

    await context.sync();
    setStatus("Acordeão inserido.", "ok");
  });
}

function addAccordionItem() {
  run(async function (context) {
    const cc = await getParentCCByTag(context, "DN-accordion");
    if (!cc) {
      setStatus(
        "Coloque o cursor dentro de um acordeão antes de adicionar um item.",
        "warning",
      );
      return;
    }

    const t = cc.insertParagraph("Título do item", "End");
    t.style = "DN-Accordion-Titulo";

    const table = cc.insertTable(1, 2, "End", [
      ["[N]", "Texto do item aqui..."],
    ]);
    table.style = "Table Grid";

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

    const table = cc.insertTable(1, 2, "End", [
      ["[N]", "Conteúdo da aba aqui..."],
    ]);
    table.style = "Table Grid";

    await context.sync();
    setStatus("Bloco de Abas inserido.", "ok");
  });
}

function addTabItem() {
  run(async function (context) {
    const cc = await getParentCCByTag(context, "DN-tabs");
    if (!cc) {
      setStatus(
        "Coloque o cursor dentro de um bloco de Abas antes de adicionar uma aba.",
        "warning",
      );
      return;
    }

    const t = cc.insertParagraph("Título da aba", "End");
    t.style = "DN-Tab-Titulo";

    const table = cc.insertTable(1, 2, "End", [
      ["[N]", "Conteúdo da aba aqui..."],
    ]);
    table.style = "Table Grid";

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

// ─── Callout ──────────────────────────────────────────────────────────

function insertCallout() {
  run(async function (context) {
    const selection = context.document.getSelection();
    const cc = selection.insertContentControl();
    cc.tag = "DN-callout";
    cc.title = "Callout";
    cc.cannotDelete = false;
    cc.cannotEdit = false;

    // Tipo: edite para "info", "alert" ou "tip"
    const tipo = cc.insertParagraph("info", "Start");
    tipo.style = "DN-Callout-Tipo";
    const titulo = cc.insertParagraph("Título do destaque", "End");
    titulo.style = "DN-Callout-Titulo";
    const conteudo = cc.insertParagraph("Conteúdo do destaque...", "End");
    conteudo.style = "DN-Callout-Conteudo";

    await context.sync();
    setStatus(
      'Callout inserido. (1ª linha: troque por "info", "alert" ou "tip")',
      "ok",
    );
  });
}

// ─── Vídeo ────────────────────────────────────────────────────────────

function insertVideo() {
  run(async function (context) {
    const selection = context.document.getSelection();
    const cc = selection.insertContentControl();
    cc.tag = "DN-video";
    cc.title = "Vídeo";
    cc.cannotDelete = false;
    cc.cannotEdit = false;

    const url = cc.insertParagraph(
      "https://www.youtube.com/watch?v=...",
      "Start",
    );
    url.style = "DN-Video-Url";
    const legenda = cc.insertParagraph("Legenda do vídeo (opcional)", "End");
    legenda.style = "DN-Video-Legenda";

    await context.sync();
    setStatus(
      "Vídeo inserido. (1ª linha = URL do YouTube/Vimeo, 2ª = legenda)",
      "ok",
    );
  });
}

// ─── Cards ────────────────────────────────────────────────────────────

function insertCards() {
  run(async function (context) {
    const selection = context.document.getSelection();
    const cc = selection.insertContentControl();
    cc.tag = "DN-cards";
    cc.title = "Cards";
    cc.cannotDelete = false;
    cc.cannotEdit = false;

    const t1 = cc.insertParagraph("Título do card", "Start");
    t1.style = "DN-Card-Titulo";

    const table = cc.insertTable(1, 2, "End", [["[N]", "Conteúdo do card..."]]);
    table.style = "Table Grid";

    await context.sync();
    setStatus("Cards inserido.", "ok");
  });
}

function addCardItem() {
  run(async function (context) {
    const cc = await getParentCCByTag(context, "DN-cards");
    if (!cc) {
      setStatus(
        "Coloque o cursor dentro de um bloco Cards antes de adicionar um card.",
        "warning",
      );
      return;
    }

    const t = cc.insertParagraph("Título do card", "End");
    t.style = "DN-Card-Titulo";

    const table = cc.insertTable(1, 2, "End", [["[N]", "Conteúdo do card..."]]);
    table.style = "Table Grid";

    await context.sync();
    setStatus("Novo card adicionado.", "ok");
  });
}

// ─── FlipCard ─────────────────────────────────────────────────────────

function insertFlipCard() {
  run(async function (context) {
    const selection = context.document.getSelection();
    const cc = selection.insertContentControl();
    cc.tag = "DN-flipcard";
    cc.title = "FlipCard";
    cc.cannotDelete = false;
    cc.cannotEdit = false;

    const ft = cc.insertParagraph("Frente — título", "Start");
    ft.style = "DN-Flip-Frente-Titulo";

    const frontTable = cc.insertTable(1, 2, "End", [
      ["[N]", "Frente — conteúdo..."],
    ]);
    frontTable.style = "Table Grid";

    const vt = cc.insertParagraph("Verso — título", "End");
    vt.style = "DN-Flip-Verso-Titulo";

    const backTable = cc.insertTable(1, 2, "End", [
      ["[N]", "Verso — conteúdo..."],
    ]);
    backTable.style = "Table Grid";

    await context.sync();
    setStatus("FlipCard inserido.", "ok");
  });
}

function addFlipCardItem() {
  run(async function (context) {
    const cc = await getParentCCByTag(context, "DN-flipcard");
    if (!cc) {
      setStatus(
        "Coloque o cursor dentro de um bloco FlipCard antes de adicionar um card.",
        "warning",
      );
      return;
    }

    const ft = cc.insertParagraph("Frente — título", "End");
    ft.style = "DN-Flip-Frente-Titulo";

    const frontTable = cc.insertTable(1, 2, "End", [
      ["[N]", "Frente — conteúdo..."],
    ]);
    frontTable.style = "Table Grid";

    const vt = cc.insertParagraph("Verso — título", "End");
    vt.style = "DN-Flip-Verso-Titulo";

    const backTable = cc.insertTable(1, 2, "End", [
      ["[N]", "Verso — conteúdo..."],
    ]);
    backTable.style = "Table Grid";

    await context.sync();
    setStatus("Novo flipcard adicionado.", "ok");
  });
}

// ─── Quiz ─────────────────────────────────────────────────────────────

function insertQuiz() {
  run(async function (context) {
    const selection = context.document.getSelection();
    const cc = selection.insertContentControl();
    cc.tag = "DN-quiz";
    cc.title = "Quiz";
    cc.cannotDelete = false;
    cc.cannotEdit = false;

    // Tipo: "single" (radio) ou "multi" (checkbox)
    const tipo = cc.insertParagraph("single", "Start");
    tipo.style = "DN-Quiz-Tipo";
    const pergunta = cc.insertParagraph("Pergunta do quiz?", "End");
    pergunta.style = "DN-Quiz-Pergunta";
    const op1 = cc.insertParagraph("Opção 1 (errada)", "End");
    op1.style = "DN-Quiz-Opcao";
    const op2 = cc.insertParagraph("Opção 2 (certa)", "End");
    op2.style = "DN-Quiz-OpcaoCerta";
    const op3 = cc.insertParagraph("Opção 3 (errada)", "End");
    op3.style = "DN-Quiz-Opcao";
    const fOk = cc.insertParagraph("Resposta correta! Parabéns.", "End");
    fOk.style = "DN-Quiz-FeedbackOk";
    const fErr = cc.insertParagraph("Não foi dessa vez. Tente de novo!", "End");
    fErr.style = "DN-Quiz-FeedbackErro";

    await context.sync();
    setStatus(
      'Quiz inserido. (1ª linha: "single" ou "multi"; marque a certa com o estilo "DN-Quiz-OpcaoCerta")',
      "ok",
    );
  });
}

function addQuizOption() {
  run(async function (context) {
    const cc = await getParentCCByTag(context, "DN-quiz");
    if (!cc) {
      setStatus(
        "Coloque o cursor dentro de um Quiz antes de adicionar uma opção.",
        "warning",
      );
      return;
    }
    const op = cc.insertParagraph("Nova opção", "End");
    op.style = "DN-Quiz-Opcao";
    await context.sync();
    setStatus(
      'Nova opção adicionada. (use o estilo "DN-Quiz-OpcaoCerta" se for a correta)',
      "ok",
    );
  });
}

// ─── Botão Continuar ──────────────────────────────────────────────────

function insertContinue() {
  run(async function (context) {
    const selection = context.document.getSelection();
    const cc = selection.insertContentControl();
    cc.tag = "DN-continue";
    cc.title = "Botão Continuar";
    cc.cannotDelete = false;
    cc.cannotEdit = false;

    const txt = cc.insertParagraph("Continuar", "Start");
    txt.style = "DN-Continue-Texto";

    await context.sync();
    setStatus("Botão Continuar inserido.", "ok");
  });
}
