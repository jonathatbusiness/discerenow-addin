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

// ─── UI ───────────────────────────────────────────────────────────────

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

// ─── Dispatcher ───────────────────────────────────────────────────────

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
  }
}

function applyStyle(styleName) {
  run(async function (context) {
    const selection = context.document.getSelection();
    selection.paragraphs.load("items");

    await context.sync();

    selection.paragraphs.items.forEach(function (p) {
      p.style = styleName;
    });

    await context.sync();
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
  });
}

// ─── Helpers ──────────────────────────────────────────────────────────

async function getParentCCByTag(context, expectedTag) {
  const sel = context.document.getSelection();
  const cc = sel.parentContentControlOrNullObject;
  cc.load("tag, isNullObject");
  await context.sync();
  if (cc.isNullObject || cc.tag !== expectedTag) return null;
  return cc;
}

async function getSafeBlockInsertionTarget(context) {
  const selection = context.document.getSelection();
  const parentCc = selection.parentContentControlOrNullObject;

  parentCc.load("isNullObject");
  await context.sync();

  if (parentCc.isNullObject) return selection;

  return parentCc.insertParagraph("", "After");
}

// ─── CONSTANTE PADRÃO ─────────────────────────────────────────────────

const IMG_PLACEHOLDER =
  '[Insira uma imagem nesta coluna ou deixe apenas "N" caso não queira adicionar imagem]';

// ─── Accordion ────────────────────────────────────────────────────────

function insertAccordion() {
  run(async function (context) {
    const target = await getSafeBlockInsertionTarget(context);
    const cc = target.insertContentControl();

    cc.tag = "DN-accordion";

    const t = cc.insertParagraph("Título do item", "Start");
    t.style = "DN-Accordion-Titulo";

    const table = cc.insertTable(1, 2, "End", [
      ["Conteúdo do item...", IMG_PLACEHOLDER],
    ]);

    await context.sync();
  });
}

function addAccordionItem() {
  run(async function (context) {
    const cc = await getParentCCByTag(context, "DN-accordion");
    if (!cc) return;

    const t = cc.insertParagraph("Título do item", "End");
    t.style = "DN-Accordion-Titulo";

    cc.insertTable(1, 2, "End", [["Conteúdo do item...", IMG_PLACEHOLDER]]);

    await context.sync();
  });
}

// ─── Tabs ─────────────────────────────────────────────────────────────

function insertTabs() {
  run(async function (context) {
    const target = await getSafeBlockInsertionTarget(context);
    const cc = target.insertContentControl();

    cc.tag = "DN-tabs";

    const t = cc.insertParagraph("Título da aba", "Start");
    t.style = "DN-Tab-Titulo";

    cc.insertTable(1, 2, "End", [["Conteúdo da aba...", IMG_PLACEHOLDER]]);

    await context.sync();
  });
}

function addTabItem() {
  run(async function (context) {
    const cc = await getParentCCByTag(context, "DN-tabs");
    if (!cc) return;

    const t = cc.insertParagraph("Título da aba", "End");
    t.style = "DN-Tab-Titulo";

    cc.insertTable(1, 2, "End", [["Conteúdo da aba...", IMG_PLACEHOLDER]]);

    await context.sync();
  });
}

// ─── Cards ────────────────────────────────────────────────────────────

function insertCards() {
  run(async function (context) {
    const target = await getSafeBlockInsertionTarget(context);
    const cc = target.insertContentControl();

    cc.tag = "DN-cards";

    const t = cc.insertParagraph("Título do card", "Start");
    t.style = "DN-Card-Titulo";

    cc.insertTable(1, 2, "End", [["Conteúdo do card...", IMG_PLACEHOLDER]]);

    await context.sync();
  });
}

function addCardItem() {
  run(async function (context) {
    const cc = await getParentCCByTag(context, "DN-cards");
    if (!cc) return;

    const t = cc.insertParagraph("Título do card", "End");
    t.style = "DN-Card-Titulo";

    cc.insertTable(1, 2, "End", [["Conteúdo do card...", IMG_PLACEHOLDER]]);

    await context.sync();
  });
}

// ─── FlipCard ─────────────────────────────────────────────────────────

function insertFlipCard() {
  run(async function (context) {
    const target = await getSafeBlockInsertionTarget(context);
    const cc = target.insertContentControl();

    cc.tag = "DN-flipcard";

    const ft = cc.insertParagraph("Frente — título", "Start");
    ft.style = "DN-Flip-Frente-Titulo";

    cc.insertTable(1, 2, "End", [["Frente — conteúdo...", IMG_PLACEHOLDER]]);

    const vt = cc.insertParagraph("Verso — título", "End");
    vt.style = "DN-Flip-Verso-Titulo";

    cc.insertTable(1, 2, "End", [["Verso — conteúdo...", IMG_PLACEHOLDER]]);

    await context.sync();
  });
}

function addFlipCardItem() {
  run(async function (context) {
    const cc = await getParentCCByTag(context, "DN-flipcard");
    if (!cc) return;

    const ft = cc.insertParagraph("Frente — título", "End");
    ft.style = "DN-Flip-Frente-Titulo";

    cc.insertTable(1, 2, "End", [["Frente — conteúdo...", IMG_PLACEHOLDER]]);

    const vt = cc.insertParagraph("Verso — título", "End");
    vt.style = "DN-Flip-Verso-Titulo";

    cc.insertTable(1, 2, "End", [["Verso — conteúdo...", IMG_PLACEHOLDER]]);

    await context.sync();
  });
}
