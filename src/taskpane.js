/* global Office, Word */

// ─── Bootstrap ────────────────────────────────────────────────────────

function dnT(key, params) {
  return window.DNI18N ? window.DNI18N.t(key, params) : key;
}

let dnUpdateInfo = null;

Office.onReady(function (info) {
  if (info.host !== Office.HostType.Word) return;

  if (window.DNI18N) window.DNI18N.applyTranslations();
  loadUpdateInfo();

  ensureStyles().then(function () {
    setStatus(dnT("ui.ready"), "ok");
  });

  attachUiHandlers();
  attachSelectionListener();
});

async function loadUpdateInfo() {
  try {
    const response = await fetch("./update-log.json?v=1.5.2", { cache: "no-store" });
    if (!response.ok) return;

    dnUpdateInfo = await response.json();
    renderUpdateInfo();
  } catch (error) {
    console.warn("Could not load update info:", error);
  }
}

function getUpdateMessage(data) {
  if (!data) return "";
  const language = window.DNI18N ? window.DNI18N.getLanguage() : "pt-BR";
  const translations = data.updateTxtI18n;
  if (translations) {
    return translations[language] || translations["pt-BR"] || translations.en || "";
  }
  return typeof data.updateTxt === "string" ? data.updateTxt : "";
}

function renderUpdateInfo() {
  if (!dnUpdateInfo) return;
  const versionEl = document.getElementById("dn-app-version");
  const badge = document.getElementById("dn-update-badge");
  const date = document.getElementById("dn-update-date");
  const message = document.getElementById("dn-update-message");
  const version = dnUpdateInfo.version || "";

  if (versionEl && version) versionEl.textContent = "v" + version;
  if (date) date.textContent = dnUpdateInfo.date || "";
  if (message) message.textContent = getUpdateMessage(dnUpdateInfo);
  if (badge && version) {
    badge.hidden = localStorage.getItem("dn-update-seen-" + version) === "1";
  }
}

function openUpdatePanel() {
  if (!dnUpdateInfo) return;
  const panel = document.getElementById("dn-update-panel");
  const badge = document.getElementById("dn-update-badge");
  if (panel) panel.hidden = false;
  if (badge) badge.hidden = true;
  if (dnUpdateInfo.version) {
    localStorage.setItem("dn-update-seen-" + dnUpdateInfo.version, "1");
  }
}

function toggleUpdatePanel() {
  const panel = document.getElementById("dn-update-panel");
  if (!panel || panel.hidden) openUpdatePanel();
  else panel.hidden = true;
}

function setResourceSectionOpen(section, shouldOpen, animate) {
  const header = section.querySelector(".dn-section-header");
  const content = section.querySelector(".dn-section-content");
  if (!header || !content) return;

  header.setAttribute("aria-expanded", String(shouldOpen));

  if (!animate) {
    section.classList.toggle("is-open", shouldOpen);
    content.style.height = shouldOpen ? "auto" : "0px";
    return;
  }

  if (shouldOpen) {
    section.classList.add("is-open");
    content.style.height = "0px";
    content.getBoundingClientRect();
    content.style.height = content.scrollHeight + "px";
  } else {
    content.style.height = content.scrollHeight + "px";
    content.getBoundingClientRect();
    section.classList.remove("is-open");
    content.style.height = "0px";
  }
}

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
    setStatus(dnT("status.errorPrefix") + err.message, "error");
    console.error(err);
  });
}

// ─── UI: bind dos cliques ─────────────────────────────────────────────

function attachUiHandlers() {
  const updateBadge = document.getElementById("dn-update-badge");
  const versionButton = document.getElementById("dn-app-version");
  const updateClose = document.getElementById("dn-update-close");
  if (updateBadge) updateBadge.addEventListener("click", openUpdatePanel);
  if (versionButton) versionButton.addEventListener("click", toggleUpdatePanel);
  if (updateClose) {
    updateClose.addEventListener("click", function () {
      const panel = document.getElementById("dn-update-panel");
      if (panel) panel.hidden = true;
    });
  }

  document.querySelectorAll(".dn-section").forEach(function (section) {
    setResourceSectionOpen(section, section.classList.contains("is-open"), false);
    const content = section.querySelector(".dn-section-content");
    if (content) {
      content.addEventListener("transitionend", function (event) {
        if (event.propertyName === "height" && section.classList.contains("is-open")) {
          content.style.height = "auto";
        }
      });
    }
  });

  document.querySelectorAll(".dn-section-header").forEach(function (header) {
    header.addEventListener("click", function () {
      const section = header.closest(".dn-section");
      if (!section) return;

      const willOpen = !section.classList.contains("is-open");

      document.querySelectorAll(".dn-section").forEach(function (otherSection) {
        const shouldOpen = otherSection === section && willOpen;
        setResourceSectionOpen(otherSection, shouldOpen, true);
      });
    });
  });

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

  document.querySelectorAll(".dn-quiz-action").forEach(function (btn) {
    btn.addEventListener("click", function (ev) {
      ev.stopPropagation();
      const action = btn.getAttribute("data-action");
      if (action) handleAction(action);
    });
  });

  const languageSelect = document.getElementById("dn-language-select");
  if (languageSelect && window.DNI18N) {
    languageSelect.value = window.DNI18N.getLanguage();
    languageSelect.addEventListener("change", function () {
      window.DNI18N.setLanguage(languageSelect.value);
      renderUpdateInfo();
      updateContextHighlight();
    });
  }
}

// ─── Despachadores ────────────────────────────────────────────────────

function handleAction(action) {
  switch (action) {
    case "apply-chapter":
      return applyStructuralStyle("DN-Capitulo", "word.chapterTitle");
    case "apply-lesson":
      return applyStructuralStyle("DN-Licao", "word.lessonTitle");
    case "apply-paragraph":
      return markParagraphBlock();
    case "insert-paragraph-heading":
      return insertParagraphWithLead("heading");
    case "insert-paragraph-subheading":
      return insertParagraphWithLead("subheading");
    case "insert-heading":
      return insertTextBlock("heading");
    case "insert-subheading":
      return insertTextBlock("subheading");
    case "insert-columns":
      return insertColumns();
    case "insert-table":
      return insertDataTable();
    case "insert-numbered-list":
      return insertListBlock("numbered");
    case "insert-checkbox-list":
      return insertListBlock("checkbox");
    case "insert-bullet-list":
      return insertListBlock("bullet");
    case "insert-callout":
      return insertCallout();
    case "insert-imgtext":
      return insertImgText();
    case "insert-image-centered":
      return insertImageCentered();
    case "insert-video":
      return insertVideo();
    case "insert-accordion":
      return insertAccordion();
    case "insert-tabs":
      return insertTabs();
    case "insert-cards":
      return insertCards();
    case "insert-process":
      return insertProcess();
    case "insert-flipcard":
      return insertFlipCard();
    case "insert-quiz":
      return insertQuiz();
    case "quiz-type-single":
      return setQuizType("single");
    case "quiz-type-multiple":
      return setQuizType("multiple");
    case "quiz-mark-correct":
      return markQuizCorrectAnswer();
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
    case "process-step":
      return addProcessStep();
    case "flipcard-item":
      return addFlipCardItem();
    case "quiz-item":
      return addQuizOption();
    case "numbered-list-item":
      return addListItem("numbered");
    case "checkbox-list-item":
      return addListItem("checkbox");
    case "bullet-list-item":
      return addListItem("bullet");
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
  Word.run(async function (context) {
    const sel = context.document.getSelection();

    // Improved logic to find the relevant Content Control
    const parentCCs = sel.getContentControls();
    parentCCs.load("items, tag");

    await context.sync();

    let tag = null;

    if (parentCCs.items.length > 0) {
      // If directly inside one, take the closest one
      tag = parentCCs.items[0].tag;
    } else {
      // Fallback to surrounding if not directly inside but selection spans it
      const surrounding = sel.getContentControls({
        selectionMode: "Surrounding",
      });
      surrounding.load("items, tag");
      await context.sync();
      if (surrounding.items.length > 0) {
        tag = surrounding.items[0].tag;
      }
    }

    applyContextualState(tag);
  }).catch(function (error) {
    console.error("Context Update Error:", error);
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
    setStatus(dnT("status.cursorIn", { name: friendlyTagName(tag) }), "info");
  } else {
    setStatus(dnT("ui.ready"), "ok");
  }
}

function friendlyTagName(tag) {
  switch (tag) {
    case "DN-paragraph":
      return dnT("ui.paragraph");
    case "DN-paragraphHeading":
      return dnT("ui.paragraphHeading");
    case "DN-paragraphSubheading":
      return dnT("ui.paragraphSubheading");
    case "DN-heading":
      return dnT("ui.heading");
    case "DN-subheading":
      return dnT("ui.subheading");
    case "DN-columns":
      return dnT("ui.columns");
    case "DN-table":
      return dnT("ui.table");
    case "DN-numberedList":
      return dnT("ui.numberedList");
    case "DN-checkboxList":
      return dnT("ui.checkboxList");
    case "DN-bulletList":
      return dnT("ui.bulletList");
    case "DN-accordion":
      return dnT("ui.accordion");
    case "DN-tabs":
      return dnT("ui.tabs");
    case "DN-imgText":
      return dnT("ui.imageText");
    case "DN-imageCentered":
      return dnT("ui.imageCentered");
    case "DN-cards":
      return dnT("ui.cards");
    case "DN-process":
      return dnT("ui.process");
    case "DN-flipcard":
      return dnT("ui.flipcard");
    case "DN-quiz":
      return dnT("ui.quiz");
    case "DN-callout":
      return dnT("ui.callout");
    case "DN-video":
      return dnT("ui.video");
    case "DN-continue":
      return dnT("ui.continueButton");
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
      // Texto
      { name: "DN-Heading", fontSize: 24, bold: true, color: "1e3c72" },
      { name: "DN-Subheading", fontSize: 18, bold: true, color: "2a5298" },
      { name: "DN-Text-Content", fontSize: 12, bold: false, color: "333333" },
      { name: "DN-Table-Header", fontSize: 12, bold: true, color: "ffffff" },
      { name: "DN-Table-Cell", fontSize: 12, bold: false, color: "333333" },
      { name: "DN-List-Item", fontSize: 12, bold: false, color: "333333" },
      { name: "DN-Image-Legenda", fontSize: 11, bold: false, color: "666666" },
      { name: "DN-Process-Passo", fontSize: 10, bold: true, color: "888888" },
      { name: "DN-Process-Titulo", fontSize: 13, bold: true, color: "1e3c72" },
      { name: "DN-Process-Texto", fontSize: 12, bold: false, color: "333333" },
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

async function getContainingContentControl(context) {
  const selection = context.document.getSelection();
  const directParent = selection.parentContentControlOrNullObject;
  directParent.load("isNullObject, tag");
  await context.sync();
  if (!directParent.isNullObject) return directParent;

  const surrounding = selection.getContentControls({
    selectionMode: "Surrounding",
  });
  surrounding.load("items, tag");
  await context.sync();
  if (surrounding.items[0]) return surrounding.items[0];

  const contained = selection.getContentControls();
  contained.load("items, tag");
  await context.sync();
  return contained.items[0] || null;
}

function applyStructuralStyle(styleName, placeholderKey) {
  run(async function (context) {
    const selection = context.document.getSelection();
    const containingCC = await getContainingContentControl(context);
    let paragraphToSelect = null;

    if (containingCC) {
      const paragraph = containingCC.insertParagraph(
        dnT(placeholderKey),
        "After",
      );
      paragraph.style = styleName;
      paragraphToSelect = paragraph;
    } else {
      selection.paragraphs.load("items");
      await context.sync();
      selection.paragraphs.items.forEach(function (p) {
        p.style = styleName;
      });
    }

    await context.sync();

    if (Office.context.requirements.isSetSupported("WordApi", "1.3")) {
      await rebuildStructuralNumbering(context);
    }

    if (paragraphToSelect) {
      paragraphToSelect.getRange("Content").select("End");
      await context.sync();
      setStatus(
        dnT("status.structureInsertedAfterBlock", {
          name: dnT(placeholderKey),
        }),
        "ok",
      );
    } else {
      setStatus(dnT("status.styleApplied", { styleName: styleName }), "ok");
    }
  });
}

async function rebuildStructuralNumbering(context) {
  const paragraphs = context.document.body.paragraphs;
  paragraphs.load("items/style");
  await context.sync();

  const structuralParagraphs = paragraphs.items.filter(function (paragraph) {
    return paragraph.style === "DN-Capitulo" || paragraph.style === "DN-Licao";
  });

  if (!structuralParagraphs.length) return;

  structuralParagraphs.forEach(function (paragraph) {
    paragraph.detachFromList();
  });
  await context.sync();

  const firstParagraph = structuralParagraphs[0];
  const list = firstParagraph.startNewList();
  list.load("id");
  await context.sync();

  // Level 0: 1.  Level 1: 1.1. The number is list metadata, not title text.
  list.setLevelNumbering(0, Word.ListNumbering.arabic, [0, "."]);
  list.setLevelNumbering(1, Word.ListNumbering.arabic, [0, ".", 1]);
  firstParagraph.listItem.level = firstParagraph.style === "DN-Licao" ? 1 : 0;

  for (let i = 1; i < structuralParagraphs.length; i += 1) {
    const paragraph = structuralParagraphs[i];
    const level = paragraph.style === "DN-Licao" ? 1 : 0;
    paragraph.attachToList(list.id, level);
  }

  await context.sync();
}

function markParagraphBlock() {
  run(async function (context) {
    const selection = context.document.getSelection();

    const existingParagraphCC = await getParentCCByTag(context, "DN-paragraph");
    if (existingParagraphCC) {
      setStatus(dnT("status.paragraphAlreadyMarked"), "info");
      return;
    }

    const containingCC = await getContainingContentControl(context);
    let cc;

    if (containingCC) {
      const paragraph = containingCC.insertParagraph(
        dnT("word.paragraphContent"),
        "After",
      );
      cc = paragraph.getRange().insertContentControl();
    } else {
      selection.load("text");
      await context.sync();
      const isSelectionEmpty = selection.text.length === 0;

      cc = selection.insertContentControl();
      if (isSelectionEmpty) {
        cc.insertText(dnT("word.paragraphContent"), "Replace");
      }
    }
    cc.tag = "DN-paragraph";
    cc.title = dnT("ui.paragraph");
    cc.cannotDelete = false;
    cc.cannotEdit = false;

    await context.sync();
    setStatus(
      dnT(
        containingCC
          ? "status.paragraphInsertedAfterBlock"
          : "status.paragraphMarked",
      ),
      "ok",
    );
    updateContextHighlight();
  });
}

async function createBlockContentControl(context, tag, title) {
  const { target, type, position } = await getSafeBlockInsertionTarget(context);
  let cc;

  if (type === "ContentControl" && position === "After") {
    const paragraphAfter = target.insertParagraph("", "After");
    cc = paragraphAfter.getRange().insertContentControl();
  } else {
    cc = target.insertContentControl();
  }

  cc.tag = tag;
  cc.title = title;
  cc.cannotDelete = false;
  cc.cannotEdit = false;
  return cc;
}

function insertTextBlock(kind) {
  run(async function (context) {
    const isHeading = kind === "heading";
    const cc = await createBlockContentControl(
      context,
      isHeading ? "DN-heading" : "DN-subheading",
      dnT(isHeading ? "ui.heading" : "ui.subheading"),
    );
    const paragraph = cc.insertParagraph(
      dnT(isHeading ? "word.heading" : "word.subheading"),
      "Start",
    );
    paragraph.style = isHeading ? "DN-Heading" : "DN-Subheading";
    await context.sync();
    setStatus(dnT(isHeading ? "status.headingInserted" : "status.subheadingInserted"), "ok");
  });
}

function insertParagraphWithLead(kind) {
  run(async function (context) {
    const isHeading = kind === "heading";
    const cc = await createBlockContentControl(
      context,
      isHeading ? "DN-paragraphHeading" : "DN-paragraphSubheading",
      dnT(isHeading ? "ui.paragraphHeading" : "ui.paragraphSubheading"),
    );
    const lead = cc.insertParagraph(
      dnT(isHeading ? "word.heading" : "word.subheading"),
      "Start",
    );
    lead.style = isHeading ? "DN-Heading" : "DN-Subheading";
    const content = cc.insertParagraph(dnT("word.paragraphContent"), "End");
    content.style = "DN-Text-Content";
    await context.sync();
    setStatus(
      dnT(isHeading ? "status.paragraphHeadingInserted" : "status.paragraphSubheadingInserted"),
      "ok",
    );
  });
}

function insertColumns() {
  run(async function (context) {
    const cc = await createBlockContentControl(context, "DN-columns", dnT("ui.columns"));
    const table = cc.insertTable(1, 2, "Start", [
      [dnT("word.columnContent"), dnT("word.columnContent")],
    ]);
    table.style = "Table Grid";
    table.getCell(0, 0).body.paragraphs.getFirst().style = "DN-Text-Content";
    table.getCell(0, 1).body.paragraphs.getFirst().style = "DN-Text-Content";
    await context.sync();
    setStatus(dnT("status.columnsInserted"), "ok");
  });
}

function insertDataTable() {
  run(async function (context) {
    const cc = await createBlockContentControl(context, "DN-table", dnT("ui.table"));
    const table = cc.insertTable(2, 2, "Start", [
      [dnT("word.tableHeader"), dnT("word.tableHeader")],
      [dnT("word.tableCell"), dnT("word.tableCell")],
    ]);
    table.style = "Table Grid";
    for (let col = 0; col < 2; col++) {
      const header = table.getCell(0, col);
      header.body.paragraphs.getFirst().style = "DN-Table-Header";
      header.shadingColor = "2a5298";
      table.getCell(1, col).body.paragraphs.getFirst().style = "DN-Table-Cell";
    }
    await context.sync();
    setStatus(dnT("status.tableInserted"), "ok");
  });
}

const listBlockConfig = {
  numbered: { tag: "DN-numberedList", uiKey: "ui.numberedList" },
  checkbox: { tag: "DN-checkboxList", uiKey: "ui.checkboxList" },
  bullet: { tag: "DN-bulletList", uiKey: "ui.bulletList" },
};

function insertListBlock(kind) {
  run(async function (context) {
    const config = listBlockConfig[kind];
    const cc = await createBlockContentControl(context, config.tag, dnT(config.uiKey));
    const table = cc.insertTable(2, 1, "Start", [
      [dnT("word.listItem1")],
      [dnT("word.listItem2")],
    ]);
    table.style = "Table Grid";
    table.getCell(0, 0).body.paragraphs.getFirst().style = "DN-List-Item";
    table.getCell(1, 0).body.paragraphs.getFirst().style = "DN-List-Item";
    await context.sync();
    setStatus(dnT("status.listInserted", { name: dnT(config.uiKey) }), "ok");
  });
}

function addListItem(kind) {
  run(async function (context) {
    const config = listBlockConfig[kind];
    const cc = await getParentCCByTag(context, config.tag);
    if (!cc) {
      setStatus(dnT("status.listItemMissing", { name: dnT(config.uiKey) }), "warning");
      return;
    }

    const tables = cc.tables;
    tables.load("items");
    await context.sync();
    if (tables.items.length === 0) return;

    const table = tables.items[0];
    table.load("rowCount");
    await context.sync();
    table.addRows("End", 1, [[dnT("word.listNewItem")]]);
    await context.sync();
    table.load("rowCount");
    await context.sync();
    table.getCell(table.rowCount - 1, 0).body.paragraphs.getFirst().style = "DN-List-Item";
    await context.sync();
    setStatus(dnT("status.listItemAdded"), "ok");
  });
}

// ─── Helper: garante estar dentro de um CC com tag esperada ───────────

async function getParentCCByTag(context, expectedTag) {
  const selection = context.document.getSelection();

  const directParent = selection.parentContentControlOrNullObject;
  directParent.load("isNullObject, tag");
  await context.sync();

  if (!directParent.isNullObject && directParent.tag === expectedTag) {
    return directParent;
  }

  const containedCCs = selection.getContentControls();
  containedCCs.load("items/tag");
  await context.sync();

  for (let i = 0; i < containedCCs.items.length; i++) {
    if (containedCCs.items[i].tag === expectedTag) {
      return containedCCs.items[i];
    }
  }

  const surroundingCCs = selection.getContentControls({
    selectionMode: "Surrounding",
  });
  surroundingCCs.load("items/tag");
  await context.sync();

  for (let i = 0; i < surroundingCCs.items.length; i++) {
    if (surroundingCCs.items[i].tag === expectedTag) {
      return surroundingCCs.items[i];
    }
  }

  return null;
}

// ─── getSafeBlockInsertionTarget ────────────────────

async function getSafeBlockInsertionTarget(context) {
  const selection = context.document.getSelection();
  const parentCc = selection.parentContentControlOrNullObject;

  parentCc.load("isNullObject");
  await context.sync();

  if (!parentCc.isNullObject) {
    return { target: parentCc, type: "ContentControl", position: "After" };
  }

  return { target: selection, type: "Selection", position: "Replace" };
}

// ─── CONSTANTE PADRÃO ────────────────────────────────────────────────

function imagePlaceholder() {
  return dnT("word.imagePlaceholder");
}

// ─── Acordeão ─────────────────────────────────────────────────────────

function insertAccordion() {
  run(async function (context) {
    const { target, type, position } =
      await getSafeBlockInsertionTarget(context);
    let cc;

    if (type === "ContentControl" && position === "After") {
      const paragraphAfter = target.insertParagraph("", "After");
      const insertionRange = paragraphAfter.getRange();
      cc = insertionRange.insertContentControl();
    } else {
      cc = target.insertContentControl();
    }

    cc.tag = "DN-accordion";
    cc.title = dnT("ui.accordion");
    cc.cannotDelete = false;
    cc.cannotEdit = false;

    const t = cc.insertParagraph(dnT("word.accordionTitle"), "Start");
    t.style = "DN-Accordion-Titulo";

    const table = cc.insertTable(1, 2, "End", [
      [dnT("word.accordionContent"), imagePlaceholder()],
    ]);
    table.style = "Table Grid";

    await context.sync();
    setStatus(dnT("status.accordionInserted"), "ok");
  });
}

function addAccordionItem() {
  run(async function (context) {
    const cc = await getParentCCByTag(context, "DN-accordion");
    if (!cc) {
      setStatus(dnT("status.accordionItemMissing"), "warning");
      return;
    }
    const t = cc.insertParagraph(dnT("word.accordionTitle"), "End");
    t.style = "DN-Accordion-Titulo";

    const table = cc.insertTable(1, 2, "End", [
      [dnT("word.accordionContent"), imagePlaceholder()],
    ]);
    table.style = "Table Grid";

    await context.sync();
    setStatus(dnT("status.accordionItemAdded"), "ok");
  });
}

// ─── Tabs ─────────────────────────────────────────────────────────────

function insertTabs() {
  run(async function (context) {
    const { target, type, position } =
      await getSafeBlockInsertionTarget(context);
    let cc;

    if (type === "ContentControl" && position === "After") {
      const paragraphAfter = target.insertParagraph("", "After");
      const insertionRange = paragraphAfter.getRange();
      cc = insertionRange.insertContentControl();
    } else {
      cc = target.insertContentControl();
    }

    cc.tag = "DN-tabs";
    cc.title = dnT("ui.tabs");
    cc.cannotDelete = false;
    cc.cannotEdit = false;

    const t = cc.insertParagraph(dnT("word.tabTitle"), "Start");
    t.style = "DN-Tab-Titulo";

    const table = cc.insertTable(1, 2, "End", [
      [dnT("word.tabContent"), imagePlaceholder()],
    ]);
    table.style = "Table Grid";

    await context.sync();
    setStatus(dnT("status.tabsInserted"), "ok");
  });
}

function addTabItem() {
  run(async function (context) {
    const cc = await getParentCCByTag(context, "DN-tabs");
    if (!cc) {
      setStatus(dnT("status.tabItemMissing"), "warning");
      return;
    }
    const t = cc.insertParagraph(dnT("word.tabTitle"), "End");
    t.style = "DN-Tab-Titulo";

    const table = cc.insertTable(1, 2, "End", [
      [dnT("word.tabContent"), imagePlaceholder()],
    ]);
    table.style = "Table Grid";

    await context.sync();
    setStatus(dnT("status.tabItemAdded"), "ok");
  });
}

// ─── Imagem + Texto ───────────────────────────────────────────────────

function insertImgText() {
  run(async function (context) {
    const { target, type, position } =
      await getSafeBlockInsertionTarget(context);
    let cc;

    if (type === "ContentControl" && position === "After") {
      const paragraphAfter = target.insertParagraph("", "After");
      const insertionRange = paragraphAfter.getRange();
      cc = insertionRange.insertContentControl();
    } else {
      cc = target.insertContentControl();
    }

    cc.tag = "DN-imgText";
    cc.title = dnT("ui.imageText");
    cc.cannotDelete = false;
    cc.cannotEdit = false;

    const table = cc.insertTable(1, 2, "End", [
      [dnT("word.imageHere"), dnT("word.imageSideText")],
    ]);
    table.style = "Table Grid";

    await context.sync();
    setStatus(dnT("status.imgTextInserted"), "ok");
  });
}

function insertImageCentered() {
  run(async function (context) {
    const cc = await createBlockContentControl(
      context,
      "DN-imageCentered",
      dnT("ui.imageCentered"),
    );
    const table = cc.insertTable(2, 1, "Start", [
      [dnT("word.imageHere")],
      [dnT("word.imageCaptionOptional")],
    ]);
    table.style = "Table Grid";
    table.getCell(1, 0).body.paragraphs.getFirst().style =
      "DN-Image-Legenda";
    await context.sync();
    setStatus(dnT("status.imageCenteredInserted"), "ok");
  });
}

// ─── Callout ──────────────────────────────────────────────────────────

function insertCallout() {
  run(async function (context) {
    const { target, type, position } =
      await getSafeBlockInsertionTarget(context);
    let cc;

    if (type === "ContentControl" && position === "After") {
      const paragraphAfter = target.insertParagraph("", "After");
      const insertionRange = paragraphAfter.getRange();
      cc = insertionRange.insertContentControl();
    } else {
      cc = target.insertContentControl();
    }

    cc.tag = "DN-callout";
    cc.title = "Callout";
    cc.cannotDelete = false;
    cc.cannotEdit = false;

    const tipo = cc.insertParagraph("info", "Start");
    tipo.style = "DN-Callout-Tipo";

    const titulo = cc.insertParagraph(dnT("word.calloutTitle"), "End");
    titulo.style = "DN-Callout-Titulo";

    const conteudo = cc.insertParagraph(dnT("word.calloutContent"), "End");
    conteudo.style = "DN-Callout-Conteudo";

    await context.sync();
    setStatus(dnT("status.calloutInserted"), "ok");
  });
}

// ─── Vídeo ────────────────────────────────────────────────────────────

function insertVideo() {
  run(async function (context) {
    const { target, type, position } =
      await getSafeBlockInsertionTarget(context);
    let cc;

    if (type === "ContentControl" && position === "After") {
      const paragraphAfter = target.insertParagraph("", "After");
      const insertionRange = paragraphAfter.getRange();
      cc = insertionRange.insertContentControl();
    } else {
      cc = target.insertContentControl();
    }

    cc.tag = "DN-video";
    cc.title = dnT("ui.video");
    cc.cannotDelete = false;
    cc.cannotEdit = false;

    const url = cc.insertParagraph(
      "https://www.youtube.com/watch?v=...",
      "Start",
    );
    url.style = "DN-Video-Url";

    const legenda = cc.insertParagraph(dnT("word.videoCaption"), "End");
    legenda.style = "DN-Video-Legenda";

    await context.sync();
    setStatus(dnT("status.videoInserted"), "ok");
  });
}

// ─── Cards ────────────────────────────────────────────────────────────

function insertCards() {
  run(async function (context) {
    const { target, type, position } =
      await getSafeBlockInsertionTarget(context);
    let cc;

    if (type === "ContentControl" && position === "After") {
      const paragraphAfter = target.insertParagraph("", "After");
      const insertionRange = paragraphAfter.getRange();
      cc = insertionRange.insertContentControl();
    } else {
      cc = target.insertContentControl();
    }

    cc.tag = "DN-cards";
    cc.title = dnT("ui.cards");
    cc.cannotDelete = false;
    cc.cannotEdit = false;

    const t = cc.insertParagraph(dnT("word.cardTitle"), "Start");
    t.style = "DN-Card-Titulo";

    const table = cc.insertTable(1, 2, "End", [
      [dnT("word.cardContent"), imagePlaceholder()],
    ]);
    table.style = "Table Grid";

    await context.sync();
    setStatus(dnT("status.cardsInserted"), "ok");
  });
}

function addCardItem() {
  run(async function (context) {
    const cc = await getParentCCByTag(context, "DN-cards");
    if (!cc) {
      setStatus(dnT("status.cardMissing"), "warning");
      return;
    }
    const t = cc.insertParagraph(dnT("word.cardTitle"), "End");
    t.style = "DN-Card-Titulo";

    const table = cc.insertTable(1, 2, "End", [
      [dnT("word.cardContent"), imagePlaceholder()],
    ]);
    table.style = "Table Grid";

    await context.sync();
    setStatus(dnT("status.cardAdded"), "ok");
  });
}

function appendProcessStep(cc, stepNumber) {
  const table = cc.insertTable(4, 2, "End", [
    [dnT("word.processStepLabel"), dnT("word.processStep", { number: stepNumber })],
    [dnT("word.processTitleLabel"), dnT("word.processTitle")],
    [dnT("word.processImageLabel"), imagePlaceholder()],
    [dnT("word.processTextLabel"), dnT("word.processText")],
  ]);
  table.style = "Table Grid";
  table.getCell(0, 1).body.paragraphs.getFirst().style = "DN-Process-Passo";
  table.getCell(1, 1).body.paragraphs.getFirst().style = "DN-Process-Titulo";
  table.getCell(3, 1).body.paragraphs.getFirst().style = "DN-Process-Texto";
}

function insertProcess() {
  run(async function (context) {
    const cc = await createBlockContentControl(context, "DN-process", dnT("ui.process"));
    appendProcessStep(cc, 1);
    await context.sync();
    setStatus(dnT("status.processInserted"), "ok");
  });
}

function addProcessStep() {
  run(async function (context) {
    const cc = await getParentCCByTag(context, "DN-process");
    if (!cc) {
      setStatus(dnT("status.processStepMissing"), "warning");
      return;
    }
    const tables = cc.tables;
    tables.load("items");
    await context.sync();
    appendProcessStep(cc, tables.items.length + 1);
    await context.sync();
    setStatus(dnT("status.processStepAdded"), "ok");
  });
}

// ─── FlipCard ─────────────────────────────────────────────────────────

function insertFlipCard() {
  run(async function (context) {
    const { target, type, position } =
      await getSafeBlockInsertionTarget(context);
    let cc;

    if (type === "ContentControl" && position === "After") {
      const paragraphAfter = target.insertParagraph("", "After");
      const insertionRange = paragraphAfter.getRange();
      cc = insertionRange.insertContentControl();
    } else {
      cc = target.insertContentControl();
    }

    cc.tag = "DN-flipcard";
    cc.title = dnT("ui.flipcard");
    cc.cannotDelete = false;
    cc.cannotEdit = false;

    const ft = cc.insertParagraph(dnT("word.flipFrontTitle"), "Start");
    ft.style = "DN-Flip-Frente-Titulo";

    const table1 = cc.insertTable(1, 2, "End", [
      [dnT("word.flipFrontContent"), imagePlaceholder()],
    ]);
    table1.style = "Table Grid";

    const vt = cc.insertParagraph(dnT("word.flipBackTitle"), "End");
    vt.style = "DN-Flip-Verso-Titulo";

    const table2 = cc.insertTable(1, 2, "End", [
      [dnT("word.flipBackContent"), imagePlaceholder()],
    ]);
    table2.style = "Table Grid";

    await context.sync();
    setStatus(dnT("status.flipcardInserted"), "ok");
  });
}

function addFlipCardItem() {
  run(async function (context) {
    const cc = await getParentCCByTag(context, "DN-flipcard");
    if (!cc) {
      setStatus(dnT("status.flipcardMissing"), "warning");
      return;
    }
    const ft = cc.insertParagraph(dnT("word.flipFrontTitle"), "End");
    ft.style = "DN-Flip-Frente-Titulo";

    const table1 = cc.insertTable(1, 2, "End", [
      [dnT("word.flipFrontContent"), imagePlaceholder()],
    ]);
    table1.style = "Table Grid";

    const vt = cc.insertParagraph(dnT("word.flipBackTitle"), "End");
    vt.style = "DN-Flip-Verso-Titulo";

    const table2 = cc.insertTable(1, 2, "End", [
      [dnT("word.flipBackContent"), imagePlaceholder()],
    ]);
    table2.style = "Table Grid";

    await context.sync();
    setStatus(dnT("status.flipcardAdded"), "ok");
  });
}

// ─── Quiz ─────────────────────────────────────────────────────────────

function insertQuiz() {
  run(async function (context) {
    const { target, type, position } =
      await getSafeBlockInsertionTarget(context);
    let cc;

    if (type === "ContentControl" && position === "After") {
      const paragraphAfter = target.insertParagraph("", "After");
      const insertionRange = paragraphAfter.getRange();
      cc = insertionRange.insertContentControl();
    } else {
      cc = target.insertContentControl();
    }

    cc.tag = "DN-quiz";
    cc.title = dnT("ui.quiz");
    cc.cannotDelete = false;
    cc.cannotEdit = false;

    const table = cc.insertTable(7, 2, "End", [
      [dnT("word.quizTypeLabel"), "single"],
      [dnT("word.quizQuestionLabel"), dnT("word.quizQuestion")],
      [dnT("word.quizOptionLabel"), dnT("word.quizOption1")],
      [dnT("word.quizOptionLabel"), dnT("word.quizOption2")],
      [dnT("word.quizOptionLabel"), dnT("word.quizOption3")],
      [dnT("word.quizCorrectFeedbackLabel"), dnT("word.quizCorrectFeedback")],
      [
        dnT("word.quizIncorrectFeedbackLabel"),
        dnT("word.quizIncorrectFeedback"),
      ],
    ]);

    table.style = "Table Grid";

    table.getCell(0, 0).value = dnT("word.quizTypeHelp");
    table.getCell(0, 1).body.paragraphs.getFirst().style = "DN-Quiz-Tipo";

    table.getCell(1, 1).body.paragraphs.getFirst().style = "DN-Quiz-Pergunta";

    table.getCell(2, 1).body.paragraphs.getFirst().style = "DN-Quiz-Opcao";
    table.getCell(3, 1).body.paragraphs.getFirst().style = "DN-Quiz-OpcaoCerta";
    table.getCell(4, 1).body.paragraphs.getFirst().style = "DN-Quiz-Opcao";

    table.getCell(5, 1).body.paragraphs.getFirst().style = "DN-Quiz-FeedbackOk";
    table.getCell(6, 1).body.paragraphs.getFirst().style =
      "DN-Quiz-FeedbackErro";

    await context.sync();
    setStatus(dnT("status.quizInserted"), "ok");
  });
}

function addQuizOption() {
  run(async function (context) {
    const cc = await getParentCCByTag(context, "DN-quiz");
    if (!cc) {
      setStatus(dnT("status.quizOptionMissing"), "warning");
      return;
    }

    const tables = cc.tables;
    tables.load("items");
    await context.sync();

    if (tables.items.length === 0) {
      const op = cc.insertParagraph(dnT("word.quizNewOption"), "End");
      op.style = "DN-Quiz-Opcao";
      await context.sync();
      setStatus(dnT("status.quizOptionAdded"), "ok");
      return;
    }

    const table = tables.items[0];
    const feedbackLabels = [
      dnT("word.quizCorrectFeedbackLabel"),
      window.DNI18N
        ? window.DNI18N.t("word.quizCorrectFeedbackLabel", null)
        : "Feedback correto",
      "Feedback correto",
      "Correct feedback",
    ];

    let feedbackCorrectCell = null;

    for (let i = 0; i < feedbackLabels.length; i++) {
      const foundCell = table.search(feedbackLabels[i]).getFirstOrNullObject();
      foundCell.load("isNullObject");
      await context.sync();

      if (!foundCell.isNullObject) {
        feedbackCorrectCell = foundCell;
        break;
      }
    }

    if (feedbackCorrectCell) {
      const row = feedbackCorrectCell.parentTable.getCell(5, 0);
      row.insertRows("Before", 1, [
        [dnT("word.quizOptionLabel"), dnT("word.quizNewOption")],
      ]);
    } else {
      table
        .getCell(0, 0)
        .insertRows("After", 1, [
          [dnT("word.quizOptionLabel"), dnT("word.quizNewOption")],
        ]);
    }

    await context.sync();

    const rows = table.rows;
    rows.load("items");
    await context.sync();

    const optionRowIndex = Math.max(2, rows.items.length - 3);
    table.getCell(optionRowIndex, 1).body.paragraphs.getFirst().style =
      "DN-Quiz-Opcao";

    await context.sync();
    setStatus(dnT("status.quizOptionAddedToQuiz"), "ok");
  });
}

function setQuizType(quizType) {
  run(async function (context) {
    const cc = await getParentCCByTag(context, "DN-quiz");
    if (!cc) {
      setStatus(dnT("status.quizTypeMissing"), "warning");
      return;
    }

    const paragraphs = cc.paragraphs;
    paragraphs.load("items/style,text");
    await context.sync();

    let typeParagraph = null;

    for (let i = 0; i < paragraphs.items.length; i++) {
      if (paragraphs.items[i].style === "DN-Quiz-Tipo") {
        typeParagraph = paragraphs.items[i];
        break;
      }
    }

    if (!typeParagraph) {
      setStatus(dnT("status.quizTypeLineNotFound"), "warning");
      return;
    }

    typeParagraph.getRange().insertText(quizType, "Replace");

    await context.sync();
    setStatus(
      quizType === "single"
        ? dnT("status.quizSingle")
        : dnT("status.quizMultiple"),
      "ok",
    );
  });
}

function markQuizCorrectAnswer() {
  run(async function (context) {
    const cc = await getParentCCByTag(context, "DN-quiz");
    if (!cc) {
      setStatus(dnT("status.quizCorrectMissing"), "warning");
      return;
    }

    const selection = context.document.getSelection();
    const selectedParagraphs = selection.paragraphs;
    selectedParagraphs.load("items/style,text");

    const quizParagraphs = cc.paragraphs;
    quizParagraphs.load("items/style,text");

    await context.sync();

    if (selectedParagraphs.items.length === 0) {
      setStatus(dnT("status.quizSelectOption"), "warning");
      return;
    }

    const selectedParagraph = selectedParagraphs.items[0];

    if (
      selectedParagraph.style !== "DN-Quiz-Opcao" &&
      selectedParagraph.style !== "DN-Quiz-OpcaoCerta"
    ) {
      setStatus(dnT("status.quizCursorNeedOption"), "warning");
      return;
    }

    let quizType = "single";

    for (let i = 0; i < quizParagraphs.items.length; i++) {
      const p = quizParagraphs.items[i];
      if (p.style === "DN-Quiz-Tipo") {
        const value = (p.text || "").trim().toLowerCase();
        quizType =
          value === "multiple" || value === "multi" ? "multiple" : "single";
        break;
      }
    }

    if (quizType === "single") {
      for (let i = 0; i < quizParagraphs.items.length; i++) {
        const p = quizParagraphs.items[i];
        if (p.style === "DN-Quiz-OpcaoCerta") {
          p.style = "DN-Quiz-Opcao";
        }
      }
    }

    selectedParagraph.style = "DN-Quiz-OpcaoCerta";

    await context.sync();

    setStatus(
      quizType === "single"
        ? dnT("status.quizCorrectSingle")
        : dnT("status.quizCorrectMultiple"),
      "ok",
    );
  });
}

// ─── Botão Continuar ──────────────────────────────────────────────────

function insertContinue() {
  run(async function (context) {
    const { target, type, position } =
      await getSafeBlockInsertionTarget(context);
    let cc;

    if (type === "ContentControl" && position === "After") {
      const paragraphAfter = target.insertParagraph("", "After");
      const insertionRange = paragraphAfter.getRange();
      cc = insertionRange.insertContentControl();
    } else {
      cc = target.insertContentControl();
    }

    cc.tag = "DN-continue";
    cc.title = dnT("ui.continueButton");
    cc.cannotDelete = false;
    cc.cannotEdit = false;

    const txt = cc.insertParagraph(dnT("word.continue"), "Start");
    txt.style = "DN-Continue-Texto";

    await context.sync();
    setStatus(dnT("status.continueInserted"), "ok");
  });
}
