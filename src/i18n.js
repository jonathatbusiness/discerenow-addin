/* global window, document, localStorage */

(function () {
  const STORAGE_KEY = "dn-language";
  const DEFAULT_LANGUAGE = "pt-BR";

  const DICTIONARY = {
    "pt-BR": {
      "ui.languageLabel": "Idioma",
      "ui.structure": "Estrutura",
      "ui.chapter": "Capítulo",
      "ui.lesson": "Lição",
      "ui.mark": "marcar",
      "ui.text": "Texto",
      "ui.paragraph": "Parágrafo",
      "ui.paragraphHeading": "Parágrafo com título",
      "ui.paragraphSubheading": "Parágrafo com subtítulo",
      "ui.heading": "Título",
      "ui.subheading": "Subtítulo",
      "ui.columns": "Colunas",
      "ui.table": "Tabela",
      "ui.list": "Lista",
      "ui.numberedList": "Lista numerada",
      "ui.checkboxList": "Lista de verificação",
      "ui.bulletList": "Lista com marcadores",
      "ui.addListItem": "Adicionar item",
      "ui.callout": "Callout",
      "ui.calloutInfo": "Informação",
      "ui.calloutAlert": "Alerta",
      "ui.calloutTip": "Dica",
      "ui.media": "Mídia",
      "ui.image": "Imagem",
      "ui.imageText": "Imagem + Texto",
      "ui.imageCentered": "Imagem centralizada",
      "ui.video": "Vídeo",
      "ui.interaction": "Interação",
      "ui.accordion": "Acordeão",
      "ui.tabs": "Abas",
      "ui.cards": "Cards",
      "ui.process": "Processo",
      "ui.addProcessStep": "Adicionar passo",
      "ui.flipcard": "FlipCard",
      "ui.assessment": "Avaliação",
      "ui.quiz": "Quiz",
      "ui.single": "Single",
      "ui.multiple": "Multiple",
      "ui.correctAnswer": "Resposta certa",
      "ui.navigation": "Navegação",
      "ui.continueButton": "Botão Continuar",
      "ui.ready": "DiscereNow pronto.",
      "ui.whatsNew": "Novidades! Clique aqui",
      "ui.updateTitle": "Novidades",
      "ui.closeUpdate": "Fechar novidades",
      "ui.addAccordionItem": "Adicionar item de acordeão",
      "ui.addTab": "Adicionar aba",
      "ui.addCard": "Adicionar card",
      "ui.addFlipcard": "Adicionar flipcard",
      "ui.addQuizOption": "Adicionar opção ao quiz",
      "ui.setSingle": "Definir quiz como resposta única",
      "ui.setMultiple": "Definir quiz como múltiplas respostas",
      "ui.markCorrect": "Marcar opção selecionada como resposta correta",

      "word.accordionTitle": "Título do item",
      "word.heading": "Título",
      "word.chapterTitle": "Novo capítulo",
      "word.lessonTitle": "Nova lição",
      "word.subheading": "Subtítulo",
      "word.paragraphContent": "Digite o texto do parágrafo...",
      "word.columnContent": "Digite o conteúdo da coluna...",
      "word.tableHeader": "Cabeçalho",
      "word.tableCell": "Conteúdo",
      "word.listItem1": "Item 1",
      "word.listItem2": "Item 2",
      "word.listNewItem": "Novo item",
      "word.accordionContent": "Conteúdo do item...",
      "word.tabTitle": "Título da aba",
      "word.tabContent": "Conteúdo da aba...",
      "word.cardTitle": "Título do card",
      "word.cardContent": "Conteúdo do card...",
      "word.processStepLabel": "Passo (opcional)",
      "word.processStep": "Passo {number}",
      "word.processTitleLabel": "Título (opcional)",
      "word.processTitle": "Título do passo (opcional)",
      "word.processImageLabel": "Imagem (opcional)",
      "word.processTextLabel": "Texto",
      "word.processText": "Texto do passo...",
      "word.flipFrontTitle": "Frente — título",
      "word.flipFrontContent": "Frente — conteúdo...",
      "word.flipBackTitle": "Verso — título",
      "word.flipBackContent": "Verso — conteúdo...",
      "word.imagePlaceholder":
        '[Insira uma imagem nesta coluna ou deixe apenas "N" caso não queira adicionar imagem]',
      "word.imageHere": "[Inserir imagem aqui]",
      "word.imageSideText": "Texto ao lado da imagem...",
      "word.imageCaptionOptional": "Legenda da imagem (opcional)",
      "word.calloutTitle": "Título do destaque",
      "word.calloutContent": "Conteúdo do destaque...",
      "word.videoCaption": "Legenda do vídeo (opcional)",
      "word.quizTypeLabel": "Tipo do quiz",
      "word.quizTypeHelp":
        'Tipo do quiz — use "single" para resposta única ou "multiple" para múltiplas respostas.',
      "word.quizQuestionLabel": "Pergunta",
      "word.quizQuestion": "Pergunta do quiz?",
      "word.quizOptionLabel": "Opção",
      "word.quizOption1": "Opção 1",
      "word.quizOption2": "Opção 2",
      "word.quizOption3": "Opção 3",
      "word.quizNewOption": "Nova opção",
      "word.quizCorrectFeedbackLabel": "Feedback correto",
      "word.quizIncorrectFeedbackLabel": "Feedback incorreto",
      "word.quizCorrectFeedback": "Resposta correta! Parabéns.",
      "word.quizIncorrectFeedback": "Não foi dessa vez. Tente de novo!",
      "word.continue": "Continuar",

      "status.errorPrefix": "Erro: ",
      "status.cursorIn": "Cursor em: {name}",
      "status.styleApplied": 'Estilo "{styleName}" aplicado.',
      "status.paragraphMarked": "Parágrafo marcado para o DiscereNow.",
      "status.paragraphAlreadyMarked":
        "Este parágrafo já está marcado para o DiscereNow.",
      "status.paragraphInsertedAfterBlock": "Novo parágrafo inserido após o bloco atual.",
      "status.structureInsertedAfterBlock": "{name} inserido após o bloco atual.",
      "status.headingInserted": "Título inserido.",
      "status.subheadingInserted": "Subtítulo inserido.",
      "status.paragraphHeadingInserted": "Parágrafo com título inserido.",
      "status.paragraphSubheadingInserted": "Parágrafo com subtítulo inserido.",
      "status.columnsInserted": "Bloco de colunas inserido.",
      "status.tableInserted": "Tabela inserida.",
      "status.listInserted": "{name} inserida.",
      "status.listItemMissing": "Coloque o cursor dentro de {name} antes de adicionar um item.",
      "status.listItemAdded": "Novo item adicionado.",
      "status.accordionInserted": "Acordeão inserido.",
      "status.accordionItemMissing":
        "Coloque o cursor dentro de um acordeão antes de adicionar um item.",
      "status.accordionItemAdded": "Novo item de acordeão adicionado.",
      "status.tabsInserted": "Bloco de Abas inserido.",
      "status.tabItemMissing":
        "Coloque o cursor dentro de um bloco de Abas antes de adicionar uma aba.",
      "status.tabItemAdded": "Nova aba adicionada.",
      "status.imgTextInserted": "Bloco Imagem+Texto inserido.",
      "status.imageCenteredInserted": "Imagem centralizada inserida.",
      "status.calloutInserted": "Destaque {name} inserido.",
      "status.videoInserted":
        "Vídeo inserido. (1ª linha = URL do YouTube/Vimeo, 2ª = legenda)",
      "status.cardsInserted": "Cards inserido.",
      "status.cardMissing":
        "Coloque o cursor dentro de um bloco Cards antes de adicionar um card.",
      "status.cardAdded": "Novo card adicionado.",
      "status.processInserted": "Processo inserido.",
      "status.processStepMissing": "Coloque o cursor dentro de um Processo antes de adicionar um passo.",
      "status.processStepAdded": "Novo passo adicionado.",
      "status.flipcardInserted": "FlipCard inserido.",
      "status.flipcardMissing":
        "Coloque o cursor dentro de um bloco FlipCard antes de adicionar um card.",
      "status.flipcardAdded": "Novo flipcard adicionado.",
      "status.quizInserted":
        "Quiz inserido. Tipo padrão: single. Use os botões Single, Multiple e Resposta certa para configurar.",
      "status.quizOptionMissing":
        "Coloque o cursor dentro de um Quiz antes de adicionar uma opção.",
      "status.quizOptionAdded": "Nova opção adicionada.",
      "status.quizOptionAddedToQuiz": "Nova opção adicionada ao Quiz.",
      "status.quizTypeMissing":
        "Coloque o cursor dentro de um Quiz antes de alterar o tipo.",
      "status.quizTypeLineNotFound": "Não encontrei a linha de tipo do Quiz.",
      "status.quizSingle":
        "Quiz configurado como Single: apenas uma resposta correta.",
      "status.quizMultiple":
        "Quiz configurado como Multiple: permite mais de uma resposta correta.",
      "status.quizCorrectMissing":
        "Coloque o cursor sobre uma opção dentro de um Quiz.",
      "status.quizSelectOption":
        "Selecione ou posicione o cursor sobre uma opção do Quiz.",
      "status.quizCursorNeedOption":
        "O cursor precisa estar em uma opção do Quiz.",
      "status.quizCorrectSingle":
        "Resposta certa definida. As outras opções foram marcadas como incorretas.",
      "status.quizCorrectMultiple":
        "Resposta certa adicionada. As outras respostas certas foram mantidas.",
      "status.continueInserted": "Botão Continuar inserido.",
    },
    en: {
      "ui.languageLabel": "Language",
      "ui.structure": "Structure",
      "ui.chapter": "Chapter",
      "ui.lesson": "Lesson",
      "ui.mark": "mark",
      "ui.text": "Text",
      "ui.paragraph": "Paragraph",
      "ui.paragraphHeading": "Paragraph with heading",
      "ui.paragraphSubheading": "Paragraph with subheading",
      "ui.heading": "Heading",
      "ui.subheading": "Subheading",
      "ui.columns": "Columns",
      "ui.table": "Table",
      "ui.list": "List",
      "ui.numberedList": "Numbered list",
      "ui.checkboxList": "Checkbox list",
      "ui.bulletList": "Bullet list",
      "ui.addListItem": "Add item",
      "ui.callout": "Callout",
      "ui.calloutInfo": "Information",
      "ui.calloutAlert": "Alert",
      "ui.calloutTip": "Tip",
      "ui.media": "Media",
      "ui.image": "Image",
      "ui.imageText": "Image + Text",
      "ui.imageCentered": "Centered image",
      "ui.video": "Video",
      "ui.interaction": "Interaction",
      "ui.accordion": "Accordion",
      "ui.tabs": "Tabs",
      "ui.cards": "Cards",
      "ui.process": "Process",
      "ui.addProcessStep": "Add step",
      "ui.flipcard": "FlipCard",
      "ui.assessment": "Assessment",
      "ui.quiz": "Quiz",
      "ui.single": "Single",
      "ui.multiple": "Multiple",
      "ui.correctAnswer": "Correct answer",
      "ui.navigation": "Navigation",
      "ui.continueButton": "Continue button",
      "ui.ready": "DiscereNow ready.",
      "ui.whatsNew": "New! Click here",
      "ui.updateTitle": "What's new",
      "ui.closeUpdate": "Close update",
      "ui.addAccordionItem": "Add accordion item",
      "ui.addTab": "Add tab",
      "ui.addCard": "Add card",
      "ui.addFlipcard": "Add flipcard",
      "ui.addQuizOption": "Add quiz option",
      "ui.setSingle": "Set quiz as single answer",
      "ui.setMultiple": "Set quiz as multiple answers",
      "ui.markCorrect": "Mark selected option as correct answer",

      "word.accordionTitle": "Item title",
      "word.heading": "Heading",
      "word.chapterTitle": "New chapter",
      "word.lessonTitle": "New lesson",
      "word.subheading": "Subheading",
      "word.paragraphContent": "Enter paragraph text...",
      "word.columnContent": "Enter column content...",
      "word.tableHeader": "Header",
      "word.tableCell": "Content",
      "word.listItem1": "Item 1",
      "word.listItem2": "Item 2",
      "word.listNewItem": "New item",
      "word.accordionContent": "Item content...",
      "word.tabTitle": "Tab title",
      "word.tabContent": "Tab content...",
      "word.cardTitle": "Card title",
      "word.cardContent": "Card content...",
      "word.processStepLabel": "Step (optional)",
      "word.processStep": "Step {number}",
      "word.processTitleLabel": "Title (optional)",
      "word.processTitle": "Step title (optional)",
      "word.processImageLabel": "Image (optional)",
      "word.processTextLabel": "Text",
      "word.processText": "Step text...",
      "word.flipFrontTitle": "Front — title",
      "word.flipFrontContent": "Front — content...",
      "word.flipBackTitle": "Back — title",
      "word.flipBackContent": "Back — content...",
      "word.imagePlaceholder":
        '[Insert an image in this column or leave only "N" if you do not want to add an image]',
      "word.imageHere": "[Insert image here]",
      "word.imageSideText": "Text beside the image...",
      "word.imageCaptionOptional": "Image caption (optional)",
      "word.calloutTitle": "Callout title",
      "word.calloutContent": "Callout content...",
      "word.videoCaption": "Video caption (optional)",
      "word.quizTypeLabel": "Quiz type",
      "word.quizTypeHelp":
        'Quiz type — use "single" for single answer or "multiple" for multiple answers.',
      "word.quizQuestionLabel": "Question",
      "word.quizQuestion": "Quiz question?",
      "word.quizOptionLabel": "Option",
      "word.quizOption1": "Option 1",
      "word.quizOption2": "Option 2",
      "word.quizOption3": "Option 3",
      "word.quizNewOption": "New option",
      "word.quizCorrectFeedbackLabel": "Correct feedback",
      "word.quizIncorrectFeedbackLabel": "Incorrect feedback",
      "word.quizCorrectFeedback": "Correct answer! Well done.",
      "word.quizIncorrectFeedback": "Not this time. Try again!",
      "word.continue": "Continue",

      "status.errorPrefix": "Error: ",
      "status.cursorIn": "Cursor in: {name}",
      "status.styleApplied": 'Style "{styleName}" applied.',
      "status.paragraphMarked": "Paragraph marked for DiscereNow.",
      "status.paragraphAlreadyMarked":
        "This paragraph is already marked for DiscereNow.",
      "status.paragraphInsertedAfterBlock": "New paragraph inserted after the current block.",
      "status.structureInsertedAfterBlock": "{name} inserted after the current block.",
      "status.headingInserted": "Heading inserted.",
      "status.subheadingInserted": "Subheading inserted.",
      "status.paragraphHeadingInserted": "Paragraph with heading inserted.",
      "status.paragraphSubheadingInserted": "Paragraph with subheading inserted.",
      "status.columnsInserted": "Columns block inserted.",
      "status.tableInserted": "Table inserted.",
      "status.listInserted": "{name} inserted.",
      "status.listItemMissing": "Place the cursor inside {name} before adding an item.",
      "status.listItemAdded": "New item added.",
      "status.accordionInserted": "Accordion inserted.",
      "status.accordionItemMissing":
        "Place the cursor inside an accordion before adding an item.",
      "status.accordionItemAdded": "New accordion item added.",
      "status.tabsInserted": "Tabs block inserted.",
      "status.tabItemMissing":
        "Place the cursor inside a Tabs block before adding a tab.",
      "status.tabItemAdded": "New tab added.",
      "status.imgTextInserted": "Image+Text block inserted.",
      "status.imageCenteredInserted": "Centered image inserted.",
      "status.calloutInserted": "{name} callout inserted.",
      "status.videoInserted":
        "Video inserted. (1st line = YouTube/Vimeo URL, 2nd = caption)",
      "status.cardsInserted": "Cards inserted.",
      "status.cardMissing":
        "Place the cursor inside a Cards block before adding a card.",
      "status.cardAdded": "New card added.",
      "status.processInserted": "Process inserted.",
      "status.processStepMissing": "Place the cursor inside a Process before adding a step.",
      "status.processStepAdded": "New step added.",
      "status.flipcardInserted": "FlipCard inserted.",
      "status.flipcardMissing":
        "Place the cursor inside a FlipCard block before adding a card.",
      "status.flipcardAdded": "New flipcard added.",
      "status.quizInserted":
        "Quiz inserted. Default type: single. Use the Single, Multiple and Correct answer buttons to configure it.",
      "status.quizOptionMissing":
        "Place the cursor inside a Quiz before adding an option.",
      "status.quizOptionAdded": "New option added.",
      "status.quizOptionAddedToQuiz": "New option added to Quiz.",
      "status.quizTypeMissing":
        "Place the cursor inside a Quiz before changing the type.",
      "status.quizTypeLineNotFound": "I could not find the Quiz type line.",
      "status.quizSingle": "Quiz set as Single: only one correct answer.",
      "status.quizMultiple":
        "Quiz set as Multiple: allows more than one correct answer.",
      "status.quizCorrectMissing":
        "Place the cursor on an option inside a Quiz.",
      "status.quizSelectOption": "Select or place the cursor on a Quiz option.",
      "status.quizCursorNeedOption": "The cursor must be on a Quiz option.",
      "status.quizCorrectSingle":
        "Correct answer set. The other options were marked as incorrect.",
      "status.quizCorrectMultiple":
        "Correct answer added. The other correct answers were kept.",
      "status.continueInserted": "Continue button inserted.",
    },
  };

  function getLanguage() {
    const stored = localStorage.getItem(STORAGE_KEY);
    return DICTIONARY[stored] ? stored : DEFAULT_LANGUAGE;
  }

  function setLanguage(language) {
    if (!DICTIONARY[language]) return;
    localStorage.setItem(STORAGE_KEY, language);
    document.documentElement.lang = language;
    applyTranslations();
  }

  function translate(key, params) {
    const language = getLanguage();
    let value =
      (DICTIONARY[language] && DICTIONARY[language][key]) ||
      (DICTIONARY[DEFAULT_LANGUAGE] && DICTIONARY[DEFAULT_LANGUAGE][key]) ||
      key;

    if (params) {
      Object.keys(params).forEach(function (paramKey) {
        value = value.replaceAll("{" + paramKey + "}", params[paramKey]);
      });
    }

    return value;
  }

  function applyTranslations() {
    const language = getLanguage();
    document.documentElement.lang = language;

    document.querySelectorAll("[data-i18n]").forEach(function (el) {
      el.textContent = translate(el.getAttribute("data-i18n"));
    });

    document.querySelectorAll("[data-i18n-title]").forEach(function (el) {
      el.setAttribute("title", translate(el.getAttribute("data-i18n-title")));
    });

    const selector = document.getElementById("dn-language-select");
    if (selector) selector.value = language;
  }

  window.DNI18N = {
    getLanguage: getLanguage,
    setLanguage: setLanguage,
    t: translate,
    applyTranslations: applyTranslations,
  };
})();
