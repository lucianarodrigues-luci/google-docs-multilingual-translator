/**
 * Google Docs Multilingual Translator
 *
 * Creates translated copies of the active Google Doc in seven languages,
 * keeps the original formatting (fonts, sizes, bold, italic, colors, links,
 * lists and tables), adds a localized disclaimer to each footer, and exports
 * every copy as a PDF. All files are saved in the same Drive folder as the
 * original document.
 *
 * Author: Luciana Rodrigues
 * License: MIT
 */

/* -------------------------------------------------------------------------- */
/*  Configuration                                                             */
/* -------------------------------------------------------------------------- */

/** Language of the original document (ISO-639 code used by LanguageApp). */
const SOURCE_LANGUAGE = 'en';

/**
 * Target languages.
 * - suffix: added to the copy's file name (e.g. "Report-fr")
 * - code:   language code passed to LanguageApp.translate
 * - footer: disclaimer added to the footer of that copy
 *
 * Add, remove or reorder entries to change the output.
 */
const LANGUAGES = [
  {
    suffix: 'es',
    name: 'Spanish',
    code: 'es',
    footer: 'Este documento es una traducción automática preliminar y puede contener imprecisiones.'
  },
  {
    suffix: 'pt-BR',
    name: 'Portuguese (Brazil)',
    code: 'pt',
    footer: 'Este documento é uma tradução automática preliminar e pode conter imprecisões.'
  },
  {
    suffix: 'fr',
    name: 'French',
    code: 'fr',
    footer: 'Ce document est une traduction automatique préliminaire et peut contenir des inexactitudes.'
  },
  {
    suffix: 'de',
    name: 'German',
    code: 'de',
    footer: 'Dieses Dokument ist eine vorläufige maschinelle Übersetzung und kann Ungenauigkeiten enthalten.'
  },
  {
    suffix: 'zh',
    name: 'Chinese (Simplified)',
    code: 'zh-CN',
    footer: '本文件为初步机器翻译，可能包含不准确之处。'
  },
  {
    suffix: 'ja',
    name: 'Japanese',
    code: 'ja',
    footer: '本書は暫定的な機械翻訳であり、不正確な内容が含まれている可能性があります。'
  },
  {
    suffix: 'ru',
    name: 'Russian',
    code: 'ru',
    footer: 'Данный документ является предварительным машинным переводом и может содержать неточности.'
  }
];

/** Footer disclaimer font size, in points. */
const FOOTER_FONT_SIZE = 10;

/* -------------------------------------------------------------------------- */
/*  Menu                                                                      */
/* -------------------------------------------------------------------------- */

/** Adds a "Translate" menu to the Google Doc when it opens. */
function onOpen() {
  DocumentApp.getUi()
    .createMenu('Translate')
    .addItem('All languages (7)', 'translateAll')
    .addSeparator()
    .addItem('European: ES · PT-BR · FR · DE', 'translateEuropean')
    .addItem('Asian & Russian: ZH · JA · RU', 'translateAsianRussian')
    .addToUi();
}

/* -------------------------------------------------------------------------- */
/*  Entry points                                                              */
/* -------------------------------------------------------------------------- */

/** Translates the active document into all configured languages. */
function translateAll() {
  translateActiveDocument_(LANGUAGES.map(function (l) { return l.suffix; }));
}

/** Spanish, Portuguese (Brazil), French and German. */
function translateEuropean() {
  translateActiveDocument_(['es', 'pt-BR', 'fr', 'de']);
}

/** Chinese, Japanese and Russian. */
function translateAsianRussian() {
  translateActiveDocument_(['zh', 'ja', 'ru']);
}

/* -------------------------------------------------------------------------- */
/*  Main workflow                                                             */
/* -------------------------------------------------------------------------- */

/**
 * Creates one translated copy (Google Doc + PDF) per selected language.
 * Splitting the work into groups helps stay under Apps Script's
 * 6-minute execution limit for long documents.
 *
 * @param {string[]} suffixes Suffixes of the LANGUAGES entries to process.
 */
function translateActiveDocument_(suffixes) {
  const originalFile = DriveApp.getFileById(DocumentApp.getActiveDocument().getId());
  const parents = originalFile.getParents();

  if (!parents.hasNext()) {
    throw new Error('The original document is not in any Drive folder.');
  }

  const folder = parents.next();
  const languages = LANGUAGES.filter(function (l) {
    return suffixes.indexOf(l.suffix) !== -1;
  });

  languages.forEach(function (language) {
    const copyName = originalFile.getName() + '-' + language.suffix;
    const copy = originalFile.makeCopy(copyName, folder);
    const document = DocumentApp.openById(copy.getId());

    translateBody_(document.getBody(), language.code);
    setFooterDisclaimer_(document, language.footer);
    document.saveAndClose();

    exportPdf_(copy.getId(), copyName, folder);
    Logger.log('Done: ' + language.name + ' → ' + copyName);
  });

  Logger.log('Finished ' + languages.length + ' translation(s).');
}

/* -------------------------------------------------------------------------- */
/*  Translation helpers                                                       */
/* -------------------------------------------------------------------------- */

/** Translates every paragraph, list item and table in the document body. */
function translateBody_(body, lang) {
  for (let i = 0; i < body.getNumChildren(); i++) {
    translateElement_(body.getChild(i), lang);
  }
}

/** Dispatches an element to the right translator based on its type. */
function translateElement_(element, lang) {
  const type = element.getType();

  if (type === DocumentApp.ElementType.PARAGRAPH) {
    translateTextChildren_(element.asParagraph(), lang);
  } else if (type === DocumentApp.ElementType.LIST_ITEM) {
    translateTextChildren_(element.asListItem(), lang);
  } else if (type === DocumentApp.ElementType.TABLE) {
    translateTable_(element.asTable(), lang);
  }
}

/** Translates the paragraphs and list items inside each table cell. */
function translateTable_(table, lang) {
  for (let r = 0; r < table.getNumRows(); r++) {
    const row = table.getRow(r);

    for (let c = 0; c < row.getNumCells(); c++) {
      const cell = row.getCell(c);

      for (let k = 0; k < cell.getNumChildren(); k++) {
        translateElement_(cell.getChild(k), lang);
      }
    }
  }
}

/**
 * Translates the text children of a paragraph or list item while keeping
 * every formatting run (font, size, bold, italic, color, link, etc.).
 *
 * Each run with its own formatting is translated separately, then the
 * original attributes are re-applied to the matching translated text.
 */
function translateTextChildren_(container, lang) {
  for (let i = 0; i < container.getNumChildren(); i++) {
    const child = container.getChild(i);

    if (child.getType() !== DocumentApp.ElementType.TEXT) {
      continue;
    }

    const textElement = child.asText();
    const originalText = textElement.getText();

    if (originalText.trim() === '') {
      continue;
    }

    const starts = textElement.getTextAttributeIndices();
    const runs = [];

    for (let r = 0; r < starts.length; r++) {
      const start = starts[r];
      const end = r + 1 < starts.length ? starts[r + 1] : originalText.length;
      const parts = originalText.substring(start, end).match(/^(\s*)([\s\S]*?)(\s*)$/);
      const content = parts[2];

      const translated = content === ''
        ? ''
        : LanguageApp.translate(content, SOURCE_LANGUAGE, lang);

      runs.push({
        text: parts[1] + translated + parts[3],
        attributes: textElement.getAttributes(start)
      });
    }

    textElement.setText(runs.map(function (run) { return run.text; }).join(''));

    let position = 0;
    runs.forEach(function (run) {
      if (run.text.length > 0) {
        textElement.setAttributes(position, position + run.text.length - 1, run.attributes);
      }
      position += run.text.length;
    });
  }
}

/* -------------------------------------------------------------------------- */
/*  Footer & export                                                           */
/* -------------------------------------------------------------------------- */

/** Replaces the footer content with a centered disclaimer. */
function setFooterDisclaimer_(document, text) {
  let footer = document.getFooter();

  if (!footer) {
    footer = document.addFooter();
  } else {
    footer.clear();
  }

  footer.appendParagraph('');
  footer.appendParagraph(text)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .setFontSize(FOOTER_FONT_SIZE);
  footer.appendParagraph('');
}

/** Saves a PDF version of a Google Doc in the given folder. */
function exportPdf_(fileId, fileName, folder) {
  const pdf = DriveApp.getFileById(fileId).getAs(MimeType.PDF).setName(fileName + '.pdf');
  folder.createFile(pdf);
}
