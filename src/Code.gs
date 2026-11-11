/**
 * Translates the active Google Doc from English into seven languages:
 * Simplified Chinese, Japanese, Russian, Spanish,
 * Portuguese, German, and French.
 *
 * Adds a localized disclaimer to each footer and exports each copy as PDF.
 * All files are saved in the same Drive folder as the original document.
 */
function translateAndExportDocSevenLanguages() {
  const ORIGINAL_DOC_ID = DocumentApp.getActiveDocument().getId();

  const TRANSLATIONS = {
    zh_CN: "zh-CN",
    ja_JP: "ja",
    ru_RU: "ru",
    es_ES: "es",
    pt_BR: "pt",
    de_DE: "de",
    fr_FR: "fr"
  };

  const FOOTER_TEXTS = {
    zh_CN:
      "本转录稿是由人工智能生成的初稿，可能存在不准确之处。本材料归德噶国际（Tergar International）所有，仅供德噶禅修教师培训项目内部使用。严禁未经授权在该项目之外复制、传播或使用本材料。",

    ja_JP:
      "この書き起こしは、人工知能によって生成された暫定版であり、不正確な内容が含まれている可能性があります。本資料はTergar Internationalに帰属し、Tergarの瞑想指導者養成プログラム内での使用のみを目的としています。同プログラム外での無断複製、配布、使用は固く禁じられています。",

    ru_RU:
      "Эта расшифровка представляет собой предварительную версию, созданную искусственным интеллектом, и может содержать неточности. Этот материал является собственностью Tergar International и предназначен исключительно для использования в рамках Программы подготовки преподавателей медитации Тергар. Несанкционированное копирование, распространение или использование этого материала за пределами программы строго запрещено.",

    es_ES:
      "Esta transcripción es un documento preliminar generado por inteligencia artificial y puede contener imprecisiones. Este material es propiedad de Tergar International y está destinado exclusivamente para su uso dentro del Programa de Formación de Profesores de Meditación de Tergar. Queda estrictamente prohibida su copia, distribución o utilización no autorizada fuera del Programa de Formación de Profesores de Meditación de Tergar.",

    pt_BR:
      "Esta transcrição é um documento preliminar gerado por inteligência artificial e pode conter imprecisões. Este material é propriedade da Tergar International e destina-se exclusivamente ao uso no Programa de Formação de Professores de Meditação da Tergar. É estritamente proibida a cópia, distribuição ou utilização não autorizada fora do Programa de Formação de Professores de Meditação da Tergar.",

    de_DE:
      "Dieses Transkript ist eine mithilfe künstlicher Intelligenz erstellte vorläufige Fassung und kann Ungenauigkeiten enthalten. Dieses Material ist Eigentum von Tergar International und ausschließlich zur Verwendung innerhalb des Tergar-Ausbildungsprogramms für Meditationslehrende bestimmt. Die unbefugte Vervielfältigung, Verbreitung oder Verwendung außerhalb dieses Programms ist strengstens untersagt.",

    fr_FR:
      "Cette transcription est un document préliminaire généré par intelligence artificielle et peut contenir des inexactitudes. Ce document est la propriété de Tergar International et est destiné exclusivement à un usage dans le cadre du Programme de formation des enseignants de méditation de Tergar. Toute copie, distribution ou utilisation non autorisée en dehors de ce programme est strictement interdite."
  };

  const originalFile = DriveApp.getFileById(ORIGINAL_DOC_ID);
  const folders = originalFile.getParents();

  if (!folders.hasNext()) {
    throw new Error("The original document is not in any folder.");
  }

  const folder = folders.next();

  /**
   * Translates text children and reapplies each formatting run's attributes,
   * including font, size, bold, italic, color, and links.
   */
  function translateTextChildren_(container, language) {
    const numberOfChildren = container.getNumChildren();

    for (let index = 0; index < numberOfChildren; index++) {
      const child = container.getChild(index);

      if (child.getType() !== DocumentApp.ElementType.TEXT) {
        continue;
      }

      const textElement = child.asText();
      const originalText = textElement.getText();

      if (originalText.trim() === "") {
        continue;
      }

      const attributeIndices = textElement.getTextAttributeIndices();
      const translatedRuns = [];

      for (
        let runIndex = 0;
        runIndex < attributeIndices.length;
        runIndex++
      ) {
        const start = attributeIndices[runIndex];

        const end =
          runIndex + 1 < attributeIndices.length
            ? attributeIndices[runIndex + 1]
            : originalText.length;

        const runText = originalText.substring(start, end);
        const whitespaceMatch = runText.match(/^(\s*)([\s\S]*?)(\s*)$/);

        const leadingWhitespace = whitespaceMatch[1];
        const content = whitespaceMatch[2];
        const trailingWhitespace = whitespaceMatch[3];

        const translatedContent =
          content === ""
            ? ""
            : LanguageApp.translate(content, "en", language);

        translatedRuns.push({
          text: leadingWhitespace + translatedContent + trailingWhitespace,
          attributes: textElement.getAttributes(start)
        });
      }

      const completeTranslation = translatedRuns
        .map(function (run) {
          return run.text;
        })
        .join("");

      textElement.setText(completeTranslation);

      let translatedPosition = 0;

      translatedRuns.forEach(function (run) {
        if (run.text.length > 0) {
          textElement.setAttributes(
            translatedPosition,
            translatedPosition + run.text.length - 1,
            run.attributes
          );
        }

        translatedPosition += run.text.length;
      });
    }
  }

  /**
   * Translates paragraphs, list items, and tables within a container.
   */
  function translateContainer_(container, language) {
    for (let index = 0; index < container.getNumChildren(); index++) {
      const element = container.getChild(index);
      const type = element.getType();

      if (type === DocumentApp.ElementType.PARAGRAPH) {
        translateTextChildren_(element.asParagraph(), language);
      } else if (type === DocumentApp.ElementType.LIST_ITEM) {
        translateTextChildren_(element.asListItem(), language);
      } else if (type === DocumentApp.ElementType.TABLE) {
        translateTable_(element.asTable(), language);
      }
    }
  }

  /**
   * Translates each table cell, including nested tables.
   */
  function translateTable_(table, language) {
    for (let rowIndex = 0; rowIndex < table.getNumRows(); rowIndex++) {
      const row = table.getRow(rowIndex);

      for (let cellIndex = 0; cellIndex < row.getNumCells(); cellIndex++) {
        translateContainer_(row.getCell(cellIndex), language);
      }
    }
  }

  /**
   * Exports a translated Google Doc as PDF to the original folder.
   */
  function exportPdf_(fileId, fileName) {
    const exportUrl =
      "https://docs.google.com/document/d/" + fileId + "/export?format=pdf";

    const response = UrlFetchApp.fetch(exportUrl, {
      headers: {
        Authorization: "Bearer " + ScriptApp.getOAuthToken()
      }
    });

    return folder.createFile(
      response.getBlob().setName(fileName + ".pdf")
    );
  }

  Object.keys(TRANSLATIONS).forEach(function (suffix) {
    const language = TRANSLATIONS[suffix];
    const translatedName = originalFile.getName() + "-" + suffix;

    Logger.log("Starting translation: " + suffix);

    const translatedCopy = originalFile.makeCopy(translatedName, folder);
    const translatedDocument = DocumentApp.openById(translatedCopy.getId());

    translateContainer_(translatedDocument.getBody(), language);

    let footer = translatedDocument.getFooter();

    if (!footer) {
      footer = translatedDocument.addFooter();
    } else {
      footer.clear();
    }

    footer.appendParagraph("");

    const disclaimer = footer.appendParagraph(FOOTER_TEXTS[suffix]);
    disclaimer.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    disclaimer.setFontSize(10);

    footer.appendParagraph("");

    translatedDocument.saveAndClose();

    const pdfFile = exportPdf_(translatedCopy.getId(), translatedName);

    Logger.log("Completed: " + suffix);
    Logger.log("Google Doc: " + translatedCopy.getUrl());
    Logger.log("PDF: " + pdfFile.getUrl());
  });

  Logger.log("All seven translations and PDF exports are complete.");
}
