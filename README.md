# Google Docs Multilingual Translator

A Google Apps Script that turns one English Google Doc into **seven translated copies** — Spanish, Portuguese (Brazil), French, German, Chinese (Simplified), Japanese and Russian — and exports each one as a **PDF**, while keeping the original formatting.

Built to speed up the publishing of multilingual learning resources: one click in the document menu replaces copying, translating, renaming, footnoting and exporting each language by hand.

## What it does

For each target language, the script:

1. Makes a copy of the active document in the same Drive folder, named `<original name>-<language>` (e.g. `Workbook-fr`).
2. Translates every paragraph, list item and table cell.
3. **Keeps the formatting** of each piece of text — font, size, bold, italic, color, links, alignment, lists and tables.
4. Replaces the footer with a disclaimer in that language, saying the document is a preliminary machine translation.
5. Exports the copy as a PDF in the same folder.

Result for a document called `Workbook`:

```
Workbook            (original)
Workbook-es   + Workbook-es.pdf
Workbook-pt-BR + Workbook-pt-BR.pdf
Workbook-fr   + Workbook-fr.pdf
Workbook-de   + Workbook-de.pdf
Workbook-zh   + Workbook-zh.pdf
Workbook-ja   + Workbook-ja.pdf
Workbook-ru   + Workbook-ru.pdf
```

## How formatting is preserved

Google Docs stores text as *runs*: stretches of text that share the same formatting. The script reads the start index of every run (`getTextAttributeIndices`), translates each run on its own, rebuilds the paragraph from the translated runs, and re-applies the original attributes to the matching new character ranges. Leading and trailing spaces are kept outside the translation so words don't run together.

## Setup

1. Open the English Google Doc you want to translate.
2. Go to **Extensions → Apps Script**.
3. Replace the contents of `Code.gs` with [`src/Code.gs`](src/Code.gs) and save.
4. Reload the document. A **Translate** menu appears.
5. The first run asks you to authorize access to Docs and Drive.

## Usage

From the **Translate** menu:

| Menu item | Languages |
|---|---|
| All languages (7) | ES · PT-BR · FR · DE · ZH · JA · RU |
| European | ES · PT-BR · FR · DE |
| Asian & Russian | ZH · JA · RU |

The two smaller groups exist for long documents, which can hit Apps Script's 6-minute execution limit when all seven languages run at once.

## Customizing

Everything is configured at the top of `Code.gs`:

- `SOURCE_LANGUAGE` — language of the original document (default `en`).
- `LANGUAGES` — one entry per target language with its file suffix, language code and footer text. Add, remove or reorder entries.
- `FOOTER_FONT_SIZE` — size of the disclaimer text.

## Limitations

- Translation uses Google's built-in `LanguageApp` service, which is machine translation and has daily quotas.
- Text is translated run by run, so a sentence that changes formatting mid-way (for example, one bold word) is translated in pieces and may read less naturally. Always review before publishing.
- Headers, footnotes, images and drawings are not translated.

## Tech

Google Apps Script (V8) · DocumentApp · DriveApp · LanguageApp

## License

MIT © Luciana Rodrigues
