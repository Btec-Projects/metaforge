# Metaforge

Edit the hidden details inside Word (.docx) files: author, last modified by, title, company, created and modified dates, and total editing time. It can also tidy a few wordy phrases in the document text.

Everything runs in your browser. Files are never uploaded.

## Use it

Open the live page (see below) or open `index.html` in a browser, drop in a `.docx` file, change what you need, and click **Save edited copy**. Your original file is not touched.

## Publish on GitHub Pages

1. Create a new repository and upload `index.html` and this `README.md`.
2. Go to **Settings > Pages**.
3. Under **Build and deployment**, choose **Deploy from a branch**, pick `main` and `/ (root)`, then save.
4. After a minute your site is live at `https://YOUR-USERNAME.github.io/YOUR-REPO/`.

## How it works

A .docx file is a zip archive. Metaforge uses [JSZip](https://stuk.github.io/jszip/) to open it, edits `docProps/core.xml` and `docProps/app.xml`, and zips it again. Phrase tidying edits `word/document.xml`.

## Good to know

- Word may overwrite some values the next time the file is saved.
- Phrases split across formatting runs (bold, italics) are skipped by the tidy tool.
- Use it honestly. Don't use it to misrepresent who wrote a document where authorship matters.

## License

MIT
