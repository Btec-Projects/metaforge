# Metaforge

A small web tool for editing the hidden details inside Word (.docx) files: author, last modified by, title, company, created and modified dates, and total editing time. It can also swap out a few overused phrases in the text.

It all runs in your browser. Your files never leave your computer.

## How to use it

Open the live page, or open `index.html` in any browser. Drop in a .docx file, change what you want, and click **Save edited copy**. Your original file stays untouched.

## Put your own copy online

1. Make a new GitHub repository and upload `index.html` and this README.
2. Go to **Settings > Pages**.
3. Under **Build and deployment**, choose **Deploy from a branch**, pick `main` and `/ (root)`, then save.
4. After a minute or two the site is live at `https://YOUR-USERNAME.github.io/YOUR-REPO/`.

## How it works

A .docx file is really a zip archive. Metaforge opens it with [JSZip](https://stuk.github.io/jszip/), edits `docProps/core.xml` and `docProps/app.xml`, then zips it back up. The phrase tool edits `word/document.xml`.

## Notes

- Word may overwrite some of these values the next time you save the file.
- Phrases split across bold or italic text are skipped by the phrase tool.

## License

MIT
