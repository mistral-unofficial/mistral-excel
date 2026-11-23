# Mistral for Excel

An Excel task-pane add-in that brings the [Mistral AI API](https://docs.mistral.ai) into your workbooks. It shares its chat panel, settings and Mistral client with the sibling PowerPoint add-in (`../mistral-ppt`).

## How it works

The pane is a single chat. Mistral is given tools (via [function calling](https://docs.mistral.ai/capabilities/function_calling/)) and decides on its own when to use them:

| Tool | Effect | Min. ExcelApi |
| --- | --- | --- |
| `get_workbook_overview` | Sheets with used range and size, tables, charts, named ranges, active sheet, selection. | 1.7 |
| `get_selection` | Address, values, formulas and number formats of the selected cells. | 1.7 |
| `read_range` | Values and formulas of a range (A1 address, `used`, table or name); capped at 2,500 cells. | 1.7 |
| `find` | Searches all sheets for text/numbers and returns matching addresses. | 1.9 |
| `write_range` | Writes a 2-D array of values/formulas starting at a cell; reports overwritten cells. | 1.7 |
| `clear_range` / `insert_cells` / `delete_cells` / `sort_range` | Clear, insert/delete cells, rows or columns, sort by columns. | 1.7 |
| `format_range` | Number format, font, fill, alignment, wrap, borders, merge, sizes, autofit. | 1.7 |
| `add_conditional_format` | Colour scales, data bars, value and text highlights. | 1.7 |
| `create_table` / `filter_table` | Excel tables and their column filters. | 1.7 |
| `create_chart` / `update_chart` / `view_chart` | Charts from a range; change or delete them; render a chart for Mistral to look at (vision). | 1.7 |
| `add_worksheet` / `rename_worksheet` / `delete_worksheet` | Sheet management (deleting asks for confirmation in a dialog). | 1.7 |
| `select_range` / `freeze_panes` / `add_named_range` | Navigate, freeze header rows/columns, define names. | 1.7 |
| `read_attachment` / `import_csv` | Preview an attached CSV/TSV, or write it straight into a sheet. | 1.7 |
| `insert_sheets_from_file` | Copies worksheets from an attached `.xlsx`. | 1.13 |
| `insert_image` | Places an attached image at a cell. | 1.9 |
| `delete_ai_content` | Removes sheets, charts and images Mistral created and converts its tables back to ranges (tracked in the workbook's add-in settings). | 1.7 |

**Attachments:** use 📎, paste, or drag & drop images, `.csv`/`.tsv`/`.txt` and `.xlsx` files into the chat. Images are also shown to Mistral (vision). Attachments live only for the current chat.

Not available: VBA/macros, pivot tables, Power Query and external data connections (Mistral offers formula-based alternatives such as SUMIFS/UNIQUE). Note that changes made by add-ins are generally **not** undoable with Ctrl+Z / Cmd+Z in Excel.

Examples: *"Add a total row with sums under this table"*, *"Which customer had the highest revenue in Q3?"*, *"Chart revenue per month as a line chart"*, *"Highlight all values above 10,000 in red"*, *"Import the attached CSV into a new sheet and format it as a table"*. The **Actions** menu next to Send holds predefined requests (summarize workbook, explain formulas, analyse selection, find errors, clean up data, chart selection, format as table, remove AI content). Settings (API key, model, temperature, base URL) are behind ⚙; **＋** starts a new chat.

The system prompt and tool definitions live in [`src/tools.ts`](src/tools.ts); Office.js helpers are in [`src/excel.ts`](src/excel.ts). The dev server runs on port **3001**, so it can run alongside the PowerPoint add-in (port 3000).

Calls go directly from the add-in to `https://api.mistral.ai/v1`. The API key is kept in the add-in's `localStorage` on your machine only.

## Requirements

- Node.js 20+
- Excel (Microsoft 365 desktop on Windows/macOS, or Excel on the web) with **ExcelApi 1.7+**. Search and images need 1.9, importing worksheets from .xlsx needs 1.13 (current Excel for Mac, e.g. 16.115, supports all of them).
- A Mistral API key from <https://console.mistral.ai/api-keys>

## Getting started

```bash
npm install
```

```bash
npm run certs
```

Installs and trusts a localhost HTTPS certificate (one-time; asks for your password).

```bash
npm run dev
```

Serves the add-in at `https://localhost:3001`. Then, in a second terminal, sideload it into Excel desktop:

```bash
npm start
```

Open Excel → **Home** tab → **Mistral** button. Paste your API key in the settings (⚙), then **Save**.

Stop sideloading with `npm run stop`.

### Excel on the web

With `npm run dev` running, open a workbook in Excel on the web → **Home → Add-ins → More add-ins → My add-ins → Upload my add-in** and choose `manifest.xml`.

## Scripts

| Script | Purpose |
| --- | --- |
| `npm run dev` | Vite dev server with HTTPS on port 3001 |
| `npm run build` | Production build into `dist/` |
| `npm run typecheck` | TypeScript type-check |
| `npm run validate` | Validate `manifest.xml` |
| `npm run manifest:prod` / `npm run validate:prod` | Generate / validate `manifest.prod.xml` for the hosted add-in |
| `npm start` / `npm run stop` | Sideload / unload in Excel desktop |

## Project layout

```
manifest.xml              Office add-in manifest (Home tab button → task pane)
manifest.prod.xml         Production manifest (GitHub Pages URLs, separate add-in ID)
.github/workflows/        GitHub Pages deployment
scripts/                  Production manifest generator
assets/                   Icons and the install landing page (served at the site root)
src/taskpane/             Task pane HTML, CSS, chat UI and agent loop
src/mistral.ts            Mistral client: chat, SSE streaming with tool calls, model listing
src/tools.ts              System prompt, tool definitions and tool execution
src/excel.ts              Office.js helpers: reading, writing, formatting, tables, charts, sheets, AI-content tracking
src/attachments.ts        Chat attachments (images, CSV/TSV, .xlsx)
src/settings.ts           Settings persistence
```

## Deploying

The add-in is hosted on GitHub Pages: **<https://mistral-unofficial.github.io/mistral-excel/>**

- Every push to `main` runs [`.github/workflows/deploy.yml`](.github/workflows/deploy.yml): typecheck, build, generate the production manifest, publish `dist/` to Pages.
- The site root is a short install page with a **Download manifest** button (`manifest.xml` next to the add-in).
- [`manifest.prod.xml`](manifest.prod.xml) is the same production manifest, checked in for admins. Regenerate it with `npm run manifest:prod` after changing `manifest.xml` (another host: `npm run manifest:prod -- https://your.host/path/`), and check it with `npm run validate:prod`.
- The production manifest has its own add-in ID, so the hosted version and your local dev version (`manifest.xml`, localhost) can be installed side by side.
- Code changes reach users the next time they open the pane. Manifest changes (icons, ribbon) need a higher `<Version>` and a redistribution of the manifest.

### Distributing to colleagues

| Option | How |
| --- | --- |
| **Microsoft 365 admin center** (recommended for teams) | Admin center → Settings → **Integrated apps** → **Upload custom apps** → upload `manifest.prod.xml` and assign users/groups. The button then appears for them automatically on Windows, Mac and web. |
| **Excel on the web** | Home → Add-ins → More add-ins → My add-ins → **Upload my add-in** → the downloaded manifest. |
| **Excel for Mac** | Copy the manifest into `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/` and restart Excel. |
| **Excel for Windows** | Put the manifest in a shared network folder, add it under File → Options → Trust Center → **Trusted Add-in Catalogs**, restart, then Add-ins → Shared Folder. |

> **Note:** Because the key lives in the browser, every user brings their own Mistral key. For an organisation-wide key, put a small backend proxy in front of the API and point **API base URL** at it.
