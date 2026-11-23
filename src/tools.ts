import type { ToolCall, ToolDefinition } from "./mistral";
import { getAttachment } from "./attachments";
import {
  addConditionalFormat,
  addNamedRange,
  addWorksheet,
  clearRange,
  createChart,
  createTable,
  deleteAiContent,
  deleteCells,
  deleteWorksheet,
  filterTable,
  findInWorkbook,
  formatRange,
  freezePanes,
  getChartImage,
  getSelection,
  getWorkbookOverview,
  insertCells,
  insertImage,
  insertSheetsFromFile,
  readRange,
  renameWorksheet,
  selectRange,
  sortRange,
  updateChart,
  worksheetStats,
  writeRange,
  type ChartSpec,
  type ConditionalFormatSpec,
  type FormatSpec,
  type SortKey,
  type TableFilter,
} from "./excel";

export const SYSTEM_PROMPT = `You are Mistral, an assistant built into Microsoft Excel. You can read and change the user's workbook with tools.

Reading:
- get_workbook_overview: all sheets with their used range, tables, charts, plus named ranges, the active sheet and the selection. Call it first when you don't yet know the workbook.
- get_selection: address and contents (values, formulas, number formats) of the selected cells. Use it when the user says "this", "these cells", "the selection", "this column", etc.
- read_range: values and formulas of a range. Addresses are A1 references ("B2:F40", "Sales!A:D"), "used" for a sheet's used range, a table name or a named range. Large ranges are truncated — read in blocks if needed.
- find: search all sheets (or one) for text/numbers and get the matching cell addresses.
- view_chart: a rendered picture of a chart, to check how it looks.
- read_attachment: preview the rows of an attached CSV file.

Writing:
- write_range: write a 2-D array of values starting at a cell. Strings starting with "=" are formulas (use English function names and comma separators, e.g. "=SUM(B2:B10)"). null leaves a cell unchanged, "" clears it.
- clear_range, insert_cells, delete_cells (rows "5:7", columns "C:D" or blocks), sort_range.
- format_range: number format, font, fill, alignment, wrap, borders, merge, column width/row height, autofit.
- add_conditional_format: colour scales, data bars, highlight by value or text.
- create_table, filter_table: Excel tables and their filters.
- create_chart, update_chart: charts from a data range (types e.g. ColumnClustered, ColumnStacked, BarClustered, Line, LineMarkers, Pie, Doughnut, Area, XYScatter, Waterfall, Histogram...).
- add_worksheet, rename_worksheet, delete_worksheet, select_range (activate a sheet / jump to cells), freeze_panes, add_named_range.
- import_csv: write an attached CSV into a sheet directly (no need to pass its data through write_range).
- insert_sheets_from_file: copy worksheets from an attached .xlsx into this workbook.
- insert_image: place an attached image on a sheet.
- delete_ai_content: remove the sheets, tables, charts and images you created earlier. The add-in asks the user to confirm in a dialog, so do not ask in chat; if they cancel, accept it.

Working rules:
- Rows and columns are as shown in Excel: row 1 is the first row, column A the first column.
- Look before you write: read the relevant range (or get_selection) first so you know the headers, data types, existing formulas and where free space is. Never overwrite data the user did not ask you to change; put new results in empty cells next to or below the data, or on a new sheet.
- Prefer live formulas over computed constants, so results update when the data changes. Use absolute references ($A$1) where a formula is copied across cells. Write one formula per cell in the 2-D array (adjust references per row yourself).
- Prefer modern functions where helpful (XLOOKUP, FILTER, UNIQUE, SORT, LET) unless the workbook clearly targets old Excel; dynamic array formulas spill, so leave room below/right of them.
- After writing formulas, read the result range back once to check for errors (#VALUE!, #REF!, #NAME?, #DIV/0!, #SPILL!) and fix them.
- Format results sensibly: headers bold, numbers with a number format (e.g. "#,##0.00", "0.0%", "yyyy-mm-dd"), autofit columns. Keep the workbook's existing style.
- Dates in values are Excel serial numbers; use read_range with display_text to see them as shown.
- Charts: pick a type that fits the data (line for time series, column/bar for categories, pie only for few parts of a whole), give a title and axis titles.
- When explaining analysis results, give concrete numbers and the cells they come from.

Behaviour:
- Act directly instead of asking for confirmation, then briefly say what you did and where (sheet and cells).
- You cannot run VBA/macros, create pivot tables, use Power Query or access external data. Say so if asked and offer the closest alternative (e.g. SUMIFS/UNIQUE formulas instead of a pivot table).
- Answer in the user's language, but keep formulas in English function syntax.`;

const sheetProp = { type: "string", description: "Worksheet name. Omit for the active sheet (or when the address contains the sheet, e.g. \"Sales!A1:C9\")." };
const addressProp = {
  type: "string",
  description: 'A1 range ("B2:F40", "A:A", "5:5"), "used" for the used range, a table name or a named range.',
};
const color = { type: "string", description: 'Hex color "#RRGGBB".' };

function tool(name: string, description: string, properties: Record<string, unknown> = {}, required: string[] = []): ToolDefinition {
  return { type: "function", function: { name, description, parameters: { type: "object", properties, required } } };
}

export const TOOLS: ToolDefinition[] = [
  tool("get_workbook_overview", "Lists all worksheets (used range, size, tables, charts), named ranges, the active sheet and the selected range."),
  tool("get_selection", "Returns the address and contents (values, formulas, number formats) of the selected cells."),
  tool(
    "read_range",
    "Reads the values (and formulas) of a range. Rows are labelled with their Excel row number, columns with their letters. Large ranges are truncated.",
    {
      sheet: sheetProp,
      address: addressProp,
      include_formats: { type: "boolean", description: "Also return number formats of non-General cells." },
      display_text: { type: "boolean", description: "Return the text as displayed in Excel (formatted dates, currencies) instead of raw values." },
    },
    ["address"],
  ),
  tool(
    "find",
    "Finds cells containing a text or number in all worksheets (or one). Returns the matching addresses per sheet.",
    {
      text: { type: "string" },
      sheet: { type: "string", description: "Limit the search to this worksheet." },
      complete_match: { type: "boolean", description: "Whole cell must match." },
      match_case: { type: "boolean" },
    },
    ["text"],
  ),
  tool(
    "write_range",
    'Writes a 2-D array of values starting at the top-left cell of address. Strings starting with "=" are formulas. null leaves a cell unchanged, "" clears it.',
    {
      sheet: sheetProp,
      address: { type: "string", description: 'Top-left cell (e.g. "B2"), or a range whose top-left cell is used.' },
      values: {
        type: "array",
        description: 'Rows of cells, e.g. [["Name","Total"],["A","=SUM(B2:B9)"]]. Each row is an array.',
        items: { type: "array", items: { type: ["string", "number", "boolean", "null"] } },
      },
      number_format: { type: "string", description: 'Optional number format for all written cells, e.g. "#,##0.00".' },
    },
    ["address", "values"],
  ),
  tool(
    "clear_range",
    "Clears a range.",
    { sheet: sheetProp, address: addressProp, what: { type: "string", enum: ["all", "contents", "formats"], description: "Default: contents." } },
    ["address"],
  ),
  tool(
    "insert_cells",
    'Inserts cells, whole rows ("5:7") or whole columns ("C:D") at the address, shifting existing cells.',
    { sheet: sheetProp, address: addressProp, shift: { type: "string", enum: ["down", "right"] } },
    ["address"],
  ),
  tool(
    "delete_cells",
    'Deletes cells, whole rows ("5:7") or whole columns ("C:D"), shifting the remaining cells.',
    { sheet: sheetProp, address: addressProp, shift: { type: "string", enum: ["up", "left"] } },
    ["address"],
  ),
  tool(
    "sort_range",
    "Sorts a range (or table) by one or more columns.",
    {
      sheet: sheetProp,
      address: addressProp,
      keys: {
        type: "array",
        items: {
          type: "object",
          properties: { column: { type: "string", description: 'Column letter as shown in Excel, e.g. "C".' }, ascending: { type: "boolean" } },
          required: ["column"],
        },
      },
      has_headers: { type: "boolean", description: "First row is a header and stays on top. Default true." },
    },
    ["address", "keys"],
  ),
  tool(
    "format_range",
    "Formats a range: number format, font, fill, alignment, wrapping, borders, merge, column width, row height, autofit.",
    {
      sheet: sheetProp,
      address: addressProp,
      number_format: { type: "string", description: 'e.g. "#,##0", "#,##0.00 €", "0.0%", "yyyy-mm-dd", "@" (text).' },
      bold: { type: "boolean" },
      italic: { type: "boolean" },
      underline: { type: "boolean" },
      font_color: color,
      font_size: { type: "number" },
      font_name: { type: "string" },
      fill_color: { type: "string", description: 'Background "#RRGGBB", or "none" to remove.' },
      horizontal_alignment: { type: "string", enum: ["General", "Left", "Center", "Right", "Justify", "CenterAcrossSelection"] },
      vertical_alignment: { type: "string", enum: ["Top", "Center", "Bottom"] },
      wrap_text: { type: "boolean" },
      borders: { type: "string", enum: ["all", "outline", "inside", "bottom", "top", "none"] },
      border_color: color,
      border_weight: { type: "string", enum: ["Hairline", "Thin", "Medium", "Thick"] },
      column_width: { type: "number", description: "Column width in points." },
      row_height: { type: "number", description: "Row height in points." },
      autofit: { type: "boolean", description: "Autofit column widths and row heights to the content." },
      merge: { type: "boolean" },
      unmerge: { type: "boolean" },
    },
    ["address"],
  ),
  tool(
    "add_conditional_format",
    "Adds conditional formatting to a range (or clears it).",
    {
      sheet: sheetProp,
      address: addressProp,
      type: { type: "string", enum: ["color_scale", "data_bar", "cell_value", "text_contains", "clear"] },
      operator: {
        type: "string",
        enum: ["Between", "NotBetween", "EqualTo", "NotEqualTo", "GreaterThan", "LessThan", "GreaterThanOrEqual", "LessThanOrEqual"],
        description: "For cell_value.",
      },
      value1: { type: ["number", "string"], description: 'For cell_value: number, text or formula like "=$B$1".' },
      value2: { type: ["number", "string"], description: "Second value for Between / NotBetween." },
      text: { type: "string", description: "For text_contains." },
      fill_color: color,
      font_color: color,
      min_color: color,
      mid_color: { type: "string", description: 'Color scale midpoint "#RRGGBB", or "none" for a 2-color scale.' },
      max_color: color,
      bar_color: color,
    },
    ["address", "type"],
  ),
  tool(
    "create_table",
    "Converts a range into an Excel table (with filter buttons and banded rows).",
    {
      sheet: sheetProp,
      address: addressProp,
      has_headers: { type: "boolean", description: "Default true." },
      name: { type: "string", description: "Table name (letters, digits, underscores)." },
      style: { type: "string", description: 'e.g. "TableStyleMedium2", "TableStyleLight9".' },
    },
    ["address"],
  ),
  tool(
    "filter_table",
    "Filters a table column by values or criteria, or clears filters.",
    {
      table: { type: "string" },
      column: { type: "string", description: "Header name of the column." },
      values: { type: "array", items: { type: "string" }, description: "Show only rows with these values." },
      criteria1: { type: "string", description: 'Custom criterion, e.g. ">100", "=*north*".' },
      criteria2: { type: "string" },
      operator: { type: "string", enum: ["And", "Or"] },
      clear: { type: "boolean", description: "Clear the filter on column, or all filters when no column is given." },
    },
    ["table"],
  ),
  tool(
    "create_chart",
    "Creates a chart from a data range (headers in the first row/column become series names/categories). Placed right of the data unless top_left is given.",
    {
      sheet: sheetProp,
      source: { ...addressProp, description: "Data range including headers." },
      type: {
        type: "string",
        description:
          "Excel chart type, e.g. ColumnClustered, ColumnStacked, ColumnStacked100, BarClustered, BarStacked, Line, LineMarkers, Pie, Doughnut, Area, AreaStacked, XYScatter, XYScatterLines, Bubble, Radar, Waterfall, Histogram, Pareto, Treemap, Funnel.",
      },
      title: { type: "string" },
      series_by: { type: "string", enum: ["Auto", "Columns", "Rows"] },
      target_sheet: { type: "string", description: "Put the chart on this sheet instead." },
      top_left: { type: "string", description: 'Cell for the chart\'s top-left corner, e.g. "H2".' },
      bottom_right: { type: "string", description: 'Cell for the bottom-right corner, e.g. "P20".' },
      legend: { type: "string", enum: ["Right", "Left", "Top", "Bottom", "None"] },
      x_axis_title: { type: "string" },
      y_axis_title: { type: "string" },
      name: { type: "string" },
    },
    ["source"],
  ),
  tool(
    "update_chart",
    "Changes an existing chart (type, data, title, legend, axis titles, position) or deletes it.",
    {
      sheet: sheetProp,
      chart: { type: "string", description: "Chart name (see get_workbook_overview)." },
      type: { type: "string" },
      source: { type: "string", description: "New data range." },
      series_by: { type: "string", enum: ["Auto", "Columns", "Rows"] },
      title: { type: "string" },
      legend: { type: "string", enum: ["Right", "Left", "Top", "Bottom", "None"] },
      x_axis_title: { type: "string" },
      y_axis_title: { type: "string" },
      top_left: { type: "string" },
      bottom_right: { type: "string" },
      delete: { type: "boolean" },
    },
    ["chart"],
  ),
  tool("view_chart", "Renders a chart to an image you can look at.", { sheet: sheetProp, chart: { type: "string" } }, ["chart"]),
  tool(
    "add_worksheet",
    "Adds a new worksheet at the end of the workbook.",
    { name: { type: "string" }, activate: { type: "boolean", description: "Default true." } },
  ),
  tool("rename_worksheet", "Renames a worksheet.", { sheet: { type: "string" }, new_name: { type: "string" } }, ["sheet", "new_name"]),
  tool(
    "delete_worksheet",
    "Deletes a worksheet. The add-in asks the user to confirm in a dialog; do not ask in chat.",
    { sheet: { type: "string" } },
    ["sheet"],
  ),
  tool("select_range", "Activates the sheet and selects a range so the user sees it.", { sheet: sheetProp, address: addressProp }, ["address"]),
  tool(
    "freeze_panes",
    "Freezes the top rows and/or left columns of a sheet (0 and 0 unfreezes).",
    { sheet: sheetProp, rows: { type: "integer" }, columns: { type: "integer" } },
  ),
  tool(
    "add_named_range",
    "Creates a workbook-level name for a range.",
    { sheet: sheetProp, address: addressProp, name: { type: "string" }, comment: { type: "string" } },
    ["address", "name"],
  ),
  tool(
    "read_attachment",
    "Returns the first rows of an attached CSV file.",
    { attachment_id: { type: "string" }, max_rows: { type: "integer", description: "Default 50." } },
    ["attachment_id"],
  ),
  tool(
    "import_csv",
    "Writes all rows of an attached CSV file into a sheet, starting at a cell. Numbers are converted automatically.",
    {
      attachment_id: { type: "string" },
      sheet: { type: "string", description: "Target sheet. Omit to create a new sheet named after the file." },
      cell: { type: "string", description: 'Top-left cell. Default "A1".' },
    },
    ["attachment_id"],
  ),
  tool(
    "insert_sheets_from_file",
    "Copies worksheets from an attached .xlsx file to the end of this workbook.",
    { attachment_id: { type: "string" }, sheets: { type: "array", items: { type: "string" }, description: "Only these sheet names (default: all)." } },
    ["attachment_id"],
  ),
  tool(
    "insert_image",
    "Places an attached image on a sheet with its top-left corner at a cell, keeping its aspect ratio.",
    {
      attachment_id: { type: "string" },
      sheet: sheetProp,
      cell: { type: "string", description: "Top-left cell. Default: the selected cell." },
      width: { type: "number", description: "Width in points (height follows)." },
      alt_text: { type: "string" },
    },
    ["attachment_id"],
  ),
  tool("delete_ai_content", "Removes the worksheets, charts and images the assistant created and converts its tables back to plain ranges."),
];

const LABELS: Record<string, string> = {
  get_workbook_overview: "Reading workbook",
  get_selection: "Reading selection",
  read_range: "Reading cells",
  find: "Searching",
  write_range: "Writing cells",
  clear_range: "Clearing cells",
  insert_cells: "Inserting cells",
  delete_cells: "Deleting cells",
  sort_range: "Sorting",
  format_range: "Formatting",
  add_conditional_format: "Adding conditional format",
  create_table: "Creating table",
  filter_table: "Filtering table",
  create_chart: "Creating chart",
  update_chart: "Updating chart",
  view_chart: "Looking at chart",
  add_worksheet: "Adding worksheet",
  rename_worksheet: "Renaming worksheet",
  delete_worksheet: "Deleting worksheet",
  select_range: "Selecting",
  freeze_panes: "Freezing panes",
  add_named_range: "Adding name",
  read_attachment: "Reading attachment",
  import_csv: "Importing CSV",
  insert_sheets_from_file: "Importing worksheets",
  insert_image: "Inserting image",
  delete_ai_content: "Removing AI content",
};

export function describeToolCall(call: ToolCall): string {
  return LABELS[call.function.name] ?? call.function.name;
}

const plural = (n: number, word: string, many = `${word}s`) => `${n} ${n === 1 ? word : many}`;
const local = (address: string) => address.slice(address.lastIndexOf("!") + 1);
const str = (v: unknown) => (v == null || v === "" ? undefined : String(v));

export interface ToolOutcome {
  result: string;
  summary: string;
  ok: boolean;
  /** Image (data URL) to show the model after the tool results, e.g. a chart picture. */
  image?: string;
}

export interface ConfirmRequest {
  title: string;
  message: string;
  confirmLabel: string;
}

/** Asks the user to confirm a destructive action; resolves true to proceed. */
export type Confirm = (request: ConfirmRequest) => Promise<boolean>;

/**
 * Counts the content the assistant created, asks the user to confirm, then removes it.
 * Used by the delete_ai_content tool and directly (without Mistral) by the Actions menu.
 */
export async function removeAiContent(confirm: Confirm): Promise<ToolOutcome> {
  const found = await deleteAiContent(true);
  const parts = [
    found.sheets ? plural(found.sheets, "worksheet") : "",
    found.charts ? plural(found.charts, "chart") : "",
    found.shapes ? plural(found.shapes, "image") : "",
    found.tables ? plural(found.tables, "table") : "",
  ].filter(Boolean);
  if (!parts.length) {
    return { result: "Nothing created by the assistant was found, so nothing was removed.", summary: "Nothing to remove", ok: true };
  }
  const ok = await confirm({
    title: "Remove AI content?",
    message:
      `This removes what Mistral created: ${parts.join(", ")}. ` +
      "Worksheets are deleted with everything on them; tables are converted back to plain ranges and keep their data. " +
      "Values, formulas and formatting Mistral wrote into cells are not undone.",
    confirmLabel: "Remove",
  });
  if (!ok) {
    return {
      result: "The user cancelled the removal. Nothing was deleted. Do not try again unless the user asks.",
      summary: "Removal cancelled",
      ok: false,
    };
  }
  const done = await deleteAiContent();
  return {
    result: `Removed ${done.sheets} worksheets, ${done.charts} charts and ${done.shapes} images; converted ${done.tables} tables to ranges.`,
    summary: `Removed ${parts.join(", ")}`,
    ok: true,
  };
}

export class ToolRunner {
  constructor(private readonly confirm: Confirm = async () => true) {}

  /** Runs a tool and returns the result for the model plus a short summary for the UI. Errors go back to the model. */
  async run(call: ToolCall): Promise<ToolOutcome> {
    try {
      const args = call.function.arguments ? JSON.parse(call.function.arguments) : {};
      return await this.dispatch(call.function.name, args);
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      return { result: `Error: ${message}`, summary: `Failed: ${message}`, ok: false };
    }
  }

  private async dispatch(name: string, args: Record<string, any>): Promise<ToolOutcome> {
    const sheet = str(args.sheet);
    const address = String(args.address ?? "");
    switch (name) {
      case "get_workbook_overview": {
        const overview = (await getWorkbookOverview()) as { sheets: unknown[] };
        return { result: JSON.stringify(overview), summary: `Read ${plural(overview.sheets.length, "sheet")}`, ok: true };
      }
      case "get_selection": {
        const sel = (await getSelection()) as { selection: string };
        return { result: JSON.stringify(sel), summary: `Read selection ${local(sel.selection)}`, ok: true };
      }
      case "read_range": {
        const data = (await readRange(sheet, address, { formats: Boolean(args.include_formats), displayText: Boolean(args.display_text) })) as {
          sheet: string;
          address: string;
          truncated?: boolean;
        };
        return { result: JSON.stringify(data), summary: `Read ${data.sheet}!${data.address}${data.truncated ? " (truncated)" : ""}`, ok: true };
      }
      case "find": {
        const hits = (await findInWorkbook(String(args.text ?? ""), sheet, Boolean(args.complete_match), Boolean(args.match_case))) as {
          cells: number;
        }[];
        const total = hits.reduce((n, h) => n + h.cells, 0);
        return { result: hits.length ? JSON.stringify(hits) : "No matches.", summary: `Found ${plural(total, "match", "matches")}`, ok: true };
      }
      case "write_range": {
        const res = await writeRange(sheet, address, args.values, str(args.number_format));
        const note = res.overwritten ? ` ${res.overwritten} non-empty cells were overwritten.` : "";
        return { result: `Wrote ${res.address}.${note}`, summary: `Wrote ${res.address}`, ok: true };
      }
      case "clear_range": {
        const what = ["all", "contents", "formats"].includes(args.what) ? args.what : "contents";
        const addr = await clearRange(sheet, address, what);
        return { result: `Cleared ${what} of ${addr}.`, summary: `Cleared ${addr}`, ok: true };
      }
      case "insert_cells": {
        const addr = await insertCells(sheet, address, args.shift === "right" ? "right" : "down");
        return { result: `Inserted ${addr}. Cells below/right have moved.`, summary: `Inserted ${addr}`, ok: true };
      }
      case "delete_cells": {
        const addr = await deleteCells(sheet, address, args.shift === "left" ? "left" : "up");
        return { result: `Deleted ${addr}. Cells below/right have moved.`, summary: `Deleted ${addr}`, ok: true };
      }
      case "sort_range": {
        const addr = await sortRange(sheet, address, (args.keys ?? []) as SortKey[], args.has_headers !== false);
        return { result: `Sorted ${addr}.`, summary: `Sorted ${addr}`, ok: true };
      }
      case "format_range": {
        const done = await formatRange(sheet, address, args as FormatSpec);
        return { result: done.length ? `Formatted ${address}: ${done.join(", ")}.` : "Nothing to change.", summary: `Formatted ${address}`, ok: true };
      }
      case "add_conditional_format": {
        await addConditionalFormat(sheet, address, args as ConditionalFormatSpec);
        return {
          result: args.type === "clear" ? `Cleared conditional formats on ${address}.` : `Added ${args.type} to ${address}.`,
          summary: args.type === "clear" ? `Cleared rules on ${address}` : `Conditional format on ${address}`,
          ok: true,
        };
      }
      case "create_table": {
        const t = await createTable(sheet, address, { hasHeaders: args.has_headers, name: str(args.name), style: str(args.style) });
        return { result: JSON.stringify({ table: t.name, range: t.address }), summary: `Created table ${t.name}`, ok: true };
      }
      case "filter_table": {
        const msg = await filterTable(args as TableFilter);
        return { result: msg, summary: msg.replace(/\.$/, ""), ok: true };
      }
      case "create_chart": {
        const c = await createChart({ ...(args as ChartSpec), sheet });
        return { result: JSON.stringify({ chart: c.name, sheet: c.sheet }), summary: `Created chart on ${c.sheet}`, ok: true };
      }
      case "update_chart": {
        const msg = await updateChart(sheet, String(args.chart), args as Partial<ChartSpec> & { delete?: boolean });
        return { result: msg, summary: msg.replace(/\.$/, ""), ok: true };
      }
      case "view_chart": {
        const image = await getChartImage(sheet, String(args.chart));
        return { result: `The picture of chart ${args.chart} follows in the next message.`, summary: `Looked at ${args.chart}`, ok: true, image };
      }
      case "add_worksheet": {
        const n = await addWorksheet(str(args.name), args.activate !== false);
        return { result: `Added worksheet "${n}".`, summary: `Added sheet ${n}`, ok: true };
      }
      case "rename_worksheet": {
        await renameWorksheet(String(args.sheet), String(args.new_name));
        return { result: `Renamed "${args.sheet}" to "${args.new_name}".`, summary: `Renamed ${args.sheet} → ${args.new_name}`, ok: true };
      }
      case "delete_worksheet": {
        const stats = await worksheetStats(String(args.sheet));
        const ok = await this.confirm({
          title: `Delete worksheet "${stats.name}"?`,
          message: `The sheet ${stats.cells ? `and its ${plural(stats.cells, "used cell")} ` : ""}will be deleted. This cannot be undone with Ctrl+Z.`,
          confirmLabel: "Delete",
        });
        if (!ok) return { result: "The user cancelled. The worksheet was not deleted. Do not try again unless asked.", summary: "Deletion cancelled", ok: false };
        await deleteWorksheet(stats.name);
        return { result: `Deleted worksheet "${stats.name}".`, summary: `Deleted sheet ${stats.name}`, ok: true };
      }
      case "select_range": {
        const addr = await selectRange(sheet, address);
        return { result: `Selected ${addr}.`, summary: `Selected ${addr}`, ok: true };
      }
      case "freeze_panes": {
        const rows = Math.max(0, Number(args.rows) || 0);
        const cols = Math.max(0, Number(args.columns) || 0);
        await freezePanes(sheet, rows, cols);
        const what = rows || cols ? `Froze ${[rows ? plural(rows, "row") : "", cols ? plural(cols, "column") : ""].filter(Boolean).join(" and ")}` : "Unfroze panes";
        return { result: `${what}.`, summary: what, ok: true };
      }
      case "add_named_range": {
        const addr = await addNamedRange(sheet, address, String(args.name), str(args.comment));
        return { result: `Named ${addr} "${args.name}".`, summary: `Named ${args.name}`, ok: true };
      }
      case "read_attachment": {
        const att = getAttachment(String(args.attachment_id));
        if (att.kind !== "csv") throw new Error(`${att.id} is not a CSV file.`);
        const max = Math.max(1, Math.min(500, Number(args.max_rows) || 50));
        const rows = att.rows ?? [];
        return {
          result: JSON.stringify({ total_rows: rows.length, rows: rows.slice(0, max) }),
          summary: `Read ${att.name}`,
          ok: true,
        };
      }
      case "import_csv": {
        const att = getAttachment(String(args.attachment_id));
        if (att.kind !== "csv") throw new Error(`${att.id} is not a CSV file.`);
        if (!att.rows?.length) throw new Error(`${att.name} is empty.`);
        let target = sheet;
        if (!target) target = await addWorksheet(att.name.replace(/\.[^.]+$/, "").replace(/[\\/?*[\]:]/g, "_").slice(0, 31), true);
        // Leading "=" in CSV data must stay text, not become a formula.
        const values = att.rows.map((r) => r.map((c) => (c.startsWith("=") ? `'${c}` : c)));
        const res = await writeRange(target, str(args.cell) ?? "A1", values);
        return { result: `Imported ${att.rows.length} rows into ${res.address}.`, summary: `Imported ${plural(att.rows.length, "row")} into ${target}`, ok: true };
      }
      case "insert_sheets_from_file": {
        const att = getAttachment(String(args.attachment_id));
        if (att.kind !== "xlsx") throw new Error(`${att.id} is not an Excel file.`);
        const count = await insertSheetsFromFile(att.base64, Array.isArray(args.sheets) ? args.sheets.map(String) : undefined);
        return { result: `Inserted ${count} worksheets from ${att.name}. Call get_workbook_overview for their names.`, summary: `Imported ${plural(count, "sheet")} from ${att.name}`, ok: true };
      }
      case "insert_image": {
        const att = getAttachment(String(args.attachment_id));
        if (att.kind !== "image") throw new Error(`${att.id} is not an image.`);
        const id = await insertImage({
          base64: att.base64,
          sheet,
          cell: str(args.cell),
          width: Number(args.width) || Math.min(400, att.width ?? 400),
          name: att.name,
          altText: str(args.alt_text),
        });
        return { result: JSON.stringify({ shape_id: id }), summary: `Inserted ${att.name}`, ok: true };
      }
      case "delete_ai_content":
        return removeAiContent(this.confirm);
      default:
        throw new Error(`Unknown tool ${name}`);
    }
  }
}
