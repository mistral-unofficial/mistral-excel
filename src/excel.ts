/** Office.js helpers for Excel: reading, writing, formatting, tables, charts, sheets and AI-content tracking. */

/** Upper bound on cells returned to the model by one read. */
export const MAX_READ_CELLS = 2500;
const MAX_READ_COLUMNS = 60;

// ---------------------------------------------------------------------------
// Requirements and addressing
// ---------------------------------------------------------------------------

/** Throws a readable error when the running Excel does not support the given ExcelApi version. */
export function requireApi(version: string, feature: string): void {
  if (typeof Office === "undefined" || !Office.context?.requirements?.isSetSupported("ExcelApi", version)) {
    throw new Error(`${feature} needs ExcelApi ${version}, which this Excel version does not support.`);
  }
}

export function columnToNumber(col: string): number {
  let n = 0;
  for (const ch of col.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

export function numberToColumn(n: number): string {
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** Splits "Sheet 1!A1:B2" / "'My sheet'!A1" into sheet and local address. */
export function splitAddress(address: string): { sheet?: string; local: string } {
  const i = address.lastIndexOf("!");
  if (i < 0) return { local: address };
  return { sheet: address.slice(0, i).replace(/^'(.*)'$/, "$1").replace(/''/g, "'"), local: address.slice(i + 1) };
}

/** First cell (column number, row number) of a local address like "B3:D9". Whole columns start at row 1, whole rows at column A. */
function startOf(local: string): { col: number; row: number } {
  const first = local.split(":")[0].replace(/\$/g, "");
  const m = /^([A-Za-z]*)(\d*)$/.exec(first);
  return { col: m?.[1] ? columnToNumber(m[1]) : 1, row: m?.[2] ? Number(m[2]) : 1 };
}

const cellName = (col: number, row: number) => `${numberToColumn(col)}${row}`;

function sheetOf(ctx: Excel.RequestContext, name?: string): Excel.Worksheet {
  return name ? ctx.workbook.worksheets.getItem(name) : ctx.workbook.worksheets.getActiveWorksheet();
}

/**
 * Resolves a range from an optional sheet name and an address. The address may contain its own sheet
 * ("Sales!A1:C9"), be "used" for the sheet's used range, or be the name of a table or named range.
 */
function rangeOf(ctx: Excel.RequestContext, sheet: string | undefined, address: string): Excel.Range {
  const trimmed = String(address ?? "").trim();
  if (!trimmed) throw new Error("No range address given.");
  const { sheet: inAddress, local } = splitAddress(trimmed);
  const ws = sheetOf(ctx, inAddress ?? sheet);
  if (local.toLowerCase() === "used") return ws.getUsedRange();
  return ws.getRange(local);
}

/** Like rangeOf, but also accepts table names and workbook names when the address is not a cell reference. */
async function resolveRange(ctx: Excel.RequestContext, sheet: string | undefined, address: string): Promise<Excel.Range> {
  const trimmed = String(address ?? "").trim();
  const looksLikeCells = /^(?:.+!)?\$?[A-Za-z]{0,3}\$?\d*(?::\$?[A-Za-z]{0,3}\$?\d*)?$/.test(trimmed) || /!?used$/i.test(trimmed);
  if (looksLikeCells) return rangeOf(ctx, sheet, trimmed);
  const table = ctx.workbook.tables.getItemOrNullObject(trimmed);
  const named = ctx.workbook.names.getItemOrNullObject(trimmed);
  await ctx.sync();
  if (!table.isNullObject) return table.getRange();
  if (!named.isNullObject) return named.getRange();
  return rangeOf(ctx, sheet, trimmed);
}

type Cell = string | number | boolean | null;

// ---------------------------------------------------------------------------
// Reading
// ---------------------------------------------------------------------------

export async function getWorkbookOverview(): Promise<unknown> {
  return Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name,items/position,items/visibility,items/id");
    const active = sheets.getActiveWorksheet().load("name");
    const selection = ctx.workbook.getSelectedRange().load("address");
    const names = ctx.workbook.names.load("items/name,items/type,items/value,items/visible");
    await ctx.sync();

    const perSheet = sheets.items.map((ws) => {
      const used = ws.getUsedRangeOrNullObject(true).load("address,rowCount,columnCount");
      const tables = ws.tables.load("items/name");
      const charts = ws.charts.load("items/name,items/chartType");
      return { ws, used, tables, charts };
    });
    await ctx.sync();

    const tableRanges = perSheet.map((s) => s.tables.items.map((t) => t.getRange().load("address")));
    await ctx.sync();

    return {
      active_sheet: active.name,
      selection: selection.address,
      sheets: perSheet.map((s, i) => ({
        name: s.ws.name,
        position: s.ws.position + 1,
        ...(s.ws.visibility !== "Visible" ? { visibility: s.ws.visibility } : {}),
        used_range: s.used.isNullObject ? null : splitAddress(s.used.address).local,
        rows: s.used.isNullObject ? 0 : s.used.rowCount,
        columns: s.used.isNullObject ? 0 : s.used.columnCount,
        tables: s.tables.items.map((t, j) => ({ name: t.name, range: splitAddress(tableRanges[i][j].address).local })),
        charts: s.charts.items.map((c) => ({ name: c.name, type: c.chartType })),
        ai_created: isTracked("sheets", s.ws.id) || undefined,
      })),
      named_ranges: names.items.filter((n) => n.visible).map((n) => ({ name: n.name, refers_to: n.value })),
    };
  });
}

export interface ReadOptions {
  formulas?: boolean;
  formats?: boolean;
  displayText?: boolean;
}

/** Reads a range (capped at MAX_READ_CELLS) as rows labelled with Excel row numbers. */
export async function readRange(sheet: string | undefined, address: string, opts: ReadOptions = {}): Promise<unknown> {
  return Excel.run(async (ctx) => {
    let range = await resolveRange(ctx, sheet, address);
    range.load("address,rowCount,columnCount");
    await ctx.sync();
    return readLoaded(ctx, range, opts);
  });
}

async function readLoaded(ctx: Excel.RequestContext, full: Excel.Range, opts: ReadOptions): Promise<unknown> {
  const totalRows = full.rowCount;
  const totalCols = full.columnCount;
  const cols = Math.min(totalCols, MAX_READ_COLUMNS);
  const rows = Math.max(1, Math.min(totalRows, Math.floor(MAX_READ_CELLS / Math.max(1, cols))));
  const truncated = rows < totalRows || cols < totalCols;
  const range = truncated ? full.getCell(0, 0).getResizedRange(rows - 1, cols - 1) : full;
  range.load(["address", "values", "formulas", opts.formats ? "numberFormat" : "", opts.displayText ? "text" : ""].filter(Boolean));
  await ctx.sync();

  const { sheet, local } = splitAddress(range.address);
  const start = startOf(local);
  const header = Array.from({ length: cols }, (_, c) => numberToColumn(start.col + c));
  const source = (opts.displayText ? range.text : range.values) as Cell[][];
  const data = source.map((row, r) => [start.row + r, ...row]);

  const formulas: Record<string, string> = {};
  const formats: Record<string, string> = {};
  if (opts.formulas !== false) {
    (range.formulas as unknown[][]).forEach((row, r) =>
      row.forEach((f, c) => {
        if (typeof f === "string" && f.startsWith("=")) formulas[cellName(start.col + c, start.row + r)] = f;
      }),
    );
  }
  if (opts.formats) {
    (range.numberFormat as unknown[][]).forEach((row, r) =>
      row.forEach((f, c) => {
        if (typeof f === "string" && f !== "General") formats[cellName(start.col + c, start.row + r)] = f;
      }),
    );
  }

  return {
    sheet,
    address: splitAddress(full.address).local,
    ...(truncated ? { truncated: true, returned: local, total_rows: totalRows, total_columns: totalCols } : {}),
    columns: ["row", ...header],
    rows: data,
    ...(Object.keys(formulas).length ? { formulas } : {}),
    ...(Object.keys(formats).length ? { number_formats: formats } : {}),
  };
}

/** The selected range: its address and (capped) contents. */
export async function getSelection(): Promise<unknown> {
  return Excel.run(async (ctx) => {
    const range = ctx.workbook.getSelectedRange().load("address,rowCount,columnCount");
    await ctx.sync();
    return { selection: range.address, ...((await readLoaded(ctx, range, { formats: true })) as object) };
  });
}

export async function findInWorkbook(text: string, sheet?: string, completeMatch = false, matchCase = false): Promise<unknown> {
  requireApi("1.9", "Search");
  return Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets.load("items/name");
    await ctx.sync();
    const targets = sheet ? sheets.items.filter((s) => s.name === sheet) : sheets.items;
    if (!targets.length) throw new Error(`No worksheet named "${sheet}".`);
    const hits = targets.map((ws) => ({ ws, found: ws.findAllOrNullObject(text, { completeMatch, matchCase }).load("address,cellCount") }));
    await ctx.sync();
    return hits
      .filter((h) => !h.found.isNullObject)
      .map((h) => ({ sheet: h.ws.name, cells: h.found.cellCount, addresses: h.found.address.split(",").map((a) => splitAddress(a).local).slice(0, 100) }));
  });
}

/** Renders a chart to a PNG data URL. */
export async function getChartImage(sheet: string | undefined, chart: string): Promise<string> {
  return Excel.run(async (ctx) => {
    const image = sheetOf(ctx, sheet).charts.getItem(chart).getImage(800);
    await ctx.sync();
    return `data:image/png;base64,${image.value}`;
  });
}

// ---------------------------------------------------------------------------
// Writing cells
// ---------------------------------------------------------------------------

/** Pads ragged rows so every row has the same number of cells. */
function rectangular(values: unknown[][]): Cell[][] {
  const width = Math.max(1, ...values.map((r) => (Array.isArray(r) ? r.length : 1)));
  return values.map((r) => {
    const row = (Array.isArray(r) ? r : [r]).map((v) => (v === undefined ? null : (v as Cell)));
    while (row.length < width) row.push(null);
    return row;
  });
}

export interface WriteResult {
  address: string;
  overwritten: number;
}

/**
 * Writes a 2-D array starting at the top-left cell of `address`. Strings starting with "=" become formulas.
 * null leaves a cell unchanged; "" clears it.
 */
export async function writeRange(
  sheet: string | undefined,
  address: string,
  values: unknown[][],
  numberFormat?: string,
): Promise<WriteResult> {
  if (!Array.isArray(values) || !values.length) throw new Error("No values given.");
  const grid = rectangular(values);
  return Excel.run(async (ctx) => {
    const anchor = (await resolveRange(ctx, sheet, address)).getCell(0, 0);
    const target = anchor.getResizedRange(grid.length - 1, grid[0].length - 1);
    target.load("address,formulas");
    await ctx.sync();
    const overwritten = (target.formulas as unknown[][]).reduce(
      (n, row, r) => n + row.filter((v, c) => v !== "" && grid[r][c] !== null && grid[r][c] !== v).length,
      0,
    );
    // Setting formulas with plain values writes constants; null entries are skipped by Excel.
    target.formulas = grid as unknown as string[][];
    if (numberFormat) target.numberFormat = grid.map((row) => row.map(() => numberFormat));
    await ctx.sync();
    return { address: target.address, overwritten };
  });
}

export async function clearRange(sheet: string | undefined, address: string, what: "all" | "contents" | "formats"): Promise<string> {
  return Excel.run(async (ctx) => {
    const range = (await resolveRange(ctx, sheet, address)).load("address");
    range.clear(what === "contents" ? Excel.ClearApplyTo.contents : what === "formats" ? Excel.ClearApplyTo.formats : Excel.ClearApplyTo.all);
    await ctx.sync();
    return range.address;
  });
}

export async function insertCells(sheet: string | undefined, address: string, shift: "down" | "right"): Promise<string> {
  return Excel.run(async (ctx) => {
    const range = (await resolveRange(ctx, sheet, address)).load("address");
    range.insert(shift === "right" ? Excel.InsertShiftDirection.right : Excel.InsertShiftDirection.down);
    await ctx.sync();
    return range.address;
  });
}

export async function deleteCells(sheet: string | undefined, address: string, shift: "up" | "left"): Promise<string> {
  return Excel.run(async (ctx) => {
    const range = (await resolveRange(ctx, sheet, address)).load("address");
    await ctx.sync();
    const addr = range.address;
    range.delete(shift === "left" ? Excel.DeleteShiftDirection.left : Excel.DeleteShiftDirection.up);
    await ctx.sync();
    return addr;
  });
}

export interface SortKey {
  column: string;
  ascending?: boolean;
}

export async function sortRange(sheet: string | undefined, address: string, keys: SortKey[], hasHeaders: boolean): Promise<string> {
  if (!keys.length) throw new Error("No sort keys given.");
  return Excel.run(async (ctx) => {
    const range = (await resolveRange(ctx, sheet, address)).load("address,columnIndex,columnCount");
    await ctx.sync();
    const fields = keys.map((k) => {
      const key = columnToNumber(String(k.column).replace(/[^A-Za-z]/g, "")) - 1 - range.columnIndex;
      if (key < 0 || key >= range.columnCount) throw new Error(`Sort column ${k.column} is outside ${range.address}.`);
      return { key, ascending: k.ascending !== false };
    });
    range.sort.apply(fields, false, hasHeaders);
    await ctx.sync();
    return range.address;
  });
}

// ---------------------------------------------------------------------------
// Formatting
// ---------------------------------------------------------------------------

export interface FormatSpec {
  number_format?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  font_color?: string;
  font_size?: number;
  font_name?: string;
  fill_color?: string;
  horizontal_alignment?: "General" | "Left" | "Center" | "Right" | "Justify" | "CenterAcrossSelection";
  vertical_alignment?: "Top" | "Center" | "Bottom";
  wrap_text?: boolean;
  borders?: "all" | "outline" | "inside" | "bottom" | "top" | "none";
  border_color?: string;
  border_weight?: "Hairline" | "Thin" | "Medium" | "Thick";
  column_width?: number;
  row_height?: number;
  autofit?: boolean;
  merge?: boolean;
  unmerge?: boolean;
}

const BORDER_SETS: Record<string, Excel.BorderIndex[]> = {
  outline: ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"] as Excel.BorderIndex[],
  inside: ["InsideHorizontal", "InsideVertical"] as Excel.BorderIndex[],
  bottom: ["EdgeBottom"] as Excel.BorderIndex[],
  top: ["EdgeTop"] as Excel.BorderIndex[],
};
BORDER_SETS.all = [...BORDER_SETS.outline, ...BORDER_SETS.inside];
BORDER_SETS.none = [...BORDER_SETS.all, "DiagonalDown", "DiagonalUp"] as Excel.BorderIndex[];

export async function formatRange(sheet: string | undefined, address: string, spec: FormatSpec): Promise<string[]> {
  return Excel.run(async (ctx) => {
    const range = (await resolveRange(ctx, sheet, address)).load("address,rowCount,columnCount");
    await ctx.sync();
    const done: string[] = [];
    const f = range.format;

    if (spec.number_format) {
      range.numberFormat = Array.from({ length: range.rowCount }, () => Array(range.columnCount).fill(spec.number_format));
      done.push("number format");
    }
    if (spec.bold !== undefined) f.font.bold = spec.bold;
    if (spec.italic !== undefined) f.font.italic = spec.italic;
    if (spec.underline !== undefined) f.font.underline = spec.underline ? "Single" : "None";
    if (spec.font_color) f.font.color = spec.font_color;
    if (spec.font_size) f.font.size = spec.font_size;
    if (spec.font_name) f.font.name = spec.font_name;
    if ([spec.bold, spec.italic, spec.underline, spec.font_color, spec.font_size, spec.font_name].some((v) => v !== undefined)) done.push("font");
    if (spec.fill_color) {
      if (spec.fill_color === "none") f.fill.clear();
      else f.fill.color = spec.fill_color;
      done.push("fill");
    }
    if (spec.horizontal_alignment) f.horizontalAlignment = spec.horizontal_alignment;
    if (spec.vertical_alignment) f.verticalAlignment = spec.vertical_alignment;
    if (spec.wrap_text !== undefined) f.wrapText = spec.wrap_text;
    if (spec.horizontal_alignment || spec.vertical_alignment || spec.wrap_text !== undefined) done.push("alignment");
    if (spec.borders) {
      for (const edge of BORDER_SETS[spec.borders] ?? []) {
        if ((edge === "InsideHorizontal" && range.rowCount < 2) || (edge === "InsideVertical" && range.columnCount < 2)) continue;
        const border = f.borders.getItem(edge);
        if (spec.borders === "none") border.style = "None";
        else {
          border.style = "Continuous";
          border.weight = spec.border_weight ?? "Thin";
          border.color = spec.border_color ?? "#000000";
        }
      }
      done.push("borders");
    }
    if (spec.merge) {
      range.merge(false);
      done.push("merged");
    }
    if (spec.unmerge) {
      range.unmerge();
      done.push("unmerged");
    }
    if (spec.column_width) {
      f.columnWidth = spec.column_width;
      done.push("column width");
    }
    if (spec.row_height) {
      f.rowHeight = spec.row_height;
      done.push("row height");
    }
    if (spec.autofit) {
      f.autofitColumns();
      f.autofitRows();
      done.push("autofit");
    }
    await ctx.sync();
    return done;
  });
}

export interface ConditionalFormatSpec {
  type: "color_scale" | "data_bar" | "cell_value" | "text_contains" | "clear";
  operator?: "Between" | "NotBetween" | "EqualTo" | "NotEqualTo" | "GreaterThan" | "LessThan" | "GreaterThanOrEqual" | "LessThanOrEqual";
  value1?: string | number;
  value2?: string | number;
  text?: string;
  fill_color?: string;
  font_color?: string;
  min_color?: string;
  mid_color?: string;
  max_color?: string;
  bar_color?: string;
}

const asFormula = (v: unknown) => {
  const s = String(v ?? "");
  return s.startsWith("=") ? s : `=${/^-?\d+(\.\d+)?$/.test(s) ? s : JSON.stringify(s)}`;
};

export async function addConditionalFormat(sheet: string | undefined, address: string, spec: ConditionalFormatSpec): Promise<void> {
  requireApi("1.6", "Conditional formatting");
  return Excel.run(async (ctx) => {
    const range = await resolveRange(ctx, sheet, address);
    const cfs = range.conditionalFormats;
    switch (spec.type) {
      case "clear":
        cfs.clearAll();
        break;
      case "color_scale": {
        const cf = cfs.add(Excel.ConditionalFormatType.colorScale);
        cf.colorScale.criteria = {
          minimum: { formula: null as unknown as string, type: Excel.ConditionalFormatColorCriterionType.lowestValue, color: spec.min_color ?? "#F8696B" },
          ...(spec.mid_color !== "none"
            ? { midpoint: { formula: "50", type: Excel.ConditionalFormatColorCriterionType.percentile, color: spec.mid_color ?? "#FFEB84" } }
            : {}),
          maximum: { formula: null as unknown as string, type: Excel.ConditionalFormatColorCriterionType.highestValue, color: spec.max_color ?? "#63BE7B" },
        };
        break;
      }
      case "data_bar": {
        const cf = cfs.add(Excel.ConditionalFormatType.dataBar);
        cf.dataBar.positiveFormat.fillColor = spec.bar_color ?? "#638EC6";
        break;
      }
      case "cell_value": {
        if (!spec.operator || spec.value1 === undefined) throw new Error("cell_value needs operator and value1.");
        const cf = cfs.add(Excel.ConditionalFormatType.cellValue);
        cf.cellValue.rule = {
          operator: spec.operator,
          formula1: asFormula(spec.value1),
          ...(spec.value2 !== undefined ? { formula2: asFormula(spec.value2) } : {}),
        };
        if (spec.fill_color) cf.cellValue.format.fill.color = spec.fill_color;
        if (spec.font_color) cf.cellValue.format.font.color = spec.font_color;
        if (!spec.fill_color && !spec.font_color) cf.cellValue.format.fill.color = "#FFC7CE";
        break;
      }
      case "text_contains": {
        if (!spec.text) throw new Error("text_contains needs text.");
        const cf = cfs.add(Excel.ConditionalFormatType.containsText);
        cf.textComparison.rule = { operator: Excel.ConditionalTextOperator.contains, text: spec.text };
        if (spec.fill_color) cf.textComparison.format.fill.color = spec.fill_color;
        if (spec.font_color) cf.textComparison.format.font.color = spec.font_color;
        if (!spec.fill_color && !spec.font_color) cf.textComparison.format.fill.color = "#FFEB9C";
        break;
      }
      default:
        throw new Error(`Unknown conditional format type "${spec.type}".`);
    }
    await ctx.sync();
  });
}

// ---------------------------------------------------------------------------
// Tables, charts, names
// ---------------------------------------------------------------------------

export async function createTable(
  sheet: string | undefined,
  address: string,
  opts: { hasHeaders?: boolean; name?: string; style?: string },
): Promise<{ name: string; address: string }> {
  return Excel.run(async (ctx) => {
    const range = (await resolveRange(ctx, sheet, address)).load("address");
    await ctx.sync();
    const table = ctx.workbook.tables.add(range.address, opts.hasHeaders !== false);
    if (opts.name) table.name = opts.name.replace(/[^\w.]/g, "_");
    if (opts.style) table.style = opts.style;
    table.load("id,name");
    await ctx.sync();
    await track("tables", table.id);
    return { name: table.name, address: range.address };
  });
}

export interface TableFilter {
  table: string;
  column?: string;
  values?: string[];
  criteria1?: string;
  criteria2?: string;
  operator?: "And" | "Or";
  clear?: boolean;
}

export async function filterTable(spec: TableFilter): Promise<string> {
  return Excel.run(async (ctx) => {
    const table = ctx.workbook.tables.getItem(spec.table);
    if (spec.clear && !spec.column) {
      table.clearFilters();
      await ctx.sync();
      return "Cleared all filters.";
    }
    if (!spec.column) throw new Error("Give the column (header name) to filter on.");
    const filter = table.columns.getItem(spec.column).filter;
    if (spec.clear) filter.clear();
    else if (spec.values?.length) filter.applyValuesFilter(spec.values.map(String));
    else if (spec.criteria1) {
      if (spec.criteria2) {
        filter.applyCustomFilter(spec.criteria1, spec.criteria2, spec.operator === "Or" ? Excel.FilterOperator.or : Excel.FilterOperator.and);
      } else filter.applyCustomFilter(spec.criteria1);
    } else throw new Error("Give values, criteria1 or clear.");
    await ctx.sync();
    return spec.clear ? `Cleared filter on ${spec.column}.` : `Filtered ${spec.table} on ${spec.column}.`;
  });
}

export interface ChartSpec {
  sheet?: string;
  source: string;
  type?: string;
  title?: string;
  series_by?: "Auto" | "Columns" | "Rows";
  target_sheet?: string;
  top_left?: string;
  bottom_right?: string;
  legend?: "Right" | "Left" | "Top" | "Bottom" | "None";
  x_axis_title?: string;
  y_axis_title?: string;
  name?: string;
}

function applyChartOptions(chart: Excel.Chart, spec: Partial<ChartSpec>): void {
  if (spec.title !== undefined) {
    chart.title.text = spec.title;
    chart.title.visible = Boolean(spec.title);
  }
  if (spec.legend) {
    if (spec.legend === "None") chart.legend.visible = false;
    else {
      chart.legend.visible = true;
      chart.legend.position = spec.legend;
    }
  }
  if (spec.x_axis_title !== undefined) {
    chart.axes.categoryAxis.title.text = spec.x_axis_title;
    chart.axes.categoryAxis.title.visible = Boolean(spec.x_axis_title);
  }
  if (spec.y_axis_title !== undefined) {
    chart.axes.valueAxis.title.text = spec.y_axis_title;
    chart.axes.valueAxis.title.visible = Boolean(spec.y_axis_title);
  }
  if (spec.top_left) chart.setPosition(spec.top_left, spec.bottom_right);
}

export async function createChart(spec: ChartSpec): Promise<{ name: string; sheet: string }> {
  return Excel.run(async (ctx) => {
    const source = await resolveRange(ctx, spec.sheet, spec.source);
    source.load("address,columnIndex,columnCount");
    await ctx.sync();
    const { sheet: sourceSheet } = splitAddress(source.address);
    const host = sheetOf(ctx, spec.target_sheet ?? sourceSheet).load("name");
    const chart = host.charts.add((spec.type ?? "ColumnClustered") as Excel.ChartType, source, spec.series_by ?? "Auto");
    if (spec.name) chart.name = spec.name;
    // Default placement: right of the source data, about 8 columns wide and 16 rows high.
    const pos = { ...spec };
    if (!pos.top_left && (!spec.target_sheet || spec.target_sheet === sourceSheet)) {
      const start = startOf(splitAddress(source.address).local);
      const col = source.columnIndex + source.columnCount + 2;
      pos.top_left = cellName(col, start.row);
      pos.bottom_right = cellName(col + 7, start.row + 15);
    }
    applyChartOptions(chart, pos);
    chart.load("id,name");
    await ctx.sync();
    await track("charts", chart.id);
    return { name: chart.name, sheet: host.name };
  });
}

export async function updateChart(sheet: string | undefined, name: string, spec: Partial<ChartSpec> & { delete?: boolean }): Promise<string> {
  return Excel.run(async (ctx) => {
    const chart = sheetOf(ctx, sheet).charts.getItem(name);
    if (spec.delete) {
      chart.delete();
      await ctx.sync();
      return `Deleted chart ${name}.`;
    }
    if (spec.type) chart.chartType = spec.type as Excel.ChartType;
    if (spec.source) chart.setData(await resolveRange(ctx, sheet, spec.source), spec.series_by ?? "Auto");
    applyChartOptions(chart, spec);
    await ctx.sync();
    return `Updated chart ${name}.`;
  });
}

export async function addNamedRange(sheet: string | undefined, address: string, name: string, comment?: string): Promise<string> {
  return Excel.run(async (ctx) => {
    const range = (await resolveRange(ctx, sheet, address)).load("address");
    ctx.workbook.names.add(name, range, comment);
    await ctx.sync();
    return range.address;
  });
}

// ---------------------------------------------------------------------------
// Worksheets and navigation
// ---------------------------------------------------------------------------

export async function addWorksheet(name?: string, activate = true): Promise<string> {
  return Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.add(name || undefined).load("id,name");
    if (activate) ws.activate();
    await ctx.sync();
    await track("sheets", ws.id);
    return ws.name;
  });
}

export async function renameWorksheet(sheet: string, newName: string): Promise<void> {
  return Excel.run(async (ctx) => {
    ctx.workbook.worksheets.getItem(sheet).name = newName;
    await ctx.sync();
  });
}

export async function worksheetStats(sheet: string): Promise<{ name: string; cells: number }> {
  return Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(sheet).load("name");
    const used = ws.getUsedRangeOrNullObject(true).load("cellCount");
    await ctx.sync();
    return { name: ws.name, cells: used.isNullObject ? 0 : used.cellCount };
  });
}

export async function deleteWorksheet(sheet: string): Promise<void> {
  return Excel.run(async (ctx) => {
    ctx.workbook.worksheets.getItem(sheet).delete();
    await ctx.sync();
  });
}

export async function selectRange(sheet: string | undefined, address: string): Promise<string> {
  return Excel.run(async (ctx) => {
    const range = (await resolveRange(ctx, sheet, address)).load("address");
    range.worksheet.activate();
    range.select();
    await ctx.sync();
    return range.address;
  });
}

export async function freezePanes(sheet: string | undefined, rows: number, columns: number): Promise<void> {
  return Excel.run(async (ctx) => {
    const panes = sheetOf(ctx, sheet).freezePanes;
    panes.unfreeze();
    if (rows > 0 && columns > 0) panes.freezeAt(sheetOf(ctx, sheet).getRange(`A1:${cellName(columns, rows)}`));
    else if (rows > 0) panes.freezeRows(rows);
    else if (columns > 0) panes.freezeColumns(columns);
    await ctx.sync();
  });
}

// ---------------------------------------------------------------------------
// Images and files
// ---------------------------------------------------------------------------

export interface ImageSpec {
  base64: string;
  sheet?: string;
  cell?: string;
  width?: number;
  height?: number;
  name?: string;
  altText?: string;
}

export async function insertImage(spec: ImageSpec): Promise<string> {
  requireApi("1.9", "Inserting images");
  return Excel.run(async (ctx) => {
    const ws = sheetOf(ctx, spec.sheet);
    const anchor = (spec.cell ? ws.getRange(spec.cell) : ctx.workbook.getSelectedRange()).load("left,top");
    await ctx.sync();
    const shape = ws.shapes.addImage(spec.base64);
    shape.left = anchor.left;
    shape.top = anchor.top;
    shape.lockAspectRatio = true;
    if (spec.width) shape.width = spec.width;
    else if (spec.height) shape.height = spec.height;
    if (spec.name) shape.name = spec.name;
    if (spec.altText) shape.altTextDescription = spec.altText;
    shape.load("id");
    await ctx.sync();
    await track("shapes", shape.id);
    return shape.id;
  });
}

export async function insertSheetsFromFile(base64: string, sheetNames?: string[]): Promise<number> {
  requireApi("1.13", "Inserting worksheets from a file");
  return Excel.run(async (ctx) => {
    const ids = ctx.workbook.insertWorksheetsFromBase64(base64, {
      positionType: Excel.WorksheetPositionType.end,
      ...(sheetNames?.length ? { sheetNamesToInsert: sheetNames } : {}),
    });
    await ctx.sync();
    for (const id of ids.value) await track("sheets", id);
    return ids.value.length;
  });
}

// ---------------------------------------------------------------------------
// AI-content tracking (stored in the workbook's add-in settings)
// ---------------------------------------------------------------------------

type TrackKind = "sheets" | "tables" | "charts" | "shapes";
type Registry = Record<TrackKind, string[]>;
const REGISTRY_KEY = "mistral.created";

function registry(): Registry {
  const empty: Registry = { sheets: [], tables: [], charts: [], shapes: [] };
  try {
    const stored = Office.context.document.settings.get(REGISTRY_KEY) as Partial<Registry> | null;
    return { ...empty, ...(stored ?? {}) };
  } catch {
    return empty;
  }
}

function saveRegistry(reg: Registry): Promise<void> {
  return new Promise((resolve) => {
    try {
      Office.context.document.settings.set(REGISTRY_KEY, reg);
      Office.context.document.settings.saveAsync(() => resolve());
    } catch {
      resolve();
    }
  });
}

function isTracked(kind: TrackKind, id: string): boolean {
  return registry()[kind].includes(id);
}

async function track(kind: TrackKind, id: string): Promise<void> {
  const reg = registry();
  if (!reg[kind].includes(id)) reg[kind].push(id);
  await saveRegistry(reg);
}

export interface AiContentCount {
  sheets: number;
  tables: number;
  charts: number;
  shapes: number;
}

/**
 * Finds (dryRun) or removes everything the assistant created: worksheets are deleted, charts and images
 * removed, tables converted back to plain ranges (their data is kept).
 */
export async function deleteAiContent(dryRun = false): Promise<AiContentCount> {
  const reg = registry();
  return Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets.load("items/id,items/name");
    const tables = ctx.workbook.tables.load("items/id");
    await ctx.sync();
    const sheetIds = new Set(sheets.items.map((s) => s.id));
    const doomedSheets = sheets.items.filter((s) => reg.sheets.includes(s.id));
    const keptSheets = sheets.items.filter((s) => !reg.sheets.includes(s.id));

    const charts = keptSheets.map((s) => s.charts.load("items/id"));
    const shapes = Office.context.requirements.isSetSupported("ExcelApi", "1.9") ? keptSheets.map((s) => s.shapes.load("items/id")) : [];
    await ctx.sync();

    const tableHits = tables.items.filter((t) => reg.tables.includes(t.id));
    const chartHits = charts.flatMap((c) => c.items.filter((ch) => reg.charts.includes(ch.id)));
    const shapeHits = shapes.flatMap((c) => c.items.filter((sh) => reg.shapes.includes(sh.id)));
    // Excel refuses to delete the last remaining sheet.
    const sheetHits = keptSheets.length ? doomedSheets : doomedSheets.slice(1);

    const count = { sheets: sheetHits.length, tables: tableHits.length, charts: chartHits.length, shapes: shapeHits.length };
    if (dryRun) return count;

    chartHits.forEach((c) => c.delete());
    shapeHits.forEach((s) => s.delete());
    tableHits.forEach((t) => t.convertToRange());
    if (sheetHits.length) {
      if (keptSheets.length) keptSheets[0].activate();
      sheetHits.forEach((s) => s.delete());
    }
    await ctx.sync();

    const remaining = new Set(sheetHits.map((s) => s.id));
    await saveRegistry({
      sheets: reg.sheets.filter((id) => sheetIds.has(id) && !remaining.has(id)),
      tables: [],
      charts: [],
      shapes: [],
    });
    return count;
  });
}
