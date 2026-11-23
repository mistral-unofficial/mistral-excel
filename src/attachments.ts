/** Files the user attached to the chat (images, CSV/TSV and .xlsx), kept in memory for the current conversation. */

export interface Attachment {
  id: string;
  name: string;
  kind: "image" | "csv" | "xlsx";
  /** Raw base64 (no data: prefix), as needed by Office.js. Images are always PNG or JPEG. Empty for CSV. */
  base64: string;
  mime: string;
  width?: number;
  height?: number;
  /** Downscaled data URL sent to the model for vision. */
  previewDataUrl?: string;
  /** Parsed rows of a CSV/TSV file. */
  rows?: string[][];
}

const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const MAX_VISION_EDGE = 1024;

const store = new Map<string, Attachment>();
let counter = 0;

export function getAttachment(id: string): Attachment {
  const a = store.get(id);
  if (!a) throw new Error(`Unknown attachment "${id}". Available: ${[...store.keys()].join(", ") || "none"}.`);
  return a;
}

export function clearAttachments(): void {
  store.clear();
}

const isCsv = (file: File) => /\.(csv|tsv|txt)$/i.test(file.name) || ["text/csv", "text/tab-separated-values", "text/plain"].includes(file.type);
const isXlsx = (file: File) => file.type === XLSX_MIME || /\.(xlsx|xlsm)$/i.test(file.name);

export function isSupportedFile(file: File): boolean {
  return file.type.startsWith("image/") || isCsv(file) || isXlsx(file);
}

function readAsDataUrl(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(reader.error ?? new Error("Could not read file."));
    reader.readAsDataURL(blob);
  });
}

function loadImage(src: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = () => reject(new Error("Could not decode image."));
    img.src = src;
  });
}

function drawToPng(img: HTMLImageElement, maxEdge?: number): string {
  const scale = maxEdge ? Math.min(1, maxEdge / Math.max(img.naturalWidth, img.naturalHeight)) : 1;
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.round(img.naturalWidth * scale));
  canvas.height = Math.max(1, Math.round(img.naturalHeight * scale));
  canvas.getContext("2d")!.drawImage(img, 0, 0, canvas.width, canvas.height);
  return canvas.toDataURL("image/png");
}

const stripPrefix = (dataUrl: string) => dataUrl.slice(dataUrl.indexOf(",") + 1);

/** Picks the delimiter (tab, semicolon or comma) that occurs most often in the first line. */
function detectDelimiter(text: string): string {
  const first = text.split(/\r?\n/, 1)[0] ?? "";
  const counts = ["\t", ";", ","].map((d) => [d, first.split(d).length - 1] as const);
  counts.sort((a, b) => b[1] - a[1]);
  return counts[0][1] > 0 ? counts[0][0] : ",";
}

/** RFC 4180-style CSV parser (quoted fields, escaped quotes, newlines inside quotes). */
export function parseCsv(text: string, delimiter = detectDelimiter(text)): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let quoted = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (quoted) {
      if (c === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i++;
        } else quoted = false;
      } else field += c;
    } else if (c === '"' && field === "") quoted = true;
    else if (c === delimiter) {
      row.push(field);
      field = "";
    } else if (c === "\n" || c === "\r") {
      if (c === "\r" && text[i + 1] === "\n") i++;
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
    } else field += c;
  }
  if (field !== "" || row.length) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}

export async function addAttachment(file: File): Promise<Attachment> {
  const id = `att_${++counter}`;
  const name = file.name || (file.type.startsWith("image/") ? `pasted-image-${counter}.png` : `file-${counter}`);

  if (file.type.startsWith("image/")) {
    const original = await readAsDataUrl(file);
    const img = await loadImage(original);
    // Excel reliably accepts PNG and JPEG; convert anything else (SVG, WebP, ...) to PNG.
    const keep = file.type === "image/png" || file.type === "image/jpeg";
    const dataUrl = keep ? original : drawToPng(img);
    const attachment: Attachment = {
      id,
      name,
      kind: "image",
      base64: stripPrefix(dataUrl),
      mime: keep ? file.type : "image/png",
      width: img.naturalWidth,
      height: img.naturalHeight,
      previewDataUrl: drawToPng(img, MAX_VISION_EDGE),
    };
    store.set(id, attachment);
    return attachment;
  }

  if (isXlsx(file)) {
    const attachment: Attachment = { id, name, kind: "xlsx", base64: stripPrefix(await readAsDataUrl(file)), mime: XLSX_MIME };
    store.set(id, attachment);
    return attachment;
  }

  if (isCsv(file)) {
    const text = (await file.text()).replace(/^﻿/, "");
    const rows = parseCsv(text).filter((r) => r.some((c) => c !== ""));
    const attachment: Attachment = { id, name, kind: "csv", base64: "", mime: file.type || "text/csv", rows };
    store.set(id, attachment);
    return attachment;
  }

  throw new Error(`${name}: only images, .csv/.tsv/.txt and .xlsx files can be attached.`);
}

/** One line per attachment, telling the model how it can be used. */
export function describeAttachment(a: Attachment): string {
  switch (a.kind) {
    case "image":
      return `${a.id}: ${a.name} (image, ${a.width}x${a.height}px) — can be placed on a sheet with insert_image`;
    case "csv": {
      const cols = Math.max(0, ...(a.rows ?? []).map((r) => r.length));
      return `${a.id}: ${a.name} (CSV, ${a.rows?.length ?? 0} rows x ${cols} columns) — preview with read_attachment, write into a sheet with import_csv`;
    }
    case "xlsx":
      return `${a.id}: ${a.name} (Excel workbook) — its worksheets can be copied in with insert_sheets_from_file`;
  }
}
