import * as XLSX from "xlsx";

type ParseRequest = {
  file: File;
  sheet?: string;
  headerRow?: number;
};

type ParseResult = {
  sheetNames: string[];
  sheet: string;
  headerRow: number;
  columns: string[];
  rows: Record<string, unknown>[];
};

type SheetCell = XLSX.CellObject | undefined;
type DenseWorksheet = XLSX.WorkSheet & { [row: number]: (SheetCell | undefined)[] };

function cellValue(worksheet: DenseWorksheet, row: number, column: number): unknown {
  const cell = worksheet[row]?.[column];
  if (!cell) return undefined;
  // Formula cells count as populated even when their cached value is empty.
  if (cell.v !== undefined && cell.v !== null && cell.v !== "") return cell.v;
  if (cell.f) return cell.f;
  return undefined;
}

function detectHeader(worksheet: DenseWorksheet, lastColumn: number): number {
  for (let rowIndex = 0; rowIndex < 15; rowIndex++) {
    let strings = 0;
    for (let columnIndex = 0; columnIndex <= lastColumn; columnIndex++) {
      const value = cellValue(worksheet, rowIndex, columnIndex);
      if (typeof value === "string" && value.trim() !== "") strings++;
    }
    if (strings >= 2) return rowIndex;
  }
  return 0;
}

function findLastPopulatedRow(
  worksheet: DenseWorksheet,
  firstDataRow: number,
  lastDeclaredRow: number,
  lastColumn: number,
): number {
  // Excel files can have a bloated used range caused by formatting empty rows.
  // Walk backwards through the parsed dense worksheet and stop at the last
  // row containing a value, instead of materializing every declared row.
  for (let rowIndex = lastDeclaredRow; rowIndex >= firstDataRow; rowIndex--) {
    const row = worksheet[rowIndex];
    if (!row) continue;
    for (let columnIndex = 0; columnIndex <= lastColumn; columnIndex++) {
      if (cellValue(worksheet, rowIndex, columnIndex) !== undefined) return rowIndex;
    }
  }
  return firstDataRow - 1;
}

type WorkerScope = {
  onmessage: ((event: MessageEvent<ParseRequest>) => void) | null;
  postMessage: (message: unknown) => void;
};

const workerScope = self as unknown as WorkerScope;

workerScope.onmessage = async (event: MessageEvent<ParseRequest>) => {
  try {
    const { file, sheet: requestedSheet, headerRow: requestedHeaderRow } = event.data;
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array", cellDates: true, dense: true });
    const sheetNames = workbook.SheetNames;
    const sheet = requestedSheet && workbook.Sheets[requestedSheet]
      ? requestedSheet
      : sheetNames[0];
    if (!sheet) throw new Error("The workbook does not contain a readable worksheet.");

    const worksheet = workbook.Sheets[sheet] as DenseWorksheet;
    const declaredRange = worksheet["!ref"]
      ? XLSX.utils.decode_range(worksheet["!ref"])
      : { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };

    const headerRow = Number.isInteger(requestedHeaderRow)
      ? Math.max(0, requestedHeaderRow as number)
      : detectHeader(worksheet, declaredRange.e.c);

    const lastDataRow = findLastPopulatedRow(
      worksheet,
      headerRow + 1,
      declaredRange.e.r,
      declaredRange.e.c,
    );

    const header = Array.from(
      { length: declaredRange.e.c - declaredRange.s.c + 1 },
      (_, index) => cellValue(worksheet, headerRow, declaredRange.s.c + index),
    );
    const seen = new Map<string, number>();
    const columns = header.map((cell, index) => {
      const base = cell == null || String(cell).trim() === ""
        ? `Column ${index + 1}`
        : String(cell).trim();
      const occurrence = seen.get(base) ?? 0;
      seen.set(base, occurrence + 1);
      return occurrence > 0 ? `${base} (${occurrence + 1})` : base;
    });

    const rows: Record<string, unknown>[] = [];
    if (lastDataRow >= headerRow + 1) {
      const matrix = XLSX.utils.sheet_to_json<unknown[]>(worksheet, {
        header: 1,
        defval: null,
        raw: true,
        blankrows: false,
        range: {
          s: { r: headerRow + 1, c: declaredRange.s.c },
          e: { r: lastDataRow, c: declaredRange.e.c },
        },
      });

      for (const source of matrix) {
        const record: Record<string, unknown> = {};
        let hasValue = false;
        for (let columnIndex = 0; columnIndex < columns.length; columnIndex++) {
          const value = source[columnIndex] ?? null;
          record[columns[columnIndex]] = value;
          if (value !== null && value !== "") hasValue = true;
        }
        if (hasValue) rows.push(record);
      }
    }

    const result: ParseResult = { sheetNames, sheet, headerRow, columns, rows };
    workerScope.postMessage({ ok: true, result });
  } catch (error) {
    workerScope.postMessage({
      ok: false,
      error: error instanceof Error ? error.message : "Unable to read this workbook.",
    });
  }
};
