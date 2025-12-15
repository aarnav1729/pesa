import * as XLSX from "xlsx";
import { HoldingRecord } from "./db";

interface RawRow {
  SNo?: number;
  DPID?: string;
  "CLIENT-ID"?: string;
  NAME?: string;
  SECOND?: string;
  THIRD?: string;
  CATEGORY?: string;
  BOUGHT?: string | number;
  SOLD?: string | number;
  [key: string]: string | number | undefined;
}

interface ParsedFile {
  firstDate: string;
  secondDate: string;
  firstAsOnCol: string;
  secondAsOnCol: string;
  rows: Array<{
    dpid: string;
    clientId: string;
    name: string;
    category: string;
    firstValue: number;
    secondValue: number;
    bought: number;
    sold: number;
  }>;
}

function extractDateFromHeader(header: string): string | null {
  const m = header.match(/(\d{2}-\d{2}-\d{4})/);
  return m ? m[1] : null;
}

function parseDate(dateStr: string): Date {
  const match = dateStr.match(/(\d{2})-(\d{2})-(\d{4})/);
  if (match) {
    const [, day, month, year] = match;
    return new Date(
      parseInt(year, 10),
      parseInt(month, 10) - 1,
      parseInt(day, 10)
    );
  }
  return new Date(0);
}

function toNumber(v: unknown): number {
  if (v === null || v === undefined) return 0;
  if (typeof v === "number") return Number.isFinite(v) ? v : 0;

  let s = String(v).trim();
  if (!s) return 0;

  // Handle accounting negatives like "(1,234)"
  let neg = false;
  if (s.startsWith("(") && s.endsWith(")")) {
    neg = true;
    s = s.slice(1, -1).trim();
  }

  // Remove commas and spaces (supports Indian grouping too)
  s = s.replace(/,/g, "").replace(/\s+/g, "");

  // Remove leading "+"
  if (s.startsWith("+")) s = s.slice(1);

  const n = Number(s);
  if (!Number.isFinite(n)) return 0;
  return neg ? -n : n;
}

export async function parseXLSXFile(file: File): Promise<ParsedFile | null> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: "array" });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        const jsonData: RawRow[] = XLSX.utils.sheet_to_json(worksheet);

        if (jsonData.length === 0) {
          resolve(null);
          return;
        }

        const headers = Object.keys(jsonData[0] ?? {});
        const asOnColumns = headers.filter((h) =>
          /^as on/i.test(String(h).trim())
        );

        const asOnWithDates = asOnColumns
          .map((header) => ({ header, date: extractDateFromHeader(header) }))
          .filter((x): x is { header: string; date: string } => Boolean(x.date))
          .sort(
            (a, b) => parseDate(a.date).getTime() - parseDate(b.date).getTime()
          );

        if (asOnWithDates.length < 2) {
          console.warn(
            "Could not find expected AS ON columns (need at least 2)"
          );
          resolve(null);
          return;
        }

        const firstAsOnCol = asOnWithDates[0].header;
        const secondAsOnCol = asOnWithDates[1].header;

        const firstDate = asOnWithDates[0].date;
        const secondDate = asOnWithDates[1].date;

        const rows = jsonData
          .map((row) => ({
            dpid: String(row.DPID || ""),
            clientId: String(row["CLIENT-ID"] || ""),
            name: String(row.NAME || ""),
            category: String(row.CATEGORY || ""),
            firstValue: toNumber(row[firstAsOnCol]),
            secondValue: toNumber(row[secondAsOnCol]),
            bought: toNumber(row.BOUGHT),
            sold: toNumber(row.SOLD),
          }))
          .filter((row) => row.name);

        resolve({
          firstDate,
          secondDate,
          firstAsOnCol,
          secondAsOnCol,
          rows,
        });
      } catch (error) {
        console.error("Error parsing XLSX:", error);
        reject(error);
      }
    };

    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

export function consolidateData(parsedFiles: ParsedFile[]): {
  holdings: HoldingRecord[];
  dates: string[];
} {
  if (parsedFiles.length === 0) {
    return { holdings: [], dates: [] };
  }

  // Sort files by their secondDate so timeline flows logically.
  const sortedFiles = [...parsedFiles].sort(
    (a, b) =>
      parseDate(a.secondDate).getTime() - parseDate(b.secondDate).getTime()
  );

  /**
   * Build column instances (2 per file).
   * Each instance gets a unique key so duplicates of the same base date are preserved.
   */
  const columns: Array<{
    key: string;
    baseDate: string;
    fileIndex: number;
    pos: 1 | 2;
  }> = [];

  sortedFiles.forEach((f, i) => {
    const key1 = `${f.firstDate}@@${i}-1`;
    const key2 = `${f.secondDate}@@${i}-2`;

    columns.push({ key: key1, baseDate: f.firstDate, fileIndex: i, pos: 1 });
    columns.push({ key: key2, baseDate: f.secondDate, fileIndex: i, pos: 2 });
  });

  // Sort column instances by base date, then by file order, then position.
  columns.sort((a, b) => {
    const d = parseDate(a.baseDate).getTime() - parseDate(b.baseDate).getTime();
    if (d !== 0) return d;
    if (a.fileIndex !== b.fileIndex) return a.fileIndex - b.fileIndex;
    return a.pos - b.pos;
  });

  const dates = columns.map((c) => c.key);
  const holdingsMap = new Map<string, HoldingRecord>();

  // Precompute each file's two keys for usage in row mapping.
  const fileKeys = sortedFiles.map((f, i) => ({
    firstKey: `${f.firstDate}@@${i}-1`,
    secondKey: `${f.secondDate}@@${i}-2`,
  }));

  for (let i = 0; i < sortedFiles.length; i++) {
    const file = sortedFiles[i];
    const { firstKey, secondKey } = fileKeys[i];

    for (const row of file.rows) {
      const rowKey = `${row.dpid}-${row.clientId}-${row.name}`;

      if (!holdingsMap.has(rowKey)) {
        holdingsMap.set(rowKey, {
          id: rowKey,
          dpid: row.dpid,
          clientId: row.clientId,
          name: row.name,
          category: row.category,
          dateValues: {},
        });
      }

      const holding = holdingsMap.get(rowKey)!;

      // First AS ON column for this file (no bought/sold attributed here)
      holding.dateValues[firstKey] = {
        value: row.firstValue,
        bought: 0,
        sold: 0,
      };

      // Second AS ON column for this file (bought/sold belong to this period)
      holding.dateValues[secondKey] = {
        value: row.secondValue,
        bought: row.bought,
        sold: row.sold,
      };

      if (row.category) {
        holding.category = row.category;
      }
    }
  }

  return {
    holdings: Array.from(holdingsMap.values()),
    dates,
  };
}
