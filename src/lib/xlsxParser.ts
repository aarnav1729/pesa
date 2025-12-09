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

        // Case-insensitive AS ON detection
        const asOnColumns = headers.filter((h) =>
          /^as on/i.test(String(h).trim())
        );

        // Extract dates and sort the AS ON columns by the date inside the header
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
            firstValue: Number(row[firstAsOnCol]) || 0,
            secondValue: Number(row[secondAsOnCol]) || 0,
            bought: Number(row.BOUGHT) || 0,
            sold: Number(row.SOLD) || 0,
          }))
          .filter((row) => row.name);

        resolve({ firstDate, secondDate, rows });
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

  // Build unique union of all dates across all files
  const dateSet = new Set<string>();
  parsedFiles.forEach((f) => {
    dateSet.add(f.firstDate);
    dateSet.add(f.secondDate);
  });

  const dates = Array.from(dateSet).sort(
    (a, b) => parseDate(a).getTime() - parseDate(b).getTime()
  );

  // Sort files by their secondDate so later snapshots apply in order
  const sortedFiles = [...parsedFiles].sort(
    (a, b) =>
      parseDate(a.secondDate).getTime() - parseDate(b.secondDate).getTime()
  );

  const holdingsMap = new Map<string, HoldingRecord>();

  for (const file of sortedFiles) {
    for (const row of file.rows) {
      const key = `${row.dpid}-${row.clientId}-${row.name}`;

      if (!holdingsMap.has(key)) {
        holdingsMap.set(key, {
          id: key,
          dpid: row.dpid,
          clientId: row.clientId,
          name: row.name,
          category: row.category,
          dateValues: {},
        });
      }

      const holding = holdingsMap.get(key)!;

      // 1) First date snapshot (no bought/sold attributed here)
      const existingFirst = holding.dateValues[file.firstDate];
      holding.dateValues[file.firstDate] = {
        value: row.firstValue,
        bought: existingFirst?.bought ?? 0,
        sold: existingFirst?.sold ?? 0,
      };

      // 2) Second date snapshot (bought/sold belongs to this file's period)
      holding.dateValues[file.secondDate] = {
        value: row.secondValue,
        bought: row.bought,
        sold: row.sold,
      };

      // Update category if it's set
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
