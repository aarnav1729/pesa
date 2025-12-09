import * as XLSX from 'xlsx';
import { HoldingRecord } from './db';

interface RawRow {
  SNo?: number;
  DPID?: string;
  'CLIENT-ID'?: string;
  NAME?: string;
  SECOND?: string;
  THIRD?: string;
  CATEGORY?: string;
  [key: string]: string | number | undefined;
}

interface ParsedFile {
  date: string;
  rows: Array<{
    dpid: string;
    clientId: string;
    name: string;
    category: string;
    asOnValue: number;
    bought: number;
    sold: number;
  }>;
}

function parseDate(dateStr: string): Date {
  // Parse date from format "AS ON DD-MM-YYYY"
  const match = dateStr.match(/(\d{2})-(\d{2})-(\d{4})/);
  if (match) {
    const [, day, month, year] = match;
    return new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
  }
  return new Date(0);
}

function formatDate(date: Date): string {
  const day = date.getDate().toString().padStart(2, '0');
  const month = (date.getMonth() + 1).toString().padStart(2, '0');
  const year = date.getFullYear();
  return `${day}-${month}-${year}`;
}

export async function parseXLSXFile(file: File): Promise<ParsedFile | null> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        const jsonData: RawRow[] = XLSX.utils.sheet_to_json(worksheet);
        
        if (jsonData.length === 0) {
          resolve(null);
          return;
        }
        
        // Find the "AS ON" columns
        const headers = Object.keys(jsonData[0]);
        const asOnColumns = headers.filter(h => h.startsWith('AS ON'));
        
        if (asOnColumns.length < 2) {
          console.warn('Could not find expected AS ON columns');
          resolve(null);
          return;
        }
        
        // Get the first date (starting position) and second date (ending position)
        const firstAsOnCol = asOnColumns[0];
        const secondAsOnCol = asOnColumns[1];
        const dateMatch = secondAsOnCol.match(/AS ON (\d{2}-\d{2}-\d{4})/);
        
        if (!dateMatch) {
          console.warn('Could not parse date from column:', secondAsOnCol);
          resolve(null);
          return;
        }
        
        const date = dateMatch[1];
        
        const rows = jsonData.map(row => ({
          dpid: String(row.DPID || ''),
          clientId: String(row['CLIENT-ID'] || ''),
          name: String(row.NAME || ''),
          category: String(row.CATEGORY || ''),
          asOnValue: Number(row[secondAsOnCol]) || 0,
          bought: Number(row.BOUGHT) || 0,
          sold: Number(row.SOLD) || 0,
        })).filter(row => row.name);
        
        resolve({ date, rows });
      } catch (error) {
        console.error('Error parsing XLSX:', error);
        reject(error);
      }
    };
    
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

export function consolidateData(parsedFiles: ParsedFile[]): { holdings: HoldingRecord[]; dates: string[] } {
  // Sort files by date chronologically
  const sortedFiles = [...parsedFiles].sort((a, b) => {
    return parseDate(a.date).getTime() - parseDate(b.date).getTime();
  });
  
  const dates = sortedFiles.map(f => f.date);
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
      holding.dateValues[file.date] = {
        value: row.asOnValue,
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
