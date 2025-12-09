import { openDB, DBSchema, IDBPDatabase } from 'idb';

export interface HoldingRecord {
  id: string;
  dpid: string;
  clientId: string;
  name: string;
  category: string;
  dateValues: Record<string, { value: number; bought: number; sold: number }>;
}

interface MetadataRecord {
  key: string;
  dates: string[];
}

interface HoldingsDB extends DBSchema {
  holdings: {
    key: string;
    value: HoldingRecord;
    indexes: { 'by-name': string; 'by-dpid': string; 'by-category': string };
  };
  metadata: {
    key: string;
    value: MetadataRecord;
  };
}

let dbPromise: Promise<IDBPDatabase<HoldingsDB>> | null = null;

export async function getDB() {
  if (!dbPromise) {
    dbPromise = openDB<HoldingsDB>('holdings-tracker', 1, {
      upgrade(db) {
        const holdingsStore = db.createObjectStore('holdings', { keyPath: 'id' });
        holdingsStore.createIndex('by-name', 'name');
        holdingsStore.createIndex('by-dpid', 'dpid');
        holdingsStore.createIndex('by-category', 'category');
        
        db.createObjectStore('metadata', { keyPath: 'key' });
      },
    });
  }
  return dbPromise;
}

export async function saveHoldings(holdings: HoldingRecord[], dates: string[]) {
  const db = await getDB();
  const tx = db.transaction(['holdings', 'metadata'], 'readwrite');
  
  // Clear existing data
  await tx.objectStore('holdings').clear();
  
  // Save new holdings
  for (const holding of holdings) {
    await tx.objectStore('holdings').put(holding);
  }
  
  // Save metadata
  await tx.objectStore('metadata').put({ key: 'dates', dates: dates });
  
  await tx.done;
}

export async function getHoldings(): Promise<{ holdings: HoldingRecord[]; dates: string[] }> {
  const db = await getDB();
  const holdings = await db.getAll('holdings');
  const metadata = await db.get('metadata', 'dates');
  return { holdings, dates: metadata?.dates || [] };
}

export async function clearAllData() {
  const db = await getDB();
  const tx = db.transaction(['holdings', 'metadata'], 'readwrite');
  await tx.objectStore('holdings').clear();
  await tx.objectStore('metadata').clear();
  await tx.done;
}
