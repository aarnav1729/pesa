"use strict";

require("dotenv").config();

const fs = require("fs");
const path = require("path");
const https = require("https");

const express = require("express");
const cookieParser = require("cookie-parser");
const jwt = require("jsonwebtoken");

// Microsoft Graph (OTP Email)
const { Client } = require("@microsoft/microsoft-graph-client");
const { ClientSecretCredential } = require("@azure/identity");
require("isomorphic-fetch");
const cors = require("cors");
const helmet = require("helmet");
const compression = require("compression");
const morgan = require("morgan");
const sql = require("mssql");
const XLSX = require("xlsx");
const ExcelJS = require("exceljs");

// HTTPS options (paths as requested)
const httpsOptions = {
  key: fs.readFileSync(path.join(__dirname, "certs", "mydomain.key")),
  cert: fs.readFileSync(path.join(__dirname, "certs", "d466aacf3db3f299.crt")),
  ca: fs.readFileSync(path.join(__dirname, "certs", "gd_bundle-g2-g1.crt")),
};

const APP_PORT = Number(process.env.PORT || 27443);
const APP_HOST = process.env.HOST || "0.0.0.0";

// ─────────────────────────────────────────────────────────────────────────────
// Auth + OTP + Graph config
// ─────────────────────────────────────────────────────────────────────────────
function mustGetEnv(name) {
  const v = process.env[name];
  if (!v || !String(v).trim()) {
    console.error(`❌ Missing required env: ${name}`);
    process.exit(1);
  }
  return String(v).trim();
}

// IMPORTANT: Do NOT hardcode secrets. Put the same values you used in INVEST into .env for PESA.
const GRAPH_CLIENT_ID = "3d310826-2173-44e5-b9a2-b21e940b67f7";
const GRAPH_TENANT_ID = "1c3de7f3-f8d1-41d3-8583-2517cf3ba3b1";
const GRAPH_CLIENT_SECRET = "2e78Q~yX92LfwTTOg4EYBjNQrXrZ2z5di1Kvebog";
const GRAPH_SENDER_EMAIL = "spot@premierenergies.com"; // e.g. spot@premierenergies.com

const SESSION_JWT_SECRET = mustGetEnv("PESA_JWT_SECRET"); // any strong random string
const COOKIE_DOMAIN = (process.env.COOKIE_DOMAIN || "").trim(); // optional (blank for localhost)
const SESSION_COOKIE = "pesa_session";

// Allow-list (server is the source of truth)
// Set: PESA_ALLOWED_EMAILS="a@premierenergies.com,b@premierenergies.com"
const ALLOWED = new Set(
  String(process.env.PESA_ALLOWED_EMAILS || "")
    .split(",")
    .map((s) => s.trim().toLowerCase())
    .filter(Boolean)
);
// Safe fallback (keep aligned with your Login.tsx)
if (ALLOWED.size === 0) {
  ["vcs@premierenergies.com", "saluja@premierenergies.com"].forEach((e) =>
    ALLOWED.add(e)
  );
}

function normalizeEmail(userInput = "") {
  const raw = String(userInput || "")
    .trim()
    .toLowerCase();
  if (!raw) return "";
  return raw.includes("@") ? raw : `${raw}@premierenergies.com`;
}
function isAllowed(email) {
  return ALLOWED.has(
    String(email || "")
      .trim()
      .toLowerCase()
  );
}

// Graph client (once)
const credential = new ClientSecretCredential(
  GRAPH_TENANT_ID,
  GRAPH_CLIENT_ID,
  GRAPH_CLIENT_SECRET
);
const graphClient = Client.initWithMiddleware({
  authProvider: {
    getAccessToken: () =>
      credential
        .getToken("https://graph.microsoft.com/.default")
        .then((t) => t.token),
  },
});

async function sendEmail(to, subject, html) {
  await graphClient.api(`/users/${GRAPH_SENDER_EMAIL}/sendMail`).post({
    message: {
      subject,
      body: { contentType: "HTML", content: html },
      toRecipients: [{ emailAddress: { address: to } }],
    },
    saveToSentItems: "true",
  });
}

function cookieBaseOptions(req) {
  // You’re running HTTPS already; secure cookies are correct.
  // If you ever terminate TLS elsewhere, keep `app.set("trust proxy", 1)` (we do below).
  const base = {
    httpOnly: true,
    secure: true,
    sameSite: "lax",
  };

  // Optional cookie domain support (e.g. ".premierenergies.com")
  const host = (req.hostname || "").toLowerCase();
  const normalizedDomain = (COOKIE_DOMAIN || "")
    .replace(/^\./, "")
    .toLowerCase();
  const shouldSetDomain =
    normalizedDomain &&
    (host === normalizedDomain || host.endsWith(`.${normalizedDomain}`));

  return shouldSetDomain ? { ...base, domain: COOKIE_DOMAIN } : base;
}

function issueSession(email) {
  return jwt.sign({ sub: email, email, apps: ["pesa"] }, SESSION_JWT_SECRET, {
    expiresIn: "12h",
  });
}

function readSession(req) {
  const token = req.cookies?.[SESSION_COOKIE];
  if (!token) return null;
  try {
    return jwt.verify(token, SESSION_JWT_SECRET);
  } catch {
    return null;
  }
}
// Use given credentials, but allow env override if present
const dbConfig = {
  user: process.env.DB_USER || "PEL_DB",
  password: process.env.DB_PASSWORD || "Pel@0184",
  server: process.env.DB_SERVER || "10.0.50.17",
  port: Number(process.env.DB_PORT || 1433),
  database: process.env.DB_NAME || "pesa",
  // --- timeouts (ms) ---
  requestTimeout: 100000,
  connectionTimeout: 10000000,
  pool: {
    max: 10,
    min: 0,
    idleTimeoutMillis: 300000,
  },
  options: {
    encrypt: false,
    trustServerCertificate: true,
  },
};

// Path to frontend build (Vite default)
const FRONTEND_DIST_PATH = path.join(__dirname, "..", "dist");

let poolPromise = null;

async function getPool() {
  if (!poolPromise) {
    poolPromise = (async () => {
      const pool = new sql.ConnectionPool(dbConfig);
      pool.on("error", (err) => {
        console.error("[MSSQL] Pool error", err);
      });
      await pool.connect();
      console.log("[MSSQL] Connected to database:", dbConfig.database);
      await ensurePesaTable(pool);
      return pool;
    })().catch((err) => {
      console.error("[MSSQL] Connection error", err);
      // Reset so next call can retry
      poolPromise = null;
      throw err;
    });
  }
  return poolPromise;
}

async function ensurePesaTable(pool) {
  const ddl = `
IF NOT EXISTS (SELECT * FROM sys.schemas WHERE name = 'dbo')
BEGIN
  EXEC('CREATE SCHEMA dbo');
END;

IF OBJECT_ID('dbo.pesa_data', 'U') IS NULL
BEGIN
  CREATE TABLE dbo.pesa_data (
    id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
    rowKey NVARCHAR(450) NOT NULL,
    dpid NVARCHAR(100) NOT NULL,
    clientId NVARCHAR(100) NOT NULL,
    name NVARCHAR(255) NOT NULL,
    category NVARCHAR(100) NULL,
    dateKey NVARCHAR(50) NOT NULL,    -- e.g. "03-12-2025@@0-2"
    baseDate DATE NOT NULL,           -- parsed from the base part of dateKey
    value BIGINT NOT NULL,
    bought BIGINT NOT NULL,
    sold BIGINT NOT NULL,
    createdAt DATETIME2(3) NOT NULL
      CONSTRAINT DF_pesa_data_createdAt DEFAULT SYSUTCDATETIME()
  );
END;

/* Unique by logical row + dateKey to prevent duplicates on re-import without truncate */
IF NOT EXISTS (
  SELECT 1 FROM sys.indexes
  WHERE name = 'IX_pesa_data_rowKey_dateKey'
    AND object_id = OBJECT_ID('dbo.pesa_data')
)
BEGIN
  CREATE UNIQUE NONCLUSTERED INDEX IX_pesa_data_rowKey_dateKey
  ON dbo.pesa_data (rowKey, dateKey);
END;

/* Helpful for lookups by dpid/client and date */
IF NOT EXISTS (
  SELECT 1 FROM sys.indexes
  WHERE name = 'IX_pesa_data_dpid_client_baseDate'
    AND object_id = OBJECT_ID('dbo.pesa_data')
)
BEGIN
  CREATE NONCLUSTERED INDEX IX_pesa_data_dpid_client_baseDate
  ON dbo.pesa_data (dpid, clientId, baseDate);
END;

/* For time-based aggregations */
IF NOT EXISTS (
  SELECT 1 FROM sys.indexes
  WHERE name = 'IX_pesa_data_baseDate'
    AND object_id = OBJECT_ID('dbo.pesa_data')
)
BEGIN
  CREATE NONCLUSTERED INDEX IX_pesa_data_baseDate
  ON dbo.pesa_data (baseDate);
END;

-- OTP storage (no SPOT dependency)
IF OBJECT_ID('dbo.PesaLoginOTP', 'U') IS NULL
BEGIN
  CREATE TABLE dbo.PesaLoginOTP (
    Email NVARCHAR(256) NOT NULL PRIMARY KEY,
    OTP NVARCHAR(6) NOT NULL,
    OTP_Expiry DATETIME2(0) NOT NULL,
    CreatedAt DATETIME2(0) NOT NULL CONSTRAINT DF_PesaLoginOTP_CreatedAt DEFAULT SYSUTCDATETIME()
  );
END;
`;

  console.log("[DB] Ensuring dbo.pesa_data table & indexes exist...");
  await pool.request().batch(ddl);
  console.log("[DB] Schema check done.");
}

// Extract base DD-MM-YYYY from "DD-MM-YYYY@@i-pos"
function getBaseDate(dateKey) {
  const str = String(dateKey || "");
  const base = str.split("@@")[0];
  return base || str;
}

// Parse "DD-MM-YYYY" to JS Date
function parseDateStringToJsDate(dateStr) {
  if (!dateStr) return null;
  const s0 = String(dateStr).trim();
  if (!s0) return null;

  // Normalize separators
  const s = s0.replace(/\./g, "-").replace(/\//g, "-");

  // 1) DD-MM-YYYY or D-M-YYYY
  let m = /^(\d{1,2})-(\d{1,2})-(\d{4})$/.exec(s);
  if (m) {
    const dd = Number(m[1]);
    const mm = Number(m[2]);
    const yyyy = Number(m[3]);
    if (!dd || !mm || !yyyy) return null;
    const d = new Date(yyyy, mm - 1, dd);
    return Number.isNaN(d.getTime()) ? null : d;
  }

  // 2) YYYY-MM-DD
  m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(s);
  if (m) {
    const yyyy = Number(m[1]);
    const mm = Number(m[2]);
    const dd = Number(m[3]);
    if (!dd || !mm || !yyyy) return null;
    const d = new Date(yyyy, mm - 1, dd);
    return Number.isNaN(d.getTime()) ? null : d;
  }

  // 3) DD-MMM-YYYY (e.g. 03-Dec-2025)
  m = /^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/.exec(s);
  if (m) {
    const dd = Number(m[1]);
    const mon = m[2].toLowerCase();
    const yyyy = Number(m[3]);
    const months = {
      jan: 0,
      feb: 1,
      mar: 2,
      apr: 3,
      may: 4,
      jun: 5,
      jul: 6,
      aug: 7,
      sep: 8,
      oct: 9,
      nov: 10,
      dec: 11,
    };
    const mm = months[mon];
    if (mm === undefined || !dd || !yyyy) return null;
    const d = new Date(yyyy, mm, dd);
    return Number.isNaN(d.getTime()) ? null : d;
  }

  return null;
}

// Normalize name same as FE (trim whitespace and leading/trailing dots)
function normalizeName(name) {
  if (!name) return "";
  let n = String(name);
  n = n.trim();
  n = n.replace(/^[.\s]+/, "");
  n = n.replace(/[.\s]+$/, "");
  n = n.trim();
  return n;
}

// Extract fileIndex from dateKey like "03-12-2025@@1-2"
function getFileIndexFromDateKeyServer(dateKey) {
  const parts = String(dateKey || "").split("@@");
  if (parts.length < 2) return null;
  const meta = parts[1]; // e.g. "1-2"
  const idxStr = meta.split("-")[0];
  const idx = Number(idxStr);
  return Number.isNaN(idx) ? null : idx;
}

// Extract snapshot position from dateKey like "03-12-2025@@1-2" (returns 1 or 2)
function getSnapshotPosFromDateKeyServer(dateKey) {
  const parts = String(dateKey || "").split("@@");
  if (parts.length < 2) return null; // <-- if no @@, always null

  const meta = parts[1]; // e.g. "1-2"
  const posStr = meta.split("-")[1];
  const pos = Number(posStr);
  return Number.isNaN(pos) ? null : pos;
}

// DD-MM-YYYY -> Excel serial (1900 system)
function toExcelSerialFromDDMMYYYY(ddmmyyyy) {
  const dt = parseDateStringToJsDate(ddmmyyyy);
  if (!dt || Number.isNaN(dt.getTime())) return null;
  const utc = Date.UTC(dt.getFullYear(), dt.getMonth(), dt.getDate());
  const excelEpoch = Date.UTC(1899, 11, 30); // Excel day 0
  return (utc - excelEpoch) / 86400000;
}

// Format DD-MM-YYYY -> dropdown label like "12 Nov" (adds year only if multiple years exist)
function formatDropdownLabelFromBase(baseStr, includeYear) {
  const dt = parseDateStringToJsDate(baseStr);
  if (!dt || Number.isNaN(dt.getTime())) return String(baseStr || "");
  const day = String(dt.getDate()).padStart(2, "0");
  const mon = dt.toLocaleString("en-US", { month: "short" }); // Nov, Dec, etc.
  const y = dt.getFullYear();
  return includeYear ? `${day} ${mon} ${y}` : `${day} ${mon}`;
}

// Build the same matrix headers + rows as FE, but from DB recordset
function buildExportMatrixFromDbRows(dbRows) {
  // 1) Build holdings map + dateKeys set
  const holdingsMap = new Map(); // key -> { dpid, clientId, name, category, dateValues }
  const dateKeySet = new Set();

  for (const row of dbRows) {
    const dpid = row.dpid || "";
    const clientId = row.clientId || "";
    const nameRaw = row.name || "";
    const name = normalizeName(nameRaw);
    const category = row.category || "";
    const dateKey = String(row.dateKey || "");
    const value = Number(row.value || 0);
    const bought = Number(row.bought || 0);
    const sold = Number(row.sold || 0);

    dateKeySet.add(dateKey);

    const key = `${dpid}|${clientId}|${name}`;
    let holding = holdingsMap.get(key);
    if (!holding) {
      holding = {
        id: key,
        dpid,
        clientId,
        name,
        category,
        dateValues: {}, // dateKey -> { value, bought, sold }
      };
      holdingsMap.set(key, holding);
    }

    const existingDv = holding.dateValues[dateKey];
    if (!existingDv) {
      holding.dateValues[dateKey] = { value, bought, sold };
    } else {
      existingDv.value += value;
      existingDv.bought += bought;
      existingDv.sold += sold;
    }
  }

  const holdings = Array.from(holdingsMap.values());
  const allDateKeys = Array.from(dateKeySet);

  if (holdings.length === 0 || allDateKeys.length === 0) {
    return {
      headers: ["DPID", "CLIENT-ID", "CATEGORY", "NAME", "INITIAL HOLDING"],
      rows: [],
    };
  }

  // 2) Build file groups similar to FE (by fileIndex, then baseDate asc)
  const grouped = new Map(); // fileIndex -> dateKeys[]
  for (const dk of allDateKeys) {
    const fi = getFileIndexFromDateKeyServer(dk);
    const key = fi ?? -1;
    const arr = grouped.get(key);
    if (arr) {
      arr.push(dk);
    } else {
      grouped.set(key, [dk]);
    }
  }

  const sortedFileIndexes = Array.from(grouped.keys()).sort((a, b) => a - b);

  function parseBaseDate(dateKey) {
    const baseStr = getBaseDate(dateKey);
    const jsDate = parseDateStringToJsDate(baseStr);
    return jsDate ? jsDate.getTime() : 0;
  }

  const fileGroups = sortedFileIndexes.map((fi) => {
    const arr = grouped.get(fi);
    arr.sort((a, b) => parseBaseDate(a) - parseBaseDate(b));
    return { fileIndex: fi, dateKeys: arr };
  });

  // 3) Compute initialHolding per holding
  function pickInitialHolding(holding) {
    let bestDateKey = null;
    let bestTime = Infinity;

    for (const dk of Object.keys(holding.dateValues || {})) {
      const t = parseBaseDate(dk);
      if (t < bestTime) {
        bestTime = t;
        bestDateKey = dk;
      }
    }

    if (!bestDateKey) return 0;
    return Number(holding.dateValues[bestDateKey]?.value || 0);
  }

  // 4) Build headers
  const headers = ["DPID", "CLIENT-ID", "CATEGORY", "NAME", "INITIAL HOLDING"];

  fileGroups.forEach((group, groupIndex) => {
    const firstKey = group.dateKeys[0];
    const secondKey = group.dateKeys[1];

    const firstBase = firstKey ? getBaseDate(firstKey) : "";
    const secondBase = secondKey ? getBaseDate(secondKey) : "";

    headers.push(
      firstBase ? `AS ON ${firstBase}` : `FILE ${groupIndex + 1} HOLDING 1`
    );
    headers.push("B/S");
    headers.push(
      secondBase ? `AS ON ${secondBase}` : `FILE ${groupIndex + 1} HOLDING 2`
    );
  });

  // 5) Build rows
  const rows = holdings.map((h) => {
    const initialHolding = pickInitialHolding(h);
    const row = [h.dpid, h.clientId, h.category, h.name, initialHolding];

    fileGroups.forEach((group, groupIndex) => {
      const firstKey = group.dateKeys[0];
      const secondKey = group.dateKeys[1];

      const dv1 = firstKey ? h.dateValues[firstKey] : undefined;
      const dv2 = secondKey ? h.dateValues[secondKey] : undefined;

      // AS ON first date
      row.push(dv1 ? Number(dv1.value || 0) : 0);

      // B/S from second date
      if (dv2) {
        const bought = Number(dv2.bought || 0);
        const sold = Number(dv2.sold || 0);
        const bs =
          (bought > 0 ? `+${bought}` : "") + (sold > 0 ? `-${sold}` : "");
        row.push(bs || "-");
      } else {
        row.push("-");
      }

      // AS ON second date
      row.push(dv2 ? Number(dv2.value || 0) : 0);
    });

    return row;
  });

  return { headers, rows };
}

function toSafeInt(v) {
  if (v === null || v === undefined) return 0;

  if (typeof v === "number") {
    return Number.isFinite(v) ? Math.trunc(v) : 0;
  }

  // if some lib returns bigint
  if (typeof v === "bigint") {
    // keep within JS safe integer range; otherwise return as Number may lose precision
    // for your use-case shares typically fit; if not, consider DECIMAL(38,0) in DB.
    const n = Number(v);
    return Number.isFinite(n) ? n : 0;
  }

  // strings like "1,23,456", "12,345", "-", ""
  const s = String(v).trim();
  if (!s || s === "-" || s.toLowerCase() === "na" || s.toLowerCase() === "null")
    return 0;

  // remove commas (both 1,234 and 1,23,456)
  const cleaned = s.replace(/,/g, "");

  // allow negative too
  const n = Number(cleaned);
  return Number.isFinite(n) ? Math.trunc(n) : 0;
}

function parsePackedBS(x) {
  // supports: "+123-45", "+123 -45", "123/45", "123-45", "B:123 S:45"
  const s = String(x || "").trim();
  if (!s) return { bought: 0, sold: 0 };

  // Case 1: explicit + and - tokens
  const plus = /(?:\+|b[:\s]?)(\d[\d,]*)/i.exec(s);
  const minus = /(?:-|s[:\s]?)(\d[\d,]*)/i.exec(s);

  const bought = plus ? toSafeInt(plus[1]) : 0;
  const sold = minus ? toSafeInt(minus[1]) : 0;

  if (bought || sold) return { bought, sold };

  // Case 2: "123/45" or "123-45" meaning bought/sold
  const m = /^(\d[\d,]*)\s*[/\-]\s*(\d[\d,]*)$/.exec(s);
  if (m) return { bought: toSafeInt(m[1]), sold: toSafeInt(m[2]) };

  return { bought: 0, sold: 0 };
}

function normalizeDateValue(dv) {
  if (dv === null || dv === undefined) return { value: 0, bought: 0, sold: 0 };

  // ✅ Arrays are VERY common from parsers: [value, bought, sold] OR [value, "B/S"]
  if (Array.isArray(dv)) {
    const v = dv[0];
    const b = dv[1];
    const s = dv[2];

    // if second is a packed B/S string like "+123-45"
    if (typeof b === "string" && dv.length === 2) {
      const bs = parsePackedBS(b);
      return { value: toSafeInt(v), bought: bs.bought, sold: bs.sold };
    }

    return {
      value: toSafeInt(v),
      bought: toSafeInt(b),
      sold: toSafeInt(s),
    };
  }

  // primitives => holding value only
  if (
    typeof dv === "number" ||
    typeof dv === "bigint" ||
    typeof dv === "string"
  ) {
    return { value: toSafeInt(dv), bought: 0, sold: 0 };
  }

  if (typeof dv !== "object") return { value: 0, bought: 0, sold: 0 };

  // object: try multiple keys
  const v =
    dv.value ??
    dv.Value ??
    dv.holding ??
    dv.Holding ??
    dv.qty ??
    dv.Qty ??
    dv.quantity ??
    dv.Quantity ??
    dv.v ??
    dv.V ??
    0;

  // bought/sold can be separate OR in a single field (bs)
  const b =
    dv.bought ??
    dv.Bought ??
    dv.buy ??
    dv.Buy ??
    dv.purchase ??
    dv.Purchase ??
    dv.b ??
    dv.B ??
    0;

  const s =
    dv.sold ??
    dv.Sold ??
    dv.sell ??
    dv.Sell ??
    dv.sale ??
    dv.Sale ??
    dv.s ??
    dv.S ??
    0;

  // if b is a packed string and sold is empty, parse it
  if (typeof b === "string" && (!s || String(s).trim() === "")) {
    const bs = parsePackedBS(b);
    return { value: toSafeInt(v), bought: bs.bought, sold: bs.sold };
  }

  return { value: toSafeInt(v), bought: toSafeInt(b), sold: toSafeInt(s) };
}

// Very basic payload validation/logging
function validateImportPayload(body) {
  if (!body || typeof body !== "object") {
    return "Body must be a JSON object";
  }
  const { holdings, dates } = body;
  if (!Array.isArray(holdings)) {
    return "`holdings` must be an array";
  }
  if (!Array.isArray(dates)) {
    return "`dates` must be an array";
  }
  return null;
}

async function persistHoldingsToDb(holdings, dates) {
  const pool = await getPool();

  console.log(
    `[DB] Import request: holdings=${holdings.length}, dateKeys=${dates.length}`
  );

  // Strategy: replace-all (same as FE clearing IndexedDB)
  console.log("[DB] Truncating dbo.pesa_data before bulk insert...");
  await pool.request().query("TRUNCATE TABLE dbo.pesa_data;");

  const table = new sql.Table("dbo.pesa_data");
  table.create = false;

  table.columns.add("rowKey", sql.NVarChar(450), { nullable: false });
  table.columns.add("dpid", sql.NVarChar(100), { nullable: false });
  table.columns.add("clientId", sql.NVarChar(100), { nullable: false });
  table.columns.add("name", sql.NVarChar(255), { nullable: false });
  table.columns.add("category", sql.NVarChar(100), { nullable: true });
  table.columns.add("dateKey", sql.NVarChar(50), { nullable: false });
  table.columns.add("baseDate", sql.Date, { nullable: false });
  table.columns.add("value", sql.BigInt, { nullable: false });
  table.columns.add("bought", sql.BigInt, { nullable: false });
  table.columns.add("sold", sql.BigInt, { nullable: false });

  let totalRows = 0;
  let skippedBadDate = 0;
  let sampleBadDate = null;

  let nonZeroBought = 0;
  let nonZeroSold = 0;

  for (const holding of holdings) {
    if (!holding || typeof holding !== "object") continue;

    const rowKey =
      holding.id ||
      `${holding.dpid || ""}-${holding.clientId || ""}-${holding.name || ""}`;

    const dpid = String(holding.dpid || "");
    const clientId = String(holding.clientId || "");
    const name = String(holding.name || "");
    const category =
      typeof holding.category === "string" && holding.category.length > 0
        ? holding.category
        : null;

    const dateValues = holding.dateValues || {};
    const entries = Object.entries(dateValues);

    for (const [dateKey, dv] of entries) {
      if (dv === null || dv === undefined) continue;

      const baseDateStr = getBaseDate(dateKey);
      const baseDateJs = parseDateStringToJsDate(baseDateStr);
      if (!baseDateJs) {
        skippedBadDate++;
        if (!sampleBadDate) sampleBadDate = { dateKey, baseDateStr };
        continue;
      }

      const ndv = normalizeDateValue(dv);
      const value = ndv.value;
      const bought = ndv.bought;
      const sold = ndv.sold;

      if (bought) nonZeroBought++;
      if (sold) nonZeroSold++;

      table.rows.add(
        rowKey,
        dpid,
        clientId,
        name,
        category,
        String(dateKey),
        baseDateJs,
        value,
        bought,
        sold
      );

      totalRows++;
    }
  }

  if (totalRows === 0) {
    console.log(
      "[DB] No rows to insert (dateValues empty). Leaving table empty."
    );
    return { rowsInserted: 0, dbStats: null, topRows: [] };
  }

  if (skippedBadDate > 0) {
    console.warn(
      `[DB] Skipped ${skippedBadDate} rows due to unparseable dates. Sample:`,
      sampleBadDate
    );
  }

  console.log(
    `[DB] Bulk inserting ${totalRows} rows into dbo.pesa_data... (nonZeroBoughtRows=${nonZeroBought}, nonZeroSoldRows=${nonZeroSold})`
  );
  await pool.request().bulk(table);
  console.log("[DB] Bulk insert complete.");

  const stats = await pool.request().query(`
    SELECT
      COUNT(1) AS rowsCount,
      SUM(CAST(value AS BIGINT)) AS sumValue,
      SUM(CAST(bought AS BIGINT)) AS sumBought,
      SUM(CAST(sold AS BIGINT)) AS sumSold,
      MAX(CAST(value AS BIGINT)) AS maxValue,
      MAX(CAST(bought AS BIGINT)) AS maxBought,
      MAX(CAST(sold AS BIGINT)) AS maxSold
    FROM dbo.pesa_data WITH (NOLOCK);
  `);

  const top = await pool.request().query(`
    SELECT TOP 5 dpid, clientId, name, category, dateKey, baseDate, value, bought, sold
    FROM dbo.pesa_data WITH (NOLOCK)
    ORDER BY (CASE WHEN value < 0 THEN -value ELSE value END) DESC;
  `);

  return {
    rowsInserted: totalRows,
    dbStats: stats.recordset?.[0] || null,
    topRows: top.recordset || [],
  };
}

const app = express();

// Middlewares
app.use(helmet());
app.set("trust proxy", 1);
app.use(
  cors({
    // reflect origin (works with credentials); same-origin requests ignore CORS anyway
    origin: true,
    credentials: true,
  })
);
app.use(compression());
app.use(express.json({ limit: "25mb" }));
app.use(cookieParser());
app.use(
  morgan("dev", {
    skip: () => process.env.NODE_ENV === "test",
  })
);

// ------------------------ API ROUTES ------------------------

const STATIC_DIR = path.resolve(__dirname, "../dist");
if (fs.existsSync(STATIC_DIR)) {
  app.use(express.static(STATIC_DIR));

  app.get("/", (req, res) => {
    res.sendFile(path.join(STATIC_DIR, "index.html"));
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// AUTH GATE for API (protect all /api/* except auth/session/health)
// ─────────────────────────────────────────────────────────────────────────────
app.use((req, res, next) => {
  if (!req.path.startsWith("/api/") && !req.path.startsWith("/auth/"))
    return next();

  // Always allow:
  if (
    req.path === "/api/health" ||
    req.path === "/health" ||
    req.path === "/api/session" ||
    req.path === "/api/send-otp" ||
    req.path === "/api/verify-otp" ||
    req.path === "/auth/logout"
  ) {
    return next();
  }

  const session = readSession(req);
  if (!session?.email) {
    return res.status(401).json({ error: "unauthenticated" });
  }

  // Optional: enforce allow-list even for existing sessions
  if (!isAllowed(String(session.email))) {
    return res.status(403).json({ error: "forbidden" });
  }

  req.user = session;
  return next();
});

// ─────────────────────────────────────────────────────────────────────────────
// Session endpoint (frontend uses it on boot)
// ─────────────────────────────────────────────────────────────────────────────
app.get("/api/session", (req, res) => {
  const s = readSession(req);
  if (!s?.email) return res.status(401).json({ error: "unauthenticated" });
  return res.json({ user: { email: s.email, apps: s.apps || ["pesa"] } });
});

app.post("/auth/logout", (req, res) => {
  const base = cookieBaseOptions(req);
  res.clearCookie(SESSION_COOKIE, { ...base, path: "/" });
  if (COOKIE_DOMAIN) {
    // clear with explicit domain too (some browsers store it that way)
    res.clearCookie(SESSION_COOKIE, {
      ...base,
      path: "/",
      domain: COOKIE_DOMAIN,
    });
  }
  res.json({ ok: true });
});

// ─────────────────────────────────────────────────────────────────────────────
// OTP: SEND
// ─────────────────────────────────────────────────────────────────────────────
app.post("/api/send-otp", async (req, res) => {
  const email = normalizeEmail(req.body?.email);
  if (!email) return res.status(400).json({ message: "Missing email" });

  if (!isAllowed(email)) {
    return res
      .status(403)
      .json({ message: "Access denied: this app is restricted." });
  }

  const otp = Math.floor(100000 + Math.random() * 900000).toString();
  const expiry = new Date(Date.now() + 5 * 60 * 1000); // 5 mins

  try {
    const pool = await getPool();

    await pool
      .request()
      .input("Email", sql.NVarChar(256), email)
      .input("OTP", sql.NVarChar(6), otp)
      .input("Exp", sql.DateTime2(0), expiry).query(`
          MERGE dbo.PesaLoginOTP AS T
          USING (SELECT @Email AS Email) AS S
          ON (T.Email = S.Email)
          WHEN MATCHED THEN
            UPDATE SET OTP=@OTP, OTP_Expiry=@Exp, CreatedAt=SYSUTCDATETIME()
          WHEN NOT MATCHED THEN
            INSERT (Email, OTP, OTP_Expiry) VALUES (@Email, @OTP, @Exp);
        `);

    const baseUrl = `${req.protocol}://${req.get("host")}`;
    const subject = "Your PESA One-Time Password (OTP)";
    const html = `
        <div style="margin:0;padding:0;background:#f5f7fb;">
          <table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="background:#f5f7fb;padding:24px 0;">
            <tr>
              <td align="center">
                <table role="presentation" cellpadding="0" cellspacing="0" width="640" style="width:640px;max-width:94vw;background:#ffffff;border:1px solid #e6eaf2;border-radius:14px;overflow:hidden;font-family:Arial,Helvetica,sans-serif;">
                  <tr>
                    <td style="background:linear-gradient(90deg,#0f172a,#1e293b);padding:18px 22px;color:#fff;">
                      <div style="font-size:14px;opacity:.9;">Premier Energies</div>
                      <div style="font-size:20px;font-weight:700;margin-top:2px;">PESA • Secure Login</div>
                    </td>
                  </tr>
                  <tr>
                    <td style="padding:22px;">
                      <div style="font-size:14px;color:#334155;">Hi,</div>
                      <div style="font-size:14px;color:#334155;margin-top:10px;line-height:1.6;">
                        Use the OTP below to sign in to <strong>PESA</strong>. This code expires in <strong>5 minutes</strong>.
                      </div>
  
                      <div style="margin:18px 0 10px;border:1px solid #e6eaf2;border-radius:12px;background:#f8fafc;padding:16px;text-align:center;">
                        <div style="font-size:12px;color:#64748b;letter-spacing:.12em;text-transform:uppercase;">One-Time Password</div>
                        <div style="font-size:34px;font-weight:800;letter-spacing:.18em;color:#0f172a;margin-top:6px;">${otp}</div>
                      </div>
  
                      <div style="font-size:12px;color:#64748b;line-height:1.6;">
                        Requested for: <span style="color:#0f172a;font-weight:600;">${email}</span><br/>
                        If you didn’t request this, you can safely ignore this email.
                      </div>
  
                      <div style="margin-top:18px;">
                        <a href="${baseUrl}/login" style="display:inline-block;background:#2563eb;color:#fff;text-decoration:none;padding:10px 14px;border-radius:10px;font-size:14px;font-weight:700;">
                          Open PESA
                        </a>
                      </div>
  
                      <div style="margin-top:20px;border-top:1px solid #eef2f7;padding-top:14px;font-size:12px;color:#94a3b8;line-height:1.6;">
                        Security tip: Never share OTPs with anyone. Premier Energies staff will never ask for your OTP.
                      </div>
                    </td>
                  </tr>
                  <tr>
                    <td style="background:#f8fafc;border-top:1px solid #eef2f7;padding:14px 22px;color:#64748b;font-size:12px;">
                      © ${new Date().getFullYear()} Premier Energies • Automated message
                    </td>
                  </tr>
                </table>
              </td>
            </tr>
          </table>
        </div>
      `;

    await sendEmail(email, subject, html);
    return res.json({ ok: true, message: "OTP sent successfully" });
  } catch (err) {
    console.error("send-otp error:", err?.stack || err);
    return res.status(500).json({ message: "Server error" });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// OTP: VERIFY
// ─────────────────────────────────────────────────────────────────────────────
app.post("/api/verify-otp", async (req, res) => {
  const email = normalizeEmail(req.body?.email);
  const otp = String(req.body?.otp || "").trim();
  if (!email) return res.status(400).json({ message: "Missing email" });
  if (!otp) return res.status(400).json({ message: "Missing OTP" });

  if (!isAllowed(email)) {
    return res
      .status(403)
      .json({ message: "Access denied: this app is restricted." });
  }

  try {
    const pool = await getPool();
    const r = await pool
      .request()
      .input("Email", sql.NVarChar(256), email)
      .input("OTP", sql.NVarChar(6), otp).query(`
          SELECT OTP_Expiry
            FROM dbo.PesaLoginOTP WITH (NOLOCK)
           WHERE Email=@Email AND OTP=@OTP;
        `);

    if (!r.recordset?.length) {
      return res.status(400).json({ message: "Invalid OTP" });
    }

    const exp = new Date(r.recordset[0].OTP_Expiry);
    if (new Date() > exp) {
      return res.status(400).json({ message: "OTP expired" });
    }

    // burn the OTP (best practice)
    await pool
      .request()
      .input("Email", sql.NVarChar(256), email)
      .query(`DELETE FROM dbo.PesaLoginOTP WHERE Email=@Email;`);

    // issue session cookie
    const token = issueSession(email);
    const base = cookieBaseOptions(req);
    res.cookie(SESSION_COOKIE, token, {
      ...base,
      path: "/",
      maxAge: 12 * 60 * 60 * 1000,
    });

    return res.json({ ok: true, user: { email } });
  } catch (err) {
    console.error("verify-otp error:", err?.stack || err);
    return res.status(500).json({ message: "Server error" });
  }
});
/**
 * POST /api/pesa/import
 *
 * Expected JSON body:
 * {
 *   "holdings": HoldingRecord[],
 *   "dates": string[]   // dateKeys as used by FE
 * }
 *
 * This is designed to be compatible with your current FE:
 * after consolidateData(), you can POST { holdings: newHoldings, dates: newDates } here.
 */
app.post("/api/pesa/import", async (req, res) => {
  const error = validateImportPayload(req.body);
  if (error) {
    return res.status(400).json({ ok: false, error });
  }

  const { holdings, dates } = req.body;

  try {
    const { rowsInserted, dbStats, topRows } = await persistHoldingsToDb(
      holdings,
      dates
    );

    res.json({
      ok: true,
      holdingsCount: holdings.length,
      dateKeysCount: dates.length,
      rowsInserted,
      dbStats,
      topRows,
    });
  } catch (err) {
    console.error("[API] /api/pesa/import error", err);
    res
      .status(500)
      .json({ ok: false, error: "Failed to persist holdings to database" });
  }
});

app.post("/api/pesa/clear", async (req, res) => {
  try {
    const pool = await getPool();
    await pool.request().query("TRUNCATE TABLE dbo.pesa_data;");
    console.log("[API] /api/pesa/clear - table truncated");
    res.json({ ok: true });
  } catch (err) {
    console.error("[API] /api/pesa/clear error", err);
    res
      .status(500)
      .json({ ok: false, error: "Failed to clear PESA data table" });
  }
});

app.get("/api/pesa/export/xlsx", async (req, res) => {
  try {
    const pool = await getPool();
    const result = await pool.request().query(
      `
        SELECT dpid, clientId, name, category, dateKey, baseDate, value, bought, sold
        FROM dbo.pesa_data
        ORDER BY dpid, clientId, name, baseDate, dateKey
        `
    );
    const dbRows = result.recordset || [];
    const { headers, rows } = buildExportMatrixFromDbRows(dbRows);

    const worksheetData = [headers, ...rows];
    const ws = XLSX.utils.aoa_to_sheet(worksheetData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "PESA");

    const buffer = XLSX.write(wb, { bookType: "xlsx", type: "buffer" });

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
      "Content-Disposition",
      'attachment; filename="pesa-export.xlsx"'
    );

    res.send(buffer);
  } catch (err) {
    console.error("[API] /api/pesa/export/xlsx error", err);
    res
      .status(500)
      .json({ ok: false, error: "Failed to generate XLSX export" });
  }
});

/**
 * GET /api/pesa/rn
 * Generates the FULL Summary workbook from DB:
 * - Summary
 * - Controls
 * - SummaryDynamic (Excel formulas driven by Controls)
 * - ByDate (long format; one row per dbo.pesa_data row)
 * - Meta
 *
 * Designed for very large datasets (streaming XLSX writer).
 */
app.get("/api/pesa/rn", async (req, res) => {
  let workbook = null;

  try {
    const pool = await getPool();
    let dbStats = null;
    try {
      const s = await pool.request().query(`
        SELECT
          COUNT(1) AS rowsCount,
          SUM(CAST(value AS BIGINT)) AS sumValue,
          SUM(CAST(bought AS BIGINT)) AS sumBought,
          SUM(CAST(sold AS BIGINT)) AS sumSold,
          MAX(CAST(value AS BIGINT)) AS maxValue,
          MAX(CAST(bought AS BIGINT)) AS maxBought,
          MAX(CAST(sold AS BIGINT)) AS maxSold
        FROM dbo.pesa_data WITH (NOLOCK);
      `);
      dbStats = s.recordset?.[0] || null;
    } catch (e) {
      console.warn("[RN] dbStats query failed:", e?.message || e);
    }

    // Download headers
    const stamp = new Date().toISOString().replace(/[:.]/g, "-");
    const filename = `pesa-summary-${stamp}.xlsx`;

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    res.setHeader("Cache-Control", "no-store");

    // Streaming workbook -> response stream
    workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
      stream: res,
      useStyles: true,
      useSharedStrings: true,
    });

    // Ensure Excel recalculates formulas on open (ExcelJS does not compute formulas)
    workbook.calcProperties = { fullCalcOnLoad: true };

    // -------------------- Create sheets in the same order as FE export --------------------
    const wsSummary = workbook.addWorksheet("Summary");
    const wsControls = workbook.addWorksheet("Controls");
    const wsSummaryDynamic = workbook.addWorksheet("SummaryDynamic");
    const wsByDate = workbook.addWorksheet("ByDate");
    const wsMeta = workbook.addWorksheet("Meta");

    const wsLists = workbook.addWorksheet("Lists");
    // Hide helper sheet (keeps workbook clean)
    wsLists.state = "veryHidden";

    // -------------------- ByDate headers (matches FE exportSummaryToXLSX) --------------------
    const byDateHeaders = [
      "BaseDate",
      "BaseDateExcel", // real Excel date serial
      "SnapshotPos",
      "FileIndex",
      "DPID",
      "ClientID",
      "Category",
      "Name",
      "Key", // h.id
      "KeyDateRank", // BaseDateExcel*10 + SnapshotPos
      "KeyRank", // Key|KeyDateRank
      "Value",
      "Bought",
      "Sold",
      "DateKey",
    ];

    wsByDate.addRow(byDateHeaders).commit();
    wsByDate.autoFilter = {
      from: { row: 1, column: 1 },
      to: { row: 1, column: byDateHeaders.length },
    };

    // Helpful column widths (optional)
    wsByDate.columns = [
      { width: 12 }, // BaseDate
      { width: 14 }, // BaseDateExcel
      { width: 12 }, // SnapshotPos
      { width: 10 }, // FileIndex
      { width: 16 }, // DPID
      { width: 16 }, // ClientID
      { width: 14 }, // Category
      { width: 30 }, // Name
      { width: 32 }, // Key
      { width: 14 }, // KeyDateRank
      { width: 40 }, // KeyRank
      { width: 14 }, // Value
      { width: 12 }, // Bought
      { width: 12 }, // Sold
      { width: 22 }, // DateKey
    ];

    // -------------------- Aggregation state for Summary/Controls --------------------
    const aggMap = new Map();
    // Collect unique AS ON base dates (Excel serial -> baseStr)
    const serialToBaseStr = new Map(); // intSerial -> "DD-MM-YYYY"

    let totalDbRows = 0;

    let cntPos1 = 0;
    let cntPos2 = 0;
    let cntPos0 = 0;

    let globalMinSerial = Infinity;
    let globalMaxSerial = -Infinity;

    // Stream query (don’t load 3L rows into memory)
    const request = pool.request();
    request.stream = true;

    const query = `
      SELECT dpid, clientId, name, category, dateKey, baseDate, value, bought, sold
      FROM dbo.pesa_data
      ORDER BY baseDate, dateKey, dpid, clientId, name
    `;

    // IMPORTANT: attach listeners BEFORE query() in streaming mode
    request.on("row", (row) => {
      totalDbRows++;

      const dpid = String(row.dpid || "");
      const clientId = String(row.clientId || "");
      const nameNorm = normalizeName(String(row.name || ""));
      const category = row.category ? String(row.category) : "";

      const dateKey = String(row.dateKey || "");
      const baseStr = getBaseDate(dateKey); // DD-MM-YYYY
      const serial = toExcelSerialFromDDMMYYYY(baseStr);
      // Track unique base dates for dropdown (keep as integer serial)
      if (typeof serial === "number" && Number.isFinite(serial)) {
        const intSerial = Math.round(serial);
        if (!serialToBaseStr.has(intSerial)) {
          serialToBaseStr.set(intSerial, baseStr);
        }
      }

      const snapshotPosRaw = getSnapshotPosFromDateKeyServer(dateKey);
      const snapshotPos =
        typeof snapshotPosRaw === "number"
          ? snapshotPosRaw
          : String(dateKey).includes("@@")
          ? 0
          : 2; // if no @@ meta, treat as snapshot 2

      if (snapshotPos === 1) cntPos1++;
      else if (snapshotPos === 2) cntPos2++;
      else cntPos0++;

      const fileIndexRaw = getFileIndexFromDateKeyServer(dateKey);
      const fileIndex =
        typeof fileIndexRaw === "number"
          ? fileIndexRaw
          : String(dateKey).includes("@@")
          ? ""
          : 1; // if no @@ meta, default fileIndex=1

      const value = toSafeInt(row.value);
      const bought = toSafeInt(row.bought);
      const sold = toSafeInt(row.sold);

      const key = `${dpid}|${clientId}|${nameNorm}`;

      const rank =
        typeof serial === "number"
          ? serial * 10 + Number(snapshotPos || 0)
          : null;
      const keyRank = typeof rank === "number" ? `${key}|${rank}` : "";

      // Track global min/max for Controls
      if (typeof serial === "number") {
        if (serial < globalMinSerial) globalMinSerial = serial;
        if (serial > globalMaxSerial) globalMaxSerial = serial;
      }

      // Update per-key aggregation for Summary
      let agg = aggMap.get(key);
      if (!agg) {
        agg = {
          id: key,
          dpid,
          clientId,
          category: category || "",
          name: nameNorm,
          minRank: Infinity,
          minPos: 0,
          minValue: 0,
          minBought: 0,
          minSold: 0,
          maxRank: -Infinity,
          boughtSum: 0,
          soldSum: 0,
        };
        aggMap.set(key, agg);
      } else {
        // Prefer non-empty category
        if (!agg.category && category) agg.category = category;
      }

      if (typeof rank === "number") {
        if (rank < agg.minRank) {
          agg.minRank = rank;
          agg.minPos = snapshotPos;
          agg.minValue = value;
          agg.minBought = bought;
          agg.minSold = sold;
        }
        if (rank > agg.maxRank) {
          agg.maxRank = rank;
        }
      }

      if (snapshotPos === 2) {
        agg.boughtSum += bought;
        agg.soldSum += sold;
      }

      // Write ByDate row
      const outRow = wsByDate.addRow([
        baseStr,
        typeof serial === "number" ? serial : "",
        snapshotPos,
        fileIndex,
        dpid,
        clientId,
        category,
        nameNorm,
        key,
        typeof rank === "number" ? rank : "",
        keyRank,
        value,
        bought,
        sold,
        dateKey,
      ]);

      // Format BaseDateExcel as date
      const c2 = outRow.getCell(2);
      if (typeof c2.value === "number") c2.numFmt = "yyyy-mm-dd";

      outRow.commit();
    });

    request.on("error", (err) => {
      console.error("[API] /api/pesa/rn stream error", err);
      // If the stream errors mid-flight, Excel file will be incomplete.
      // Close the workbook to end response.
      try {
        wsByDate.commit();
        wsSummary.commit();
        wsControls.commit();
        wsSummaryDynamic.commit();
        wsMeta.commit();
      } catch {}
      try {
        workbook.commit();
      } catch {}
    });

    request.on("done", async () => {
      try {
        // -------------------- Build Summary rows --------------------
        const summaryHeaders = [
          "DPID",
          "ClientID",
          "Category",
          "Sold",
          "Name",
          "Bought",
          "Initial Holding",
          "Net B/S (Bought - Sold)",
          "Still Holding",
        ];

        wsSummary.columns = [
          { width: 16 },
          { width: 16 },
          { width: 14 },
          { width: 12 },
          { width: 30 },
          { width: 12 },
          { width: 16 },
          { width: 18 },
          { width: 16 },
        ];

        wsSummary.addRow(summaryHeaders).commit();
        wsSummary.autoFilter = {
          from: { row: 1, column: 1 },
          to: { row: 1, column: summaryHeaders.length },
        };

        // Compute per-key summary
        const rows = Array.from(aggMap.values()).map((a) => {
          // initial reconstruction matches FE logic:
          // if earliest snapshot is pos1 => initial=value
          // if earliest snapshot is pos2 => initial = value - (bought - sold)
          const initial =
            a.minPos === 1
              ? Number(a.minValue || 0)
              : Number(a.minValue || 0) -
                (Number(a.minBought || 0) - Number(a.minSold || 0));

          const boughtSum = Number(a.boughtSum || 0);
          const soldSum = Number(a.soldSum || 0);
          const net = boughtSum - soldSum;
          const still = initial + net;

          return {
            id: a.id,
            dpid: a.dpid,
            clientId: a.clientId,
            category: a.category || "",
            name: a.name,
            initial,
            bought: boughtSum,
            sold: soldSum,
            net,
            still,
          };
        });

        // Sort same default feel as UI (net desc)
        rows.sort((x, y) => (y.net || 0) - (x.net || 0));

        // Totals
        let tInitial = 0,
          tBought = 0,
          tSold = 0,
          tNet = 0,
          tStill = 0;

        for (const r of rows) {
          tInitial += r.initial;
          tBought += r.bought;
          tSold += r.sold;
          tNet += r.net;
          tStill += r.still;
        }

        // Totals row (row 2)
        wsSummary
          .addRow(["", "", "", tSold, "Total", tBought, tInitial, tNet, tStill])
          .commit();

        // Data rows
        for (const r of rows) {
          wsSummary
            .addRow([
              r.dpid,
              r.clientId,
              r.category,
              r.sold,
              r.name,
              r.bought,
              r.initial,
              r.net,
              r.still,
            ])
            .commit();
        }

        // -------------------- Controls sheet --------------------
        // Build sorted unique list of AS ON dates for dropdowns
        const serials = Array.from(serialToBaseStr.keys()).sort(
          (a, b) => a - b
        );

        // Determine if multiple years exist; if yes, include year in labels to avoid ambiguity
        const yearSet = new Set();
        for (const s of serials) {
          const baseStr = serialToBaseStr.get(s);
          const dt = parseDateStringToJsDate(baseStr);
          if (dt && !Number.isNaN(dt.getTime())) yearSet.add(dt.getFullYear());
        }
        const includeYearInLabel = yearSet.size > 1;

        // Populate Lists sheet (Label -> Serial)
        wsLists.columns = [{ width: 18 }, { width: 14 }];
        wsLists.addRow(["Label", "Serial"]).commit();

        for (const s of serials) {
          const baseStr = serialToBaseStr.get(s) || "";
          const label = formatDropdownLabelFromBase(
            baseStr,
            includeYearInLabel
          );

          const rr = wsLists.addRow([label, s]);
          // Serial column is an Excel date serial; format nicely (even though sheet is hidden)
          rr.getCell(2).numFmt = "yyyy-mm-dd";
          rr.commit();
        }

        // Default From/To = global min/max (as Excel date serial)
        const minSerial = Number.isFinite(globalMinSerial)
          ? Math.round(globalMinSerial)
          : "";
        const maxSerial = Number.isFinite(globalMaxSerial)
          ? Math.round(globalMaxSerial)
          : "";

        // Compute default labels from min/max (fallback to first/last in list)
        const effectiveMinSerial =
          typeof minSerial === "number" && serialToBaseStr.has(minSerial)
            ? minSerial
            : serials[0] ?? "";
        const effectiveMaxSerial =
          typeof maxSerial === "number" && serialToBaseStr.has(maxSerial)
            ? maxSerial
            : serials[serials.length - 1] ?? "";

        const defaultFromBase = serialToBaseStr.get(effectiveMinSerial) || "";
        const defaultToBase = serialToBaseStr.get(effectiveMaxSerial) || "";

        const defaultFromLabel = formatDropdownLabelFromBase(
          defaultFromBase,
          includeYearInLabel
        );
        const defaultToLabel = formatDropdownLabelFromBase(
          defaultToBase,
          includeYearInLabel
        );

        wsControls.columns = [{ width: 18 }, { width: 26 }];

        wsControls.addRow(["Parameter", "Value"]).commit();

        const listLastRow = serials.length + 1; // header row is 1
        const labelRange = `Lists!$A$2:$A$${listLastRow}`;
        const serialRange = `Lists!$B$2:$B$${listLastRow}`;

        // From/To as LABELS (dropdown) — MUST set validation BEFORE row.commit() in streaming mode
        const rFrom = wsControls.addRow(["From", defaultFromLabel]);
        rFrom.getCell(2).dataValidation = {
          type: "list",
          allowBlank: false,
          showInputMessage: true,
          promptTitle: "Select AS ON date",
          prompt:
            "Pick a value from the dropdown (pulled from all uploaded sheets).",
          showErrorMessage: true,
          errorTitle: "Invalid selection",
          error: "Please select a value from the dropdown list.",
          // ExcelJS expects range formula WITHOUT leading "="
          formulae: [labelRange],
        };
        rFrom.commit();

        const rTo = wsControls.addRow(["To", defaultToLabel]);
        rTo.getCell(2).dataValidation = {
          type: "list",
          allowBlank: false,
          showInputMessage: true,
          promptTitle: "Select AS ON date",
          prompt:
            "Pick a value from the dropdown (pulled from all uploaded sheets).",
          showErrorMessage: true,
          errorTitle: "Invalid selection",
          error: "Please select a value from the dropdown list.",
          formulae: [labelRange],
        };
        rTo.commit();

        wsControls.addRow([]).commit();

        // Start/End remain NUMERIC serials (used by SummaryDynamic formulas)
        const fromSerialLookup = `XLOOKUP(B2,${labelRange},${serialRange},"")`;
        const toSerialLookup = `XLOOKUP(B3,${labelRange},${serialRange},"")`;

        const rStart = wsControls.addRow([
          "Start",
          { formula: `MIN(${fromSerialLookup},${toSerialLookup})` },
        ]);
        rStart.getCell(2).numFmt = "yyyy-mm-dd";
        rStart.commit();

        const rEnd = wsControls.addRow([
          "End",
          { formula: `MAX(${fromSerialLookup},${toSerialLookup})` },
        ]);
        rEnd.getCell(2).numFmt = "yyyy-mm-dd";
        rEnd.commit();

        wsControls.addRow([]).commit();
        wsControls
          .addRow([
            "Tip",
            "Use the From/To dropdowns above. SummaryDynamic updates automatically.",
          ])
          .commit();

        // -------------------- SummaryDynamic sheet (formulas) --------------------
        // Use full-column ranges so we don't need last-row counts.
        const keyR = `ByDate!$I:$I`; // Key
        const dateR = `ByDate!$B:$B`; // BaseDateExcel
        const rankR = `ByDate!$J:$J`; // KeyDateRank
        const keyRankR = `ByDate!$K:$K`; // KeyRank
        const valR = `ByDate!$L:$L`; // Value
        const buyR = `ByDate!$M:$M`; // Bought
        const soldR = `ByDate!$N:$N`; // Sold
        const posR = `ByDate!$C:$C`; // SnapshotPos

        const dynHeaders = [
          "Key",
          "DPID",
          "ClientID",
          "Category",
          "Name",
          "Initial Holding",
          "Bought",
          "Sold",
          "Net (Bought-Sold)",
          "Still Holding",
        ];

        wsSummaryDynamic.columns = [
          { width: 32 },
          { width: 16 },
          { width: 16 },
          { width: 14 },
          { width: 30 },
          { width: 16 },
          { width: 12 },
          { width: 12 },
          { width: 14 },
          { width: 16 },
        ];

        wsSummaryDynamic.addRow(dynHeaders).commit();

        // Totals row at row 2; data starts at row 3
        const totalRow = wsSummaryDynamic.addRow([
          "",
          "",
          "",
          "",
          "Total",
          { formula: `SUM(F3:F1048576)` },
          { formula: `SUM(G3:G1048576)` },
          { formula: `SUM(H3:H1048576)` },
          { formula: `G2-H2` },
          { formula: `F2+I2` },
        ]);
        totalRow.commit();

        const startCell = "Controls!$B$5";
        const endCell = "Controls!$B$6";

        let excelRow = 3;
        for (const r of rows) {
          const keyCell = `A${excelRow}`;

          // FE-equivalent initial formula (LET + MINIFS + XLOOKUP)
          const initialFormula =
            `IFERROR(LET(` +
            `k,${keyCell},` +
            `st,${startCell},` +
            `en,${endCell},` +
            `sr,MINIFS(${rankR},${keyR},k,${dateR},">="&st,${dateR},"<="&en),` +
            `sp,XLOOKUP(k&"|"&sr,${keyRankR},${posR},0),` +
            `sv,XLOOKUP(k&"|"&sr,${keyRankR},${valR},0),` +
            `sb,XLOOKUP(k&"|"&sr,${keyRankR},${buyR},0),` +
            `ss,XLOOKUP(k&"|"&sr,${keyRankR},${soldR},0),` +
            `IF(sp=1,sv,sv-(sb-ss))` +
            `),0)`;

          const boughtSumFormula = `SUMIFS(${buyR},${keyR},${keyCell},${dateR},">="&${startCell},${dateR},"<="&${endCell},${posR},2)`;

          const soldSumFormula = `SUMIFS(${soldR},${keyR},${keyCell},${dateR},">="&${startCell},${dateR},"<="&${endCell},${posR},2)`;

          const rowOut = wsSummaryDynamic.addRow([
            r.id,
            r.dpid,
            r.clientId,
            r.category || "",
            r.name,
            { formula: initialFormula },
            { formula: boughtSumFormula },
            { formula: soldSumFormula },
            { formula: `G${excelRow}-H${excelRow}` },
            { formula: `F${excelRow}+I${excelRow}` },
          ]);
          rowOut.commit();
          excelRow++;
        }

        wsSummaryDynamic.autoFilter = {
          from: { row: 1, column: 1 },
          to: { row: 1, column: dynHeaders.length },
        };

        // -------------------- Meta sheet --------------------
        wsMeta.columns = [{ width: 24 }, { width: 60 }];
        wsMeta.addRow(["ExportedAt", new Date().toISOString()]).commit();
        wsMeta.addRow(["DB Rows (dbo.pesa_data)", totalDbRows]).commit();
        if (dbStats) {
          wsMeta
            .addRow(["DB Sum Value", String(dbStats.sumValue ?? "")])
            .commit();
          wsMeta
            .addRow(["DB Sum Bought", String(dbStats.sumBought ?? "")])
            .commit();
          wsMeta
            .addRow(["DB Sum Sold", String(dbStats.sumSold ?? "")])
            .commit();
          wsMeta
            .addRow(["DB Max Value", String(dbStats.maxValue ?? "")])
            .commit();
          wsMeta
            .addRow(["DB Max Bought", String(dbStats.maxBought ?? "")])
            .commit();
          wsMeta
            .addRow(["DB Max Sold", String(dbStats.maxSold ?? "")])
            .commit();
        }

        wsMeta.addRow(["SnapshotPos=1 rows", cntPos1]).commit();
        wsMeta.addRow(["SnapshotPos=2 rows", cntPos2]).commit();
        wsMeta.addRow(["SnapshotPos=other rows", cntPos0]).commit();

        wsMeta.addRow(["Unique Keys", rows.length]).commit();
        wsMeta.addRow(["Default Range From (serial)", minSerial]).commit();
        wsMeta.addRow(["Default Range To (serial)", maxSerial]).commit();
        wsMeta
          .addRow([
            "Endpoint",
            "/api/pesa/rn (generates Summary + Controls + SummaryDynamic + ByDate + Meta)",
          ])
          .commit();

        // -------------------- Commit sheets + workbook --------------------
        wsByDate.commit();
        wsSummary.commit();
        wsControls.commit();
        wsSummaryDynamic.commit();
        wsMeta.commit();
        wsLists.commit();

        await workbook.commit();
      } catch (err) {
        console.error("[API] /api/pesa/rn finalize error", err);
        try {
          wsByDate.commit();
          wsSummary.commit();
          wsControls.commit();
          wsSummaryDynamic.commit();
          wsMeta.commit();
        } catch {}
        try {
          await workbook.commit();
        } catch {}
      }
    });
    // Start streaming AFTER listeners are attached
    request.query(query);
  } catch (err) {
    console.error("[API] /api/pesa/rn error", err);
    // If we haven't started streaming XLSX properly, return JSON.
    // If headers were already sent, just end.
    if (!res.headersSent) {
      return res
        .status(500)
        .json({ ok: false, error: "Failed to generate RN summary XLSX" });
    }
    try {
      res.end();
    } catch {}
  }
});

// 404 fallback for unknown API routes
app.use("/api", (req, res) => {
  res.status(404).json({ ok: false, error: "API route not found" });
});

// ------------------------ FRONTEND STATIC ------------------------

// Serve frontend build (Vite dist) as static files
if (fs.existsSync(FRONTEND_DIST_PATH)) {
  console.log("[FE] Serving static files from:", FRONTEND_DIST_PATH);
  app.use(express.static(FRONTEND_DIST_PATH));
} else {
  console.warn(
    "[FE] Frontend dist folder not found at",
    FRONTEND_DIST_PATH,
    "— build FE before running server."
  );
}

// ------------------------ START HTTPS SERVER ------------------------

const httpsServer = https.createServer(httpsOptions, app);

httpsServer.listen(APP_PORT, APP_HOST, () => {
  console.log(
    `[PESA] HTTPS server listening on https://${APP_HOST}:${APP_PORT} (env: ${
      process.env.NODE_ENV || "development"
    })`
  );
});

// ------------------------ GRACEFUL SHUTDOWN ------------------------

function shutdown(signal) {
  console.log(`[PESA] Received ${signal}, shutting down gracefully...`);
  httpsServer.close(() => {
    console.log("[PESA] HTTPS server closed.");
    if (poolPromise) {
      poolPromise
        .then((pool) => pool.close())
        .then(() => {
          console.log("[MSSQL] Pool closed.");
          process.exit(0);
        })
        .catch((err) => {
          console.error("[MSSQL] Error closing pool", err);
          process.exit(1);
        });
    } else {
      process.exit(0);
    }
  });
}

process.on("SIGINT", () => shutdown("SIGINT"));
process.on("SIGTERM", () => shutdown("SIGTERM"));
