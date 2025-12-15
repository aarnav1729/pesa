"use strict";

require("dotenv").config();

const fs = require("fs");
const path = require("path");
const https = require("https");

const express = require("express");
const cors = require("cors");
const helmet = require("helmet");
const compression = require("compression");
const morgan = require("morgan");
const sql = require("mssql");
const XLSX = require("xlsx");

// HTTPS options (paths as requested)
const httpsOptions = {
  key: fs.readFileSync(path.join(__dirname, "certs", "mydomain.key")),
  cert: fs.readFileSync(path.join(__dirname, "certs", "d466aacf3db3f299.crt")),
  ca: fs.readFileSync(path.join(__dirname, "certs", "gd_bundle-g2-g1.crt")),
};

const APP_PORT = Number(process.env.PORT || 27443);
const APP_HOST = process.env.HOST || "0.0.0.0";

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
  const m = /(\d{2})-(\d{2})-(\d{4})/.exec(dateStr);
  if (!m) return null;
  const [, dd, mm, yyyy] = m;
  const day = Number(dd);
  const month = Number(mm);
  const year = Number(yyyy);
  if (!day || !month || !year) return null;
  return new Date(year, month - 1, day);
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
  // Do not auto-create table, it already exists
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
  // createdAt is defaulted by DB, no need to send

  let totalRows = 0;

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
      if (!dv) continue;
      const baseDateStr = getBaseDate(dateKey);
      const baseDateJs = parseDateStringToJsDate(baseDateStr);
      if (!baseDateJs) {
        console.warn(
          `[DB] Skipping invalid dateKey "${dateKey}" (base "${baseDateStr}") for rowKey="${rowKey}"`
        );
        continue;
      }

      const value = Number(dv.value || 0);
      const bought = Number(dv.bought || 0);
      const sold = Number(dv.sold || 0);

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
    return { rowsInserted: 0 };
  }

  console.log(`[DB] Bulk inserting ${totalRows} rows into dbo.pesa_data...`);
  await pool.request().bulk(table);
  console.log("[DB] Bulk insert complete.");

  return { rowsInserted: totalRows };
}

const app = express();

// Middlewares
app.use(helmet());
app.use(
  cors({
    // When FE is served from same origin, this is effectively a no-op.
    // You can restrict to specific origins if needed.
    origin: process.env.CORS_ORIGIN || "*",
  })
);
app.use(compression());
app.use(express.json({ limit: "25mb" }));
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
    const { rowsInserted } = await persistHoldingsToDb(holdings, dates);
    res.json({
      ok: true,
      holdingsCount: holdings.length,
      dateKeysCount: dates.length,
      rowsInserted,
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
