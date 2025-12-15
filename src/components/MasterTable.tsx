import {
  useState,
  useMemo,
  Fragment,
  ReactNode,
  useRef,
  useEffect,
  useDeferredValue,
} from "react";
import { createPortal } from "react-dom";
import {
  Search,
  ArrowUpDown,
  ArrowUp,
  ArrowDown,
  Filter,
  X,
  Trash2,
  Download,
  RotateCcw,
  ChevronRight,
  ChevronDown,
  Info,
} from "lucide-react";
import { Input } from "@/components/ui/input";
import { Button } from "@/components/ui/button";
import { HoldingRecord } from "@/lib/db";
import { cn } from "@/lib/utils";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";

interface MasterTableProps {
  holdings: HoldingRecord[];
  dates: string[]; // dateKeys (may include duplicates of base date)
  onClearData: () => void;
}

type SortConfig = {
  key: string;
  direction: "asc" | "desc";
} | null;

type SummaryTypeFilter = "all" | "buyers" | "sellers";
type SummarySortKey =
  | "initialHolding"
  | "bought"
  | "sold"
  | "net"
  | "stillHolding";

type SummaryFieldKey = "dpid" | "clientId" | "category" | "name";
type MatrixFieldKey = "dpid" | "clientId" | "category" | "name";

type FileGroup = {
  fileIndex: number;
  dateKeys: string[];
};

const PAGE_SIZE_OPTIONS = [50, 100, 200, 500, 1000];

// API base:
// - If you're serving FE from the same Express server, this works out of the box.
// - If FE is hosted elsewhere, set VITE_API_BASE_URL (e.g. https://your-domain.com)
const API_BASE_URL = (() => {
  try {
    const v = (import.meta as any)?.env?.VITE_API_BASE_URL;
    const base =
      typeof v === "string" && v.trim().length
        ? v.trim().replace(/\/+$/, "")
        : window.location.origin;
    return `${base}/api`;
  } catch {
    return `${window.location.origin}/api`;
  }
})();

function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

/** Extract base date from a dateKey like "03-12-2025@@1-1" */
function getBaseDate(dateKey: string): string {
  return String(dateKey).split("@@")[0] || dateKey;
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

/** Extract fileIndex from dateKey like "03-12-2025@@1-2" */
function getFileIndexFromDateKey(dateKey: string): number | null {
  const parts = String(dateKey).split("@@");
  if (parts.length < 2) return null;
  const meta = parts[1]; // e.g. "1-2"
  const idxStr = meta.split("-")[0];
  const idx = Number(idxStr);
  return Number.isNaN(idx) ? null : idx;
}

/** Extract snapshot position from dateKey like "03-12-2025@@1-2" (returns 1 or 2 if present). */
function getSnapshotPosFromDateKey(dateKey: string): number | null {
  const parts = String(dateKey).split("@@");
  if (parts.length < 2) return null;
  const meta = parts[1]; // e.g. "1-2"
  const posStr = meta.split("-")[1];
  const pos = Number(posStr);
  return Number.isNaN(pos) ? null : pos;
}

/** True if this dateKey represents the "second snapshot" column (where bought/sold deltas belong). */
function isSecondSnapshotKey(dateKey: string): boolean {
  return getSnapshotPosFromDateKey(dateKey) === 2;
}

function getValueColorClass(
  current: number,
  previous: number | undefined
): string {
  if (previous === undefined || previous === 0) return "";

  const change = ((current - previous) / previous) * 100;

  if (change > 50) return "cell-positive-4";
  if (change > 25) return "cell-positive-3";
  if (change > 10) return "cell-positive-2";
  if (change > 0) return "cell-positive-1";
  if (change < -50) return "cell-negative-4";
  if (change < -25) return "cell-negative-3";
  if (change < -10) return "cell-negative-2";
  if (change < 0) return "cell-negative-1";

  return "";
}

function getBuySellColorClass(bought: number, sold: number): string {
  const net = bought - sold;
  if (net > 0) return "cell-positive-2";
  if (net < 0) return "cell-negative-2";
  return "";
}

/** Name normalizer: trim periods and then whitespaces (leading/trailing). */
function normalizeName(name: string): string {
  if (!name) return "";
  let n = String(name);

  // First trim whitespace
  n = n.trim();
  // Then trim leading periods + surrounding whitespace
  n = n.replace(/^[.\s]+/, "");
  // Then trim trailing periods + surrounding whitespace
  n = n.replace(/[.\s]+$/, "");
  // Final whitespace trim
  n = n.trim();

  return n;
}

/** Combine duplicate holdings by normalized (DPID + CLIENT-ID + NAME). */
function combineHoldings(raw: HoldingRecord[]): HoldingRecord[] {
  const map = new Map<string, HoldingRecord>();

  for (const h of raw) {
    const normalized = normalizeName(h.name);
    const key = `${h.dpid}|${h.clientId}|${normalized}`;

    const existing = map.get(key);

    if (!existing) {
      // Clone to avoid mutating the original
      map.set(key, {
        ...h,
        id: key,
        name: normalized,
        dateValues: { ...h.dateValues },
      });
    } else {
      // Merge into existing
      for (const [dateKey, dv] of Object.entries(h.dateValues)) {
        const existingDV = existing.dateValues[dateKey];
        if (existingDV) {
          existingDV.value += dv.value;
          existingDV.bought += dv.bought;
          existingDV.sold += dv.sold;
        } else {
          existing.dateValues[dateKey] = { ...dv };
        }
      }

      // Prefer non-empty category
      if (!existing.category && h.category) {
        existing.category = h.category;
      }
    }
  }

  return Array.from(map.values());
}

/** Small inline hover info tooltip rendered in a portal (never clipped) */
function InfoHint({ content }: { content: ReactNode }) {
  const triggerRef = useRef<HTMLSpanElement | null>(null);
  const [open, setOpen] = useState(false);
  const [pos, setPos] = useState<{
    top: number;
    left: number;
    width: number;
  } | null>(null);

  const computePos = () => {
    const el = triggerRef.current;
    if (!el) return;

    const rect = el.getBoundingClientRect();

    const WIDTH = 320;
    const GAP = 8;
    const PADDING = 12;

    // Right-aligned to icon by default
    let left = rect.right - WIDTH;
    left = Math.max(
      PADDING,
      Math.min(left, window.innerWidth - WIDTH - PADDING)
    );

    // Prefer drop-down; if too low, flip upward
    let top = rect.bottom + GAP;
    const estimatedHeight = 220; // safe approx for your content
    if (top + estimatedHeight > window.innerHeight - PADDING) {
      top = Math.max(PADDING, rect.top - GAP - estimatedHeight);
    }

    setPos({ top, left, width: WIDTH });
  };

  const onEnter = () => {
    computePos();
    setOpen(true);
  };

  const onLeave = () => setOpen(false);

  const onFocus = () => {
    computePos();
    setOpen(true);
  };

  const onBlur = () => setOpen(false);

  const attachWindowListeners = open;

  if (attachWindowListeners && typeof window !== "undefined") {
    window.requestAnimationFrame(() => {
      computePos();
    });
  }

  return (
    <span
      ref={triggerRef}
      className="inline-flex items-center"
      onMouseEnter={onEnter}
      onMouseLeave={onLeave}
      onFocus={onFocus}
      onBlur={onBlur}
      tabIndex={0}
      role="button"
      aria-label="Info"
    >
      <Info className="w-4 h-4 text-muted-foreground hover:text-foreground transition-colors" />

      {open &&
        pos &&
        typeof document !== "undefined" &&
        createPortal(
          <div
            style={{
              position: "fixed",
              top: pos.top,
              left: pos.left,
              width: pos.width,
            }}
            className={cn(
              "z-[9999]",
              "rounded-lg border border-border bg-card shadow-enterprise-lg",
              "p-3 text-xs leading-relaxed text-muted-foreground",
              "animate-in fade-in-0 zoom-in-95"
            )}
          >
            {content}
          </div>,
          document.body
        )}
    </span>
  );
}

/** Reusable section header for collapsible blocks */
function SectionHeader({
  title,
  subtitle,
  open,
  onToggle,
  right,
  info,
}: {
  title: string;
  subtitle?: ReactNode;
  open: boolean;
  onToggle: () => void;
  right?: ReactNode;
  info?: ReactNode;
}) {
  return (
    <div className="px-4 py-3 bg-muted/30 border-b border-border relative overflow-visible">
      <div className="flex flex-col lg:flex-row lg:items-center gap-2 justify-between">
        <div className="flex items-start gap-2">
          <button
            onClick={onToggle}
            className={cn(
              "inline-flex items-center gap-2",
              "text-sm font-semibold text-foreground",
              "hover:text-primary transition-colors"
            )}
            title={open ? "Collapse section" : "Expand section"}
          >
            {open ? (
              <ChevronDown className="w-4 h-4 text-muted-foreground" />
            ) : (
              <ChevronRight className="w-4 h-4 text-muted-foreground" />
            )}
            <span>{title}</span>
          </button>

          {info && (
            <span className="mt-0.5">
              <InfoHint content={info} />
            </span>
          )}
        </div>

        <div className="flex items-center gap-2 flex-wrap">
          {subtitle && (
            <div className="text-xs text-muted-foreground">{subtitle}</div>
          )}
          {right}
        </div>
      </div>
    </div>
  );
}

/** Simple reusable pagination controls */
function PaginationControls({
  page,
  totalPages,
  pageSize,
  onPageChange,
  onPageSizeChange,
  pageSizeOptions = PAGE_SIZE_OPTIONS,
}: {
  page: number;
  totalPages: number;
  pageSize: number;
  onPageChange: (page: number) => void;
  onPageSizeChange: (size: number) => void;
  pageSizeOptions?: number[];
}) {
  const canPrev = page > 1;
  const canNext = page < totalPages;

  return (
    <div className="flex items-center gap-3 text-xs">
      <div className="flex items-center gap-1">
        <span className="text-muted-foreground">Rows per page:</span>
        <select
          className="h-7 rounded-md bg-background text-xs border border-border px-2"
          value={pageSize}
          onChange={(e) => onPageSizeChange(Number(e.target.value))}
        >
          {pageSizeOptions.map((opt) => (
            <option key={opt} value={opt}>
              {opt}
            </option>
          ))}
        </select>
      </div>

      <div className="flex items-center gap-1 text-muted-foreground">
        <span>
          Page <span className="text-foreground font-medium">{page}</span> of{" "}
          <span className="text-foreground font-medium">{totalPages}</span>
        </span>
        <Button
          variant="outline"
          size="icon"
          className="h-7 w-7"
          disabled={!canPrev}
          onClick={() => canPrev && onPageChange(page - 1)}
        >
          {"<"}
        </Button>
        <Button
          variant="outline"
          size="icon"
          className="h-7 w-7"
          disabled={!canNext}
          onClick={() => canNext && onPageChange(page + 1)}
        >
          {">"}
        </Button>
      </div>
    </div>
  );
}

export function MasterTable({
  holdings,
  dates,
  onClearData,
}: MasterTableProps) {
  // ---------------- Normalized / combined holdings ----------------
  const combinedHoldings = useMemo(() => combineHoldings(holdings), [holdings]);

  const [isClearing, setIsClearing] = useState(false);
  const [isExportingXlsx, setIsExportingXlsx] = useState(false);
  const [isExportingSummaryXlsx, setIsExportingSummaryXlsx] = useState(false);
  // Map of id -> {value, dateBase} for "Initial Holding"
  const initialHoldingMap = useMemo(() => {
    const map = new Map<string, { value: number; dateBase: string }>();

    for (const h of combinedHoldings) {
      let initialValue = 0;
      let initialDateBase = "";

      for (const dk of dates) {
        const dv = h.dateValues[dk];
        if (dv) {
          initialValue = dv.value;
          initialDateBase = getBaseDate(dk);
          break;
        }
      }

      map.set(h.id, { value: initialValue, dateBase: initialDateBase });
    }

    return map;
  }, [combinedHoldings, dates]);

  // ---------------- Collapsible states (default collapsed) ----------------
  const [summaryOpen, setSummaryOpen] = useState(false);
  const [matrixOpen, setMatrixOpen] = useState(false);

  // ---------------- Main matrix state ----------------
  const [searchTerm, setSearchTerm] = useState("");
  const deferredSearchTerm = useDeferredValue(searchTerm);

  const [sortConfig, setSortConfig] = useState<SortConfig>(null);
  const [filters, setFilters] = useState<Record<MatrixFieldKey, string>>({
    dpid: "",
    clientId: "",
    category: "",
    name: "",
  });

  const [dateOrder, setDateOrder] = useState<"asc" | "desc">("asc");

  // ---------------- Summary-only state ----------------
  const [summaryType, setSummaryType] = useState<SummaryTypeFilter>("all");
  const [summarySortKey, setSummarySortKey] = useState<SummarySortKey>("net");
  const [summarySortDir, setSummarySortDir] = useState<"asc" | "desc">("desc");
  const [summaryFromBase, setSummaryFromBase] = useState<string>("");
  const [summaryToBase, setSummaryToBase] = useState<string>("");

  // NEW: Excel-like filters for DPID / ClientID / Category / Name in summary
  const [summaryFieldFilters, setSummaryFieldFilters] = useState<
    Record<SummaryFieldKey, string>
  >({
    dpid: "",
    clientId: "",
    category: "",
    name: "",
  });

  // ---------------- Derived dates for summary ----------------
  const baseDatesSorted = useMemo(() => {
    const uniq = Array.from(new Set(dates.map(getBaseDate)));
    uniq.sort((a, b) => parseDate(a).getTime() - parseDate(b).getTime());
    return uniq;
  }, [dates]);

  const effectiveSummaryFrom = summaryFromBase || baseDatesSorted[0] || "";
  const effectiveSummaryTo =
    summaryToBase || baseDatesSorted[baseDatesSorted.length - 1] || "";

  const normalizedRange = useMemo(() => {
    if (!effectiveSummaryFrom || !effectiveSummaryTo) {
      return { from: "", to: "" };
    }
    const fromT = parseDate(effectiveSummaryFrom).getTime();
    const toT = parseDate(effectiveSummaryTo).getTime();
    if (fromT <= toT)
      return { from: effectiveSummaryFrom, to: effectiveSummaryTo };
    return { from: effectiveSummaryTo, to: effectiveSummaryFrom };
  }, [effectiveSummaryFrom, effectiveSummaryTo]);

  const dateKeysInSummaryRange = useMemo(() => {
    if (!normalizedRange.from || !normalizedRange.to) return dates;

    const fromT = parseDate(normalizedRange.from).getTime();
    const toT = parseDate(normalizedRange.to).getTime();

    return dates.filter((dk) => {
      const bd = getBaseDate(dk);
      const t = parseDate(bd).getTime();
      return t >= fromT && t <= toT;
    });
  }, [dates, normalizedRange.from, normalizedRange.to]);

  // ---------------- File groups for master matrix (per sheet) ----------------
  const fileGroupsAsc = useMemo<FileGroup[]>(() => {
    if (!dates.length) return [];
    const grouped = new Map<number, string[]>();

    for (const dk of dates) {
      const fi = getFileIndexFromDateKey(dk);
      const key = fi ?? -1;
      const existing = grouped.get(key);
      if (existing) {
        existing.push(dk);
      } else {
        grouped.set(key, [dk]);
      }
    }

    // Sort groups by fileIndex (ascending), and within each group by base date
    const sortedKeys = Array.from(grouped.keys()).sort((a, b) => a - b);
    const groups: FileGroup[] = sortedKeys.map((fi) => {
      const arr = grouped.get(fi)!;
      arr.sort(
        (a, b) =>
          parseDate(getBaseDate(a)).getTime() -
          parseDate(getBaseDate(b)).getTime()
      );
      return { fileIndex: fi, dateKeys: arr };
    });

    return groups;
  }, [dates]);

  const fileGroups = useMemo<FileGroup[]>(() => {
    return dateOrder === "asc" ? fileGroupsAsc : [...fileGroupsAsc].reverse();
  }, [fileGroupsAsc, dateOrder]);

  // ---------------- Summary period groups (range-aware, chronological) ----------------
  // We keep summary math in strict time order regardless of "Reverse Dates" UI.
  const summaryPeriodGroups = useMemo(() => {
    const inRange = new Set(dateKeysInSummaryRange);

    const groups = fileGroupsAsc
      .map((g) => {
        const k1 =
          g.dateKeys.find((k) => getSnapshotPosFromDateKey(k) === 1) || null;
        const k2 =
          g.dateKeys.find((k) => getSnapshotPosFromDateKey(k) === 2) || null;

        const firstKey = k1 && inRange.has(k1) ? k1 : null;
        const secondKey = k2 && inRange.has(k2) ? k2 : null;

        if (!firstKey && !secondKey) return null;
        return { fileIndex: g.fileIndex, firstKey, secondKey };
      })
      .filter(Boolean) as Array<{
      fileIndex: number;
      firstKey: string | null;
      secondKey: string | null;
    }>;

    return groups;
  }, [fileGroupsAsc, dateKeysInSummaryRange]);
  // ---------------- Unique values for main matrix filters ----------------
  const uniqueValues = useMemo(() => {
    return {
      dpid: [...new Set(combinedHoldings.map((h) => h.dpid))],
      clientId: [...new Set(combinedHoldings.map((h) => h.clientId))],
      category: [...new Set(combinedHoldings.map((h) => h.category))].filter(
        Boolean
      ),
      name: [...new Set(combinedHoldings.map((h) => h.name).filter(Boolean))],
    };
  }, [combinedHoldings]);

  // ---------------- Main matrix filtering/sorting ----------------
  const filteredData = useMemo(() => {
    let result = [...combinedHoldings];

    if (deferredSearchTerm) {
      const term = deferredSearchTerm.toLowerCase();
      result = result.filter(
        (h) =>
          h.name.toLowerCase().includes(term) ||
          h.dpid.toLowerCase().includes(term) ||
          h.clientId.toLowerCase().includes(term) ||
          h.category.toLowerCase().includes(term)
      );
    }

    Object.entries(filters).forEach(([key, value]) => {
      if (!value) return;
      const lower = value.toLowerCase();

      result = result.filter((h) => {
        if (key === "dpid") {
          return h.dpid.toLowerCase().includes(lower);
        }
        if (key === "clientId") {
          return h.clientId.toLowerCase().includes(lower);
        }
        if (key === "category") {
          return h.category ? h.category.toLowerCase().includes(lower) : false;
        }
        if (key === "name") {
          return h.name.toLowerCase().includes(lower);
        }
        return true;
      });
    });

    if (sortConfig) {
      result.sort((a, b) => {
        let aVal: string | number;
        let bVal: string | number;

        if (sortConfig.key === "name") {
          aVal = a.name;
          bVal = b.name;
        } else if (sortConfig.key === "dpid") {
          aVal = a.dpid;
          bVal = b.dpid;
        } else if (sortConfig.key === "clientId") {
          aVal = a.clientId;
          bVal = b.clientId;
        } else if (sortConfig.key === "category") {
          aVal = a.category;
          bVal = b.category;
        } else if (sortConfig.key === "initialHolding") {
          const ai = initialHoldingMap.get(a.id)?.value ?? 0;
          const bi = initialHoldingMap.get(b.id)?.value ?? 0;
          aVal = ai;
          bVal = bi;
        } else if (sortConfig.key.startsWith("date-")) {
          const dateKey = sortConfig.key.replace("date-", "");
          aVal = a.dateValues[dateKey]?.value || 0;
          bVal = b.dateValues[dateKey]?.value || 0;
        } else {
          return 0;
        }

        if (aVal < bVal) return sortConfig.direction === "asc" ? -1 : 1;
        if (aVal > bVal) return sortConfig.direction === "asc" ? 1 : -1;
        return 0;
      });
    }

    return result;
  }, [
    combinedHoldings,
    deferredSearchTerm,
    sortConfig,
    filters,
    initialHoldingMap,
  ]);

  // ---------------- Matrix pagination ----------------
  const [matrixPageSize, setMatrixPageSize] = useState(200);
  const [matrixPage, setMatrixPage] = useState(1);

  const matrixTotalPages = useMemo(
    () => Math.max(1, Math.ceil((filteredData.length || 0) / matrixPageSize)),
    [filteredData.length, matrixPageSize]
  );

  useEffect(() => {
    // reset to first page when data or size changes
    setMatrixPage(1);
  }, [matrixPageSize, filteredData.length]);

  const matrixPageRows = useMemo(() => {
    if (filteredData.length === 0) return [];
    const start = (matrixPage - 1) * matrixPageSize;
    return filteredData.slice(start, start + matrixPageSize);
  }, [filteredData, matrixPage, matrixPageSize]);

  const handleSort = (key: string) => {
    setSortConfig((prev) => {
      if (prev?.key === key) {
        if (prev.direction === "asc") return { key, direction: "desc" };
        return null;
      }
      return { key, direction: "asc" };
    });
  };

  const SortIcon = ({ columnKey }: { columnKey: string }) => {
    if (sortConfig?.key !== columnKey) {
      return <ArrowUpDown className="w-3.5 h-3.5 text-muted-foreground" />;
    }
    return sortConfig.direction === "asc" ? (
      <ArrowUp className="w-3.5 h-3.5 text-primary" />
    ) : (
      <ArrowDown className="w-3.5 h-3.5 text-primary" />
    );
  };

  const MatrixFilterButton = ({
    columnKey,
    label,
    values,
  }: {
    columnKey: MatrixFieldKey;
    label: string;
    values: string[];
  }) => {
    const currentValue = filters[columnKey] || "";

    const visibleOptions = values.filter((opt) =>
      currentValue
        ? opt.toLowerCase().includes(currentValue.toLowerCase())
        : true
    );

    return (
      <DropdownMenu>
        <DropdownMenuTrigger asChild>
          <button
            className={cn(
              "p-1 rounded hover:bg-muted transition-colors",
              currentValue && "text-primary bg-primary/10"
            )}
            title={`Filter ${label}`}
          >
            <Filter className="w-3 h-3" />
          </button>
        </DropdownMenuTrigger>
        <DropdownMenuContent
          align="start"
          className="w-48 max-h-72 overflow-hidden p-0"
        >
          {/* Search box */}
          <div className="p-2 border-b border-border">
            <Input
              autoFocus
              placeholder={`Search ${label}`}
              value={currentValue}
              onChange={(e) =>
                setFilters((prev) => ({
                  ...prev,
                  [columnKey]: e.target.value,
                }))
              }
              className="h-7 text-xs"
            />
          </div>

          {/* Clear filter */}
          <DropdownMenuItem
            onClick={() =>
              setFilters((prev) => ({
                ...prev,
                [columnKey]: "",
              }))
            }
          >
            <span className="text-muted-foreground text-xs">Clear filter</span>
          </DropdownMenuItem>

          {/* Values list */}
          <div className="max-h-52 overflow-y-auto">
            {visibleOptions.length === 0 && (
              <div className="px-3 py-2 text-xs text-muted-foreground">
                No matches
              </div>
            )}
            {visibleOptions.map((opt) => (
              <DropdownMenuItem
                key={opt}
                onClick={() =>
                  setFilters((prev) => ({
                    ...prev,
                    [columnKey]: opt,
                  }))
                }
              >
                <span className="text-xs">{opt}</span>
              </DropdownMenuItem>
            ))}
          </div>
        </DropdownMenuContent>
      </DropdownMenu>
    );
  };

  // ---------------- Export helpers (mirror sheet grouping) ----------------
  const buildExportMatrix = () => {
    const headers: string[] = [
      "DPID",
      "CLIENT-ID",
      "CATEGORY",
      "NAME",
      "INITIAL HOLDING",
    ];

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

    const rows = filteredData.map((h) => {
      const initialInfo = initialHoldingMap.get(h.id);
      const row: (string | number)[] = [
        h.dpid,
        h.clientId,
        h.category,
        h.name,
        initialInfo?.value ?? 0,
      ];

      fileGroups.forEach((group, groupIndex) => {
        const firstKey = group.dateKeys[0];
        const secondKey = group.dateKeys[1];

        const dv1 = firstKey ? h.dateValues[firstKey] : undefined;
        const dv2 = secondKey ? h.dateValues[secondKey] : undefined;

        row.push(dv1?.value ?? 0);

        if (dv2) {
          const bs =
            (dv2.bought > 0 ? `+${dv2.bought}` : "") +
            (dv2.sold > 0 ? `-${dv2.sold}` : "");
          row.push(bs || "-");
        } else {
          row.push("-");
        }

        row.push(dv2?.value ?? 0);
      });

      return row;
    });

    return { headers, rows };
  };

  const exportToCSV = () => {
    const { headers, rows } = buildExportMatrix();
    const csvContent = [
      headers.join(","),
      ...rows.map((r) =>
        r
          .map((cell) =>
            typeof cell === "string" && cell.includes(",")
              ? `"${cell.replace(/"/g, '""')}"`
              : String(cell)
          )
          .join(",")
      ),
    ].join("\n");

    const blob = new Blob([csvContent], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "pesa-export.csv";
    a.click();
    URL.revokeObjectURL(url);
  };

  const exportToXLSX = async () => {
    try {
      setIsExportingXlsx(true);

      const resp = await fetch("/api/pesa/export/xlsx", {
        method: "GET",
        cache: "no-store",
      });

      if (!resp.ok) {
        throw new Error(`Server responded with ${resp.status}`);
      }

      const blob = await resp.blob();
      const url = URL.createObjectURL(blob);

      const a = document.createElement("a");
      a.href = url;
      a.download = "pesa-export.xlsx";
      document.body.appendChild(a);
      a.click();
      a.remove();

      URL.revokeObjectURL(url);
    } catch (err) {
      console.error("Failed to export XLSX:", err);
      alert("Failed to export XLSX from server. Check console for details.");
    } finally {
      setIsExportingXlsx(false);
    }
  };

  const exportSummaryToXLSX = async () => {
    try {
      setIsExportingSummaryXlsx(true);

      // Lazy-load xlsx only when needed (keeps initial bundle lighter)
      const mod: any = await import("xlsx");
      const XLSX = mod?.default ? mod.default : mod;

      // Keep the same column order as the UI table
      const headers = [
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

      const totalsRow: (string | number)[] = [
        "",
        "",
        "",
        summaryTotals.sold,
        "Total",
        summaryTotals.bought,
        summaryTotals.initial,
        summaryTotals.net,
        summaryTotals.still,
      ];

      const dataRows = summaryRows.map((r) => [
        r.dpid,
        r.clientId,
        r.category || "",
        r.sold,
        r.name,
        r.bought,
        r.initialHolding,
        r.net,
        r.stillHolding,
      ]);

      const metaRows: (string | number)[][] = [
        ["Exported At", new Date().toISOString()],
        ["Summary Type", summaryType],
        ["Sort Key", summarySortKey],
        ["Sort Direction", summarySortDir],
        [
          "Range",
          normalizedRange.from && normalizedRange.to
            ? `${normalizedRange.from} → ${normalizedRange.to}`
            : "All dates",
        ],
        ["Filter: DPID", summaryFieldFilters.dpid || ""],
        ["Filter: ClientID", summaryFieldFilters.clientId || ""],
        ["Filter: Category", summaryFieldFilters.category || ""],
        ["Filter: Name", summaryFieldFilters.name || ""],
      ];

      const wsSummary = XLSX.utils.aoa_to_sheet([
        headers,
        totalsRow,
        ...dataRows,
      ]);
      const wsMeta = XLSX.utils.aoa_to_sheet(metaRows);

      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, wsSummary, "Summary");
      XLSX.utils.book_append_sheet(wb, wsMeta, "Meta");

      const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      const blob = new Blob([out], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });

      downloadBlob(blob, "pesa-summary.xlsx");
    } catch (err) {
      console.error("Failed to export Summary XLSX:", err);
      alert("Failed to export Summary XLSX. Check console for details.");
    } finally {
      setIsExportingSummaryXlsx(false);
    }
  };

  const toggleDateOrder = () => {
    setDateOrder((o) => (o === "asc" ? "desc" : "asc"));
  };

  const handleClearDataClick = async () => {
    const confirmed = window.confirm(
      "This will clear all imported PESA data from the server and this browser. Continue?"
    );
    if (!confirmed) return;

    try {
      setIsClearing(true);

      const resp = await fetch("/api/pesa/clear", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        cache: "no-store",
      });

      if (!resp.ok) {
        throw new Error(`Server responded with ${resp.status}`);
      }

      // Clear local/IndexedDB via parent callback
      onClearData();
    } catch (err) {
      console.error("Failed to clear PESA data:", err);
      alert("Failed to clear data on server. Check console for details.");
    } finally {
      setIsClearing(false);
    }
  };

  // ---------------- Latest date key (global) ----------------
  const latestDateKey = useMemo(() => {
    if (dates.length === 0) return "";

    let latestBase = "";
    let latestTime = -Infinity;

    for (const dk of dates) {
      const bd = getBaseDate(dk);
      const t = parseDate(bd).getTime();
      if (t > latestTime) {
        latestTime = t;
        latestBase = bd;
      }
    }

    for (let i = dates.length - 1; i >= 0; i--) {
      if (getBaseDate(dates[i]) === latestBase) return dates[i];
    }

    return dates[dates.length - 1];
  }, [dates]);

  // ---------------- Summary rows computation ----------------
  const rawSummaryRows = useMemo(() => {
    const groups = summaryPeriodGroups;

    return filteredData.map((h) => {
      // We compute summary by FILE-PERIODS (pos1 -> pos2), which is where deltas live.
      // This fixes the case where a name first appears in pos2 (no pos1 in-range):
      // baseline is reconstructed as: initial = dv2.value - (dv2.bought - dv2.sold)

      let startIdx = -1;
      let endIdx = -1;
      let endMode: "pos1" | "pos2" = "pos1";

      let initialHolding = 0;
      let endSnapshotValue = 0;

      for (let i = 0; i < groups.length; i++) {
        const g = groups[i];
        const dv1 = g.firstKey ? h.dateValues[g.firstKey] : undefined;
        const dv2 = g.secondKey ? h.dateValues[g.secondKey] : undefined;

        // Pick the earliest baseline inside range:
        // - Prefer pos1 value when present
        // - Else, if pos2 exists, reconstruct pos1 using dv2 - (bought - sold)
        if (startIdx === -1) {
          if (dv1 && dv1.value !== undefined) {
            startIdx = i;
            initialHolding = Number(dv1.value || 0);
          } else if (dv2 && dv2.value !== undefined) {
            startIdx = i;
            const b = Number(dv2.bought || 0);
            const s = Number(dv2.sold || 0);
            const v2 = Number(dv2.value || 0);
            initialHolding = v2 - (b - s);
          }
        }

        // Track the latest snapshot in-range (prefer pos2, else pos1)
        if (dv2 && dv2.value !== undefined) {
          endIdx = i;
          endMode = "pos2";
          endSnapshotValue = Number(dv2.value || 0);
        } else if (dv1 && dv1.value !== undefined) {
          endIdx = i;
          endMode = "pos1";
          endSnapshotValue = Number(dv1.value || 0);
        }
      }

      // If this name never appears in the selected range, keep zeros.
      if (startIdx === -1 || endIdx === -1) {
        return {
          id: h.id,
          dpid: h.dpid,
          clientId: h.clientId,
          category: h.category,
          name: h.name,
          initialHolding: 0,
          bought: 0,
          sold: 0,
          net: 0,
          stillHolding: 0,
        };
      }

      let totalBought = 0;
      let totalSold = 0;

      // Include dv2 deltas from start period through end snapshot.
      // If the end snapshot is pos1 for the last period, do NOT include that period's dv2 deltas.
      for (let i = startIdx; i <= endIdx; i++) {
        const g = groups[i];
        const dv2 = g.secondKey ? h.dateValues[g.secondKey] : undefined;
        if (!dv2) continue;

        const include = i < endIdx || (i === endIdx && endMode === "pos2");
        if (!include) continue;

        totalBought += Number(dv2.bought || 0);
        totalSold += Number(dv2.sold || 0);
      }

      const net = totalBought - totalSold;
      const expectedStill = initialHolding + net;

      // Airtight reconciliation:
      // we show "Still Holding" as the reconciled number so Initial + Bought - Sold always matches.
      // (If you ever want to expose the raw end snapshot too, it’s in `endSnapshotValue`.)
      const stillHolding = expectedStill;

      return {
        id: h.id,
        dpid: h.dpid,
        clientId: h.clientId,
        category: h.category,
        name: h.name,
        initialHolding,
        bought: totalBought,
        sold: totalSold,
        net,
        stillHolding,
      };
    });
  }, [filteredData, summaryPeriodGroups]);

  // Unique values for summary header filters (from visible summary base data)
  const summaryUniqueValues = useMemo(
    () => ({
      dpid: [...new Set(rawSummaryRows.map((r) => r.dpid))],
      clientId: [...new Set(rawSummaryRows.map((r) => r.clientId))],
      category: [
        ...new Set(rawSummaryRows.map((r) => r.category).filter(Boolean)),
      ] as string[],
      name: [
        ...new Set(rawSummaryRows.map((r) => r.name).filter(Boolean)),
      ] as string[],
    }),
    [rawSummaryRows]
  );

  const summaryRows = useMemo(() => {
    let result = [...rawSummaryRows];

    // 1) Apply Excel-like header filters (DPID / ClientID / Category / Name)
    Object.entries(summaryFieldFilters).forEach(([key, value]) => {
      if (!value) return;
      const lower = value.toLowerCase();
      if (key === "dpid") {
        result = result.filter((r) => r.dpid.toLowerCase().includes(lower));
      } else if (key === "clientId") {
        result = result.filter((r) => r.clientId.toLowerCase().includes(lower));
      } else if (key === "category") {
        result = result.filter(
          (r) => r.category && r.category.toLowerCase().includes(lower)
        );
      } else if (key === "name") {
        result = result.filter((r) => r.name.toLowerCase().includes(lower));
      }
    });

    // 2) Buyers / Sellers filter
    if (summaryType === "buyers") {
      result = result.filter((r) => r.bought > 0);
    } else if (summaryType === "sellers") {
      result = result.filter((r) => r.sold > 0);
    }

    // 3) Sort
    result.sort((a, b) => {
      const aVal = a[summarySortKey];
      const bVal = b[summarySortKey];

      if (aVal < bVal) return summarySortDir === "asc" ? -1 : 1;
      if (aVal > bVal) return summarySortDir === "asc" ? 1 : -1;
      return 0;
    });

    return result;
  }, [
    rawSummaryRows,
    summaryType,
    summarySortKey,
    summarySortDir,
    summaryFieldFilters,
  ]);

  // ---------------- Summary pagination ----------------
  const [summaryPageSize, setSummaryPageSize] = useState(200);
  const [summaryPage, setSummaryPage] = useState(1);

  const summaryTotalPages = useMemo(
    () => Math.max(1, Math.ceil((summaryRows.length || 0) / summaryPageSize)),
    [summaryRows.length, summaryPageSize]
  );

  useEffect(() => {
    setSummaryPage(1);
  }, [
    summaryRows.length,
    summaryPageSize,
    summaryType,
    summarySortKey,
    summarySortDir,
  ]);

  const summaryPageRows = useMemo(() => {
    if (summaryRows.length === 0) return [];
    const start = (summaryPage - 1) * summaryPageSize;
    return summaryRows.slice(start, start + summaryPageSize);
  }, [summaryRows, summaryPage, summaryPageSize]);

  const summaryTotals = useMemo(() => {
    let initial = 0;
    let bought = 0;
    let sold = 0;
    let net = 0;
    let still = 0;

    for (const r of summaryRows) {
      initial += r.initialHolding;
      bought += r.bought;
      sold += r.sold;
      net += r.net;
      still += r.stillHolding;
    }

    return { initial, bought, sold, net, still };
  }, [summaryRows]);

  const clearSummaryFilters = () => {
    setSummaryType("all");
    setSummarySortKey("net");
    setSummarySortDir("desc");
    setSummaryFromBase("");
    setSummaryToBase("");
    setSummaryFieldFilters({
      dpid: "",
      clientId: "",
      category: "",
      name: "",
    });
  };

  const SummarySortMenu = () => (
    <DropdownMenu>
      <DropdownMenuTrigger asChild>
        <Button variant="outline" size="sm" className="shadow-enterprise">
          <ArrowUpDown className="w-4 h-4" />
          Sort: {summarySortKey}
        </Button>
      </DropdownMenuTrigger>
      <DropdownMenuContent align="end">
        {(
          [
            "initialHolding",
            "bought",
            "sold",
            "net",
            "stillHolding",
          ] as SummarySortKey[]
        ).map((k) => (
          <DropdownMenuItem key={k} onClick={() => setSummarySortKey(k)}>
            {k}
          </DropdownMenuItem>
        ))}
        <DropdownMenuItem
          onClick={() =>
            setSummarySortDir((d) => (d === "asc" ? "desc" : "asc"))
          }
        >
          Direction: {summarySortDir === "asc" ? "Ascending" : "Descending"}
        </DropdownMenuItem>
      </DropdownMenuContent>
    </DropdownMenu>
  );

  const handleSummaryHeaderSort = (key: SummarySortKey) => {
    setSummarySortKey((prevKey) => {
      if (prevKey === key) {
        setSummarySortDir((prevDir) => (prevDir === "asc" ? "desc" : "asc"));
        return prevKey;
      }
      setSummarySortDir("desc");
      return key;
    });
  };

  const SummarySortIcon = ({ columnKey }: { columnKey: SummarySortKey }) => {
    if (summarySortKey !== columnKey) {
      return (
        <ArrowUpDown className="w-3.5 h-3.5 text-muted-foreground inline-block ml-1" />
      );
    }

    return summarySortDir === "asc" ? (
      <ArrowUp className="w-3.5 h-3.5 text-primary inline-block ml-1" />
    ) : (
      <ArrowDown className="w-3.5 h-3.5 text-primary inline-block ml-1" />
    );
  };

  const SummaryFilterButton = ({
    columnKey,
    label,
  }: {
    columnKey: SummaryFieldKey;
    label: string;
  }) => {
    const currentValue = summaryFieldFilters[columnKey];
    const allOptions = summaryUniqueValues[columnKey];

    const visibleOptions = allOptions.filter((opt) =>
      currentValue
        ? opt.toLowerCase().includes(currentValue.toLowerCase())
        : true
    );

    return (
      <DropdownMenu>
        <DropdownMenuTrigger asChild>
          <button
            className={cn(
              "p-1 rounded hover:bg-muted transition-colors",
              currentValue && "text-primary bg-primary/10"
            )}
            title={`Filter ${label}`}
          >
            <Filter className="w-3 h-3" />
          </button>
        </DropdownMenuTrigger>
        <DropdownMenuContent
          align="start"
          className="w-48 max-h-72 overflow-hidden p-0"
        >
          {/* Search box */}
          <div className="p-2 border-b border-border">
            <Input
              autoFocus
              placeholder={`Search ${label}`}
              value={currentValue}
              onChange={(e) =>
                setSummaryFieldFilters((prev) => ({
                  ...prev,
                  [columnKey]: e.target.value,
                }))
              }
              className="h-7 text-xs"
            />
          </div>

          {/* Clear filter */}
          <DropdownMenuItem
            onClick={() =>
              setSummaryFieldFilters((prev) => ({ ...prev, [columnKey]: "" }))
            }
          >
            <span className="text-muted-foreground text-xs">Clear filter</span>
          </DropdownMenuItem>

          {/* Values list */}
          <div className="max-h-52 overflow-y-auto">
            {visibleOptions.length === 0 && (
              <div className="px-3 py-2 text-xs text-muted-foreground">
                No matches
              </div>
            )}
            {visibleOptions.map((opt) => (
              <DropdownMenuItem
                key={opt}
                onClick={() =>
                  setSummaryFieldFilters((prev) => ({
                    ...prev,
                    [columnKey]: opt,
                  }))
                }
              >
                <span className="text-xs">{opt}</span>
              </DropdownMenuItem>
            ))}
          </div>
        </DropdownMenuContent>
      </DropdownMenu>
    );
  };

  const summaryInfo = (
    <div className="space-y-2">
      <div className="font-medium text-foreground">What this Summary shows</div>
      <ul className="list-disc pl-4 space-y-1">
        <li>
          <span className="text-foreground">Initial Holding</span> is the value
          on the{" "}
          <span className="text-foreground">
            first date where the name appears
          </span>
          .
        </li>
        <li>
          Aggregated <span className="text-foreground">Bought</span> and{" "}
          <span className="text-foreground">Sold</span> totals across the{" "}
          <span className="text-foreground">selected AS ON date range</span>.
        </li>
        <li>
          <span className="text-foreground">Net B/S</span> = Bought - Sold
          (netted over the range).
        </li>
        <li>
          <span className="text-foreground">Still Holding</span> uses the{" "}
          <span className="text-foreground">
            last AS ON snapshot inside the selected range
          </span>
          .
        </li>
      </ul>
      <div className="font-medium text-foreground mt-2">
        Controls explanation
      </div>
      <ul className="list-disc pl-4 space-y-1">
        <li>
          Click header cells (like in Excel) to sort by Sold, Bought, Initial,
          Net B/S, or Still Holding.
        </li>
        <li>
          <span className="text-foreground">All / Buyers / Sellers</span>{" "}
          filters rows by whether totals have bought/sold &gt; 0.
        </li>
        <li>
          <span className="text-foreground">Sort menu</span> changes ordering by
          any summary metric and direction.
        </li>
        <li>
          <span className="text-foreground">Range</span> picks base dates pulled
          from all uploaded “AS ON DD-MM-YYYY” columns.
        </li>
      </ul>
    </div>
  );

  const matrixInfo = (
    <div className="space-y-2">
      <div className="font-medium text-foreground">
        What this Master Matrix shows
      </div>
      <ul className="list-disc pl-4 space-y-1">
        <li>
          A <span className="text-foreground">time-ordered</span> view of
          holdings for each DPID/Client/Name.
        </li>
        <li>
          Each uploaded sheet contributes{" "}
          <span className="text-foreground">
            three columns: Holding as on first date, B/S, Holding as on second
            date
          </span>
          .
        </li>
        <li>
          Duplicate entries by name are{" "}
          <span className="text-foreground">combined</span> after trimming
          periods and whitespace at the edges.
        </li>
        <li>
          The <span className="text-foreground">B/S column</span> displays
          bought and sold for the sheet (change from the first AS ON to the
          second AS ON date).
        </li>
        <li>
          Cell colors show % change vs the previous displayed{" "}
          <span className="text-foreground">holding value</span>.
        </li>
      </ul>
      <div className="font-medium text-foreground mt-2">How to use</div>
      <ul className="list-disc pl-4 space-y-1">
        <li>
          Click headers (DPID, Client, Category, Name, Initial, AS ON dates) to
          sort, similar to Excel.
        </li>
        <li>Use filter icons on DPID/Client/Category to narrow rows.</li>
        <li>
          “Reverse Dates” flips sheet order (rightmost/leftmost) without
          modifying underlying data.
        </li>
      </ul>
    </div>
  );

  // ---------------- UI ----------------
  return (
    <div className="w-full animate-fade-in">
      {/* Global Header */}
      <div className="flex flex-col sm:flex-row gap-4 mb-6 items-start sm:items-center justify-between">
        <div className="relative flex-1 max-w-md">
          <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-muted-foreground" />
          <Input
            placeholder="Search by name, DPID, client ID, or category..."
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            className="pl-10 bg-card border-border shadow-enterprise"
          />
          {searchTerm && (
            <button
              onClick={() => setSearchTerm("")}
              className="absolute right-3 top-1/2 -translate-y-1/2 p-1 hover:bg-muted rounded"
            >
              <X className="w-4 h-4 text-muted-foreground" />
            </button>
          )}
        </div>

        <div className="flex gap-2 flex-wrap">
          <Button
            variant="outline"
            size="sm"
            onClick={toggleDateOrder}
            className="shadow-enterprise"
          >
            <RotateCcw className="w-4 h-4" />
            Reverse Dates
          </Button>

          <Button
            variant="outline"
            size="sm"
            onClick={exportToXLSX}
            disabled={isExportingXlsx}
            className="shadow-enterprise"
          >
            <Download className="w-4 h-4" />
            {isExportingXlsx ? "Exporting XLSX..." : "Export XLSX"}
          </Button>

          <Button
            variant="outline"
            size="sm"
            onClick={() => {
              setSummaryOpen(true); // so user also sees the Summary controls
              exportSummaryToXLSX();
            }}
            disabled={isExportingSummaryXlsx || summaryRows.length === 0}
            className="shadow-enterprise"
            title="Export the currently computed Summary table (respects filters + range)"
          >
            <Download className="w-4 h-4" />
            {isExportingSummaryXlsx ? "Exporting Summary..." : "Export Summary"}
          </Button>

          <Button
            variant="outline"
            size="sm"
            onClick={handleClearDataClick}
            disabled={isClearing}
            className="text-destructive hover:text-destructive shadow-enterprise"
          >
            <Trash2 className="w-4 h-4" />
            {isClearing ? "Clearing..." : "Clear Data"}
          </Button>
        </div>
      </div>

      {/* Active Filters (main matrix) */}
      {Object.entries(filters).some(([, v]) => v) && (
        <div className="flex gap-2 mb-4 flex-wrap">
          {Object.entries(filters).map(
            ([key, value]) =>
              value && (
                <span
                  key={key}
                  className="inline-flex items-center gap-1.5 px-2.5 py-1 text-xs font-medium rounded-full bg-primary/10 text-primary border border-primary/20"
                >
                  {key}: {value}
                  <button
                    onClick={() => setFilters((f) => ({ ...f, [key]: "" }))}
                    className="hover:text-foreground transition-colors"
                  >
                    <X className="w-3 h-3" />
                  </button>
                </span>
              )
          )}
        </div>
      )}

      {/* Stats */}
      <div className="flex gap-4 mb-4 text-sm text-muted-foreground flex-wrap">
        <span>
          Showing{" "}
          <span className="text-foreground font-semibold">
            {filteredData.length}
          </span>{" "}
          of{" "}
          <span className="text-foreground font-semibold">
            {combinedHoldings.length}
          </span>{" "}
          combined records
        </span>
        <span className="text-border hidden sm:inline-block">|</span>
        <span>
          <span className="text-foreground font-semibold">{dates.length}</span>{" "}
          AS ON columns
        </span>
      </div>

      {/* ===================== SUMMARY (COLLAPSIBLE) ===================== */}
      <div className="rounded-lg border border-border overflow-hidden shadow-enterprise-md bg-card mb-6">
        <SectionHeader
          title="Summary (Initial Holding, Bought/Sold totals, Latest Holding)"
          open={summaryOpen}
          onToggle={() => setSummaryOpen((o) => !o)}
          info={summaryInfo}
          subtitle={
            <span>
              Range end snapshot:{" "}
              <span className="text-foreground font-medium">
                {normalizedRange.to ||
                  (latestDateKey ? getBaseDate(latestDateKey) : "-")}
              </span>
            </span>
          }
          right={
            summaryOpen ? (
              <div className="flex flex-wrap items-center gap-2">
                {/* Buyer/Seller quick filter */}
                <div className="flex items-center gap-1 rounded-md border border-border bg-card p-0.5">
                  <Button
                    size="sm"
                    variant="ghost"
                    className={cn(
                      "h-7 px-2 text-xs",
                      summaryType === "all" && "bg-primary/10 text-primary"
                    )}
                    onClick={() => setSummaryType("all")}
                  >
                    All
                  </Button>
                  <Button
                    size="sm"
                    variant="ghost"
                    className={cn(
                      "h-7 px-2 text-xs",
                      summaryType === "buyers" && "bg-primary/10 text-primary"
                    )}
                    onClick={() => setSummaryType("buyers")}
                  >
                    Buyers
                  </Button>
                  <Button
                    size="sm"
                    variant="ghost"
                    className={cn(
                      "h-7 px-2 text-xs",
                      summaryType === "sellers" && "bg-primary/10 text-primary"
                    )}
                    onClick={() => setSummaryType("sellers")}
                  >
                    Sellers
                  </Button>
                </div>

                {/* Sort menu */}
                <SummarySortMenu />

                <Button
                  variant="outline"
                  size="sm"
                  className="shadow-enterprise"
                  onClick={exportSummaryToXLSX}
                  disabled={isExportingSummaryXlsx || summaryRows.length === 0}
                  title="Export the currently filtered Summary table"
                >
                  <Download className="w-4 h-4" />
                  {isExportingSummaryXlsx
                    ? "Exporting Summary..."
                    : "Export Summary"}
                </Button>

                {/* Direction quick toggle */}
                <Button
                  variant="outline"
                  size="sm"
                  className="shadow-enterprise"
                  onClick={() =>
                    setSummarySortDir((d) => (d === "asc" ? "desc" : "asc"))
                  }
                  title="Toggle sort direction"
                >
                  {summarySortDir === "asc" ? (
                    <ArrowUp className="w-4 h-4" />
                  ) : (
                    <ArrowDown className="w-4 h-4" />
                  )}
                  {summarySortDir === "asc" ? "Asc" : "Desc"}
                </Button>

                {/* Date range selectors */}
                <div className="flex items-center gap-2 rounded-md border border-border bg-card px-2 py-1">
                  <span className="text-[10px] uppercase tracking-wider text-muted-foreground">
                    Range
                  </span>

                  <select
                    className="h-7 rounded-md bg-transparent text-xs border border-border px-2"
                    value={summaryFromBase}
                    onChange={(e) => setSummaryFromBase(e.target.value)}
                  >
                    <option value="">From</option>
                    {baseDatesSorted.map((d) => (
                      <option key={`from-${d}`} value={d}>
                        {d}
                      </option>
                    ))}
                  </select>

                  <span className="text-xs text-muted-foreground">→</span>

                  <select
                    className="h-7 rounded-md bg-transparent text-xs border border-border px-2"
                    value={summaryToBase}
                    onChange={(e) => setSummaryToBase(e.target.value)}
                  >
                    <option value="">To</option>
                    {baseDatesSorted.map((d) => (
                      <option key={`to-${d}`} value={d}>
                        {d}
                      </option>
                    ))}
                  </select>
                </div>

                <Button
                  variant="ghost"
                  size="sm"
                  className="text-xs"
                  onClick={clearSummaryFilters}
                >
                  Reset Summary Filters
                </Button>
              </div>
            ) : (
              <div className="text-xs text-muted-foreground">
                Collapsed • click to view totals & filters
              </div>
            )
          }
        />

        {summaryOpen && (
          <>
            <div className="overflow-x-auto">
              <table className="w-full">
                <thead>
                  <tr className="bg-muted/50 border-b border-border">
                    {/* DPID */}
                    <th className="px-4 py-2 text-left text-xs font-semibold uppercase tracking-wider text-muted-foreground">
                      <div className="flex items-center gap-2">
                        <span>DPID</span>
                        <SummaryFilterButton columnKey="dpid" label="DPID" />
                      </div>
                    </th>

                    {/* ClientID */}
                    <th className="px-4 py-2 text-left text-xs font-semibold uppercase tracking-wider text-muted-foreground">
                      <div className="flex items-center gap-2">
                        <span>ClientID</span>
                        <SummaryFilterButton
                          columnKey="clientId"
                          label="Client ID"
                        />
                      </div>
                    </th>

                    {/* Category */}
                    <th className="px-4 py-2 text-left text-xs font-semibold uppercase tracking-wider text-muted-foreground">
                      <div className="flex items-center gap-2">
                        <span>Category</span>
                        <SummaryFilterButton
                          columnKey="category"
                          label="Category"
                        />
                      </div>
                    </th>
                    {/* Sold */}
                    <th
                      className="px-4 py-2 text-right text-xs font-semibold uppercase tracking-wider text-muted-foreground cursor-pointer"
                      onClick={() => handleSummaryHeaderSort("sold")}
                    >
                      Sold
                      <SummarySortIcon columnKey="sold" />
                    </th>
                    {/* Name */}
                    <th className="px-4 py-2 text-left text-xs font-semibold uppercase tracking-wider text-muted-foreground">
                      <div className="flex items-center gap-2">
                        <span>Name</span>
                        <SummaryFilterButton columnKey="name" label="Name" />
                      </div>
                    </th>

                    {/* Bought */}
                    <th
                      className="px-4 py-2 text-right text-xs font-semibold uppercase tracking-wider text-muted-foreground cursor-pointer"
                      onClick={() => handleSummaryHeaderSort("bought")}
                    >
                      Bought
                      <SummarySortIcon columnKey="bought" />
                    </th>
                    {/* Initial */}
                    <th
                      className="px-4 py-2 text-right text-xs font-semibold uppercase tracking-wider text-muted-foreground cursor-pointer"
                      onClick={() => handleSummaryHeaderSort("initialHolding")}
                    >
                      Initial Holding
                      <SummarySortIcon columnKey="initialHolding" />
                    </th>
                    {/* Net B/S */}
                    <th
                      className="px-4 py-2 text-right text-xs font-semibold uppercase tracking-wider text-muted-foreground cursor-pointer"
                      onClick={() => handleSummaryHeaderSort("net")}
                    >
                      Net B/S (Bought or Sold)
                      <SummarySortIcon columnKey="net" />
                    </th>
                    {/* Still Holding */}
                    <th
                      className="px-4 py-2 text-right text-xs font-semibold uppercase tracking-wider text-muted-foreground cursor-pointer"
                      onClick={() => handleSummaryHeaderSort("stillHolding")}
                    >
                      Still Holding
                      <SummarySortIcon columnKey="stillHolding" />
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {/* Totals row */}
                  {summaryRows.length > 0 && (
                    <tr className="bg-muted/40 border-b border-border/70">
                      <td className="px-4 py-2 font-mono text-xs text-muted-foreground"></td>
                      <td className="px-4 py-2 font-mono text-xs text-muted-foreground"></td>
                      <td className="px-4 py-2 text-xs text-muted-foreground"></td>
                      <td className="px-4 py-2 text-right font-mono text-sm text-destructive">
                        {summaryTotals.sold.toLocaleString()}
                      </td>
                      <td className="px-4 py-2 text-sm font-semibold text-foreground">
                        Total
                      </td>
                      <td className="px-4 py-2 text-right font-mono text-sm text-success">
                        {summaryTotals.bought.toLocaleString()}
                      </td>
                      <td className="px-4 py-2 text-right font-mono text-sm">
                        {summaryTotals.initial.toLocaleString()}
                      </td>
                      <td
                        className={cn(
                          "px-4 py-2 text-right font-mono text-sm",
                          summaryTotals.net > 0
                            ? "text-success"
                            : summaryTotals.net < 0
                            ? "text-destructive"
                            : "text-muted-foreground"
                        )}
                      >
                        {summaryTotals.net.toLocaleString()}
                      </td>
                      <td className="px-4 py-2 text-right font-mono text-sm">
                        {summaryTotals.still.toLocaleString()}
                      </td>
                    </tr>
                  )}

                  {summaryPageRows.map((r, idx) => (
                    <tr
                      key={r.id}
                      className={cn(
                        "border-b border-border/50 hover:bg-muted/30 transition-colors",
                        idx % 2 === 0 ? "bg-card" : "bg-muted/10"
                      )}
                    >
                      <td className="px-4 py-2 font-mono text-sm">{r.dpid}</td>
                      <td className="px-4 py-2 font-mono text-sm">
                        {r.clientId}
                      </td>
                      <td className="px-4 py-2">
                        {r.category && (
                          <span className="inline-flex px-2 py-0.5 text-xs font-medium rounded-md bg-accent/10 text-accent border border-accent/20">
                            {r.category}
                          </span>
                        )}
                      </td>
                      <td className="px-4 py-2 text-right font-mono text-sm text-destructive">
                        {r.sold.toLocaleString()}
                      </td>
                      <td className="px-4 py-2 text-sm font-medium text-foreground">
                        {r.name}
                      </td>
                      <td className="px-4 py-2 text-right font-mono text-sm text-success">
                        {r.bought.toLocaleString()}
                      </td>
                      <td className="px-4 py-2 text-right font-mono text-sm">
                        {r.initialHolding.toLocaleString()}
                      </td>
                      <td
                        className={cn(
                          "px-4 py-2 text-right font-mono text-sm",
                          r.net > 0
                            ? "text-success"
                            : r.net < 0
                            ? "text-destructive"
                            : "text-muted-foreground"
                        )}
                      >
                        {r.net.toLocaleString()}
                      </td>
                      <td className="px-4 py-2 text-right font-mono text-sm">
                        {r.stillHolding.toLocaleString()}
                      </td>
                    </tr>
                  ))}

                  {summaryRows.length === 0 && (
                    <tr>
                      <td
                        colSpan={9}
                        className="p-8 text-center text-muted-foreground"
                      >
                        No summary data available for the selected
                        filters/range.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>

            {summaryRows.length > 0 && (
              <div className="flex items-center justify-between px-4 py-2 border-t border-border bg-card/60 text-[11px]">
                <div className="text-muted-foreground">
                  Showing{" "}
                  <span className="text-foreground font-medium">
                    {(summaryPage - 1) * summaryPageSize + 1}
                  </span>{" "}
                  –{" "}
                  <span className="text-foreground font-medium">
                    {Math.min(
                      summaryPage * summaryPageSize,
                      summaryRows.length
                    )}
                  </span>{" "}
                  of{" "}
                  <span className="text-foreground font-medium">
                    {summaryRows.length}
                  </span>{" "}
                  matching names
                </div>
                <PaginationControls
                  page={summaryPage}
                  totalPages={summaryTotalPages}
                  pageSize={summaryPageSize}
                  onPageChange={setSummaryPage}
                  onPageSizeChange={setSummaryPageSize}
                />
              </div>
            )}

            <div className="px-4 py-2 text-[11px] text-muted-foreground bg-muted/20 border-t border-border">
              Range applied to totals using AS ON base dates:{" "}
              <span className="text-foreground">
                {normalizedRange.from && normalizedRange.to
                  ? `${normalizedRange.from} → ${normalizedRange.to}`
                  : "All dates"}
              </span>
            </div>
          </>
        )}
      </div>
      {/* =================== END SUMMARY =================== */}

      {/* ===================== MASTER MATRIX (COLLAPSIBLE) ===================== */}
      <div className="rounded-lg border border-border shadow-enterprise-lg bg-card overflow-visible">
        <SectionHeader
          title="Master Holdings Matrix"
          open={matrixOpen}
          onToggle={() => setMatrixOpen((o) => !o)}
          info={matrixInfo}
          subtitle={
            <span>
              Sheets:{" "}
              <span className="text-foreground font-medium">
                {fileGroups.length}
              </span>{" "}
              • AS ON columns:{" "}
              <span className="text-foreground font-medium">
                {dates.length}
              </span>{" "}
              • Order:{" "}
              <span className="text-foreground font-medium">
                {dateOrder === "asc" ? "Old → New" : "New → Old (by sheet)"}
              </span>
            </span>
          }
          right={
            !matrixOpen ? (
              <div className="text-xs text-muted-foreground">
                Collapsed • click to view full AS ON grid
              </div>
            ) : null
          }
        />

        {matrixOpen && (
          <>
            {/* Color Legend */}
            <div className="px-4 pt-4">
              <div className="flex gap-6 mb-4 p-3 bg-card rounded-lg border border-border shadow-enterprise text-xs overflow-x-auto">
                <div className="flex items-center gap-2">
                  <span className="text-muted-foreground font-medium">
                    Value Change:
                  </span>
                  <div className="flex items-center gap-1">
                    <span className="w-4 h-4 rounded cell-positive-1"></span>
                    <span className="w-4 h-4 rounded cell-positive-2"></span>
                    <span className="w-4 h-4 rounded cell-positive-3"></span>
                    <span className="w-4 h-4 rounded cell-positive-4"></span>
                    <span className="text-muted-foreground ml-1">Increase</span>
                  </div>
                  <div className="flex items-center gap-1 ml-3">
                    <span className="w-4 h-4 rounded cell-negative-1"></span>
                    <span className="w-4 h-4 rounded cell-negative-2"></span>
                    <span className="w-4 h-4 rounded cell-negative-3"></span>
                    <span className="w-4 h-4 rounded cell-negative-4"></span>
                    <span className="text-muted-foreground ml-1">Decrease</span>
                  </div>
                </div>
              </div>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full">
                <thead>
                  <tr className="bg-muted/50 border-b border-border">
                    {/* Fixed group: DPID / CLIENT / CATEGORY / NAME / INITIAL (bordered as a block) */}
                    <th className="sticky left-0 z-10 bg-muted/50 px-4 py-3 text-left">
                      <div className="flex items-center gap-2">
                        <button
                          onClick={() => handleSort("dpid")}
                          className="flex items-center gap-1.5 text-xs font-semibold uppercase tracking-wider text-muted-foreground hover:text-foreground transition-colors"
                        >
                          DPID
                          <SortIcon columnKey="dpid" />
                        </button>
                        <MatrixFilterButton
                          columnKey="dpid"
                          label="DPID"
                          values={uniqueValues.dpid}
                        />
                      </div>
                    </th>

                    <th className="px-4 py-3 text-left">
                      <div className="flex items-center gap-2">
                        <button
                          onClick={() => handleSort("clientId")}
                          className="flex items-center gap-1.5 text-xs font-semibold uppercase tracking-wider text-muted-foreground hover:text-foreground transition-colors"
                        >
                          CLIENT-ID
                          <SortIcon columnKey="clientId" />
                        </button>
                        <MatrixFilterButton
                          columnKey="clientId"
                          label="Client ID"
                          values={uniqueValues.clientId}
                        />
                      </div>
                    </th>

                    <th className="px-4 py-3 text-left">
                      <div className="flex items-center gap-2">
                        <button
                          onClick={() => handleSort("category")}
                          className="flex items-center gap-1.5 text-xs font-semibold uppercase tracking-wider text-muted-foreground hover:text-foreground transition-colors"
                        >
                          CATEGORY
                          <SortIcon columnKey="category" />
                        </button>
                        <MatrixFilterButton
                          columnKey="category"
                          label="Category"
                          values={uniqueValues.category}
                        />
                      </div>
                    </th>

                    <th className="px-4 py-3 text-left min-w-[200px]">
                      <div className="flex items-center gap-2">
                        <button
                          onClick={() => handleSort("name")}
                          className="flex items-center gap-1.5 text-xs font-semibold uppercase tracking-wider text-muted-foreground hover:text-foreground transition-colors"
                        >
                          NAME
                          <SortIcon columnKey="name" />
                        </button>
                        <MatrixFilterButton
                          columnKey="name"
                          label="Name"
                          values={uniqueValues.name}
                        />
                      </div>
                    </th>

                    <th className="px-4 py-3 text-right border-r-2 border-border/80">
                      <button
                        onClick={() => handleSort("initialHolding")}
                        className="flex items-center gap-1.5 text-xs font-semibold uppercase tracking-wider text-muted-foreground hover:text-foreground transition-colors"
                        title="Initial holding (first date this name appears)"
                      >
                        Initial
                        <SortIcon columnKey="initialHolding" />
                      </button>
                    </th>

                    {/* Now per-sheet group: [AS ON D1] | [B/S] | [AS ON D2] with a border around each trio */}
                    {fileGroups.map((group, groupIndex) => {
                      const firstKey = group.dateKeys[0];
                      const secondKey = group.dateKeys[1];

                      const firstBase = firstKey ? getBaseDate(firstKey) : "";
                      const secondBase = secondKey
                        ? getBaseDate(secondKey)
                        : "";

                      return (
                        <Fragment
                          key={`fg-header-${group.fileIndex}-${groupIndex}`}
                        >
                          <th
                            className={cn(
                              "px-3 py-3 text-center text-xs font-semibold uppercase tracking-wider text-muted-foreground hover:text-foreground transition-colors border-l-2 border-border/80"
                            )}
                          >
                            {firstKey ? (
                              <button
                                onClick={() => handleSort(`date-${firstKey}`)}
                                className="flex items-center gap-1.5 justify-center"
                              >
                                {firstBase}
                                <SortIcon columnKey={`date-${firstKey}`} />
                              </button>
                            ) : (
                              firstBase || "-"
                            )}
                          </th>
                          <th className="px-3 py-3 text-center text-xs font-semibold uppercase tracking-wider text-muted-foreground">
                            B/S
                          </th>
                          <th
                            className={cn(
                              "px-3 py-3 text-center text-xs font-semibold uppercase tracking-wider text-muted-foreground hover:text-foreground transition-colors border-r-2 border-border/80"
                            )}
                          >
                            {secondKey ? (
                              <button
                                onClick={() => handleSort(`date-${secondKey}`)}
                                className="flex items-center gap-1.5 justify-center"
                              >
                                {secondBase}
                                <SortIcon columnKey={`date-${secondKey}`} />
                              </button>
                            ) : (
                              secondBase || "-"
                            )}
                          </th>
                        </Fragment>
                      );
                    })}
                  </tr>
                </thead>

                <tbody>
                  {matrixPageRows.map((holding, index) => {
                    const initialInfo = initialHoldingMap.get(holding.id);

                    return (
                      <tr
                        key={holding.id}
                        className={cn(
                          "border-b border-border/50 hover:bg-muted/30 transition-colors",
                          index % 2 === 0 ? "bg-card" : "bg-muted/10"
                        )}
                      >
                        <td className="sticky left-0 z-10 px-4 py-3 font-mono text-sm bg-inherit border-r border-border/30">
                          {holding.dpid}
                        </td>
                        <td className="px-4 py-3 font-mono text-sm">
                          {holding.clientId}
                        </td>
                        <td className="px-4 py-3">
                          {holding.category && (
                            <span className="inline-flex px-2 py-0.5 text-xs font-medium rounded-md bg-accent/10 text-accent border border-accent/20">
                              {holding.category}
                            </span>
                          )}
                        </td>
                        <td className="px-4 py-3 font-medium text-foreground">
                          {holding.name}
                        </td>
                        <td className="px-4 py-3 text-right font-mono text-sm border-r-2 border-border/80">
                          {initialInfo?.value
                            ? initialInfo.value.toLocaleString()
                            : "-"}
                        </td>

                        {/* Per-sheet blocks: [holding D1] | [B/S using D2] | [holding D2] */}
                        {(() => {
                          let prevValue: number | undefined = undefined;
                          const cells: ReactNode[] = [];

                          fileGroups.forEach((group, groupIndex) => {
                            const firstKey = group.dateKeys[0];
                            const secondKey = group.dateKeys[1];

                            const dv1 = firstKey
                              ? holding.dateValues[firstKey]
                              : undefined;
                            const dv2 = secondKey
                              ? holding.dateValues[secondKey]
                              : undefined;

                            // First holding (AS ON D1)
                            const value1Color =
                              dv1 && dv1.value !== undefined
                                ? getValueColorClass(dv1.value, prevValue)
                                : "";
                            if (dv1 && dv1.value !== undefined) {
                              prevValue = dv1.value;
                            }

                            cells.push(
                              <td
                                key={`v1-${holding.id}-${groupIndex}`}
                                className={cn(
                                  "px-3 py-3 text-center font-mono text-sm transition-colors border-l-2 border-border/80",
                                  value1Color
                                )}
                              >
                                {dv1 && dv1.value !== undefined
                                  ? dv1.value.toLocaleString()
                                  : "-"}
                              </td>
                            );

                            // B/S for the sheet (using second date's bought/sold)
                            const bsColorClass =
                              dv2 && (dv2.bought || dv2.sold)
                                ? getBuySellColorClass(
                                    dv2.bought || 0,
                                    dv2.sold || 0
                                  )
                                : "";

                            cells.push(
                              <td
                                key={`bs-${holding.id}-${groupIndex}`}
                                className={cn(
                                  "px-3 py-3 text-center text-sm",
                                  bsColorClass
                                )}
                              >
                                {dv2 ? (
                                  <div className="flex flex-col items-center gap-0.5">
                                    {dv2.bought > 0 && (
                                      <span className="text-success font-semibold">
                                        +{dv2.bought.toLocaleString()}
                                      </span>
                                    )}
                                    {dv2.sold > 0 && (
                                      <span className="text-destructive font-semibold">
                                        -{dv2.sold.toLocaleString()}
                                      </span>
                                    )}
                                    {!dv2.bought && !dv2.sold && (
                                      <span className="text-muted-foreground">
                                        -
                                      </span>
                                    )}
                                  </div>
                                ) : (
                                  <span className="text-muted-foreground">
                                    -
                                  </span>
                                )}
                              </td>
                            );

                            // Second holding (AS ON D2)
                            const value2Color =
                              dv2 && dv2.value !== undefined
                                ? getValueColorClass(dv2.value, prevValue)
                                : "";
                            if (dv2 && dv2.value !== undefined) {
                              prevValue = dv2.value;
                            }

                            cells.push(
                              <td
                                key={`v2-${holding.id}-${groupIndex}`}
                                className={cn(
                                  "px-3 py-3 text-center font-mono text-sm transition-colors border-r-2 border-border/80",
                                  value2Color
                                )}
                              >
                                {dv2 && dv2.value !== undefined
                                  ? dv2.value.toLocaleString()
                                  : "-"}
                              </td>
                            );
                          });

                          return cells;
                        })()}
                      </tr>
                    );
                  })}

                  {filteredData.length === 0 && (
                    <tr>
                      <td
                        colSpan={5 + fileGroups.length * 3}
                        className="p-12 text-center text-muted-foreground bg-muted/20"
                      >
                        <p className="font-medium">
                          No records found matching your search criteria.
                        </p>
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>

            {filteredData.length > 0 && (
              <div className="flex items-center justify-between px-4 py-2 border-t border-border bg-card/60 text-[11px]">
                <div className="text-muted-foreground">
                  Showing{" "}
                  <span className="text-foreground font-medium">
                    {(matrixPage - 1) * matrixPageSize + 1}
                  </span>{" "}
                  –{" "}
                  <span className="text-foreground font-medium">
                    {Math.min(matrixPage * matrixPageSize, filteredData.length)}
                  </span>{" "}
                  of{" "}
                  <span className="text-foreground font-medium">
                    {filteredData.length}
                  </span>{" "}
                  matching rows
                </div>
                <PaginationControls
                  page={matrixPage}
                  totalPages={matrixTotalPages}
                  pageSize={matrixPageSize}
                  onPageChange={setMatrixPage}
                  onPageSizeChange={setMatrixPageSize}
                />
              </div>
            )}
          </>
        )}
      </div>
      {/* =================== END MATRIX =================== */}
    </div>
  );
}
