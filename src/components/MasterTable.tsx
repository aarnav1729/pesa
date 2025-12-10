import { useState, useMemo, Fragment, ReactNode, useRef } from "react";
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

type SummaryTypeFilter = "all" | "buyers" | "sellers";
type SummarySortKey = "bought" | "sold" | "net" | "stillHolding";

/** Small inline hover info tooltip without external deps */
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

  // Reposition on scroll/resize while open
  const attachWindowListeners = open;

  if (attachWindowListeners && typeof window !== "undefined") {
    // lightweight passive listeners
    window.requestAnimationFrame(() => {
      // Avoid repeated layout thrash by computing once per frame
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

export function MasterTable({
  holdings,
  dates,
  onClearData,
}: MasterTableProps) {
  // ---------------- Collapsible states (default collapsed) ----------------
  const [summaryOpen, setSummaryOpen] = useState(false);
  const [matrixOpen, setMatrixOpen] = useState(false);

  // ---------------- Main matrix state ----------------
  const [searchTerm, setSearchTerm] = useState("");
  const [sortConfig, setSortConfig] = useState<SortConfig>(null);
  const [filters, setFilters] = useState<Record<string, string>>({});
  const [dateOrder, setDateOrder] = useState<"asc" | "desc">("asc");

  // ---------------- Summary-only state ----------------
  const [summaryType, setSummaryType] = useState<SummaryTypeFilter>("all");
  const [summarySortKey, setSummarySortKey] = useState<SummarySortKey>("net");
  const [summarySortDir, setSummarySortDir] = useState<"asc" | "desc">("desc");
  const [summaryFromBase, setSummaryFromBase] = useState<string>("");
  const [summaryToBase, setSummaryToBase] = useState<string>("");

  // ---------------- Derived dates ----------------
  const orderedDates = useMemo(() => {
    return dateOrder === "asc" ? dates : [...dates].reverse();
  }, [dates, dateOrder]);

  const baseDatesSorted = useMemo(() => {
    const uniq = Array.from(new Set(dates.map(getBaseDate)));
    uniq.sort((a, b) => parseDate(a).getTime() - parseDate(b).getTime());
    return uniq;
  }, [dates]);

  // Initialize default summary range (first -> last) when dates arrive.
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

  // ---------------- Unique values for main matrix filters ----------------
  const uniqueValues = useMemo(() => {
    return {
      dpid: [...new Set(holdings.map((h) => h.dpid))],
      clientId: [...new Set(holdings.map((h) => h.clientId))],
      category: [...new Set(holdings.map((h) => h.category))].filter(Boolean),
    };
  }, [holdings]);

  // ---------------- Main matrix filtering/sorting ----------------
  const filteredData = useMemo(() => {
    let result = [...holdings];

    if (searchTerm) {
      const term = searchTerm.toLowerCase();
      result = result.filter(
        (h) =>
          h.name.toLowerCase().includes(term) ||
          h.dpid.toLowerCase().includes(term) ||
          h.clientId.toLowerCase().includes(term) ||
          h.category.toLowerCase().includes(term)
      );
    }

    Object.entries(filters).forEach(([key, value]) => {
      if (value) {
        result = result.filter((h) => {
          if (key === "dpid") return h.dpid === value;
          if (key === "clientId") return h.clientId === value;
          if (key === "category") return h.category === value;
          return true;
        });
      }
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
  }, [holdings, searchTerm, sortConfig, filters]);

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

  const FilterButton = ({
    columnKey,
    values,
  }: {
    columnKey: string;
    values: string[];
  }) => (
    <DropdownMenu>
      <DropdownMenuTrigger asChild>
        <button
          className={cn(
            "p-1 rounded hover:bg-muted transition-colors",
            filters[columnKey] && "text-primary bg-primary/10"
          )}
        >
          <Filter className="w-3 h-3" />
        </button>
      </DropdownMenuTrigger>
      <DropdownMenuContent align="start" className="max-h-64 overflow-y-auto">
        <DropdownMenuItem
          onClick={() => setFilters((f) => ({ ...f, [columnKey]: "" }))}
        >
          <span className="text-muted-foreground">Clear filter</span>
        </DropdownMenuItem>
        {values.map((value) => (
          <DropdownMenuItem
            key={value}
            onClick={() => setFilters((f) => ({ ...f, [columnKey]: value }))}
          >
            {value}
          </DropdownMenuItem>
        ))}
      </DropdownMenuContent>
    </DropdownMenu>
  );

  const exportToCSV = () => {
    const headers = ["DPID", "CLIENT-ID", "CATEGORY", "NAME"];
    orderedDates.forEach((dateKey) => {
      headers.push(`AS ON ${getBaseDate(dateKey)}`);
      headers.push("BOUGHT/SOLD");
    });

    const rows = filteredData.map((h) => {
      const row = [h.dpid, h.clientId, h.category, h.name];
      orderedDates.forEach((dateKey) => {
        const dv = h.dateValues[dateKey];
        row.push(String(dv?.value || 0));
        const boughtSold = dv
          ? (dv.bought > 0 ? `+${dv.bought}` : "") +
            (dv.sold > 0 ? `-${dv.sold}` : "")
          : "";
        row.push(boughtSold || "-");
      });
      return row;
    });

    const csvContent = [
      headers.join(","),
      ...rows.map((r) => r.join(",")),
    ].join("\n");
    const blob = new Blob([csvContent], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "pesa-export.csv";
    a.click();
    URL.revokeObjectURL(url);
  };

  const toggleDateOrder = () => {
    setDateOrder((o) => (o === "asc" ? "desc" : "asc"));
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
    const keysForTotals = dateKeysInSummaryRange;

    return filteredData.map((h) => {
      let totalBought = 0;
      let totalSold = 0;

      for (const dateKey of keysForTotals) {
        const dv = h.dateValues[dateKey];
        if (dv) {
          totalBought += dv.bought || 0;
          totalSold += dv.sold || 0;
        }
      }

      const net = totalBought - totalSold;
      const stillHolding = latestDateKey
        ? h.dateValues[latestDateKey]?.value ?? 0
        : 0;

      return {
        id: h.id,
        dpid: h.dpid,
        clientId: h.clientId,
        category: h.category,
        name: h.name,
        bought: totalBought,
        sold: totalSold,
        net,
        stillHolding,
      };
    });
  }, [filteredData, dateKeysInSummaryRange, latestDateKey]);

  const summaryRows = useMemo(() => {
    let result = [...rawSummaryRows];

    if (summaryType === "buyers") {
      result = result.filter((r) => r.bought > 0);
    } else if (summaryType === "sellers") {
      result = result.filter((r) => r.sold > 0);
    }

    result.sort((a, b) => {
      const aVal = a[summarySortKey];
      const bVal = b[summarySortKey];

      if (aVal < bVal) return summarySortDir === "asc" ? -1 : 1;
      if (aVal > bVal) return summarySortDir === "asc" ? 1 : -1;
      return 0;
    });

    return result;
  }, [rawSummaryRows, summaryType, summarySortKey, summarySortDir]);

  const clearSummaryFilters = () => {
    setSummaryType("all");
    setSummarySortKey("net");
    setSummarySortDir("desc");
    setSummaryFromBase("");
    setSummaryToBase("");
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
        {(["bought", "sold", "net", "stillHolding"] as SummarySortKey[]).map(
          (k) => (
            <DropdownMenuItem key={k} onClick={() => setSummarySortKey(k)}>
              {k}
            </DropdownMenuItem>
          )
        )}
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

  const summaryInfo = (
    <div className="space-y-2">
      <div className="font-medium text-foreground">What this Summary shows</div>
      <ul className="list-disc pl-4 space-y-1">
        <li>
          Aggregated <span className="text-foreground">Bought</span> and{" "}
          <span className="text-foreground">Sold</span> totals across the{" "}
          <span className="text-foreground">selected AS ON date range</span>.
        </li>
        <li>
          <span className="text-foreground">Net</span> = Bought - Sold.
        </li>
        <li>
          <span className="text-foreground">Still Holding</span> uses the{" "}
          <span className="text-foreground">latest overall AS ON snapshot</span>{" "}
          (not limited by the range).
        </li>
      </ul>
      <div className="font-medium text-foreground mt-2">
        Controls explanation
      </div>
      <ul className="list-disc pl-4 space-y-1">
        <li>
          <span className="text-foreground">All / Buyers / Sellers</span>{" "}
          filters rows by whether totals have bought/sold &gt; 0.
        </li>
        <li>
          <span className="text-foreground">Sort</span> changes ordering by
          Bought/Sold/Net/Still Holding.
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
          Each file contributes{" "}
          <span className="text-foreground">two AS ON columns</span>.
        </li>
        <li>
          Duplicate dates are preserved using internal keys (e.g.{" "}
          <span className="text-foreground">03-12-2025 appears twice</span> if
          it exists across two files).
        </li>
        <li>
          The <span className="text-foreground">B/S column</span> displays
          bought and sold for that period’s second AS ON snapshot.
        </li>
        <li>
          Cell colors show % change vs the previous displayed column instance.
        </li>
      </ul>
      <div className="font-medium text-foreground mt-2">How to use</div>
      <ul className="list-disc pl-4 space-y-1">
        <li>Use column header clicks for sorting by that column.</li>
        <li>Use filter icons on DPID/Client/Category to narrow rows.</li>
        <li>
          “Reverse Dates” flips visual ordering without modifying underlying
          data.
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
            onClick={exportToCSV}
            className="shadow-enterprise"
          >
            <Download className="w-4 h-4" />
            Export CSV
          </Button>
          <Button
            variant="outline"
            size="sm"
            onClick={onClearData}
            className="text-destructive hover:text-destructive shadow-enterprise"
          >
            <Trash2 className="w-4 h-4" />
            Clear Data
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
      <div className="flex gap-4 mb-4 text-sm text-muted-foreground">
        <span>
          Showing{" "}
          <span className="text-foreground font-semibold">
            {filteredData.length}
          </span>{" "}
          of{" "}
          <span className="text-foreground font-semibold">
            {holdings.length}
          </span>{" "}
          records
        </span>
        <span className="text-border">|</span>
        <span>
          <span className="text-foreground font-semibold">{dates.length}</span>{" "}
          AS ON columns
        </span>
      </div>

      {/* ===================== SUMMARY (COLLAPSIBLE) ===================== */}
      <div className="rounded-lg border border-border overflow-hidden shadow-enterprise-md bg-card mb-6">
        <SectionHeader
          title="Summary (Bought/Sold totals + latest holding)"
          open={summaryOpen}
          onToggle={() => setSummaryOpen((o) => !o)}
          info={summaryInfo}
          subtitle={
            <span>
              Latest snapshot:{" "}
              <span className="text-foreground font-medium">
                {latestDateKey ? getBaseDate(latestDateKey) : "-"}
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
                    <th className="px-4 py-2 text-left text-xs font-semibold uppercase tracking-wider text-muted-foreground">
                      DPID
                    </th>
                    <th className="px-4 py-2 text-left text-xs font-semibold uppercase tracking-wider text-muted-foreground">
                      ClientID
                    </th>
                    <th className="px-4 py-2 text-left text-xs font-semibold uppercase tracking-wider text-muted-foreground">
                      Category
                    </th>
                    <th className="px-4 py-2 text-right text-xs font-semibold uppercase tracking-wider text-muted-foreground">
                      Sold
                    </th>
                    <th className="px-4 py-2 text-left text-xs font-semibold uppercase tracking-wider text-muted-foreground">
                      Name
                    </th>
                    <th className="px-4 py-2 text-right text-xs font-semibold uppercase tracking-wider text-muted-foreground">
                      Bought
                    </th>
                    <th className="px-4 py-2 text-right text-xs font-semibold uppercase tracking-wider text-muted-foreground">
                      Net
                    </th>
                    <th className="px-4 py-2 text-right text-xs font-semibold uppercase tracking-wider text-muted-foreground">
                      Still Holding
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {summaryRows.map((r, idx) => (
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
                        colSpan={8}
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
              Columns shown:{" "}
              <span className="text-foreground font-medium">
                {orderedDates.length}
              </span>{" "}
              • Order:{" "}
              <span className="text-foreground font-medium">
                {dateOrder === "asc" ? "Old → New" : "New → Old"}
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
            {/* Color Legend (kept inside matrix section for cleanliness) */}
            <div className="px-4 pt-4">
              <div className="flex gap-6 mb-4 p-3 bg-card rounded-lg border border-border shadow-enterprise text-xs">
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
                    <th className="sticky left-0 z-10 bg-muted/50 px-4 py-3 text-left">
                      <div className="flex items-center gap-2">
                        <button
                          onClick={() => handleSort("dpid")}
                          className="flex items-center gap-1.5 text-xs font-semibold uppercase tracking-wider text-muted-foreground hover:text-foreground transition-colors"
                        >
                          DPID
                          <SortIcon columnKey="dpid" />
                        </button>
                        <FilterButton
                          columnKey="dpid"
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
                        <FilterButton
                          columnKey="clientId"
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
                        <FilterButton
                          columnKey="category"
                          values={uniqueValues.category}
                        />
                      </div>
                    </th>

                    <th className="px-4 py-3 text-left min-w-[200px]">
                      <button
                        onClick={() => handleSort("name")}
                        className="flex items-center gap-1.5 text-xs font-semibold uppercase tracking-wider text-muted-foreground hover:text-foreground transition-colors"
                      >
                        NAME
                        <SortIcon columnKey="name" />
                      </button>
                    </th>

                    {orderedDates.map((dateKey) => (
                      <th
                        key={dateKey}
                        className="px-4 py-3 text-center"
                        colSpan={2}
                      >
                        <button
                          onClick={() => handleSort(`date-${dateKey}`)}
                          className="flex items-center gap-1.5 text-xs font-semibold uppercase tracking-wider text-muted-foreground hover:text-foreground transition-colors mx-auto"
                        >
                          {getBaseDate(dateKey)}
                          <SortIcon columnKey={`date-${dateKey}`} />
                        </button>
                      </th>
                    ))}
                  </tr>

                  <tr className="bg-muted/30 border-b border-border">
                    <th className="sticky left-0 z-10 bg-muted/30"></th>
                    <th></th>
                    <th></th>
                    <th></th>
                    {orderedDates.map((dateKey) => (
                      <Fragment key={dateKey}>
                        <th className="px-3 py-2 text-xs text-muted-foreground font-medium">
                          Value
                        </th>
                        <th className="px-3 py-2 text-xs text-muted-foreground font-medium">
                          B/S
                        </th>
                      </Fragment>
                    ))}
                  </tr>
                </thead>

                <tbody>
                  {filteredData.map((holding, index) => (
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

                      {orderedDates.map((dateKey, dateIndex) => {
                        const dv = holding.dateValues[dateKey];

                        const prevKey =
                          dateIndex > 0
                            ? orderedDates[dateIndex - 1]
                            : undefined;
                        const prevValue = prevKey
                          ? holding.dateValues[prevKey]?.value
                          : undefined;

                        const valueColorClass = dv
                          ? getValueColorClass(dv.value, prevValue)
                          : "";
                        const bsColorClass = dv
                          ? getBuySellColorClass(dv.bought, dv.sold)
                          : "";

                        return (
                          <Fragment key={dateKey}>
                            <td
                              className={cn(
                                "px-3 py-3 text-center font-mono text-sm transition-colors",
                                valueColorClass
                              )}
                            >
                              {dv?.value?.toLocaleString() || "-"}
                            </td>
                            <td
                              className={cn(
                                "px-3 py-3 text-center text-sm",
                                bsColorClass
                              )}
                            >
                              {dv ? (
                                <div className="flex flex-col items-center gap-0.5">
                                  {dv.bought > 0 && (
                                    <span className="text-success font-semibold">
                                      +{dv.bought.toLocaleString()}
                                    </span>
                                  )}
                                  {dv.sold > 0 && (
                                    <span className="text-destructive font-semibold">
                                      -{dv.sold.toLocaleString()}
                                    </span>
                                  )}
                                  {!dv.bought && !dv.sold && (
                                    <span className="text-muted-foreground">
                                      -
                                    </span>
                                  )}
                                </div>
                              ) : (
                                <span className="text-muted-foreground">-</span>
                              )}
                            </td>
                          </Fragment>
                        );
                      })}
                    </tr>
                  ))}

                  {filteredData.length === 0 && (
                    <tr>
                      <td
                        colSpan={4 + orderedDates.length * 2}
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
          </>
        )}
      </div>
      {/* =================== END MATRIX =================== */}
    </div>
  );
}
