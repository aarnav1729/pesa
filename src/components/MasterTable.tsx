import { useState, useMemo, Fragment } from "react";
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

export function MasterTable({
  holdings,
  dates,
  onClearData,
}: MasterTableProps) {
  const [searchTerm, setSearchTerm] = useState("");
  const [sortConfig, setSortConfig] = useState<SortConfig>(null);
  const [filters, setFilters] = useState<Record<string, string>>({});
  const [dateOrder, setDateOrder] = useState<"asc" | "desc">("asc");

  const orderedDates = useMemo(() => {
    return dateOrder === "asc" ? dates : [...dates].reverse();
  }, [dates, dateOrder]);

  const uniqueValues = useMemo(() => {
    return {
      dpid: [...new Set(holdings.map((h) => h.dpid))],
      clientId: [...new Set(holdings.map((h) => h.clientId))],
      category: [...new Set(holdings.map((h) => h.category))].filter(Boolean),
    };
  }, [holdings]);

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

  /** Determine latest base date + use the LAST column instance of that date */
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

    // pick last occurrence of that base date in the original order
    for (let i = dates.length - 1; i >= 0; i--) {
      if (getBaseDate(dates[i]) === latestBase) return dates[i];
    }

    return dates[dates.length - 1];
  }, [dates]);

  /** Summary rows table above main matrix */
  const summaryRows = useMemo(() => {
    return filteredData.map((h) => {
      let totalBought = 0;
      let totalSold = 0;

      for (const dateKey of dates) {
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
  }, [filteredData, dates, latestDateKey]);

  return (
    <div className="w-full animate-fade-in">
      {/* Header */}
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

      {/* Active Filters */}
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

      {/* ===================== NEW SUMMARY TABLE ===================== */}
      <div className="rounded-lg border border-border overflow-hidden shadow-enterprise-md bg-card mb-6">
        <div className="px-4 py-3 bg-muted/30 border-b border-border">
          <div className="flex items-center justify-between">
            <div className="text-sm font-semibold text-foreground">
              Summary (Bought/Sold totals + latest holding)
            </div>
            <div className="text-xs text-muted-foreground">
              Latest snapshot key:{" "}
              <span className="text-foreground font-medium">
                {latestDateKey ? getBaseDate(latestDateKey) : "-"}
              </span>
            </div>
          </div>
        </div>

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
                  <td className="px-4 py-2 font-mono text-sm">{r.clientId}</td>
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
                    No summary data available.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
      {/* =================== END SUMMARY TABLE =================== */}

      {/* Color Legend */}
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

      {/* Main Matrix Table */}
      <div className="rounded-lg border border-border overflow-hidden shadow-enterprise-lg bg-card">
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
                    <FilterButton columnKey="dpid" values={uniqueValues.dpid} />
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
                      dateIndex > 0 ? orderedDates[dateIndex - 1] : undefined;
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
                                <span className="text-muted-foreground">-</span>
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
            </tbody>
          </table>
        </div>

        {filteredData.length === 0 && (
          <div className="p-12 text-center text-muted-foreground bg-muted/20">
            <p className="font-medium">
              No records found matching your search criteria.
            </p>
          </div>
        )}
      </div>
    </div>
  );
}
