import { useState, useEffect } from "react";
import {
  Database,
  Upload,
  BarChart3,
  TrendingUp,
  Shield,
  Zap,
} from "lucide-react";
import { FileUpload } from "@/components/FileUpload";
import { MasterTable } from "@/components/MasterTable";
import { parseXLSXFile, consolidateData } from "@/lib/xlsxParser";
import {
  saveHoldings,
  getHoldings,
  clearAllData,
  HoldingRecord,
} from "@/lib/db";
import { toast } from "sonner";
import { cn } from "@/lib/utils";

const Index = () => {
  const [holdings, setHoldings] = useState<HoldingRecord[]>([]);
  const [dates, setDates] = useState<string[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [isLoading, setIsLoading] = useState(true);

  useEffect(() => {
    const loadData = async () => {
      try {
        const { holdings: savedHoldings, dates: savedDates } =
          await getHoldings();
        setHoldings(savedHoldings);
        setDates(savedDates);
      } catch (error) {
        console.error("Error loading data:", error);
      } finally {
        setIsLoading(false);
      }
    };
    loadData();
  }, []);

  const handleFilesSelected = async (files: File[]) => {
    setIsProcessing(true);

    try {
      const parsedFiles = [];

      for (const file of files) {
        const parsed = await parseXLSXFile(file);
        if (parsed) {
          parsedFiles.push(parsed);
        }
      }

      if (parsedFiles.length === 0) {
        toast.error("No valid data found in the uploaded files");
        return;
      }

      const { holdings: newHoldings, dates: newDates } =
        consolidateData(parsedFiles);

      // 1) Persist locally (IndexedDB)
      await saveHoldings(newHoldings, newDates);
      setHoldings(newHoldings);
      setDates(newDates);

      // 2) Fire-and-forget sync to backend
      try {
        const res = await fetch("/api/pesa/import", {
          method: "POST",
          credentials: "include",
          headers: {
            "Content-Type": "application/json",
          },
          body: JSON.stringify({
            holdings: newHoldings,
            dates: newDates,
          }),
        });

        if (!res.ok) {
          console.error("Backend sync failed with status:", res.status);
          toast.warning(
            `Processed ${parsedFiles.length} file(s) locally, but server sync failed (status ${res.status}).`
          );
        } else {
          const data = await res.json().catch(() => null);
          const rowsInserted = data?.rowsInserted ?? "unknown";
          toast.success(
            `Processed ${parsedFiles.length} file(s) with ${newHoldings.length} unique records • Synced ${rowsInserted} rows to server`
          );
          return; // avoid double toast below
        }
      } catch (syncErr) {
        console.error("Error syncing to backend:", syncErr);
        toast.warning(
          `Processed ${parsedFiles.length} file(s) locally, but could not reach the server.`
        );
        return; // avoid double toast below
      }

      // Fallback toast if we didn’t return earlier
      toast.success(
        `Successfully processed ${parsedFiles.length} file(s) with ${newHoldings.length} unique records`
      );
    } catch (error) {
      console.error("Error processing files:", error);
      toast.error("Error processing files. Please check the format.");
    } finally {
      setIsProcessing(false);
    }
  };

  const handleClearData = async () => {
    try {
      await clearAllData();
      setHoldings([]);
      setDates([]);
      toast.success("All local data cleared successfully");
    } catch (error) {
      console.error("Error clearing data:", error);
      toast.error("Error clearing data");
    }
  };

  if (isLoading) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-background">
        <div className="flex items-center gap-3 text-primary">
          <div className="w-8 h-8 border-2 border-primary border-t-transparent rounded-full animate-spin" />
          <span className="text-lg font-medium">Loading PESA...</span>
        </div>
      </div>
    );
  }

  const hasData = holdings.length > 0;

  return (
    <div className="min-h-screen bg-background flex flex-col">
      <header className="border-b border-border bg-card shadow-enterprise sticky top-0 z-50">
        <div className="container mx-auto px-4 py-4">
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-3">
              <div className="p-2 rounded-xl bg-card border border-border shadow-enterprise-md">
                <img
                  src="/l.png"
                  alt="PESA logo"
                  className="w-10 h-10 object-contain"
                  draggable={false}
                />
              </div>
              <div>
                <h1 className="text-xl font-bold text-foreground tracking-tight">
                  PESA
                </h1>
                <p className="text-xs text-muted-foreground">
                  Premier Energies Stock Analysis
                </p>
              </div>
            </div>

            {hasData && (
              <button
                onClick={() => {
                  setHoldings([]);
                  setDates([]);
                }}
                className="flex items-center gap-2 px-4 py-2 rounded-lg bg-primary text-primary-foreground text-sm font-medium transition-all hover:bg-primary/90 shadow-enterprise"
              >
                <Upload className="w-4 h-4" />
                Upload More Files
              </button>
            )}
          </div>
        </div>
      </header>

      <main
        className={
          hasData
            ? "mx-auto w-[95vw] px-2 sm:px-4 py-6"
            : "container mx-auto px-4 py-8"
        }
      >
        {!hasData ? (
          <div className="max-w-4xl mx-auto">
            <div className="text-center mb-12 animate-fade-in">
              <div className="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-primary/10 text-primary text-sm font-medium mb-6 border border-primary/20">
                <Shield className="w-4 h-4" />
                PESA
              </div>
              <h2 className="text-4xl font-bold mb-4 text-foreground">
                Premier Energies Stock Analysis
                <br />
              </h2>
              <p className="text-lg text-muted-foreground max-w-2xl mx-auto leading-relaxed">
                Upload multiple XLSX files to consolidate shareholding data,
                track bought/sold quantities, and analyze patterns across dates
                with color-coded insights.
              </p>
            </div>

            <div className="grid md:grid-cols-3 gap-6 mb-12">
              {[
                {
                  icon: Upload,
                  title: "Bulk Upload",
                  description:
                    "Upload multiple XLSX files at once for quick data consolidation",
                },
                {
                  icon: Database,
                  title: "Local Storage + Sync",
                  description:
                    "Data is stored in your browser and optionally synced to the PESA database",
                },
                {
                  icon: Zap,
                  title: "Smart Analysis",
                  description:
                    "Color-coded cells highlight changes across dates instantly",
                },
              ].map((feature, i) => (
                <div
                  key={feature.title}
                  className="p-6 rounded-xl bg-card border border-border shadow-enterprise-md animate-slide-up hover:shadow-enterprise-lg transition-shadow"
                  style={{ animationDelay: `${i * 100}ms` }}
                >
                  <div className="p-3 rounded-lg bg-primary/10 w-fit mb-4">
                    <feature.icon className="w-5 h-5 text-primary" />
                  </div>
                  <h3 className="font-semibold text-foreground mb-2">
                    {feature.title}
                  </h3>
                  <p className="text-sm text-muted-foreground leading-relaxed">
                    {feature.description}
                  </p>
                </div>
              ))}
            </div>

            <FileUpload
              onFilesSelected={handleFilesSelected}
              isProcessing={isProcessing}
            />

            <div className="mt-12 p-6 rounded-xl bg-card border border-border shadow-enterprise w-full">
              <h4 className="font-semibold text-foreground mb-3 flex items-center gap-2">
                <BarChart3 className="w-4 h-4 text-primary" />
                Expected File Format
              </h4>
              <div className="overflow-x-auto bg-muted/50 rounded-lg p-6">
                <code className="text-xs text-muted-foreground font-mono block whitespace-pre">
                  {`SNo | DPID | CLIENT-ID | NAME | SECOND | THIRD | AS ON {DD-MM-YYYY} | BOUGHT | SOLD | AS ON {DD-MM-YYYY} | CATEGORY`}
                </code>
              </div>
              <p className="text-sm text-muted-foreground mt-4 leading-relaxed">
                Each file should contain the columns above. The "AS ON" date
                columns will be automatically detected and merged
                chronologically.
              </p>
            </div>
          </div>
        ) : (
          <MasterTable
            holdings={holdings}
            dates={dates}
            onClearData={handleClearData}
          />
        )}
      </main>

      <footer className="mt-auto border-t border-border bg-card">
        <div
          className={cn("container mx-auto px-4 py-6", hasData && "max-w-none")}
        >
          <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3">
            <div className="flex items-center gap-2">
              <img
                src="/l.png"
                alt="PESA"
                className="w-6 h-6 object-contain"
                draggable={false}
              />
              <span className="text-sm font-semibold text-foreground">
                PESA
              </span>
              <span className="text-xs text-muted-foreground">
                Premier Energies Stock Analysis
              </span>
            </div>

            <div className="text-xs text-muted-foreground">
              Local-first shareholding analysis • IndexedDB in-browser storage •
              Optional sync to PESA SQL backend
            </div>
          </div>
        </div>
      </footer>
    </div>
  );
};

export default Index;
