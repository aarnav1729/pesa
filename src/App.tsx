import { Toaster } from "@/components/ui/toaster";
import { Toaster as Sonner } from "@/components/ui/sonner";
import { TooltipProvider } from "@/components/ui/tooltip";
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import {
  BrowserRouter,
  Routes,
  Route,
  Navigate,
  useLocation,
} from "react-router-dom";
import Index from "./pages/Index";
import NotFound from "./pages/NotFound";
import Login from "./pages/Login";
import React from "react";
const queryClient = new QueryClient();
function RequireAuth({ children }: { children: React.ReactNode }) {
  const location = useLocation();
  const [ok, setOk] = React.useState<null | boolean>(null);

  React.useEffect(() => {
    let mounted = true;
    (async () => {
      try {
        const r = await fetch("/api/session", { credentials: "include" });
        if (!mounted) return;

        // IMPORTANT: Avoid false-positives when a dev server / SPA fallback returns index.html (200)
        const ct = (r.headers.get("content-type") || "").toLowerCase();
        if (!r.ok || !ct.includes("application/json")) {
          setOk(false);
          return;
        }

        const data = await r.json().catch(() => null as any);
        const hasUser = Boolean(data?.user?.email);
        setOk(hasUser);
      } catch {
        if (!mounted) return;
        setOk(false);
      }
    })();
    return () => {
      mounted = false;
    };
  }, []);

  if (ok === null) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-background">
        <div className="flex items-center gap-3 text-primary">
          <div className="w-8 h-8 border-2 border-primary border-t-transparent rounded-full animate-spin" />
          <span className="text-lg font-medium">Checking session…</span>
        </div>
      </div>
    );
  }

  if (!ok) {
    // remember where user tried to go
    const intended = location.pathname + location.search + location.hash;
    localStorage.setItem("redirectAfterLogin", intended || "/dashboard");
    return <Navigate to="/login" replace />;
  }

  return <>{children}</>;
}
const App = () => (
  <QueryClientProvider client={queryClient}>
    <TooltipProvider>
      <Toaster />
      <Sonner />
      <BrowserRouter>
        <Routes>
          <Route path="/login" element={<Login />} />
          <Route
            path="/"
            element={
              <RequireAuth>
                <Index />
              </RequireAuth>
            }
          />
          <Route
            path="/dashboard"
            element={
              <RequireAuth>
                <Index />
              </RequireAuth>
            }
          />
          {/* ADD ALL CUSTOM ROUTES ABOVE THE CATCH-ALL "*" ROUTE */}
          <Route path="*" element={<NotFound />} />
        </Routes>
      </BrowserRouter>
    </TooltipProvider>
  </QueryClientProvider>
);

export default App;
