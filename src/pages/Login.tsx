// pesa/src/pages/Login.tsx
import React, { useEffect, useMemo, useRef, useState } from "react";
import { useNavigate } from "react-router-dom";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Label } from "@/components/ui/label";
import { cn } from "@/lib/utils";

type Step = "enterEmail" | "enterOtp";

const API_BASE_URL = window.location.origin;

// Allow list (client-side convenience; server must enforce too)
const ALLOWED = new Set<string>([
  "vcs@premierenergies.com",
  "saluja@premierenergies.com",
  "mdo@premierenergies.com",
  "aarnav.singh@premierenergies.com"
]);

function normalizeEmail(userInput: string) {
  const raw = String(userInput || "").trim().toLowerCase();
  if (!raw) return "";
  return raw.includes("@") ? raw : `${raw}@premierenergies.com`;
}

function isAllowed(email: string) {
  return ALLOWED.has(String(email || "").trim().toLowerCase());
}

const Login: React.FC = () => {
  const navigate = useNavigate();

  const [step, setStep] = useState<Step>("enterEmail");
  const [email, setEmail] = useState("");
  const [otp, setOtp] = useState("");

  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const otpRef = useRef<HTMLInputElement | null>(null);
  const emailRef = useRef<HTMLInputElement | null>(null);

  const normalizedEmail = useMemo(() => normalizeEmail(email), [email]);

  // If already authenticated (valid cookie), go to dashboard
  useEffect(() => {
    (async () => {
      try {
        const r = await fetch("/api/session", { credentials: "include" });
        if (r.ok) {
          const redirectUrl =
            localStorage.getItem("redirectAfterLogin") || "/dashboard";
          localStorage.removeItem("redirectAfterLogin");
          window.location.replace(redirectUrl);
        }
      } catch {
        // ignore
      }
    })();
  }, []);

  useEffect(() => {
    if (step === "enterOtp") {
      setTimeout(() => otpRef.current?.focus(), 0);
    } else {
      setTimeout(() => emailRef.current?.focus(), 0);
    }
  }, [step]);

  const sendOtp = async () => {
    const fullEmail = normalizedEmail;

    setError(null);

    if (!fullEmail) {
      setError("Please enter your email / username.");
      return;
    }

    if (!isAllowed(fullEmail)) {
      setError("Access denied: this app is restricted.");
      return;
    }

    setLoading(true);
    try {
      const res = await fetch(`${API_BASE_URL}/api/send-otp`, {
        method: "POST",
        credentials: "include",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ email: fullEmail }),
      });

      const data = await res.json().catch(() => ({} as any));

      if (!res.ok) {
        setError(data?.message || "Failed to send OTP");
        return;
      }

      setStep("enterOtp");
    } catch (err: any) {
      setError(err?.message || "Network error");
    } finally {
      setLoading(false);
    }
  };

  const verifyOtp = async () => {
    const fullEmail = normalizedEmail;

    setError(null);

    if (!fullEmail) {
      setError("Missing email.");
      return;
    }
    if (!isAllowed(fullEmail)) {
      setError("Access denied: this app is restricted.");
      return;
    }

    const code = String(otp || "").trim();
    if (!code) {
      setError("Please enter the OTP.");
      return;
    }

    setLoading(true);
    try {
      const res = await fetch(`${API_BASE_URL}/api/verify-otp`, {
        method: "POST",
        credentials: "include",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ email: fullEmail, otp: code }),
      });

      const data = await res.json().catch(() => ({} as any));

      if (!res.ok) {
        setError(data?.message || "Invalid OTP");
        return;
      }

      const redirectUrl =
        localStorage.getItem("redirectAfterLogin") || "/dashboard";
      localStorage.removeItem("redirectAfterLogin");
      navigate(redirectUrl, { replace: true });
    } catch (err: any) {
      setError(err?.message || "Network error");
    } finally {
      setLoading(false);
    }
  };

  const onKeyDownEmail: React.KeyboardEventHandler<HTMLInputElement> = (e) => {
    if (e.key === "Enter") sendOtp();
  };

  const onKeyDownOtp: React.KeyboardEventHandler<HTMLInputElement> = (e) => {
    if (e.key === "Enter") verifyOtp();
  };

  return (
    <div className="min-h-screen bg-background flex flex-col">
      {/* Header — MATCH Index.tsx */}
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

            
          </div>
        </div>
      </header>

      {/* Main */}
      <main className="flex-grow flex items-center justify-center p-6">
        <Card className="w-full max-w-md shadow-enterprise-md border-border bg-card">
          <CardHeader>
            <CardTitle className="text-center">Login</CardTitle>
          </CardHeader>

          <CardContent className="space-y-4">
            {error && (
              <div className="rounded-md border border-red-200 bg-red-50 px-3 py-2 text-sm text-red-700">
                {error}
              </div>
            )}

            {step === "enterEmail" ? (
              <>
                <div className="space-y-1">
                  <Label htmlFor="email">Email / Username</Label>
                  <Input
                    id="email"
                    ref={emailRef}
                    value={email}
                    onChange={(e) => setEmail(e.target.value)}
                    onKeyDown={onKeyDownEmail}
                    placeholder="vcs  (or vcs@premierenergies.com)"
                    autoComplete="username"
                    spellCheck={false}
                  />
                  <p className="text-xs text-muted-foreground">
                    If you type only the username, we’ll append{" "}
                    <span className="font-medium">@premierenergies.com</span>.
                  </p>
                  {normalizedEmail && (
                    <p className="text-xs text-muted-foreground">
                      Using:{" "}
                      <span className="font-medium">{normalizedEmail}</span>
                      {!isAllowed(normalizedEmail) ? (
                        <span className="text-red-600"> (not in allow list)</span>
                      ) : null}
                    </p>
                  )}
                </div>

                <Button className="w-full" onClick={sendOtp} disabled={loading}>
                  {loading ? "Sending OTP..." : "Send OTP"}
                </Button>
              </>
            ) : (
              <>
                <div className="space-y-1">
                  <Label htmlFor="otp">Enter OTP</Label>
                  <Input
                    id="otp"
                    ref={otpRef}
                    value={otp}
                    onChange={(e) =>
                      setOtp(e.target.value.replace(/[^\d]/g, "").slice(0, 6))
                    }
                    onKeyDown={onKeyDownOtp}
                    placeholder="6-digit code"
                    inputMode="numeric"
                    autoComplete="one-time-code"
                  />
                  <p className="text-xs text-muted-foreground">
                    OTP was sent to{" "}
                    <span className="font-medium">{normalizedEmail}</span>.
                  </p>
                </div>

                <Button
                  className="w-full"
                  onClick={verifyOtp}
                  disabled={loading}
                >
                  {loading ? "Verifying..." : "Verify & Login"}
                </Button>

                <Button
                  variant="ghost"
                  className="w-full"
                  onClick={() => {
                    setOtp("");
                    setError(null);
                    setStep("enterEmail");
                  }}
                  disabled={loading}
                >
                  Back
                </Button>
              </>
            )}
          </CardContent>
        </Card>
      </main>

      {/* Footer — MATCH Index.tsx */}
      <footer className="mt-auto border-t border-border bg-card">
        <div className={cn("container mx-auto px-4 py-6")}>
          <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3">
            <div className="flex items-center gap-2">
              <img
                src="/l.png"
                alt="PESA"
                className="w-6 h-6 object-contain"
                draggable={false}
              />
              <span className="text-sm font-semibold text-foreground">PESA</span>
              <span className="text-xs text-muted-foreground">
                Premier Energies Stock Analysis
              </span>
            </div>

            <div className="text-xs text-muted-foreground">
              Premier Energies © {new Date().getFullYear()}. All rights reserved.
            </div>
          </div>
        </div>
      </footer>
    </div>
  );
};

export default Login;
