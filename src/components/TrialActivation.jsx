import React, { useState } from "react";
import {
  Key,
  Lock,
  Unlock,
  Sparkles,
  Clock,
  User,
  Mail,
  Building,
  CheckCircle2,
  AlertCircle,
  ArrowRight,
  ArrowLeft,
  Copy,
  ExternalLink,
  ShieldCheck,
  Activity,
  Zap,
} from "lucide-react";

export default function TrialActivation({
  onActivate,
  trialLockEnabled,
  setTrialLockEnabled,
}) {
  const [view, setView] = useState("activate"); // 'activate' | 'request' | 'success'
  const [trialKey, setTrialKey] = useState("");
  const [error, setError] = useState("");
  const [copied, setCopied] = useState(false);
  const [generatedKey, setGeneratedKey] = useState("");

  // Request form state
  const [formData, setFormData] = useState({
    name: "",
    email: "",
    company: "",
    duration: "14",
    purpose: "Deep energy network validation and alarms mismatch audit.",
  });
  const [requestLoading, setRequestLoading] = useState(false);

  // Key validation rule
  const validateKey = (key) => {
    const cleanKey = key.trim().toUpperCase();
    if (cleanKey === "RovZd2+@24l5!N8$" || cleanKey === "MOB-ENERGY-FREE") {
      return true;
    }
    // Match generated format: NET-TRIAL-XXXX-XXXX
    const regex = /^NET-TRIAL-[A-Z0-9]{4}-[A-Z0-9]{4}$/;
    return regex.test(cleanKey);
  };

  const handleActivate = (e) => {
    e.preventDefault();
    if (!trialKey.trim()) {
      setError("Please enter a trial key.");
      return;
    }

    if (validateKey(trialKey)) {
      setError("");
      onActivate();
    } else {
      setError("Invalid activation key. Please request a new trial key.");
    }
  };

  const handleRequestSubmit = (e) => {
    e.preventDefault();
    if (!formData.name.trim() || !formData.email.trim()) {
      alert("Name and Email are required fields.");
      return;
    }

    setRequestLoading(true);
    // Simulate beautiful server-side key generation
    setTimeout(() => {
      // Generate a mock key
      const randHex = () =>
        Math.random().toString(36).substring(2, 6).toUpperCase();
      // const newKey = `NET-TRIAL-${randHex()}-${randHex()}`;
      const newKey = `RovZd2+@24l5!N8$`;
      setGeneratedKey(newKey);
      setRequestLoading(false);
      setView("success");
    }, 1800);
  };

  const copyToClipboard = () => {
    navigator.clipboard.writeText(generatedKey);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const handleAutoFillAndActivate = () => {
    setTrialKey(generatedKey);
    setView("activate");
    setError("");
    // Automatically trigger activation!
    setTimeout(() => {
      onActivate();
    }, 400);
  };

  return (
    <div className="fixed inset-0 z-[9999] overflow-y-auto bg-gradient-to-br from-slate-950 via-premium-950 to-blue-950 text-white flex flex-col items-center justify-center p-6 sm:p-12 font-sans selection:bg-blue-600/30">
      {/* Background grid light effects */}
      <div className="absolute inset-0 bg-[linear-gradient(to_right,#0f172a_1px,transparent_1px),linear-gradient(to_bottom,#0f172a_1px,transparent_1px)] bg-[size:4rem_4rem] [mask-image:radial-gradient(ellipse_60%_50%_at_50%_50%,#000_70%,transparent_100%)] opacity-40 pointer-events-none" />
      <div className="absolute top-1/4 left-1/2 -translate-x-1/2 w-[500px] h-[500px] bg-blue-500/10 rounded-full blur-[120px] pointer-events-none" />
      <div className="absolute bottom-1/4 right-1/4 w-[300px] h-[300px] bg-emerald-500/5 rounded-full blur-[100px] pointer-events-none" />

      {/* Main Glassmorphic Wrapper */}
      <div className="relative w-full max-w-xl z-10 transition-all duration-500">
        {/* Logo and Brand Title */}
        <div className="flex flex-col items-center mb-10 text-center animate-fade-in">
          <div className="w-16 h-16 bg-gradient-to-br from-blue-600 to-indigo-600 rounded-3xl flex items-center justify-center shadow-2xl shadow-blue-500/30 mb-4 border border-blue-400/20 transform hover:rotate-12 transition-transform duration-300">
            <Key size={32} className="text-white" />
          </div>
          <h1 className="text-3xl sm:text-4xl font-black tracking-tighter leading-none mb-2">
            NET <span className="text-blue-500">Energizer</span>
          </h1>
          <div className="flex items-center gap-2 px-3 py-1 bg-white/5 border border-white/10 rounded-full">
            <ShieldCheck size={12} className="text-emerald-400" />
            <span className="text-[9px] font-black tracking-[0.2em] text-premium-300 uppercase">
              Secure Trial Licensing Protocol
            </span>
          </div>
        </div>

        {/* Dynamic Inner Card Content */}
        <div className="bg-slate-900/80 backdrop-blur-2xl border border-white/10 rounded-[2.5rem] shadow-2xl p-8 sm:p-12 relative overflow-hidden transition-all duration-300">
          {/* Subtle neon glowing light bars */}
          <div className="absolute top-0 inset-x-0 h-[2px] bg-gradient-to-r from-transparent via-blue-500 to-transparent opacity-60" />

          {/* VIEW: ACTIVATE */}
          {view === "activate" && (
            <div className="space-y-8 animate-slide-up">
              <div className="text-center">
                <h3 className="text-2xl font-black tracking-tight mb-2 flex items-center justify-center gap-2">
                  <Lock size={20} className="text-blue-400" /> Software
                  Activation
                </h3>
                <p className="text-sm text-premium-400 font-medium">
                  Enter your 16-character trial license key to unlock deep
                  energy diagnostics, alarms audit, and mismatch correlation.
                </p>
              </div>

              {error && (
                <div className="flex items-center gap-3 p-4 bg-rose-500/10 border border-rose-500/20 rounded-2xl text-rose-300 text-xs font-semibold animate-shake">
                  <AlertCircle size={16} className="shrink-0 text-rose-400" />
                  <span>{error}</span>
                </div>
              )}

              <form onSubmit={handleActivate} className="space-y-6">
                <div className="space-y-2">
                  <label className="text-[10px] font-black uppercase text-premium-400 tracking-widest pl-1">
                    Trial License Key
                  </label>
                  <div className="relative group">
                    <input
                      type="text"
                      value={trialKey}
                      onChange={(e) => setTrialKey(e.target.value)}
                      placeholder="NET-TRIAL-XXXX-XXXX"
                      className="w-full bg-slate-950/60 border-2 border-white/5 group-hover:border-white/15 focus:border-blue-500 rounded-2xl pl-12 pr-6 py-4 font-mono text-sm tracking-widest focus:ring-4 ring-blue-500/10 outline-none transition-all placeholder:text-slate-600 text-center uppercase"
                    />
                    <Key
                      size={18}
                      className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-500 group-focus-within:text-blue-500 transition-colors"
                    />
                  </div>
                </div>

                <button
                  type="submit"
                  className="w-full bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-500 hover:to-indigo-500 text-white font-black text-xs uppercase tracking-[0.2em] py-5 rounded-2xl flex items-center justify-center gap-3 shadow-xl shadow-blue-500/10 hover:shadow-blue-500/20 hover:scale-[1.01] active:scale-[0.99] transition-all"
                >
                  Unlock Software <ArrowRight size={16} />
                </button>
              </form>

              {/* Info Tips Section */}
              <div className="bg-blue-500/5 border border-blue-500/10 rounded-2xl p-5 text-xs text-blue-300 leading-relaxed">
                <div className="flex items-center gap-2 mb-2 font-black uppercase tracking-wider text-[10px] text-blue-400">
                  <Zap size={12} /> Sandbox Activation Code
                </div>
                Use master evaluation key{" "}
                <code className="bg-blue-950/80 px-2 py-0.5 rounded text-white font-mono font-bold text-sm tracking-wider mx-1 select-all border border-blue-800/30">
                  NET-TRIAL-2026
                </code>{" "}
                for instant, zero-delay review of features.
              </div>

              {/* Link to Request form */}
              <div className="text-center pt-2">
                <button
                  onClick={() => {
                    setError("");
                    setView("request");
                  }}
                  className="inline-flex items-center gap-2 text-xs font-black uppercase tracking-widest text-indigo-400 hover:text-indigo-300 hover:underline transition-all"
                >
                  <Sparkles size={14} /> Request a trial activation key
                </button>
              </div>
            </div>
          )}

          {/* VIEW: REQUEST FORM */}
          {view === "request" && (
            <div className="space-y-8 animate-slide-up">
              <div className="text-center">
                <h3 className="text-2xl font-black tracking-tight mb-2 flex items-center justify-center gap-2">
                  <Sparkles
                    size={22}
                    className="text-indigo-400 animate-pulse"
                  />{" "}
                  Request Evaluation Key
                </h3>
                <p className="text-sm text-premium-400 font-medium">
                  Submit a trial request to generate a temporary, custom
                  pre-authorized activation code.
                </p>
              </div>

              <form onSubmit={handleRequestSubmit} className="space-y-5">
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-5">
                  <div className="space-y-2">
                    <label className="text-[10px] font-black uppercase text-premium-400 tracking-widest pl-1">
                      Full Name
                    </label>
                    <div className="relative">
                      <input
                        type="text"
                        required
                        value={formData.name}
                        onChange={(e) =>
                          setFormData({ ...formData, name: e.target.value })
                        }
                        placeholder="John Doe"
                        className="w-full bg-slate-950/60 border border-white/10 focus:border-indigo-500 rounded-2xl pl-10 pr-4 py-3.5 text-xs font-semibold focus:ring-4 ring-indigo-500/10 outline-none transition-all"
                      />
                      <User
                        size={14}
                        className="absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-500"
                      />
                    </div>
                  </div>

                  <div className="space-y-2">
                    <label className="text-[10px] font-black uppercase text-premium-400 tracking-widest pl-1">
                      Business Email
                    </label>
                    <div className="relative">
                      <input
                        type="email"
                        required
                        value={formData.email}
                        onChange={(e) =>
                          setFormData({ ...formData, email: e.target.value })
                        }
                        placeholder="john@company.com"
                        className="w-full bg-slate-950/60 border border-white/10 focus:border-indigo-500 rounded-2xl pl-10 pr-4 py-3.5 text-xs font-semibold focus:ring-4 ring-indigo-500/10 outline-none transition-all"
                      />
                      <Mail
                        size={14}
                        className="absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-500"
                      />
                    </div>
                  </div>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-5">
                  <div className="space-y-2">
                    <label className="text-[10px] font-black uppercase text-premium-400 tracking-widest pl-1">
                      Organization
                    </label>
                    <div className="relative">
                      <input
                        type="text"
                        value={formData.company}
                        onChange={(e) =>
                          setFormData({ ...formData, company: e.target.value })
                        }
                        placeholder="Mobi-Egypt"
                        className="w-full bg-slate-950/60 border border-white/10 focus:border-indigo-500 rounded-2xl pl-10 pr-4 py-3.5 text-xs font-semibold focus:ring-4 ring-indigo-500/10 outline-none transition-all"
                      />
                      <Building
                        size={14}
                        className="absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-500"
                      />
                    </div>
                  </div>

                  <div className="space-y-2">
                    <label className="text-[10px] font-black uppercase text-premium-400 tracking-widest pl-1">
                      Trial Duration
                    </label>
                    <div className="relative">
                      <select
                        value={formData.duration}
                        onChange={(e) =>
                          setFormData({ ...formData, duration: e.target.value })
                        }
                        className="w-full bg-slate-950 border border-white/10 focus:border-indigo-500 rounded-2xl px-4 py-3.5 text-xs font-semibold focus:ring-4 ring-indigo-500/10 outline-none transition-all text-white appearance-none cursor-pointer"
                      >
                        <option value="7">7 Days - Short Evaluation</option>
                        <option value="14">
                          14 Days - Full Ingestion Trial
                        </option>
                        <option value="30">30 Days - Corporate Testing</option>
                      </select>
                      <div className="absolute right-4 top-1/2 -translate-y-1/2 pointer-events-none text-slate-400">
                        ▼
                      </div>
                    </div>
                  </div>
                </div>

                <div className="space-y-2">
                  <label className="text-[10px] font-black uppercase text-premium-400 tracking-widest pl-1">
                    Describe Use Case
                  </label>
                  <textarea
                    rows={2}
                    value={formData.purpose}
                    onChange={(e) =>
                      setFormData({ ...formData, purpose: e.target.value })
                    }
                    placeholder="Brief description of your energy audit needs..."
                    className="w-full bg-slate-950/60 border border-white/10 focus:border-indigo-500 rounded-2xl p-4 text-xs font-semibold focus:ring-4 ring-indigo-500/10 outline-none transition-all resize-none"
                  />
                </div>

                <div className="pt-4 flex flex-col sm:flex-row gap-4">
                  <button
                    type="button"
                    onClick={() => setView("activate")}
                    className="flex-1 border border-white/10 hover:border-white/20 text-white py-4 rounded-2xl font-black text-xs uppercase tracking-widest flex items-center justify-center gap-2 hover:bg-white/5 transition-all"
                  >
                    <ArrowLeft size={14} /> Back
                  </button>

                  <button
                    type="submit"
                    disabled={requestLoading}
                    className="flex-[2] bg-indigo-600 hover:bg-indigo-500 text-white py-4 rounded-2xl font-black text-xs uppercase tracking-[0.15em] flex items-center justify-center gap-2 hover:scale-[1.01] active:scale-[0.99] transition-all disabled:opacity-50"
                  >
                    {requestLoading ? (
                      <>
                        <div className="w-4 h-4 border-2 border-white/20 border-t-white rounded-full animate-spin" />
                        Generating Code...
                      </>
                    ) : (
                      <>
                        Generate Key <ArrowRight size={14} />
                      </>
                    )}
                  </button>
                </div>
              </form>
            </div>
          )}

          {/* VIEW: SUCCESS (KEY GENERATED) */}
          {view === "success" && (
            <div className="space-y-8 animate-slide-up text-center">
              <div className="flex justify-center">
                <div className="w-20 h-20 bg-emerald-500/10 text-emerald-400 rounded-full flex items-center justify-center shadow-inner animate-bounce border border-emerald-500/20">
                  <CheckCircle2 size={44} />
                </div>
              </div>

              <div>
                <h3 className="text-2xl font-black tracking-tight mb-2 text-emerald-400">
                  Trial Pre-Approved!
                </h3>
                <p className="text-sm text-premium-400 font-medium max-w-sm mx-auto">
                  Your simulated trial license key has been successfully
                  generated and is ready for use.
                </p>
              </div>

              {/* Key Box Panel */}
              <div className="bg-slate-950/80 border border-white/10 rounded-[2rem] p-6 space-y-4">
                <span className="text-[9px] font-black uppercase text-premium-400 tracking-[0.2em]">
                  Generated License Key ({formData.duration} Days Trial)
                </span>

                <div className="flex items-center justify-between bg-black/40 border border-white/5 rounded-2xl p-4 font-mono font-black text-base sm:text-lg tracking-widest text-emerald-300 relative group">
                  <span className="select-all truncate">{generatedKey}</span>
                  <button
                    onClick={copyToClipboard}
                    className="ml-3 shrink-0 p-2.5 rounded-xl bg-white/5 border border-white/10 hover:bg-white/10 text-white hover:scale-105 active:scale-95 transition-all"
                    title="Copy to clipboard"
                  >
                    {copied ? (
                      <span className="text-[10px] font-black text-emerald-400 uppercase tracking-widest font-sans">
                        Copied!
                      </span>
                    ) : (
                      <Copy size={16} />
                    )}
                  </button>
                </div>

                <div className="text-[10px] font-bold text-premium-400 text-left px-1 leading-normal">
                  User: <span className="text-white">{formData.name}</span> |
                  Exp:{" "}
                  <span className="text-emerald-400">
                    Pre-Approved Simulation
                  </span>
                </div>
              </div>

              <div className="pt-4 flex flex-col gap-3">
                <button
                  onClick={handleAutoFillAndActivate}
                  className="w-full bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-500 hover:to-teal-500 text-white py-5 rounded-2xl font-black text-xs uppercase tracking-[0.2em] flex items-center justify-center gap-2 hover:scale-[1.01] active:scale-[0.99] transition-all shadow-xl shadow-emerald-500/10"
                >
                  Auto-Fill & Activate Application <Unlock size={16} />
                </button>

                <button
                  onClick={() => setView("activate")}
                  className="text-xs font-black uppercase tracking-widest text-premium-400 hover:text-white transition-colors"
                >
                  Manually Input Key
                </button>
              </div>
            </div>
          )}
        </div>

        {/* DEVELOPER BYPASS CONTROLS badge at bottom */}
        <div className="mt-8 flex justify-center animate-fade-in">
          <div className="bg-slate-900/60 border border-white/10 rounded-full px-6 py-3 flex items-center gap-4 text-xs font-bold text-premium-300 shadow-xl backdrop-blur-md">
            <span className="flex items-center gap-2">
              <span className="w-2 h-2 rounded-full bg-indigo-500 animate-pulse" />
              Dev Switch (Web Mode)
            </span>
            <div className="h-4 w-[1px] bg-white/10" />
            <button
              onClick={() => {
                const newState = !trialLockEnabled;
                setTrialLockEnabled(newState);
                localStorage.setItem(
                  "energy_review_trial_lock_enabled",
                  JSON.stringify(newState),
                );
                if (!newState) {
                  // Reload or trigger instant callback bypass
                  onActivate();
                }
              }}
              className="bg-indigo-600/20 text-indigo-300 hover:bg-indigo-600/40 px-3 py-1 rounded-full text-[10px] font-black uppercase tracking-widest transition-all"
            >
              Toggle Lock:{" "}
              <span
                className={
                  trialLockEnabled ? "text-rose-400" : "text-emerald-400"
                }
              >
                {trialLockEnabled ? "ACTIVE" : "BYPASSED"}
              </span>
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}
