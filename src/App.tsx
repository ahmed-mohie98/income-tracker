import { useState, useMemo } from "react";
import { exportToExcel } from "./utils/exportExcel";

const MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

const defaultExpenses = {
  personal: [
    { label: "Rent / Mortgage", amount: 0 },
    { label: "Food & Groceries", amount: 0 },
    { label: "Transport", amount: 0 },
    { label: "Phone & Internet", amount: 0 },
    { label: "Subscriptions", amount: 0 },
    { label: "Health Insurance", amount: 0 },
    { label: "Other Personal", amount: 0 },
  ],
  business: [
    { label: "Software & Tools", amount: 0 },
    { label: "Equipment", amount: 0 },
    { label: "Marketing", amount: 0 },
    { label: "Platform Fees", amount: 0 },
    { label: "Professional Services", amount: 0 },
    { label: "Education & Courses", amount: 0 },
    { label: "Other Business", amount: 0 },
  ],
};

const defaultIncome = MONTHS.map((m) => ({ month: m, amount: 0 }));

// ── Tiny design tokens ─────────────────────────────────────────────────────────
const T = {
  green: "#1D9E75",
  amber: "#BA7517",
  red:   "#E05252",
  blue:  "#2563EB",
  bg:    "#F8FAFC",
  card:  "#FFFFFF",
  border:"#E5E7EB",
  text:  "#111827",
  muted: "#6B7280",
};

// ── Shared card style ──────────────────────────────────────────────────────────
const card: React.CSSProperties = {
  background: T.card,
  borderRadius: 12,
  padding: "14px 16px",
  border: `1px solid ${T.border}`,
};

export default function App() {
  const [tab, setTab]       = useState("dashboard");
  const [income, setIncome] = useState(defaultIncome);
  const [expenses, setExpenses] = useState(defaultExpenses);
  const [taxRate, setTaxRate]   = useState(28);
  const [name, setName]         = useState("Ahmed Mohie");
  const [exporting, setExporting] = useState(false);

  // ── Derived values ────────────────────────────────────────────────────────────
  const totalPersonal = useMemo(() => expenses.personal.reduce((s, e) => s + Number(e.amount), 0), [expenses]);
  const totalBusiness = useMemo(() => expenses.business.reduce((s, e) => s + Number(e.amount), 0), [expenses]);
  const breakEven     = totalPersonal + totalBusiness;
  const salary        = Math.round(breakEven * 1.15);
  const totalAnnual   = useMemo(() => income.reduce((s, i) => s + Number(i.amount), 0), [income]);
  const avgMonthly    = Math.round(totalAnnual / 12);
  const taxReserve    = Math.round(avgMonthly * taxRate / 100);
  const monthlyProfit = avgMonthly - breakEven;
  const savingsTarget = Math.round(avgMonthly * 0.10);
  const lowestMonth   = useMemo(() => Math.min(...income.map(i => Number(i.amount))), [income]);
  const highestMonth  = useMemo(() => Math.max(...income.map(i => Number(i.amount))), [income]);

  const updateIncome = (idx: number, val: string) => {
    const n = [...income];
    n[idx] = { ...n[idx], amount: Number(val) || 0 };
    setIncome(n);
  };

  const updateExpense = (type: "personal"|"business", idx: number, field: string, val: string) => {
    setExpenses(prev => ({
      ...prev,
      [type]: prev[type].map((e, i) => i === idx ? { ...e, [field]: field === "amount" ? (Number(val) || 0) : val } : e),
    }));
  };

  const statusColor = (val: number, target: number) =>
    val >= target ? T.green : val >= target * 0.7 ? T.amber : T.red;

  // ── Export handler ────────────────────────────────────────────────────────────
  const handleExport = async () => {
    setExporting(true);
    try {
      await exportToExcel(income, expenses, taxRate, name);
    } finally {
      setExporting(false);
    }
  };

  const tabs = ["dashboard", "income", "expenses", "tax", "savings"];

  return (
    <div style={{
      fontFamily: "system-ui, -apple-system, sans-serif",
      maxWidth: 740,
      margin: "0 auto",
      padding: "16px",
      background: T.bg,
      minHeight: "100vh",
      color: T.text,
    }}>

      {/* ── Header ──────────────────────────────────────────────────────────── */}
      <div style={{ marginBottom: 16 }}>
        <div style={{ fontSize: 10, color: T.muted, letterSpacing: "0.08em", fontWeight: 600, marginBottom: 4 }}>
          THE FREELANCER'S FINANCIAL PLAYBOOK
        </div>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 12 }}>
          <div style={{ fontSize: 20, fontWeight: 700, color: T.text }}>Income Tracker</div>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <input
              value={name}
              onChange={e => setName(e.target.value)}
              style={{
                fontSize: 13, padding: "5px 12px", borderRadius: 20,
                border: `1px solid ${T.border}`, background: T.card, color: T.text,
                textAlign: "right", outline: "none",
              }}
            />
            {/* Excel Export Button */}
            <button
              onClick={handleExport}
              disabled={exporting}
              title="Download as Excel (.xlsx)"
              style={{
                display: "flex", alignItems: "center", gap: 6,
                padding: "6px 14px", borderRadius: 20, fontSize: 12, fontWeight: 600,
                cursor: exporting ? "wait" : "pointer",
                border: "none",
                background: exporting ? "#16A34A88" : "#16A34A",
                color: "#fff",
                boxShadow: "0 1px 4px #0002",
                transition: "background .2s",
                whiteSpace: "nowrap",
              }}
            >
              {/* Excel icon */}
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" style={{ flexShrink: 0 }}>
                <rect x="2" y="3" width="20" height="18" rx="3" fill="#fff" fillOpacity=".25"/>
                <path d="M14 3v6h6" stroke="#fff" strokeWidth="1.5" strokeLinejoin="round"/>
                <path d="M14 3L20 9v12a1 1 0 01-1 1H5a1 1 0 01-1-1V4a1 1 0 011-1h9z" stroke="#fff" strokeWidth="1.5" strokeLinejoin="round"/>
                <path d="M8 13l3 5M11 13l-3 5" stroke="#fff" strokeWidth="1.5" strokeLinecap="round"/>
              </svg>
              {exporting ? "Generating…" : "Export Excel"}
            </button>
          </div>
        </div>
      </div>

      {/* ── Tabs ─────────────────────────────────────────────────────────────── */}
      <div style={{ display: "flex", gap: 6, marginBottom: 16, flexWrap: "wrap" }}>
        {tabs.map(t => (
          <button
            key={t}
            onClick={() => setTab(t)}
            style={{
              padding: "5px 16px", borderRadius: 20, fontSize: 12, fontWeight: 500,
              cursor: "pointer", border: `1px solid ${tab === t ? T.blue : T.border}`,
              background: tab === t ? T.blue : T.card,
              color: tab === t ? "#fff" : T.muted,
              textTransform: "capitalize", transition: "all .15s",
            }}
          >
            {t}
          </button>
        ))}
      </div>

      {/* ─────────────────────────────── DASHBOARD ──────────────────────────── */}
      {tab === "dashboard" && (
        <div>
          {/* KPI row 1 */}
          <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 10, marginBottom: 10 }}>
            {[
              { label: "Avg Monthly Income", value: `$${avgMonthly.toLocaleString()}`, color: T.green },
              { label: "Break-Even / mo",    value: `$${breakEven.toLocaleString()}`,  color: T.amber },
              { label: "Monthly Profit",     value: `$${monthlyProfit.toLocaleString()}`, color: monthlyProfit >= 0 ? T.green : T.red },
            ].map(m => (
              <div key={m.label} style={card}>
                <div style={{ fontSize: 10, color: T.muted, marginBottom: 6 }}>{m.label}</div>
                <div style={{ fontSize: 22, fontWeight: 700, color: m.color }}>{m.value}</div>
              </div>
            ))}
          </div>

          {/* KPI row 2 */}
          <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 10, marginBottom: 10 }}>
            {[
              { label: "Tax Reserve / mo",   value: `$${taxReserve.toLocaleString()}`,  color: T.red },
              { label: "Savings Target / mo", value: `$${savingsTarget.toLocaleString()}`, color: T.text },
              { label: "Freelancer Salary",  value: `$${salary.toLocaleString()}`,       color: T.text },
            ].map(m => (
              <div key={m.label} style={card}>
                <div style={{ fontSize: 10, color: T.muted, marginBottom: 6 }}>{m.label}</div>
                <div style={{ fontSize: 20, fontWeight: 700, color: m.color }}>{m.value}</div>
              </div>
            ))}
          </div>

          {/* Health indicators */}
          <div style={{ ...card, marginBottom: 10 }}>
            <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 14 }}>Financial Health Check</div>
            {[
              { label: "Income vs Break-Even",        val: avgMonthly,    target: breakEven,    display: `$${avgMonthly.toLocaleString()} / $${breakEven.toLocaleString()}` },
              { label: "Lowest Month vs Break-Even",  val: lowestMonth,   target: breakEven,    display: `$${lowestMonth.toLocaleString()} / $${breakEven.toLocaleString()}` },
              { label: "Annual Income",               val: totalAnnual,   target: breakEven*12, display: `$${totalAnnual.toLocaleString()}` },
            ].map(h => (
              <div key={h.label} style={{ marginBottom: 12 }}>
                <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, marginBottom: 5 }}>
                  <span style={{ color: T.muted }}>{h.label}</span>
                  <span style={{ fontWeight: 600, color: statusColor(h.val, h.target) }}>{h.display}</span>
                </div>
                <div style={{ background: "#F3F4F6", borderRadius: 6, height: 6, overflow: "hidden" }}>
                  <div style={{
                    height: 6, borderRadius: 6,
                    width: `${Math.min(100, h.target > 0 ? (h.val / h.target) * 100 : 0)}%`,
                    background: statusColor(h.val, h.target),
                    transition: "width .5s ease",
                  }} />
                </div>
              </div>
            ))}
          </div>

          {/* Income chart */}
          <div style={card}>
            <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 14 }}>Monthly Income</div>
            <div style={{ display: "flex", alignItems: "flex-end", gap: 4, height: 90 }}>
              {income.map((m, i) => {
                const h = highestMonth > 0 ? (Number(m.amount) / highestMonth) * 78 : 0;
                const above = Number(m.amount) >= breakEven && breakEven > 0;
                return (
                  <div key={i} style={{ flex: 1, display: "flex", flexDirection: "column", alignItems: "center", gap: 4 }}>
                    <div style={{
                      width: "100%", height: `${h}px`,
                      background: Number(m.amount) === 0 ? "#E5E7EB" : above ? T.green : T.red,
                      borderRadius: "4px 4px 0 0",
                      minHeight: Number(m.amount) > 0 ? 4 : 0,
                      transition: "height .3s ease",
                    }} />
                    <div style={{ fontSize: 9, color: T.muted }}>{m.month}</div>
                  </div>
                );
              })}
            </div>
            <div style={{ display: "flex", gap: 14, marginTop: 10 }}>
              {[["Above break-even", T.green], ["Below break-even", T.red]].map(([label, c]) => (
                <div key={label} style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 10, color: T.muted }}>
                  <div style={{ width: 8, height: 8, borderRadius: 2, background: c as string }} />
                  {label}
                </div>
              ))}
            </div>
          </div>

          {/* Export CTA */}
          <div style={{ ...card, marginTop: 10, background: "#EFF6FF", border: `1px solid #BFDBFE`, display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12 }}>
            <div>
              <div style={{ fontSize: 13, fontWeight: 600, color: T.blue, marginBottom: 2 }}>📊 Download as Excel Spreadsheet</div>
              <div style={{ fontSize: 11, color: T.muted }}>6 sheets — Dashboard · Income · Expenses · Tax · Savings · Instructions — all with live formulas.</div>
            </div>
            <button
              onClick={handleExport}
              disabled={exporting}
              style={{
                padding: "8px 18px", borderRadius: 20, fontSize: 12, fontWeight: 700,
                cursor: exporting ? "wait" : "pointer",
                border: "none", background: "#16A34A", color: "#fff",
                whiteSpace: "nowrap", flexShrink: 0,
              }}
            >
              {exporting ? "…" : "⬇ Export"}
            </button>
          </div>
        </div>
      )}

      {/* ─────────────────────────────── INCOME ─────────────────────────────── */}
      {tab === "income" && (
        <div>
          <div style={{ fontSize: 12, color: T.muted, marginBottom: 12 }}>Enter your monthly income. Green = above break-even.</div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(2,1fr)", gap: 8 }}>
            {income.map((m, i) => {
              const above = Number(m.amount) >= breakEven && breakEven > 0;
              return (
                <div key={i} style={{ ...card, border: `1.5px solid ${above ? T.green+"44" : T.border}` }}>
                  <div style={{ fontSize: 10, color: T.muted, marginBottom: 6 }}>{m.month} 2026</div>
                  <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
                    <span style={{ fontSize: 14, color: T.muted }}>$</span>
                    <input
                      type="number"
                      value={m.amount || ""}
                      onChange={e => updateIncome(i, e.target.value)}
                      placeholder="0"
                      style={{
                        flex: 1, fontSize: 17, fontWeight: 700,
                        border: "none", background: "transparent",
                        color: above ? T.green : T.text, outline: "none", width: "100%",
                      }}
                    />
                  </div>
                </div>
              );
            })}
          </div>

          <div style={{ ...card, marginTop: 12 }}>
            {[
              ["Total Annual", `$${totalAnnual.toLocaleString()}`, T.text],
              ["Monthly Average", `$${avgMonthly.toLocaleString()}`, T.text],
              ["Lowest Month", `$${lowestMonth.toLocaleString()}`, T.red],
              ["Highest Month", `$${highestMonth.toLocaleString()}`, T.green],
            ].map(([label, val, color]) => (
              <div key={label as string} style={{ display: "flex", justifyContent: "space-between", fontSize: 13, padding: "7px 0", borderBottom: `1px solid ${T.border}` }}>
                <span style={{ color: T.muted }}>{label}</span>
                <span style={{ fontWeight: 700, color: color as string }}>{val}</span>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* ────────────────────────────── EXPENSES ────────────────────────────── */}
      {tab === "expenses" && (
        <div>
          {(["personal", "business"] as const).map(type => (
            <div key={type} style={{ marginBottom: 18 }}>
              <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 8, textTransform: "capitalize", color: T.text }}>
                {type} Expenses
              </div>
              {expenses[type].map((e, i) => (
                <div key={i} style={{ display: "flex", alignItems: "center", gap: 8, ...card, marginBottom: 6, padding: "9px 12px" }}>
                  <input
                    value={e.label}
                    onChange={ev => updateExpense(type, i, "label", ev.target.value)}
                    style={{ flex: 1, fontSize: 12, border: "none", background: "transparent", color: T.text, outline: "none" }}
                  />
                  <span style={{ color: T.muted, fontSize: 13 }}>$</span>
                  <input
                    type="number"
                    value={e.amount || ""}
                    onChange={ev => updateExpense(type, i, "amount", ev.target.value)}
                    placeholder="0"
                    style={{ width: 80, fontSize: 13, fontWeight: 700, border: "none", background: "transparent", color: T.text, outline: "none", textAlign: "right" }}
                  />
                </div>
              ))}
              <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, padding: "7px 12px", color: T.muted }}>
                <span>Subtotal</span>
                <span style={{ fontWeight: 700, color: T.text }}>${(type === "personal" ? totalPersonal : totalBusiness).toLocaleString()}</span>
              </div>
            </div>
          ))}

          <div style={{ ...card, borderLeft: `4px solid ${T.amber}` }}>
            <div style={{ display: "flex", justifyContent: "space-between", fontSize: 15, fontWeight: 700 }}>
              <span>Break-Even Point</span>
              <span style={{ color: T.amber }}>${breakEven.toLocaleString()}/mo</span>
            </div>
            <div style={{ fontSize: 11, color: T.muted, marginTop: 4 }}>The minimum you must earn every month to cover all costs.</div>
          </div>
        </div>
      )}

      {/* ───────────────────────────────── TAX ──────────────────────────────── */}
      {tab === "tax" && (
        <div>
          <div style={{ fontSize: 12, color: T.muted, marginBottom: 12 }}>Set aside this amount every time a payment arrives. Never touch it.</div>

          <div style={{ ...card, marginBottom: 10 }}>
            <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 12 }}>Your Tax Rate</div>
            <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
              <input
                type="range" min={15} max={40} value={taxRate}
                onChange={e => setTaxRate(Number(e.target.value))}
                style={{ flex: 1, accentColor: T.blue }}
              />
              <div style={{ fontSize: 26, fontWeight: 800, color: T.blue, minWidth: 60 }}>{taxRate}%</div>
            </div>
            <div style={{ fontSize: 11, color: T.muted, marginTop: 4 }}>Recommended: 25–30% for most freelancers</div>
          </div>

          <div style={{ display: "grid", gridTemplateColumns: "repeat(2,1fr)", gap: 10, marginBottom: 10 }}>
            {[
              { label: "Set aside per month", value: `$${taxReserve.toLocaleString()}` },
              { label: "Annual tax estimate", value: `$${(taxReserve * 12).toLocaleString()}` },
            ].map(m => (
              <div key={m.label} style={card}>
                <div style={{ fontSize: 10, color: T.muted, marginBottom: 6 }}>{m.label}</div>
                <div style={{ fontSize: 22, fontWeight: 700, color: T.red }}>{m.value}</div>
              </div>
            ))}
          </div>

          <div style={card}>
            <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 10 }}>The Rule (from Chapter 4)</div>
            {[500, 1000, 2000, 3000, 5000].map(r => (
              <div key={r} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", fontSize: 13, padding: "7px 0", borderBottom: `1px solid ${T.border}` }}>
                <span style={{ color: T.muted }}>Receive ${r.toLocaleString()}</span>
                <span>→ Set aside <strong style={{ color: T.red }}>${Math.round(r * taxRate / 100)}</strong></span>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* ──────────────────────────────── SAVINGS ───────────────────────────── */}
      {tab === "savings" && (
        <div>
          <div style={{ fontSize: 12, color: T.muted, marginBottom: 12 }}>Track your 4-layer safety net from Chapter 7.</div>

          <div style={{ ...card, marginBottom: 10 }}>
            <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 4 }}>10% Savings Rule</div>
            <div style={{ fontSize: 11, color: T.muted, marginBottom: 12 }}>Save 10% of every payment automatically.</div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 8 }}>
              {[500, 1000, 2000, 3000, 4000, 5000].map(amt => (
                <div key={amt} style={{ background: T.bg, borderRadius: 8, padding: "8px 10px", textAlign: "center" }}>
                  <div style={{ fontSize: 10, color: T.muted }}>Receive ${amt.toLocaleString()}</div>
                  <div style={{ fontSize: 15, fontWeight: 700, color: T.green }}>Save ${(amt * 0.1).toLocaleString()}</div>
                </div>
              ))}
            </div>
          </div>

          <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 8 }}>Your 4-Layer Safety Net</div>
          {[
            { layer: "Layer 1", title: "Tax Reserve",       desc: "25–30% of every payment",       target: taxReserve * 3, color: T.red },
            { layer: "Layer 2", title: "Short-Term Buffer", desc: "1 month of break-even expenses", target: breakEven,      color: T.amber },
            { layer: "Layer 3", title: "Emergency Fund",    desc: "3–6 months of living expenses",  target: breakEven * 4,  color: T.green },
            { layer: "Layer 4", title: "Retirement Fund",   desc: "5–10% of income monthly",        target: avgMonthly * 0.07 * 12, color: T.blue },
          ].map((l, i) => (
            <div key={i} style={{ ...card, marginBottom: 8, borderLeft: `4px solid ${l.color}` }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <div>
                  <div style={{ fontSize: 10, color: T.muted }}>{l.layer}</div>
                  <div style={{ fontSize: 13, fontWeight: 600 }}>{l.title}</div>
                  <div style={{ fontSize: 11, color: T.muted, marginTop: 2 }}>{l.desc}</div>
                </div>
                <div style={{ fontSize: 18, fontWeight: 700, color: l.color }}>
                  ${Math.round(l.target).toLocaleString()}
                </div>
              </div>
            </div>
          ))}

          <div style={card}>
            <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13 }}>
              <span style={{ color: T.muted }}>Your monthly savings target</span>
              <span style={{ fontWeight: 700, color: T.green }}>${savingsTarget.toLocaleString()}</span>
            </div>
            <div style={{ fontSize: 11, color: T.muted, marginTop: 4 }}>10% of your average monthly income</div>
          </div>
        </div>
      )}

      {/* ── Footer ──────────────────────────────────────────────────────────── */}
      <div style={{ textAlign: "center", fontSize: 10, color: T.muted, marginTop: 24, paddingTop: 14, borderTop: `1px solid ${T.border}` }}>
        © 2026 Ahmed Mohie — The Freelancer's Financial Playbook Bonus Template
      </div>
    </div>
  );
}
