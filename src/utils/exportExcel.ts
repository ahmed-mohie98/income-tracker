import ExcelJS from "exceljs";
import { saveAs } from "file-saver";

// ─── Colour palette ───────────────────────────────────────────────────────────
const GREEN = "FF1D9E75";
const AMBER = "FFBA7517";
const RED = "FFE05252";
const BLUE = "FF2563EB";
const DARK = "FF111827";
const LIGHT_BG = "FFF9FAFB";
const CARD_BG = "FFFFFFFF";
const BORDER = "FFE5E7EB";
const HEADER_BG = "FF1E293B";
const HEADER_FG = "FFFFFFFF";
const YELLOW_EDIT = "FFFFF9C4";

// ─── Helper: apply a thin border to a range ───────────────────────────────────
function borderRange(
  ws: ExcelJS.Worksheet,
  fromRow: number, toRow: number,
  fromCol: number, toCol: number
) {
  const side: ExcelJS.BorderStyle = "thin";
  for (let r = fromRow; r <= toRow; r++) {
    for (let c = fromCol; c <= toCol; c++) {
      ws.getRow(r).getCell(c).border = {
        top: { style: side, color: { argb: BORDER } },
        bottom: { style: side, color: { argb: BORDER } },
        left: { style: side, color: { argb: BORDER } },
        right: { style: side, color: { argb: BORDER } },
        diagonal: { style: undefined },
      };
    }
  }
}

// ─── Helper: fill a range ────────────────────────────────────────────────────
function fillRange(
  ws: ExcelJS.Worksheet,
  fromRow: number, toRow: number,
  fromCol: number, toCol: number,
  argb: string
) {
  for (let r = fromRow; r <= toRow; r++) {
    for (let c = fromCol; c <= toCol; c++) {
      ws.getRow(r).getCell(c).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb },
      };
    }
  }
}

// ─── Style helper ─────────────────────────────────────────────────────────────
type CellStyle = {
  bold?: boolean;
  color?: string;
  fill?: string;
  size?: number;
  align?: ExcelJS.Alignment["horizontal"];
  italic?: boolean;
  numFmt?: string;
  wrap?: boolean;
};

function applyStyle(cell: ExcelJS.Cell, s: CellStyle) {
  cell.font = {
    name: "Calibri",
    bold: s.bold ?? false,
    italic: s.italic ?? false,
    size: s.size ?? 11,
    color: s.color ? { argb: s.color } : undefined,
  };
  if (s.fill) {
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: s.fill } };
  }
  if (s.align) {
    cell.alignment = { horizontal: s.align, vertical: "middle", wrapText: s.wrap };
  }
  if (s.numFmt) {
    cell.numFmt = s.numFmt;
  }
}

function sc(ws: ExcelJS.Worksheet, addr: string, value: ExcelJS.CellValue, s: CellStyle = {}) {
  const cell = ws.getCell(addr);
  cell.value = value;
  applyStyle(cell, s);
  return cell;
}

function sf(
  ws: ExcelJS.Worksheet, addr: string,
  formula: string, result: number | string,
  s: CellStyle = {}
) {
  const cell = ws.getCell(addr);
  cell.value = { formula, result: typeof result === "number" ? result : 0 };
  applyStyle(cell, s);
  return cell;
}

// ─── INSTRUCTIONS sheet ───────────────────────────────────────────────────────
function buildInstructions(wb: ExcelJS.Workbook) {
  const ws = wb.addWorksheet("📋 Instructions", {
    properties: { tabColor: { argb: BLUE } },
  });
  ws.columns = [{ width: 2 }, { width: 65 }, { width: 2 }];
  fillRange(ws, 1, 50, 1, 3, "FFFAFAFA");

  ws.getRow(1).height = 14;
  sc(ws, "B2", "📋  How to use this workbook", { bold: true, size: 16, color: DARK });

  const lines: [string, CellStyle][] = [
    ["", {}],
    ["WELCOME — The Freelancer's Financial Playbook Income Tracker", { bold: true, size: 13, color: DARK }],
    ["", {}],
    ["HOW TO USE:", { bold: true, size: 12, color: DARK }],
    ["", {}],
    ["  1.  📊  Income tab  —  Enter your monthly income in the yellow cells (Jan–Dec).", { size: 11, color: DARK }],
    ["  2.  💰  Expenses tab  —  Edit personal & business expense labels and amounts in yellow cells.", { size: 11, color: DARK }],
    ["  3.  🧾  Tax tab  —  Set your tax rate % in cell B3 (yellow). All tables auto-update.", { size: 11, color: DARK }],
    ["  4.  🏠  Dashboard tab  —  All KPIs update automatically from your inputs.", { size: 11, color: GREEN, bold: true }],
    ["  5.  💼  Savings tab  —  Shows your 4-layer safety net targets.", { size: 11, color: DARK }],
    ["", {}],
    ["⚠️  IMPORTANT: Only edit YELLOW cells. Do NOT overwrite white formula cells.", { size: 12, color: "FF92400E", bold: true }],
    ["", {}],
    ["COLOR GUIDE:", { size: 11, bold: true, color: DARK }],
    ["  🟩 Green  = healthy / above target", { size: 11, color: GREEN }],
    ["  🟧 Amber  = warning / at risk", { size: 11, color: AMBER }],
    ["  🟥 Red    = below target / needs attention", { size: 11, color: RED }],
    ["  🟦 Blue   = informational", { size: 11, color: BLUE }],
    ["  🟨 Yellow = editable input cell", { size: 11, color: "FF92400E" }],
    ["", {}],
    ["SHEET ORDER:", { size: 11, bold: true, color: DARK }],
    ["  Instructions → Dashboard → Income → Expenses → Tax → Savings", { size: 11, color: DARK }],
    ["", {}],
    ["All formulas link across sheets — changes in Income/Expenses/Tax instantly update Dashboard.", { size: 11, color: "FF6B7280", italic: true }],
    ["", {}],
    ["© 2026 Ahmed Mohie — The Freelancer's Financial Playbook", { size: 9, color: "FF6B7280", italic: true }],
  ];

  for (let i = 0; i < lines.length; i++) {
    const row = 3 + i;
    ws.getRow(row).height = 20;
    sc(ws, `B${row}`, lines[i][0], lines[i][1]);
  }
  ws.views = [{ showGridLines: false }];
}

// ─── DASHBOARD sheet ─────────────────────────────────────────────────────────
function buildDashboard(wb: ExcelJS.Workbook, name: string) {
  const ws = wb.addWorksheet("Dashboard", {
    properties: { tabColor: { argb: BLUE } },
  });

  ws.columns = [
    { width: 2 },   // A – margin
    { width: 24 },  // B
    { width: 20 },  // C
    { width: 20 },  // D
    { width: 20 },  // E
    { width: 2 },   // F – margin
  ];

  fillRange(ws, 1, 60, 1, 6, "FFFAFAFA");

  // ── Title ──
  ws.getRow(1).height = 14;
  ws.mergeCells("B2:E2");
  sc(ws, "B2", "THE FREELANCER'S FINANCIAL PLAYBOOK", { size: 9, italic: true, color: "FF6B7280" });

  ws.getRow(3).height = 34;
  ws.mergeCells("B3:C3");
  sc(ws, "B3", "Income Tracker — Dashboard", { bold: true, size: 18, color: DARK });
  ws.mergeCells("D3:E3");
  sc(ws, "D3", name, { bold: true, size: 13, color: DARK, align: "right" });

  ws.getRow(4).height = 10;

  // ── KPI row 1: labels ──
  ws.getRow(5).height = 18;
  sc(ws, "B5", "Avg Monthly Income", { size: 9, color: "FF6B7280", fill: LIGHT_BG });
  sc(ws, "C5", "Break-Even / mo", { size: 9, color: "FF6B7280", fill: LIGHT_BG });
  sc(ws, "D5", "Monthly Profit", { size: 9, color: "FF6B7280", fill: LIGHT_BG });

  // ── KPI row 1: values ──
  ws.getRow(6).height = 40;
  // B6 = avg monthly income (referenced by other sheets)
  sf(ws, "B6", "ROUND(Income!N2/12,0)", 0, { bold: true, size: 22, color: GREEN, fill: LIGHT_BG, numFmt: '"$"#,##0' });
  // C6 = break-even (= Expenses D{beRow})
  sf(ws, "C6", "Expenses!D32", 0, { bold: true, size: 22, color: AMBER, fill: LIGHT_BG, numFmt: '"$"#,##0' });
  sf(ws, "D6", "B6-C6", 0, { bold: true, size: 22, color: GREEN, fill: LIGHT_BG, numFmt: '"$"#,##0' });
  borderRange(ws, 5, 6, 2, 4);

  ws.getRow(7).height = 8;

  // ── KPI row 2 ──
  ws.getRow(8).height = 18;
  sc(ws, "B8", "Tax Reserve / mo", { size: 9, color: "FF6B7280", fill: LIGHT_BG });
  sc(ws, "C8", "Savings Target / mo", { size: 9, color: "FF6B7280", fill: LIGHT_BG });
  sc(ws, "D8", "Freelancer Salary", { size: 9, color: "FF6B7280", fill: LIGHT_BG });

  ws.getRow(9).height = 40;
  // B9 = tax reserve (referenced by Savings)
  sf(ws, "B9", "ROUND(B6*Tax!B3/100,0)", 0, { bold: true, size: 22, color: RED, fill: LIGHT_BG, numFmt: '"$"#,##0' });
  sf(ws, "C9", "ROUND(B6*0.1,0)", 0, { bold: true, size: 22, color: DARK, fill: LIGHT_BG, numFmt: '"$"#,##0' });
  sf(ws, "D9", "ROUND(C6*1.15,0)", 0, { bold: true, size: 22, color: DARK, fill: LIGHT_BG, numFmt: '"$"#,##0' });
  borderRange(ws, 8, 9, 2, 4);

  ws.getRow(10).height = 12;

  // ── Financial Health Check ──
  ws.getRow(11).height = 24;
  ws.mergeCells("B11:E11");
  sc(ws, "B11", "Financial Health Check", { bold: true, size: 13, color: DARK });

  const healthItems = [
    {
      label: "Income vs Break-Even",
      valFormula: "B6",
      targetFormula: "C6",
      displayFormula: `"$"&TEXT(B6,"#,##0")&" / $"&TEXT(C6,"#,##0")`,
    },
    {
      label: "Lowest Month vs Break-Even",
      valFormula: "Income!O2",
      targetFormula: "C6",
      displayFormula: `"$"&TEXT(Income!O2,"#,##0")&" / $"&TEXT(C6,"#,##0")`,
    },
    {
      label: "Annual Income",
      valFormula: "Income!N2",
      targetFormula: "C6*12",
      displayFormula: `"$"&TEXT(Income!N2,"#,##0")`,
    },
  ];

  let r = 12;
  for (const h of healthItems) {
    ws.getRow(r).height = 22;

    sc(ws, `B${r}`, h.label, { size: 11, color: DARK, fill: CARD_BG });

    sf(
      ws, `C${r}`,
      `IF(${h.valFormula}>=${h.targetFormula},"✅  On Track",IF(${h.valFormula}>=${h.targetFormula}*0.7,"⚠️  At Risk","❌  Below"))`,
      "—",
      { size: 10, fill: CARD_BG }
    );

    sf(ws, `D${r}`, h.displayFormula, "—", {
      bold: true, size: 11, fill: CARD_BG, color: GREEN, align: "right",
    });

    borderRange(ws, r, r, 2, 4);
    r++;
  }

  ws.getRow(r).height = 10;
  r++;

  // ── Quick Tips ──
  ws.mergeCells(`B${r}:E${r}`);
  sc(ws, `B${r}`, "📊  Income chart is auto-generated on the Income sheet — see columns B–G rows 20–37", {
    size: 10, color: "FF6B7280", italic: true,
  });

  ws.views = [{ showGridLines: false }];
}

// ─── INCOME sheet ─────────────────────────────────────────────────────────────
function buildIncome(wb: ExcelJS.Workbook, income: { month: string; amount: number }[]) {
  const ws = wb.addWorksheet("Income", {
    properties: { tabColor: { argb: GREEN } },
  });

  ws.columns = [
    { width: 2 },   // A
    { width: 14 },  // B – Month
    { width: 18 },  // C – Amount (editable)
    { width: 22 },  // D – Status
    { width: 2 },   // E
    { width: 20 },  // F – Summary label
    { width: 18 },  // G – Summary value
    { width: 2 },   // H
    // hidden reference cols
    { width: 0.1 }, // I
    { width: 0.1 }, // J
    { width: 0.1 }, // K
    { width: 0.1 }, // L
    { width: 0.1 }, // M
    { width: 0.1 }, // N – total annual
    { width: 0.1 }, // O – lowest month
    { width: 0.1 }, // P – highest month
  ];

  fillRange(ws, 1, 80, 1, 16, "FFFAFAFA");

  // Title
  ws.getRow(1).height = 14;
  ws.mergeCells("B2:G2");
  sc(ws, "B2", "Monthly Income — Enter your earnings for each month (yellow cells)", { bold: true, size: 13, color: DARK });
  ws.getRow(3).height = 10;

  // Header
  ws.getRow(4).height = 22;
  sc(ws, "B4", "Month", { bold: true, size: 10, color: HEADER_FG, fill: HEADER_BG, align: "center" });
  sc(ws, "C4", "Income ($)", { bold: true, size: 10, color: HEADER_FG, fill: HEADER_BG, align: "center" });
  sc(ws, "D4", "Status vs Break-Even", { bold: true, size: 10, color: HEADER_FG, fill: HEADER_BG, align: "center" });
  borderRange(ws, 4, 4, 2, 4);

  sc(ws, "F4", "Summary", { bold: true, size: 10, color: HEADER_FG, fill: HEADER_BG });
  sc(ws, "G4", "Value", { bold: true, size: 10, color: HEADER_FG, fill: HEADER_BG, align: "right" });
  borderRange(ws, 4, 4, 6, 7);

  const MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

  for (let i = 0; i < 12; i++) {
    const row = 5 + i;
    ws.getRow(row).height = 22;
    const bg = i % 2 === 0 ? CARD_BG : LIGHT_BG;

    // Month label
    sc(ws, `B${row}`, MONTHS[i], { bold: true, size: 12, color: DARK, fill: bg, align: "center" });

    // Income amount — yellow editable
    const amtCell = ws.getCell(`C${row}`);
    amtCell.value = income[i]?.amount ?? 0;
    amtCell.numFmt = '"$"#,##0';
    amtCell.font = { name: "Calibri", size: 13, bold: true, color: { argb: DARK } };
    amtCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: YELLOW_EDIT } };
    amtCell.alignment = { horizontal: "right", vertical: "middle" };

    // Status formula
    sf(
      ws, `D${row}`,
      `IF(C${row}>=Expenses!D32,"✅ Above Break-Even",IF(C${row}>0,"⚠️ Below Break-Even","— No Data"))`,
      "— No Data",
      { size: 10, color: "FF6B7280", fill: bg, align: "center" }
    );

    borderRange(ws, row, row, 2, 4);
  }

  // Summary section
  const summaryItems: { label: string; formula: string; color: string }[] = [
    { label: "Total Annual Income", formula: "SUM(C5:C16)", color: DARK },
    { label: "Monthly Average", formula: "ROUND(SUM(C5:C16)/12,0)", color: DARK },
    { label: "Lowest Month", formula: "MIN(C5:C16)", color: RED },
    { label: "Highest Month", formula: "MAX(C5:C16)", color: GREEN },
    { label: "Months ≥ Break-Even", formula: `COUNTIF(C5:C16,">="&Expenses!D32)`, color: BLUE },
    { label: "Months Below Break-Even", formula: `COUNTIF(C5:C16,"<"&Expenses!D32)-COUNTIF(C5:C16,0)`, color: AMBER },
  ];

  for (let i = 0; i < summaryItems.length; i++) {
    const row = 5 + i;
    const bg = i % 2 === 0 ? CARD_BG : LIGHT_BG;
    ws.getRow(row).height = 22;
    sc(ws, `F${row}`, summaryItems[i].label, { size: 10, color: DARK, fill: bg });
    sf(ws, `G${row}`, summaryItems[i].formula, 0, {
      bold: true, size: 12, color: summaryItems[i].color,
      fill: bg, align: "right",
      numFmt: i < 4 ? '"$"#,##0' : "0",
    });
    borderRange(ws, row, row, 6, 7);
  }

  ws.getRow(17).height = 8;

  // Reference cells for Dashboard (hidden columns)
  sf(ws, "N2", "SUM(C5:C16)", 0, { numFmt: '"$"#,##0' });
  sf(ws, "O2", "MIN(C5:C16)", 0, { numFmt: '"$"#,##0' });
  sf(ws, "P2", "MAX(C5:C16)", 0, { numFmt: '"$"#,##0' });

  // Notes row
  ws.getRow(18).height = 18;
  ws.mergeCells("B18:G18");
  sc(ws, "B18", "💡 Tip: Green = above break-even | Red = below | Yellow cells are editable", {
    size: 9, color: "FF6B7280", italic: true,
  });

  ws.views = [{ showGridLines: false }];
}

// ─── EXPENSES sheet ───────────────────────────────────────────────────────────
function buildExpenses(
  wb: ExcelJS.Workbook,
  expenses: {
    personal: { label: string; amount: number }[];
    business: { label: string; amount: number }[];
  }
) {
  const ws = wb.addWorksheet("Expenses", {
    properties: { tabColor: { argb: AMBER } },
  });

  ws.columns = [
    { width: 2 },   // A
    { width: 28 },  // B – Label
    { width: 18 },  // C – Amount
    { width: 0.1 }, // D – break-even reference (D32)
    { width: 2 },   // E
  ];

  fillRange(ws, 1, 80, 1, 5, "FFFAFAFA");

  ws.getRow(1).height = 14;
  ws.mergeCells("B2:C2");
  sc(ws, "B2", "Monthly Expenses — Edit labels and amounts in yellow cells", { bold: true, size: 13, color: DARK });
  ws.getRow(3).height = 10;

  // ── PERSONAL ──
  ws.getRow(4).height = 22;
  sc(ws, "B4", "Personal Expenses", { bold: true, size: 10, color: HEADER_FG, fill: HEADER_BG });
  sc(ws, "C4", "Monthly Amount ($)", { bold: true, size: 10, color: HEADER_FG, fill: HEADER_BG, align: "right" });
  borderRange(ws, 4, 4, 2, 3);

  const personalStart = 5;
  for (let i = 0; i < expenses.personal.length; i++) {
    const row = personalStart + i;
    ws.getRow(row).height = 22;
    const bg = i % 2 === 0 ? CARD_BG : LIGHT_BG;

    const labelCell = ws.getCell(`B${row}`);
    labelCell.value = expenses.personal[i].label;
    labelCell.font = { name: "Calibri", size: 11, color: { argb: DARK } };
    labelCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: YELLOW_EDIT } };
    labelCell.alignment = { vertical: "middle" };

    const amtCell = ws.getCell(`C${row}`);
    amtCell.value = Number(expenses.personal[i].amount) || 0;
    amtCell.numFmt = '"$"#,##0';
    amtCell.font = { name: "Calibri", size: 12, bold: true };
    amtCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: YELLOW_EDIT } };
    amtCell.alignment = { horizontal: "right", vertical: "middle" };

    // put bg on D
    ws.getCell(`D${row}`).fill = { type: "pattern", pattern: "solid", fgColor: { argb: bg } };
    borderRange(ws, row, row, 2, 3);
  }

  const personalEnd = personalStart + expenses.personal.length - 1;
  const personalTotalRow = personalEnd + 1;
  ws.getRow(personalTotalRow).height = 24;
  sc(ws, `B${personalTotalRow}`, "Personal Subtotal", { bold: true, size: 11, color: DARK, fill: "FFECFDF5" });
  sf(ws, `C${personalTotalRow}`, `SUM(C${personalStart}:C${personalEnd})`, 0, {
    bold: true, size: 13, color: GREEN, fill: "FFECFDF5", numFmt: '"$"#,##0', align: "right",
  });
  borderRange(ws, personalTotalRow, personalTotalRow, 2, 3);

  ws.getRow(personalTotalRow + 1).height = 10;

  // ── BUSINESS ──
  const bizHeaderRow = personalTotalRow + 2;
  ws.getRow(bizHeaderRow).height = 22;
  sc(ws, `B${bizHeaderRow}`, "Business Expenses", { bold: true, size: 10, color: HEADER_FG, fill: HEADER_BG });
  sc(ws, `C${bizHeaderRow}`, "Monthly Amount ($)", { bold: true, size: 10, color: HEADER_FG, fill: HEADER_BG, align: "right" });
  borderRange(ws, bizHeaderRow, bizHeaderRow, 2, 3);

  const bizStart = bizHeaderRow + 1;
  for (let i = 0; i < expenses.business.length; i++) {
    const row = bizStart + i;
    ws.getRow(row).height = 22;

    const labelCell = ws.getCell(`B${row}`);
    labelCell.value = expenses.business[i].label;
    labelCell.font = { name: "Calibri", size: 11, color: { argb: DARK } };
    labelCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: YELLOW_EDIT } };
    labelCell.alignment = { vertical: "middle" };

    const amtCell = ws.getCell(`C${row}`);
    amtCell.value = Number(expenses.business[i].amount) || 0;
    amtCell.numFmt = '"$"#,##0';
    amtCell.font = { name: "Calibri", size: 12, bold: true };
    amtCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: YELLOW_EDIT } };
    amtCell.alignment = { horizontal: "right", vertical: "middle" };

    borderRange(ws, row, row, 2, 3);
  }

  const bizEnd = bizStart + expenses.business.length - 1;
  const bizTotalRow = bizEnd + 1;
  ws.getRow(bizTotalRow).height = 24;
  sc(ws, `B${bizTotalRow}`, "Business Subtotal", { bold: true, size: 11, color: DARK, fill: "FFEFF6FF" });
  sf(ws, `C${bizTotalRow}`, `SUM(C${bizStart}:C${bizEnd})`, 0, {
    bold: true, size: 13, color: BLUE, fill: "FFEFF6FF", numFmt: '"$"#,##0', align: "right",
  });
  borderRange(ws, bizTotalRow, bizTotalRow, 2, 3);

  // Break-even row — this MUST be row 32 for Dashboard reference
  // personalStart=5, personal.length=7 → personalEnd=11, personalTotalRow=12
  // bizHeaderRow=14, bizStart=15, business.length=7 → bizEnd=21, bizTotalRow=22
  // beRow = 22+2 = 24  ... but Dashboard refs D32. Let's pad to 32.
  const beRow = 32;
  // fill any gap rows
  for (let rr = bizTotalRow + 1; rr < beRow; rr++) {
    ws.getRow(rr).height = 8;
  }

  ws.getRow(beRow).height = 30;
  sc(ws, `B${beRow}`, "🎯  Break-Even Point / month", { bold: true, size: 13, color: DARK, fill: "FFFFF7ED" });
  sf(ws, `C${beRow}`, `C${personalTotalRow}+C${bizTotalRow}`, 0, {
    bold: true, size: 16, color: AMBER, fill: "FFFFF7ED", numFmt: '"$"#,##0', align: "right",
  });
  // D32 referenced by Dashboard!C6 and Income status formulas
  sf(ws, `D${beRow}`, `C${personalTotalRow}+C${bizTotalRow}`, 0, { numFmt: '"$"#,##0' });
  borderRange(ws, beRow, beRow, 2, 3);

  ws.getRow(beRow + 1).height = 16;
  sc(ws, `B${beRow + 1}`, "= The minimum you must earn every month to cover all costs", {
    size: 9, color: "FF6B7280", italic: true,
  });

  ws.views = [{ showGridLines: false }];
}

// ─── TAX sheet ────────────────────────────────────────────────────────────────
function buildTax(wb: ExcelJS.Workbook, taxRate: number) {
  const ws = wb.addWorksheet("Tax", {
    properties: { tabColor: { argb: RED } },
  });

  ws.columns = [
    { width: 2 },   // A
    { width: 22 },  // B  (B3 = tax rate, referenced by Dashboard)
    { width: 22 },  // C
    { width: 2 },   // D
  ];

  fillRange(ws, 1, 60, 1, 4, "FFFAFAFA");

  ws.getRow(1).height = 14;
  ws.mergeCells("B2:C2");
  sc(ws, "B2", "Tax Planning — Set your tax rate and track reserves", { bold: true, size: 13, color: DARK });

  // B3 = tax rate (MUST be B3 for Dashboard formula)
  ws.getRow(3).height = 38;
  const rateCell = ws.getCell("B3");
  rateCell.value = taxRate;
  rateCell.numFmt = "0";
  rateCell.font = { name: "Calibri", size: 26, bold: true, color: { argb: RED } };
  rateCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: YELLOW_EDIT } };
  rateCell.alignment = { horizontal: "center", vertical: "middle" };

  sc(ws, "C3", "← Your Tax Rate (%) — edit this yellow cell", { size: 10, color: "FF6B7280", italic: true });

  ws.getRow(4).height = 14;
  sc(ws, "B4", "Recommended: 25–30% for most freelancers", { size: 9, color: "FF6B7280", italic: true });

  ws.getRow(5).height = 10;

  // KPI cards
  ws.getRow(6).height = 22;
  sc(ws, "B6", "Monthly Tax Reserve", { bold: true, size: 10, color: HEADER_FG, fill: HEADER_BG });
  sc(ws, "C6", "Annual Tax Estimate", { bold: true, size: 10, color: HEADER_FG, fill: HEADER_BG, align: "right" });
  borderRange(ws, 6, 6, 2, 3);

  ws.getRow(7).height = 38;
  sf(ws, "B7", "ROUND(Dashboard!B6*B3/100,0)", 0, { bold: true, size: 22, color: RED, fill: LIGHT_BG, numFmt: '"$"#,##0' });
  sf(ws, "C7", "ROUND(Dashboard!B6*B3/100,0)*12", 0, { bold: true, size: 22, color: RED, fill: LIGHT_BG, numFmt: '"$"#,##0', align: "right" });
  borderRange(ws, 7, 7, 2, 3);

  ws.getRow(8).height = 10;

  // Reference table
  ws.getRow(9).height = 22;
  sc(ws, "B9", "If you receive…", { bold: true, size: 10, color: HEADER_FG, fill: HEADER_BG, align: "center" });
  sc(ws, "C9", `Set aside (${taxRate}%)…`, { bold: true, size: 10, color: HEADER_FG, fill: HEADER_BG, align: "center" });
  borderRange(ws, 9, 9, 2, 3);

  const amounts = [500, 1000, 1500, 2000, 3000, 4000, 5000, 7500, 10000];
  for (let i = 0; i < amounts.length; i++) {
    const row = 10 + i;
    ws.getRow(row).height = 22;
    const bg = i % 2 === 0 ? CARD_BG : LIGHT_BG;

    const receiveCell = ws.getCell(`B${row}`);
    receiveCell.value = amounts[i];
    receiveCell.numFmt = '"$"#,##0';
    receiveCell.font = { name: "Calibri", size: 12, color: { argb: DARK } };
    receiveCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: bg } };
    receiveCell.alignment = { horizontal: "center", vertical: "middle" };

    sf(ws, `C${row}`, `ROUND(B${row}*$B$3/100,0)`, Math.round(amounts[i] * taxRate / 100), {
      bold: true, size: 12, color: RED, fill: bg, numFmt: '"$"#,##0', align: "center",
    });
    borderRange(ws, row, row, 2, 3);
  }

  ws.getRow(19).height = 10;
  ws.mergeCells("B20:C20");
  sc(ws, "B20", `💡 The Rule: Every time money arrives, immediately transfer ${taxRate}% to a dedicated tax account. Never touch it.`, {
    size: 10, color: "FF92400E", italic: true, wrap: true,
  });
  ws.getRow(20).height = 30;

  ws.views = [{ showGridLines: false }];
}

// ─── SAVINGS sheet ────────────────────────────────────────────────────────────
function buildSavings(wb: ExcelJS.Workbook) {
  const ws = wb.addWorksheet("Savings", {
    properties: { tabColor: { argb: GREEN } },
  });

  ws.columns = [
    { width: 2 },   // A
    { width: 22 },  // B
    { width: 26 },  // C
    { width: 20 },  // D
    { width: 2 },   // E
  ];

  fillRange(ws, 1, 80, 1, 5, "FFFAFAFA");

  ws.getRow(1).height = 14;
  ws.mergeCells("B2:D2");
  sc(ws, "B2", "Savings Strategy — Your 4-Layer Safety Net", { bold: true, size: 13, color: DARK });
  ws.getRow(3).height = 10;

  // 10% rule header
  ws.getRow(4).height = 22;
  sc(ws, "B4", "10% Savings Rule", { bold: true, size: 10, color: HEADER_FG, fill: HEADER_BG });
  sc(ws, "C4", "If you receive…", { bold: true, size: 10, color: HEADER_FG, fill: HEADER_BG, align: "center" });
  sc(ws, "D4", "Save at least…", { bold: true, size: 10, color: HEADER_FG, fill: HEADER_BG, align: "center" });
  borderRange(ws, 4, 4, 2, 4);

  const amounts = [500, 1000, 1500, 2000, 3000, 4000, 5000, 7500, 10000];
  for (let i = 0; i < amounts.length; i++) {
    const row = 5 + i;
    ws.getRow(row).height = 22;
    const bg = i % 2 === 0 ? CARD_BG : LIGHT_BG;

    const receiveCell = ws.getCell(`C${row}`);
    receiveCell.value = amounts[i];
    receiveCell.numFmt = '"$"#,##0';
    receiveCell.font = { name: "Calibri", size: 12, color: { argb: DARK } };
    receiveCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: bg } };
    receiveCell.alignment = { horizontal: "center", vertical: "middle" };

    sf(ws, `D${row}`, `C${row}*0.1`, amounts[i] * 0.1, {
      bold: true, size: 12, color: GREEN, fill: bg, numFmt: '"$"#,##0', align: "center",
    });
    borderRange(ws, row, row, 2, 4);
  }

  ws.getRow(14).height = 12;

  // 4-Layer Safety Net
  ws.getRow(15).height = 26;
  ws.mergeCells("B15:D15");
  sc(ws, "B15", "Your 4-Layer Safety Net", { bold: true, size: 13, color: DARK });

  ws.getRow(16).height = 22;
  sc(ws, "B16", "Layer", { bold: true, size: 10, color: HEADER_FG, fill: HEADER_BG });
  sc(ws, "C16", "Description", { bold: true, size: 10, color: HEADER_FG, fill: HEADER_BG });
  sc(ws, "D16", "Target Amount", { bold: true, size: 10, color: HEADER_FG, fill: HEADER_BG, align: "right" });
  borderRange(ws, 16, 16, 2, 4);

  const layers = [
    { layer: "Layer 1", title: "Tax Reserve",       desc: "25–30% of every payment (3 mo buffer)", formula: "ROUND(Dashboard!B9*3,0)", color: RED },
    { layer: "Layer 2", title: "Short-Term Buffer", desc: "1 month of break-even expenses",          formula: "Expenses!D32",             color: AMBER },
    { layer: "Layer 3", title: "Emergency Fund",    desc: "3–6 months of living expenses",           formula: "Expenses!D32*4",           color: GREEN },
    { layer: "Layer 4", title: "Retirement Fund",   desc: "5–10% of income × 12 months",            formula: "Dashboard!B6*0.07*12",     color: BLUE },
  ];

  for (let i = 0; i < layers.length; i++) {
    const row = 17 + i;
    ws.getRow(row).height = 26;
    const l = layers[i];
    sc(ws, `B${row}`, `${l.layer}: ${l.title}`, { bold: true, size: 11, color: l.color, fill: CARD_BG });
    sc(ws, `C${row}`, l.desc, { size: 10, color: "FF6B7280", fill: CARD_BG });
    sf(ws, `D${row}`, l.formula, 0, { bold: true, size: 13, color: l.color, fill: CARD_BG, numFmt: '"$"#,##0', align: "right" });
    borderRange(ws, row, row, 2, 4);
  }

  ws.getRow(21).height = 10;

  ws.getRow(22).height = 28;
  sc(ws, "B22", "Your Monthly Savings Target (10%):", { bold: true, size: 11, color: DARK, fill: LIGHT_BG });
  sf(ws, "D22", "Dashboard!C9", 0, { bold: true, size: 15, color: GREEN, fill: LIGHT_BG, numFmt: '"$"#,##0', align: "right" });
  borderRange(ws, 22, 22, 2, 4);

  ws.getRow(23).height = 16;
  sc(ws, "B23", "= 10% of your average monthly income", { size: 9, color: "FF6B7280", italic: true });

  ws.views = [{ showGridLines: false }];
}

// ─── Main export function ─────────────────────────────────────────────────────
export async function exportToExcel(
  income: { month: string; amount: number }[],
  expenses: {
    personal: { label: string; amount: number }[];
    business: { label: string; amount: number }[];
  },
  taxRate: number,
  name: string
) {
  const wb = new ExcelJS.Workbook();
  wb.creator = name;
  wb.created = new Date();
  wb.modified = new Date();
  wb.title = "Freelancer Income Tracker";
  wb.subject = "The Freelancer's Financial Playbook";

  buildInstructions(wb);
  buildDashboard(wb, name);
  buildIncome(wb, income);
  buildExpenses(wb, expenses);
  buildTax(wb, taxRate);
  buildSavings(wb);

  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  saveAs(blob, `Freelancer_Income_Tracker_${name.replace(/\s+/g, "_")}.xlsx`);
}
