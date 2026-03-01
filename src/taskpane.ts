import {
  bondPrice,
  bondYtm,
  capm,
  dcfValue,
  discountedPayback,
  eacFromNpv,
  expectedReturn,
  formatNumber,
  formatPercent,
  frontierTwoAssetPoints,
  gordonGrowth,
  irr,
  mirr,
  normalizeRate,
  npv,
  parseNumberList,
  paybackPeriod,
  portfolioExpectedReturn,
  portfolioStd2Asset,
  portfolioStdNAsset,
  ratioAssetTurnover,
  ratioCurrent,
  ratioDebtToEquity,
  ratioInventoryTurnover,
  ratioProfitMargin,
  ratioQuick,
  ratioRoa,
  ratioRoe,
  ratioTie,
  sampleStdDev,
  sampleVariance,
  tvmLumpSum,
  wacc
} from "./finance";

type Field =
  | { key: string; label: string; type: "number"; defaultValue?: number }
  | { key: string; label: string; type: "text"; placeholder?: string }
  | { key: string; label: string; type: "select"; options: { value: string; label: string }[]; defaultValue?: string }
  | { key: string; label: string; type: "checkbox"; defaultValue?: boolean };

type Topic = {
  id: string;
  title: string;
  formula: string;
  fields: Field[];
  solve: (data: Record<string, string>) => { result: string; substituted: string; numeric?: number; details?: string[] };
};

const topics: Topic[] = [
  {
    id: "tvm",
    title: "TVM (Lump Sum)",
    formula: "PV = FV/(1+r)^n; FV = PV(1+r)^n; n = ln(FV/PV)/ln(1+r); r = (FV/PV)^(1/n)-1",
    fields: [
      { key: "solveFor", label: "Solve for", type: "select", defaultValue: "FV", options: [{ value: "PV", label: "PV" }, { value: "FV", label: "FV" }, { value: "N", label: "N" }, { value: "R", label: "R" }] },
      { key: "pv", label: "PV", type: "number", defaultValue: 1000 },
      { key: "fv", label: "FV", type: "number", defaultValue: 1200 },
      { key: "rate", label: "Rate (8 or 0.08)", type: "number", defaultValue: 8 },
      { key: "n", label: "Periods", type: "number", defaultValue: 2 }
    ],
    solve: (d) => {
      const result = tvmLumpSum(Number(d.pv), Number(d.fv), Number(d.rate), Number(d.n), d.solveFor);
      return {
        result: d.solveFor === "R" ? formatPercent(result) : formatNumber(result),
        substituted: `${d.solveFor} computed from PV=${d.pv}, FV=${d.fv}, r=${d.rate}, n=${d.n}`,
        numeric: result
      };
    }
  },
  {
    id: "npv_irr",
    title: "NPV / IRR",
    formula: "NPV = Σ CF_t/(1+r)^t; IRR is r where NPV(r)=0",
    fields: [
      { key: "mode", label: "Mode", type: "select", defaultValue: "NPV", options: [{ value: "NPV", label: "NPV" }, { value: "IRR", label: "IRR" }] },
      { key: "rate", label: "Discount rate (NPV only)", type: "number", defaultValue: 10 },
      { key: "cashflows", label: "Cash flows (comma/newline)", type: "text", placeholder: "-1000, 400, 500, 600" }
    ],
    solve: (d) => {
      const cashflows = parseNumberList(d.cashflows);
      if (!cashflows.length) throw new Error("Enter cash flows.");
      if (d.mode === "NPV") {
        const result = npv(Number(d.rate), cashflows);
        return { result: formatNumber(result), substituted: `NPV(${d.rate}, [${cashflows.join(", ")}])`, numeric: result };
      }
      const result = irr(cashflows);
      return { result: formatPercent(result), substituted: `IRR([${cashflows.join(", ")}])`, numeric: result };
    }
  },
  {
    id: "mirr",
    title: "MIRR",
    formula: "MIRR = (FV_positive / -PV_negative)^(1/n) - 1",
    fields: [
      { key: "financeRate", label: "Finance rate (8 or 0.08)", type: "number", defaultValue: 8 },
      { key: "reinvestRate", label: "Reinvest rate (8 or 0.08)", type: "number", defaultValue: 8 },
      { key: "cashflows", label: "Cash flows", type: "text", placeholder: "-1000, 200, 300, 700" }
    ],
    solve: (d) => {
      const cashflows = parseNumberList(d.cashflows);
      const result = mirr(cashflows, Number(d.financeRate), Number(d.reinvestRate));
      return { result: formatPercent(result), substituted: `MIRR([${cashflows.join(", ")}], ${d.financeRate}, ${d.reinvestRate})`, numeric: result };
    }
  },
  {
    id: "payback",
    title: "Payback / Discounted Payback",
    formula: "Payback is first period where cumulative cash flows become non-negative.",
    fields: [
      { key: "mode", label: "Mode", type: "select", defaultValue: "simple", options: [{ value: "simple", label: "Simple" }, { value: "discounted", label: "Discounted" }] },
      { key: "rate", label: "Discount rate (if discounted)", type: "number", defaultValue: 8 },
      { key: "cashflows", label: "Cash flows", type: "text", placeholder: "-1000, 200, 300, 500, 600" }
    ],
    solve: (d) => {
      const cfs = parseNumberList(d.cashflows);
      const result = d.mode === "simple" ? paybackPeriod(cfs) : discountedPayback(cfs, Number(d.rate));
      return { result: `${formatNumber(result)} periods`, substituted: `${d.mode} payback of [${cfs.join(", ")}]`, numeric: result };
    }
  },
  {
    id: "eac",
    title: "EAC",
    formula: "EAC = NPV / [(1-(1+r)^-n)/r]",
    fields: [
      { key: "npv", label: "NPV (typically PV of costs)", type: "number", defaultValue: 798.42 },
      { key: "rate", label: "Rate", type: "number", defaultValue: 10 },
      { key: "periods", label: "Periods", type: "number", defaultValue: 3 }
    ],
    solve: (d) => {
      const result = eacFromNpv(Number(d.npv), Number(d.rate), Number(d.periods));
      return { result: formatNumber(result), substituted: `EAC(${d.npv}, ${d.rate}, ${d.periods})`, numeric: result };
    }
  },
  {
    id: "wacc",
    title: "WACC",
    formula: "WACC = Wd*Rd*(1-T) + We*Re + Wp*Rp",
    fields: [
      { key: "wd", label: "Debt weight (0.4 or 40)", type: "number", defaultValue: 40 },
      { key: "we", label: "Equity weight (0.6 or 60)", type: "number", defaultValue: 60 },
      { key: "wp", label: "Preferred weight (optional)", type: "number", defaultValue: 0 },
      { key: "rd", label: "Cost of debt", type: "number", defaultValue: 6 },
      { key: "re", label: "Cost of equity", type: "number", defaultValue: 11 },
      { key: "rp", label: "Cost of preferred", type: "number", defaultValue: 0 },
      { key: "tax", label: "Tax rate", type: "number", defaultValue: 25 }
    ],
    solve: (d) => {
      const result = wacc(Number(d.wd), Number(d.we), Number(d.rd), Number(d.re), Number(d.tax), Number(d.wp), Number(d.rp));
      return { result: formatPercent(result), substituted: `WACC(${d.wd},${d.we},${d.rd},${d.re},${d.tax},${d.wp},${d.rp})`, numeric: result };
    }
  },
  {
    id: "dcf",
    title: "DCF (FCF + terminal growth)",
    formula: "V = Σ FCF_t/(1+r)^t + [FCF_(T+1)/(r-g)]/(1+r)^T",
    fields: [
      { key: "fcf", label: "FCF list (years 1..T)", type: "text", placeholder: "100, 110, 121, 133.1" },
      { key: "discount", label: "Discount rate", type: "number", defaultValue: 10 },
      { key: "growth", label: "Terminal growth", type: "number", defaultValue: 3 }
    ],
    solve: (d) => {
      const fcf = parseNumberList(d.fcf);
      const result = dcfValue(fcf, Number(d.discount), Number(d.growth));
      return { result: formatNumber(result), substituted: `DCF([${fcf.join(", ")}], ${d.discount}, ${d.growth})`, numeric: result };
    }
  },
  {
    id: "relative",
    title: "Relative Valuation",
    formula: "Implied value = multiple * target metric",
    fields: [
      { key: "multipleType", label: "Multiple type", type: "select", defaultValue: "P/E", options: [{ value: "P/E", label: "P/E" }, { value: "EV/EBITDA", label: "EV/EBITDA" }, { value: "P/S", label: "P/S" }] },
      { key: "multiple", label: "Comparable multiple", type: "number", defaultValue: 15 },
      { key: "metric", label: "Target metric", type: "number", defaultValue: 2.4 }
    ],
    solve: (d) => {
      const result = Number(d.multiple) * Number(d.metric);
      return { result: formatNumber(result), substituted: `${d.multipleType}: ${d.multiple} x ${d.metric}`, numeric: result };
    }
  },
  {
    id: "risk_one",
    title: "Risk/Return (One Asset)",
    formula: "E(R)=mean(R); Var sample=Σ(R-mean)^2/(n-1); SD=sqrt(Var)",
    fields: [{ key: "returns", label: "Returns list", type: "text", placeholder: "0.05, -0.02, 0.03, 0.04" }],
    solve: (d) => {
      const returns = parseNumberList(d.returns);
      const mean = expectedReturn(returns);
      const variance = sampleVariance(returns);
      const sd = sampleStdDev(returns);
      return {
        result: `E(R)=${formatPercent(mean)}, SD=${formatPercent(sd)}`,
        substituted: `n=${returns.length}, var=${formatNumber(variance)}`,
        numeric: sd
      };
    }
  },
  {
    id: "risk_two",
    title: "Two-Asset Portfolio",
    formula: "E(Rp)=w1E1+w2E2; SDp=sqrt(w1^2s1^2+w2^2s2^2+2w1w2s1s2ρ)",
    fields: [
      { key: "w1", label: "Weight 1", type: "number", defaultValue: 0.5 },
      { key: "w2", label: "Weight 2", type: "number", defaultValue: 0.5 },
      { key: "er1", label: "Expected return 1", type: "number", defaultValue: 12 },
      { key: "er2", label: "Expected return 2", type: "number", defaultValue: 8 },
      { key: "sd1", label: "Std dev 1", type: "number", defaultValue: 20 },
      { key: "sd2", label: "Std dev 2", type: "number", defaultValue: 10 },
      { key: "corr", label: "Correlation", type: "number", defaultValue: 0.2 }
    ],
    solve: (d) => {
      const w1 = Number(d.w1);
      const w2 = Number(d.w2);
      const er = portfolioExpectedReturn([w1, w2], [normalizeRate(Number(d.er1)), normalizeRate(Number(d.er2))]);
      const sd = portfolioStd2Asset(w1, Number(d.sd1), w2, Number(d.sd2), Number(d.corr));
      return {
        result: `E(Rp)=${formatPercent(er)}, SDp=${formatPercent(sd)}`,
        substituted: `w1=${w1}, w2=${w2}, corr=${d.corr}`,
        numeric: sd
      };
    }
  },
  {
    id: "risk_many",
    title: "N-Asset Portfolio",
    formula: "Var_p = Σ_i Σ_j w_i w_j σ_i σ_j ρ_ij",
    fields: [
      { key: "weights", label: "Weights (comma)", type: "text", placeholder: "0.4,0.35,0.25" },
      { key: "returns", label: "Expected returns (comma)", type: "text", placeholder: "0.10,0.08,0.12" },
      { key: "stddevs", label: "Std devs (comma)", type: "text", placeholder: "0.18,0.12,0.22" },
      { key: "corr", label: "Corr matrix rows (; separated)", type: "text", placeholder: "1,0.2,0.1;0.2,1,0.15;0.1,0.15,1" }
    ],
    solve: (d) => {
      const weights = parseNumberList(d.weights);
      const returns = parseNumberList(d.returns);
      const stddevs = parseNumberList(d.stddevs);
      const corrRows = d.corr.split(";").map((r) => parseNumberList(r));
      const er = portfolioExpectedReturn(weights, returns);
      const sd = portfolioStdNAsset(weights, stddevs, corrRows);
      return {
        result: `E(Rp)=${formatPercent(er)}, SDp=${formatPercent(sd)}`,
        substituted: `weights=[${weights.join(",")}], n=${weights.length}`,
        numeric: sd
      };
    }
  },
  {
    id: "frontier",
    title: "Efficient Frontier (2 assets)",
    formula: "Compute E(Rp), SDp over w1 in [0,1] to trace frontier.",
    fields: [
      { key: "er1", label: "Expected return 1", type: "number", defaultValue: 12 },
      { key: "er2", label: "Expected return 2", type: "number", defaultValue: 8 },
      { key: "sd1", label: "Std dev 1", type: "number", defaultValue: 20 },
      { key: "sd2", label: "Std dev 2", type: "number", defaultValue: 10 },
      { key: "corr", label: "Correlation", type: "number", defaultValue: 0.2 },
      { key: "step", label: "Weight step", type: "number", defaultValue: 0.1 }
    ],
    solve: (d) => {
      const points = frontierTwoAssetPoints(Number(d.er1), Number(d.er2), Number(d.sd1), Number(d.sd2), Number(d.corr), Number(d.step));
      const min = points.reduce((a, b) => (b.sd < a.sd ? b : a), points[0]);
      return {
        result: `Min SD portfolio: w1=${formatNumber(min.w1, 2)}, E(R)=${formatPercent(min.ret)}, SD=${formatPercent(min.sd)}`,
        substituted: `Generated ${points.length} points`,
        numeric: min.sd
      };
    }
  },
  {
    id: "capm",
    title: "CAPM",
    formula: "Re = Rf + β*(Rm-Rf) or Re = Rf + β*MRP",
    fields: [
      { key: "rf", label: "Risk-free rate", type: "number", defaultValue: 4 },
      { key: "beta", label: "Beta", type: "number", defaultValue: 1.2 },
      { key: "rmOrMrp", label: "Market return or MRP", type: "number", defaultValue: 9.5 },
      { key: "useMrp", label: "Input is MRP (not Rm)", type: "checkbox", defaultValue: false }
    ],
    solve: (d) => {
      const result = capm(Number(d.rf), Number(d.beta), Number(d.rmOrMrp), d.useMrp === "true");
      return { result: formatPercent(result), substituted: `Rf=${d.rf}, beta=${d.beta}, x=${d.rmOrMrp}`, numeric: result };
    }
  },
  {
    id: "bond",
    title: "Bond Price / YTM",
    formula: "Price = Σ C/(1+y)^t + FV/(1+y)^T; solve y for YTM.",
    fields: [
      { key: "mode", label: "Mode", type: "select", defaultValue: "price", options: [{ value: "price", label: "Price" }, { value: "ytm", label: "YTM" }] },
      { key: "par", label: "Par value", type: "number", defaultValue: 1000 },
      { key: "coupon", label: "Coupon rate", type: "number", defaultValue: 6 },
      { key: "years", label: "Years", type: "number", defaultValue: 5 },
      { key: "freq", label: "Payments/year", type: "number", defaultValue: 2 },
      { key: "ytm", label: "YTM (for price)", type: "number", defaultValue: 7 },
      { key: "price", label: "Price (for ytm)", type: "number", defaultValue: 960 }
    ],
    solve: (d) => {
      if (d.mode === "price") {
        const result = bondPrice(Number(d.par), Number(d.coupon), Number(d.years), Number(d.ytm), Number(d.freq));
        return { result: formatNumber(result, 2), substituted: `Bond price with y=${d.ytm}`, numeric: result };
      }
      const result = bondYtm(Number(d.price), Number(d.par), Number(d.coupon), Number(d.years), Number(d.freq));
      return { result: formatPercent(result), substituted: `Bond YTM from price=${d.price}`, numeric: result };
    }
  },
  {
    id: "stock",
    title: "Stock (Gordon Growth)",
    formula: "P0 = D1/(r-g)",
    fields: [
      { key: "d1", label: "Next dividend D1", type: "number", defaultValue: 2 },
      { key: "r", label: "Required return", type: "number", defaultValue: 10 },
      { key: "g", label: "Growth", type: "number", defaultValue: 4 }
    ],
    solve: (d) => {
      const result = gordonGrowth(Number(d.d1), Number(d.r), Number(d.g));
      return { result: formatNumber(result), substituted: `P0 = ${d.d1}/(${d.r}-${d.g})`, numeric: result };
    }
  },
  {
    id: "ratios",
    title: "Financial Ratios",
    formula: "Current, Quick, ROE, ROA, D/E, TIE, Margin, Asset Turnover, Inventory Turnover",
    fields: [
      { key: "ca", label: "Current assets", type: "number", defaultValue: 500 },
      { key: "inv", label: "Inventory", type: "number", defaultValue: 120 },
      { key: "cl", label: "Current liabilities", type: "number", defaultValue: 250 },
      { key: "ni", label: "Net income", type: "number", defaultValue: 80 },
      { key: "equity", label: "Equity", type: "number", defaultValue: 400 },
      { key: "assets", label: "Total assets", type: "number", defaultValue: 900 },
      { key: "debt", label: "Total debt", type: "number", defaultValue: 500 },
      { key: "ebit", label: "EBIT", type: "number", defaultValue: 140 },
      { key: "interest", label: "Interest expense", type: "number", defaultValue: 20 },
      { key: "revenue", label: "Revenue", type: "number", defaultValue: 1200 },
      { key: "cogs", label: "COGS", type: "number", defaultValue: 700 }
    ],
    solve: (d) => {
      const out = [
        `Current=${formatNumber(ratioCurrent(Number(d.ca), Number(d.cl)), 3)}`,
        `Quick=${formatNumber(ratioQuick(Number(d.ca), Number(d.inv), Number(d.cl)), 3)}`,
        `ROE=${formatPercent(ratioRoe(Number(d.ni), Number(d.equity)))}`,
        `ROA=${formatPercent(ratioRoa(Number(d.ni), Number(d.assets)))}`,
        `D/E=${formatNumber(ratioDebtToEquity(Number(d.debt), Number(d.equity)), 3)}`,
        `TIE=${formatNumber(ratioTie(Number(d.ebit), Number(d.interest)), 3)}`,
        `Margin=${formatPercent(ratioProfitMargin(Number(d.ni), Number(d.revenue)))}`,
        `AssetTurnover=${formatNumber(ratioAssetTurnover(Number(d.revenue), Number(d.assets)), 3)}`,
        `InvTurnover=${formatNumber(ratioInventoryTurnover(Number(d.cogs), Number(d.inv)), 3)}`
      ];
      return { result: out.join(" | "), substituted: "Ratios from provided statement inputs" };
    }
  }
];

function render(): void {
  const app = document.getElementById("app");
  if (!app) return;
  app.innerHTML = `
  <style>
    body { margin: 0; background: #f5f7fb; }
    .wrap { font-family: Segoe UI, system-ui, sans-serif; padding: 12px; color: #18253d; }
    .card { background: white; border-radius: 10px; box-shadow: 0 1px 4px rgba(0,0,0,.08); padding: 12px; margin-bottom: 10px; }
    .grid { display: grid; grid-template-columns: 1fr; gap: 8px; }
    label { font-weight: 600; font-size: 12px; }
    input, select, textarea, button { font: inherit; width: 100%; box-sizing: border-box; border: 1px solid #ccd5e1; border-radius: 6px; padding: 6px; }
    textarea { min-height: 56px; }
    button { cursor: pointer; background: #1559cf; color: white; border: none; font-weight: 600; }
    button.secondary { background: #edf2fa; color: #22334f; border: 1px solid #ccd5e1; }
    .btns { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }
    .mono { font-family: Consolas, ui-monospace, monospace; font-size: 12px; }
    .small { font-size: 12px; opacity: 0.9; }
  </style>
  <div class="wrap">
    <div class="card">
      <h2 style="margin:0 0 8px 0;">Managerial Finance Tools</h2>
      <div class="small">Mac + Windows Office Add-in | Formula + Result</div>
    </div>
    <div class="card grid">
      <label for="topicSelect">Topic</label>
      <select id="topicSelect">${topics.map((t) => `<option value="${t.id}">${t.title}</option>`).join("")}</select>
      <div id="fields" class="grid"></div>
      <div class="btns">
        <button id="solveBtn">Solve</button>
        <button id="insertBtn" class="secondary">Insert result in active cell</button>
      </div>
      <button id="useSelectionBtn" class="secondary">Use current selection as cash-flow list</button>
    </div>
    <div class="card">
      <label>Formula</label>
      <div id="formula" class="mono"></div>
      <label style="margin-top:8px; display:block;">Formula with numbers</label>
      <div id="subFormula" class="mono"></div>
      <label style="margin-top:8px; display:block;">Result</label>
      <div id="result" class="mono"></div>
      <div id="details" class="small"></div>
    </div>
  </div>
  `;
}

function renderFields(topic: Topic): void {
  const fields = document.getElementById("fields");
  const formula = document.getElementById("formula");
  if (!fields || !formula) return;
  formula.textContent = topic.formula;

  fields.innerHTML = topic.fields
    .map((f) => {
      if (f.type === "number") {
        return `<div><label>${f.label}</label><input type="number" data-key="${f.key}" value="${f.defaultValue ?? ""}" /></div>`;
      }
      if (f.type === "text") {
        return `<div><label>${f.label}</label><textarea data-key="${f.key}" placeholder="${f.placeholder ?? ""}"></textarea></div>`;
      }
      if (f.type === "checkbox") {
        return `<div><label><input type="checkbox" data-key="${f.key}" ${f.defaultValue ? "checked" : ""}/> ${f.label}</label></div>`;
      }
      return `<div><label>${f.label}</label><select data-key="${f.key}">${f.options
        .map((o) => `<option value="${o.value}" ${o.value === f.defaultValue ? "selected" : ""}>${o.label}</option>`)
        .join("")}</select></div>`;
    })
    .join("");
}

function readFormData(): Record<string, string> {
  const data: Record<string, string> = {};
  document.querySelectorAll<HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>("[data-key]").forEach((el) => {
    const key = el.getAttribute("data-key");
    if (!key) return;
    if (el instanceof HTMLInputElement && el.type === "checkbox") data[key] = String(el.checked);
    else data[key] = el.value;
  });
  return data;
}

async function useSelectionForCashflows(): Promise<void> {
  const cashflowEl = document.querySelector<HTMLTextAreaElement>('textarea[data-key="cashflows"]');
  if (!cashflowEl || typeof Excel === "undefined") return;
  await Excel.run(async (context: Excel.RequestContext) => {
    const range = context.workbook.getSelectedRange();
    range.load("values,address");
    await context.sync();
    const values = range.values.flat().map((v) => Number(v)).filter((v) => Number.isFinite(v));
    cashflowEl.value = values.join(", ");
  });
}

async function insertResultIntoCell(resultText: string): Promise<void> {
  if (typeof Excel === "undefined") return;
  await Excel.run(async (context: Excel.RequestContext) => {
    const range = context.workbook.getSelectedRange();
    range.values = [[resultText]];
    await context.sync();
  });
}

function bind(): void {
  const topicSelect = document.getElementById("topicSelect") as HTMLSelectElement;
  const solveBtn = document.getElementById("solveBtn") as HTMLButtonElement;
  const insertBtn = document.getElementById("insertBtn") as HTMLButtonElement;
  const useSelectionBtn = document.getElementById("useSelectionBtn") as HTMLButtonElement;
  const resultEl = document.getElementById("result") as HTMLDivElement;
  const subFormula = document.getElementById("subFormula") as HTMLDivElement;
  const detailsEl = document.getElementById("details") as HTMLDivElement;

  const getTopic = () => topics.find((t) => t.id === topicSelect.value) ?? topics[0];
  renderFields(getTopic());

  topicSelect.addEventListener("change", () => {
    renderFields(getTopic());
    resultEl.textContent = "";
    subFormula.textContent = "";
    detailsEl.textContent = "";
  });

  solveBtn.addEventListener("click", () => {
    try {
      const t = getTopic();
      const solved = t.solve(readFormData());
      resultEl.textContent = solved.result;
      subFormula.textContent = solved.substituted;
      detailsEl.textContent = solved.details?.join(" | ") ?? "";
    } catch (err) {
      resultEl.textContent = `Error: ${(err as Error).message}`;
    }
  });

  insertBtn.addEventListener("click", async () => {
    if (!resultEl.textContent) return;
    try {
      await insertResultIntoCell(resultEl.textContent);
    } catch (err) {
      resultEl.textContent = `Insert error: ${(err as Error).message}`;
    }
  });

  useSelectionBtn.addEventListener("click", async () => {
    try {
      await useSelectionForCashflows();
    } catch (err) {
      resultEl.textContent = `Selection error: ${(err as Error).message}`;
    }
  });
}

render();
bind();
