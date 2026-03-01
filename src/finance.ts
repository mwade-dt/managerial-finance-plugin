export type SolveResult = {
  formula: string;
  substitutedFormula?: string;
  resultText: string;
  resultValue?: number;
  details?: string[];
};

const EPS = 1e-10;

export function normalizeRate(input: number): number {
  if (!Number.isFinite(input)) return NaN;
  if (Math.abs(input) <= 1) return input;
  return input / 100;
}

export function formatPercent(value: number, decimals = 2): string {
  return `${(value * 100).toFixed(decimals)}%`;
}

export function formatNumber(value: number, decimals = 4): string {
  if (!Number.isFinite(value)) return "NaN";
  return value.toFixed(decimals);
}

export function parseNumberList(input: string): number[] {
  return input
    .split(/[\n,;\t ]+/)
    .map((v) => v.trim())
    .filter(Boolean)
    .map((v) => Number(v))
    .filter((v) => Number.isFinite(v));
}

export function npv(rateInput: number, cashflows: number[]): number {
  const r = normalizeRate(rateInput);
  return cashflows.reduce((sum, cf, t) => sum + cf / Math.pow(1 + r, t), 0);
}

function npvAtRate(rate: number, cashflows: number[]): number {
  return cashflows.reduce((sum, cf, t) => sum + cf / Math.pow(1 + rate, t), 0);
}

export function irr(cashflows: number[], guess = 0.1): number {
  if (cashflows.length < 2) return NaN;
  const hasPositive = cashflows.some((v) => v > 0);
  const hasNegative = cashflows.some((v) => v < 0);
  if (!hasPositive || !hasNegative) return NaN;

  let r = normalizeRate(guess);
  for (let i = 0; i < 80; i++) {
    let f = 0;
    let df = 0;
    for (let t = 0; t < cashflows.length; t++) {
      const denom = Math.pow(1 + r, t);
      f += cashflows[t] / denom;
      if (t > 0) df -= (t * cashflows[t]) / (denom * (1 + r));
    }
    if (Math.abs(df) < EPS) break;
    const step = f / df;
    r -= step;
    if (!Number.isFinite(r)) break;
    if (Math.abs(step) < EPS) return r;
  }

  let low = -0.99;
  let high = 10;
  let fLow = npvAtRate(low, cashflows);
  let fHigh = npvAtRate(high, cashflows);
  if (!Number.isFinite(fLow) || !Number.isFinite(fHigh) || fLow * fHigh > 0) return NaN;

  for (let i = 0; i < 200; i++) {
    const mid = (low + high) / 2;
    const fMid = npvAtRate(mid, cashflows);
    if (Math.abs(fMid) < EPS) return mid;
    if (fLow * fMid < 0) {
      high = mid;
      fHigh = fMid;
    } else {
      low = mid;
      fLow = fMid;
    }
  }
  return (low + high) / 2;
}

export function mirr(cashflows: number[], financeRateInput: number, reinvestRateInput: number): number {
  const n = cashflows.length - 1;
  if (n <= 0) return NaN;
  const fr = normalizeRate(financeRateInput);
  const rr = normalizeRate(reinvestRateInput);

  let pvNeg = 0;
  let fvPos = 0;
  for (let t = 0; t < cashflows.length; t++) {
    const cf = cashflows[t];
    if (cf < 0) pvNeg += cf / Math.pow(1 + fr, t);
    if (cf > 0) fvPos += cf * Math.pow(1 + rr, n - t);
  }
  if (pvNeg >= 0 || fvPos <= 0) return NaN;
  return Math.pow(-fvPos / pvNeg, 1 / n) - 1;
}

export function paybackPeriod(cashflows: number[]): number {
  let cumulative = 0;
  for (let t = 0; t < cashflows.length; t++) {
    const prev = cumulative;
    cumulative += cashflows[t];
    if (cumulative >= 0) {
      if (t === 0) return 0;
      const needed = -prev;
      if (Math.abs(cashflows[t]) < EPS) return t;
      return (t - 1) + needed / cashflows[t];
    }
  }
  return NaN;
}

export function discountedPayback(cashflows: number[], rateInput: number): number {
  const r = normalizeRate(rateInput);
  const discounted = cashflows.map((cf, t) => cf / Math.pow(1 + r, t));
  return paybackPeriod(discounted);
}

export function eacFromNpv(npvValue: number, rateInput: number, periods: number): number {
  const r = normalizeRate(rateInput);
  if (periods <= 0) return NaN;
  if (Math.abs(r) < EPS) return npvValue / periods;
  const annuityFactor = (1 - Math.pow(1 + r, -periods)) / r;
  return npvValue / annuityFactor;
}

export function tvmLumpSum(pv: number, fv: number, rateInput: number, n: number, solveFor: string): number {
  const r = normalizeRate(rateInput);
  switch (solveFor) {
    case "PV":
      return fv / Math.pow(1 + r, n);
    case "FV":
      return pv * Math.pow(1 + r, n);
    case "N":
      if (pv <= 0 || fv <= 0 || 1 + r <= 0) return NaN;
      return Math.log(fv / pv) / Math.log(1 + r);
    case "R":
      if (pv <= 0 || fv <= 0 || n <= 0) return NaN;
      return Math.pow(fv / pv, 1 / n) - 1;
    default:
      return NaN;
  }
}

export function tvmAnnuity(
  pmt: number,
  rateInput: number,
  n: number,
  annuityDue: boolean,
  solveFor: string
): number {
  const r = normalizeRate(rateInput);
  const dueFactor = annuityDue ? 1 + r : 1;

  switch (solveFor) {
    case "PV":
      if (Math.abs(r) < EPS) return pmt * n * dueFactor;
      return (pmt * (1 - Math.pow(1 + r, -n)) * dueFactor) / r;
    case "FV":
      if (Math.abs(r) < EPS) return pmt * n * dueFactor;
      return (pmt * (Math.pow(1 + r, n) - 1) * dueFactor) / r;
    case "PMT":
      if (Math.abs(r) < EPS) return NaN;
      return NaN;
    default:
      return NaN;
  }
}

export function capm(rfInput: number, beta: number, rmOrMrpInput: number, useMrp: boolean): number {
  const rf = normalizeRate(rfInput);
  const mrpOrRm = normalizeRate(rmOrMrpInput);
  const mrp = useMrp ? mrpOrRm : mrpOrRm - rf;
  return rf + beta * mrp;
}

export function wacc(
  debtWeightInput: number,
  equityWeightInput: number,
  debtCostInput: number,
  equityCostInput: number,
  taxRateInput: number,
  prefWeightInput = 0,
  prefCostInput = 0
): number {
  const wdRaw = normalizeRate(debtWeightInput);
  const weRaw = normalizeRate(equityWeightInput);
  const wpRaw = normalizeRate(prefWeightInput);
  const total = wdRaw + weRaw + wpRaw;
  if (total <= 0) return NaN;
  const wd = wdRaw / total;
  const we = weRaw / total;
  const wp = wpRaw / total;
  const rd = normalizeRate(debtCostInput);
  const re = normalizeRate(equityCostInput);
  const rp = normalizeRate(prefCostInput);
  const t = normalizeRate(taxRateInput);
  return wd * rd * (1 - t) + we * re + wp * rp;
}

export function bondPrice(par: number, couponRateInput: number, years: number, ytmInput: number, freq: number): number {
  const c = normalizeRate(couponRateInput) * par / freq;
  const y = normalizeRate(ytmInput) / freq;
  const periods = Math.round(years * freq);
  let pvCoupons = 0;
  for (let t = 1; t <= periods; t++) {
    pvCoupons += c / Math.pow(1 + y, t);
  }
  return pvCoupons + par / Math.pow(1 + y, periods);
}

export function bondYtm(
  price: number,
  par: number,
  couponRateInput: number,
  years: number,
  freq: number
): number {
  const c = normalizeRate(couponRateInput) * par / freq;
  const periods = Math.round(years * freq);
  let low = 0;
  let high = 2;
  for (let i = 0; i < 200; i++) {
    const mid = (low + high) / 2;
    let pv = 0;
    for (let t = 1; t <= periods; t++) {
      pv += c / Math.pow(1 + mid / freq, t);
    }
    pv += par / Math.pow(1 + mid / freq, periods);
    if (Math.abs(pv - price) < 1e-7) return mid;
    if (pv > price) low = mid;
    else high = mid;
  }
  return (low + high) / 2;
}

export function gordonGrowth(dividend1: number, requiredReturnInput: number, growthInput: number): number {
  const r = normalizeRate(requiredReturnInput);
  const g = normalizeRate(growthInput);
  if (r <= g) return NaN;
  return dividend1 / (r - g);
}

export function dcfValue(fcfList: number[], discountRateInput: number, terminalGrowthInput: number): number {
  const r = normalizeRate(discountRateInput);
  const g = normalizeRate(terminalGrowthInput);
  if (fcfList.length < 2 || r <= g) return NaN;
  const forecast = fcfList.slice(0, -1);
  const finalFcf = fcfList[fcfList.length - 1];
  let pv = 0;
  for (let t = 0; t < forecast.length; t++) {
    pv += forecast[t] / Math.pow(1 + r, t + 1);
  }
  const terminal = (finalFcf * (1 + g)) / (r - g);
  pv += terminal / Math.pow(1 + r, forecast.length);
  return pv;
}

export function expectedReturn(returns: number[]): number {
  if (!returns.length) return NaN;
  return returns.reduce((a, b) => a + b, 0) / returns.length;
}

export function sampleVariance(returns: number[]): number {
  if (returns.length < 2) return NaN;
  const mean = expectedReturn(returns);
  return returns.reduce((a, b) => a + Math.pow(b - mean, 2), 0) / (returns.length - 1);
}

export function sampleStdDev(returns: number[]): number {
  const v = sampleVariance(returns);
  return Math.sqrt(v);
}

export function portfolioExpectedReturn(weights: number[], expectedReturns: number[]): number {
  if (weights.length !== expectedReturns.length || !weights.length) return NaN;
  return weights.reduce((sum, w, i) => sum + w * expectedReturns[i], 0);
}

export function portfolioStd2Asset(
  w1: number,
  sd1Input: number,
  w2: number,
  sd2Input: number,
  corr: number
): number {
  const sd1 = normalizeRate(sd1Input);
  const sd2 = normalizeRate(sd2Input);
  const variance = Math.pow(w1 * sd1, 2) + Math.pow(w2 * sd2, 2) + 2 * w1 * w2 * sd1 * sd2 * corr;
  return Math.sqrt(Math.max(variance, 0));
}

export function portfolioStdNAsset(
  weights: number[],
  stdDevsInput: number[],
  corrMatrix: number[][]
): number {
  const n = weights.length;
  if (n === 0 || stdDevsInput.length !== n || corrMatrix.length !== n) return NaN;
  const stdDevs = stdDevsInput.map(normalizeRate);
  let variance = 0;
  for (let i = 0; i < n; i++) {
    if (!corrMatrix[i] || corrMatrix[i].length !== n) return NaN;
    for (let j = 0; j < n; j++) {
      variance += weights[i] * weights[j] * stdDevs[i] * stdDevs[j] * corrMatrix[i][j];
    }
  }
  return Math.sqrt(Math.max(variance, 0));
}

export function frontierTwoAssetPoints(
  er1Input: number,
  er2Input: number,
  sd1Input: number,
  sd2Input: number,
  corr: number,
  step = 0.05
): { w1: number; ret: number; sd: number }[] {
  const er1 = normalizeRate(er1Input);
  const er2 = normalizeRate(er2Input);
  const points: { w1: number; ret: number; sd: number }[] = [];
  for (let w1 = 0; w1 <= 1 + EPS; w1 += step) {
    const w2 = 1 - w1;
    points.push({
      w1,
      ret: w1 * er1 + w2 * er2,
      sd: portfolioStd2Asset(w1, sd1Input, w2, sd2Input, corr)
    });
  }
  return points;
}

export function ratioCurrent(ca: number, cl: number): number {
  return ca / cl;
}

export function ratioQuick(ca: number, inventory: number, cl: number): number {
  return (ca - inventory) / cl;
}

export function ratioRoe(ni: number, equity: number): number {
  return ni / equity;
}

export function ratioRoa(ni: number, assets: number): number {
  return ni / assets;
}

export function ratioDebtToEquity(debt: number, equity: number): number {
  return debt / equity;
}

export function ratioTie(ebit: number, interest: number): number {
  return ebit / interest;
}

export function ratioProfitMargin(ni: number, revenue: number): number {
  return ni / revenue;
}

export function ratioAssetTurnover(revenue: number, assets: number): number {
  return revenue / assets;
}

export function ratioInventoryTurnover(cogs: number, inventory: number): number {
  return cogs / inventory;
}
