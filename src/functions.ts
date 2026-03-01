import {
  capm,
  eacFromNpv,
  irr,
  normalizeRate,
  npv,
  parseNumberList,
  portfolioStd2Asset,
  sampleStdDev,
  wacc
} from "./finance";

export function NPV_CASHFLOWS(rate: number, cashflows: number[][]): number {
  const flat = cashflows.flat().map(Number).filter(Number.isFinite);
  return npv(rate, flat);
}

export function IRR_CASHFLOWS(cashflows: number[][], guess = 0.1): number {
  const flat = cashflows.flat().map(Number).filter(Number.isFinite);
  return irr(flat, guess);
}

export function EAC_FROM_NPV(npvValue: number, rate: number, nPeriods: number): number {
  return eacFromNpv(npvValue, rate, nPeriods);
}

export function CAPM_COST_OF_EQUITY(rf: number, beta: number, rm: number): number {
  return capm(rf, beta, rm, false);
}

export function WACC(E: number, D: number, Re: number, Rd: number, taxRate: number): number {
  const total = E + D;
  if (total <= 0) return NaN;
  const wd = D / total;
  const we = E / total;
  return wacc(wd, we, Rd, Re, taxRate);
}

export function PORTFOLIO_STDDEV_2ASSET(w1: number, sd1: number, w2: number, sd2: number, corr: number): number {
  return portfolioStd2Asset(w1, sd1, w2, sd2, corr);
}

export function PORT_RETURNS_STATS(returnsRange: number[][]): (string | number)[][] {
  const values = returnsRange.flat().map(Number).filter((x) => Number.isFinite(x));
  if (!values.length) return [["mean", "stdev", "n"], ["", "", 0]];
  const mean = values.reduce((a, b) => a + b, 0) / values.length;
  const sd = sampleStdDev(values);
  return [["mean", "stdev", "n"], [mean, sd, values.length]];
}

export function NORMALIZE_RATE(value: number): number {
  return normalizeRate(value);
}

export function PARSE_CASHFLOWS(cashflowsText: string): number[] {
  return parseNumberList(cashflowsText);
}
