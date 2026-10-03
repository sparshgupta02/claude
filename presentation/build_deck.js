// Builds trexsim-alpha-deck.pptx: three TrexSim alphas, benchmarked against the 80-alpha pool.
// All numbers come from trexsim-submitted-alphas-families.md (Slow period, 2007-03-15 to 2021-12-30).
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");

const THEME = {
  name: "TrexSim Alpha Deck",
  headFontFace: "Calibri",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "0B1220", lt1: "FFFFFF", dk2: "556070", lt2: "F1F4F8",
    accent1: "1F4FD8", // momentum   (361878)
    accent2: "C2410C", // reversion  (320000)
    accent3: "0E7C6B", // composite  (336690) / better than benchmark
    accent4: "8A94A6", // benchmark grey
    accent5: "B42318", // worse than benchmark
    accent6: "D5DBE3", // hairline
    hlink: "1F4FD8", folHlink: "556070",
  },
};
const HEX = { ink: "0B1220", muted: "556070", panel: "F1F4F8", hair: "D5DBE3", mom: "1F4FD8", rev: "C2410C", comp: "0E7C6B", bench: "8A94A6", night: "0B1220", nightText: "C9D2E0" };
const MONO = "Consolas";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5 in
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "TrexSim: three alphas, benchmarked";
pres.author = "sparshgupta";
const C = pres.SchemeColor;
const W = 13.333, M = 0.6;

// ---------- layouts ----------
pres.defineSlideMaster({
  title: "DARK",
  background: { color: HEX.night },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: M, y: 2.2, w: 11.5, h: 1.6, fontSize: 44, bold: true, color: C.background1, align: "left", valign: "top", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: HEX.panel === "F1F4F8" ? "FFFFFF" : "FFFFFF" },
  margin: [0.5, 0.6, 0.6, 0.6],
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: M, y: 0.4, w: W - 2 * M, h: 0.8, fontSize: 30, bold: true, color: C.text1, align: "left", valign: "top", margin: 0 }, text: "" } },
    { text: { text: "TrexSim · US TOP1000 · delay 0 · industry-neutral · Slow period 2007–2021 · in-sample", options: { x: M, y: 7.05, w: 9, h: 0.3, fontSize: 10, color: C.text2, margin: 0 } } },
  ],
  slideNumber: { x: W - M - 0.5, y: 7.05, w: 0.5, h: 0.3, fontSize: 10, color: "556070", align: "right", margin: 0 },
});

// ---------- helpers ----------
let n = 0;
const nm = (s) => `${s}-${++n}`;
function txt(slide, text, o) { slide.addText(text, { isTextBox: true, margin: 0, fontSize: 15, color: C.text1, valign: "top", objectName: nm("text"), ...o }); }
function card(slide, x, y, w, h, fill = C.background2) { slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { color: fill }, rectRadius: 0.08, objectName: nm("card") }); }
function tag(slide, x, y, label, color, w = 1.9) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 0.34, fill: { color }, line: { color }, rectRadius: 0.04, objectName: nm("tag") });
  txt(slide, label, { x, y, w, h: 0.34, fontSize: 12, bold: true, color: C.background1, align: "center", valign: "middle" });
}
function formula(slide, code, x, y, w, h) {
  card(slide, x, y, w, h);
  txt(slide, code, { x: x + 0.2, y: y + 0.15, w: w - 0.4, h: h - 0.3, fontFace: MONO, fontSize: 12, color: C.text1, valign: "middle" });
}
// one step of a numbered mechanism list
function step(slide, i, x, y, w, head, body, color) {
  slide.addShape(pres.shapes.OVAL, { x, y, w: 0.42, h: 0.42, fill: { color }, line: { color }, objectName: nm("step") });
  txt(slide, String(i), { x, y, w: 0.42, h: 0.42, fontSize: 14, bold: true, color: C.background1, align: "center", valign: "middle" });
  txt(slide, [{ text: head, options: { bold: true, breakLine: true } }, { text: body, options: { color: C.text2 } }], { x: x + 0.6, y: y - 0.04, w: w - 0.6, h: 0.95, fontSize: 14 });
}
// benchmark table: rows = [metric, alpha, family, pool, rank, better(bool|null)]
function benchTable(slide, x, y, w, alphaLabel, familyLabel, rows, color) {
  const hdr = (t) => ({ text: t, options: { bold: true, color: "556070", fontSize: 11 } });
  const data = [[hdr("Metric"), { text: alphaLabel, options: { bold: true, color, fontSize: 11 } }, hdr(familyLabel), hdr("Pool median"), hdr("Rank / 80")]];
  for (const [m, a, f, p, r, better] of rows) {
    const mark = better === null ? "" : better ? "  ▲" : "  ▼";
    const mc = better === null ? "0B1220" : better ? "0E7C6B" : "B42318";
    data.push([
      { text: m, options: { color: "0B1220" } },
      [{ text: a, options: { bold: true, color: "0B1220" } }, { text: mark, options: { color: mc, bold: true } }],
      { text: f, options: { color: "556070" } },
      { text: p, options: { color: "556070" } },
      { text: r, options: { color: "0B1220" } },
    ]);
  }
  slide.addTable(data.map((row) => row.map((c) => (Array.isArray(c) ? { text: c } : c))), {
    x, y, w, colW: [w * 0.22, w * 0.25, w * 0.19, w * 0.17, w * 0.17], fontSize: 14, fontFace: "Calibri",
    border: { type: "solid", pt: 0.75, color: HEX.hair }, rowH: 0.42, valign: "middle", margin: [0, 0.08, 0, 0.08], objectName: nm("table"),
  });
}
const axis = { catAxisLabelColor: "556070", valAxisLabelColor: "556070", catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", titleFontFace: "+mn-lt", catAxisLabelFontSize: 12, valAxisLabelFontSize: 11, dataLabelFontSize: 12, dataLabelColor: "0B1220", valGridLine: { color: "E3E8EF", size: 0.75 }, catGridLine: { style: "none" }, showLegend: false, showValue: true, dataLabelPosition: "outEnd", valAxisLineShow: false, catAxisLineShow: false, titleColor: "0B1220", titleFontSize: 13 };

// =====================================================================
// 1. Title
pres.addSection({ title: "Introduction" });
{
  const s = pres.addSlide({ masterName: "DARK", sectionTitle: "Introduction" });
  txt(s, "TRExSIM ALPHA RESEARCH  ·  US TOP1000  ·  2007–2021", { x: M, y: 1.5, w: 11, h: 0.4, fontSize: 14, bold: true, color: HEX.nightText, charSpacing: 2 });
  s.addText("Three alphas, three mechanisms", { placeholder: "title" });
  txt(s, "Event-drift momentum, two-sided price reversion and an orthogonal two-leg composite, each benchmarked against the account's 80 submitted alphas.", { x: M, y: 3.35, w: 9.2, h: 1.0, fontSize: 18, color: HEX.nightText });
  tag(s, M, 5.2, "361878 · Momentum", HEX.mom, 2.4);
  tag(s, M + 2.7, 5.2, "320000 · Mean reversion", HEX.rev, 2.6);
  tag(s, M + 5.6, 5.2, "336690 · Composite", HEX.comp, 2.4);
  txt(s, "All statistics are in-sample (Slow period) and in the platform's own units.", { x: M, y: 6.7, w: 9, h: 0.35, fontSize: 11, color: HEX.nightText });
  s.addNotes("Three alphas, each built on a different economic mechanism: one rides information (earnings drift), one is paid for providing liquidity (reversion), one combines two independent views of the closing auction. Every number shown is benchmarked against the 80 alphas already submitted on the account.");
}

// 2. Setup and benchmarks
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Introduction" });
  s.addText("How every alpha is benchmarked", { placeholder: "title" });
  // left: setup
  card(s, M, 1.45, 5.2, 4.0);
  txt(s, "The test bed", { x: M + 0.3, y: 1.65, w: 4.6, h: 0.4, fontSize: 20, bold: true });
  const setup = [
    ["Universe", "US TOP1000, industry-neutral"],
    ["Period", "15 Mar 2007 – 30 Dec 2021 (Slow)"],
    ["Frequency", "Daily rebalance, delay 0 (trade the close)"],
    ["Inputs", "Daily bars + daily summaries of intraday data (last-hour VWAP, volume share, skew)"],
    ["Book", "Long-short, dollar-neutral: the right yardstick is peer alphas, not the S&P 500"],
  ];
  txt(s, setup.map(([k, v], i) => [{ text: k + "  ", options: { bold: true } }, { text: v, options: { color: C.text2, breakLine: i < setup.length - 1 } }]).flat(), { x: M + 0.3, y: 2.2, w: 4.6, h: 3.1, fontSize: 14, paraSpaceAfter: 8 });
  // right: three benchmark layers
  const L = [
    ["1", "Pool", "Median and rank among all 80 submitted alphas"],
    ["2", "Family peers", "Median of alphas built on the same mechanism"],
    ["3", "Construction", "For the composite: each leg alone, and what the maths predicts for the pair"],
  ];
  L.forEach(([i, h, b], k) => step(s, i, 6.3, 1.6 + k * 1.15, 6.4, h, b, HEX.ink));
  // bottom: pool stat tiles
  const tiles = [["0.102", "Median IR"], ["0.96", "Median turnover (TVR)"], ["0.118", "Median IR/√TVR"], ["10.8", "Median drawdown"]];
  tiles.forEach(([v, l], k) => {
    const x = M + k * 3.08;
    txt(s, v, { x, y: 5.75, w: 2.8, h: 0.6, fontSize: 36, bold: true, color: HEX.bench });
    txt(s, l, { x, y: 6.35, w: 2.8, h: 0.35, fontSize: 12, color: C.text2 });
  });
  txt(s, "IR = mean ÷ st. dev. of daily PnL. TVR = daily turnover. IR/√TVR = the platform score. Ret/DD = return ÷ max drawdown.", { x: 6.9, y: 5.0, w: 5.8, h: 0.5, fontSize: 11, color: C.text2 });
  s.addNotes("The book is dollar-neutral and industry-neutral, so a market index is not a fair comparison: market beta is close to zero by construction. Instead each alpha is compared on three levels: the full pool of 80 submitted alphas, its own family of similar alphas, and, for the composite, its own legs. The grey numbers at the bottom are the pool medians every later slide refers back to.");
}

// 3. Momentum
pres.addSection({ title: "The three alphas" });
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "The three alphas" });
  s.addText("Momentum: ride earnings reactions that volume confirmed", { placeholder: "title" });
  tag(s, M, 1.3, "361878 · Momentum", HEX.mom, 2.4);
  step(s, 1, M, 1.95, 5.8, "Spot the event", "Days-to-next-earnings resets by >10, or net income changes: a report just landed", HEX.mom);
  step(s, 2, M, 2.95, 5.8, "Score the reaction", "Return ÷ its 60-day volatility × volume ÷ its 60-day average", HEX.mom);
  step(s, 3, M, 3.95, 5.8, "Hold it for 20 days, same direction", "Under-reaction to earnings news: post-earnings drift (Bernard & Thomas, 1989)", HEX.mom);
  formula(s, "flow  = ts_sum(event × (ret1/ts_std(ret1,60))\n               × (volume/ts_mean(volume,60)), 20)\nalpha = cs_zscore(2 × cs_zscore(flow) + earnings_prox)", M, 5.05, 5.8, 1.2);
  txt(s, "earnings_prox = cs_remove_middle(1/days_to_next_earnings, 0.5): a smaller pre-announcement risk premium", { x: M, y: 6.35, w: 5.8, h: 0.5, fontSize: 11, color: C.text2 });
  benchTable(s, 6.85, 1.3, 5.85, "361878", "Family C med.", [
    ["IR", "0.074", "0.088", "0.102", "76", false],
    ["Turnover", "0.368", "0.984", "0.958", "14", true],
    ["IR/√TVR", "0.121", "0.099", "0.118", "34", true],
    ["Drawdown", "22.4", "11.7", "10.8", "76", false],
    ["Ret/DD", "0.0027", "0.0053", "0.0062", "79", false],
  ], HEX.mom);
  card(s, 6.85, 4.0, 5.85, 2.5);
  txt(s, [
    { text: "Verdict: weak alone, valuable in the book", options: { bold: true, breakLine: true } },
    { text: "Lowest IR in its family and a deep drawdown, but it turns over less than a third as much as its peers, so it still beats the pool median score. Its sign is continuation, the opposite of every reversal signal on the account, which is what makes it worth holding next to them.", options: { color: C.text2 } },
  ], { x: 7.1, y: 4.2, w: 5.4, h: 2.2, fontSize: 14, paraSpaceAfter: 6 });
  s.addNotes("Why momentum shows up here and not as plain price momentum: once returns are industry-neutralised, 12-month price momentum gave IR close to zero on this platform. Continuation survives when the move is tied to news (an earnings report) and confirmed by volume. Example: a stock with 2% daily vol jumps 6% on 3x normal volume after earnings, r = (6/2) x 3 = 9, so it is held long for about 20 days. Be upfront that IR rank is 76 of 80 and drawdown 22.4 is among the worst; the case for it is low turnover and diversification.");
}

// 4. Mean reversion
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "The three alphas" });
  s.addText("Mean reversion: get paid to absorb forced flow", { placeholder: "title" });
  tag(s, M, 1.3, "320000 · Mean reversion", HEX.rev, 2.6);
  step(s, 1, M, 1.95, 5.6, "Displaced at the open", "Open far from the 10-day VWAP: overnight order imbalance", HEX.rev);
  step(s, 2, M, 2.95, 5.6, "Still stretched at the close", "Close far from the day's VWAP: pressure, not news, if both legs agree", HEX.rev);
  step(s, 3, M, 3.95, 5.6, "Take the other side for 1–5 days", "The liquidity provider is paid as price normalises (Nagel, 2012)", HEX.rev);
  formula(s, "alpha = -cs_rank(open - ts_mean(vwap, 10))\n        × abs(cs_rank(close - vwap))", M, 5.05, 5.6, 1.0);
  txt(s, "Structure of Formulaic Alpha #5 (Kakushadze, 2016); the saved TrexSim expression is an adaptation.", { x: M, y: 6.15, w: 5.6, h: 0.5, fontSize: 11, color: C.text2 });
  // family A IR chart, 320000 highlighted
  const fam = [["320000", 0.12], ["322875", 0.096], ["322818", 0.094], ["316611", 0.093], ["316394", 0.091], ["317914", 0.085], ["319868", 0.081], ["334032", 0.078], ["320022", 0.076], ["323084", 0.075], ["316910", 0.071]].reverse();
  s.addChart(pres.charts.BAR, [{ name: "IR", labels: fam.map((f) => f[0]), values: fam.map((f) => f[1]) }], {
    x: 6.6, y: 1.25, w: 6.1, h: 3.85, barDir: "bar", ...axis, showTitle: true, title: "IR of the 11 single intraday signals (Family A)",
    chartColors: fam.map((f) => (f[0] === "320000" ? HEX.rev : HEX.bench)), valAxisMinVal: 0, valAxisMaxVal: 0.14, valAxisLabelFormatCode: "0.00", dataLabelFormatCode: "0.000", barGapWidthPct: 45, objectName: "famA-ir-chart",
  });
  const tiles = [["1 / 80", "Return 0.145 (#1 on the account)"], ["15 / 80", "IR 0.120 (pool med. 0.102)"], ["9 / 80", "Ret/DD 0.011 (med. 0.006)"]];
  tiles.forEach(([v, l], k) => {
    const x = 6.6 + k * 2.07;
    card(s, x, 5.3, 1.9, 1.35);
    txt(s, v, { x: x + 0.15, y: 5.4, w: 1.65, h: 0.55, fontSize: 26, bold: true, color: HEX.rev });
    txt(s, l, { x: x + 0.15, y: 5.95, w: 1.65, h: 0.6, fontSize: 11, color: C.text2 });
  });
  s.addNotes("Highest IR of the 11 single intraday signals (0.120 vs a family median of 0.085) and the highest raw return of all 80 alphas. Weak spot: turnover 0.99 is about the pool median, so its score IR/sqrt(TVR) = 0.121 is only at the pool median 0.118. Drawdown 13.0 equals the family median. Mechanism: large time-constrained traders push the price away from fair value at both ends of the day; whoever absorbs that flow earns the reversal.");
}

// 5. Composite: the legs
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "The three alphas" });
  s.addText("Composite: two weak, unrelated views of the close", { placeholder: "title" });
  tag(s, M, 1.3, "336690 · Composite", HEX.comp, 2.4);
  const legs = [
    ["Leg 1 · Who traded", "vol-conc  (alone: 323084)", "cs_zscore(volume_last_hour / volume)", "Benchmark-driven and institutional orders cluster at the close and are worked over several days, so their pressure persists.", "Persists  (+)", "IR 0.075 · rank 75 / 80"],
    ["Leg 2 · Where it closed", "day-range  (alone: 316611)", "-cs_zscore((close - low) / (high - low))", "A close pinned at the high means buyers were still pushing at the bell. Supplying that liquidity pays as it fades.", "Reverts  (−)", "IR 0.093 · rank 52 / 80"],
  ];
  legs.forEach(([h, sub, f, why, beh, stat], k) => {
    const x = M + k * 6.15;
    card(s, x, 1.9, 5.9, 3.55);
    txt(s, h, { x: x + 0.3, y: 2.05, w: 5.3, h: 0.4, fontSize: 20, bold: true });
    txt(s, sub, { x: x + 0.3, y: 2.45, w: 5.3, h: 0.3, fontSize: 12, color: C.text2 });
    txt(s, f, { x: x + 0.3, y: 2.85, w: 5.3, h: 0.35, fontFace: MONO, fontSize: 12 });
    txt(s, why, { x: x + 0.3, y: 3.3, w: 5.3, h: 1.2, fontSize: 14, color: C.text2 });
    txt(s, [{ text: beh, options: { bold: true, color: HEX.comp } }, { text: "     " + stat, options: { color: C.text1 } }], { x: x + 0.3, y: 4.85, w: 5.3, h: 0.4, fontSize: 14 });
  });
  formula(s, "alpha 336690 = cs_zscore( vol_conc + day_range_reversal )      equal weights, no fitting", M, 5.7, 12.05, 0.55);
  txt(s, "Why they are orthogonal: different input (volume share vs price location), different behaviour (persists vs reverts), different question (who vs where).", { x: M, y: 6.4, w: 12.05, h: 0.5, fontSize: 14, color: C.text1 });
  s.addNotes("Each leg alone is weak: 323084 ranks 75 of 80 on IR and 80 of 80 on score; 316611 ranks 52 of 80. One measures participation, the other price location, and one is a persistence effect while the other is a reversal. Nothing in one leg's input tells you the other's, which is why they should be close to uncorrelated. The next slide checks that with the numbers.");
}

// 6. Composite: the proof
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "The three alphas" });
  s.addText("The numbers say the legs are independent", { placeholder: "title" });
  const bars = [["Leg 1 alone (vol-conc)", 0.075, HEX.bench], ["Leg 2 alone (day-range)", 0.093, HEX.bench], ["Predicted if ρ = 1", 0.084, HEX.hair], ["Predicted if ρ = 0", 0.1188, HEX.hair], ["Actual 336690", 0.119, HEX.comp]].reverse();
  s.addChart(pres.charts.BAR, [{ name: "IR", labels: bars.map((b) => b[0]), values: bars.map((b) => b[1]) }], {
    x: M, y: 1.3, w: 7.0, h: 4.6, barDir: "bar", ...axis, showTitle: true, title: "IR: legs, predictions and the actual composite",
    chartColors: bars.map((b) => b[2]), valAxisMinVal: 0, valAxisMaxVal: 0.14, valAxisLabelFormatCode: "0.00", dataLabelFormatCode: "0.000", barGapWidthPct: 40, objectName: "quadrature-chart",
  });
  card(s, 7.95, 1.3, 4.75, 2.45);
  txt(s, "Equal-weight sum of two z-scored legs", { x: 8.2, y: 1.45, w: 4.3, h: 0.35, fontSize: 14, bold: true });
  txt(s, "IR = (IR₁ + IR₂) ÷ √(2 + 2ρ)", { x: 8.2, y: 1.9, w: 4.3, h: 0.45, fontFace: MONO, fontSize: 15 });
  txt(s, "Solve for ρ with IR = 0.119:", { x: 8.2, y: 2.45, w: 4.3, h: 0.35, fontSize: 14, color: C.text2 });
  txt(s, "ρ ≈ 0.00", { x: 8.2, y: 2.85, w: 4.3, h: 0.7, fontSize: 36, bold: true, color: HEX.comp });
  txt(s, [
    { text: "+59% IR over the weaker leg, +28% over the stronger", options: { bold: true, breakLine: true } },
    { text: "IR rank 16 / 80 and return rank 7 / 80, from two legs ranked 75 and 52.", options: { color: C.text2, breakLine: true } },
    { text: "The cost: turnover 1.37 (rank 74 / 80) leaves its score at 0.102, below the 0.118 pool median.", options: { color: C.text2 } },
  ], { x: 7.95, y: 4.0, w: 4.75, h: 2.3, fontSize: 14, paraSpaceAfter: 6 });
  txt(s, "Assumes both legs have equal volatility after z-scoring. Leg IRs are from their own submitted runs (323084, 316611).", { x: M, y: 6.25, w: 7.0, h: 0.5, fontSize: 11, color: C.text2 });
  s.addNotes("If the two legs were the same signal (rho = 1), adding them would just average their IRs: 0.084. If they are independent (rho = 0), the noise partly cancels and IR rises to 0.119. The measured IR is 0.119, so the implied correlation is essentially zero. That is the empirical case for calling them orthogonal. Weak spot: both legs are high-turnover, so the composite inherits a turnover of 1.37 and its platform score is below the pool median.");
}

// 7. Scorecard
pres.addSection({ title: "Benchmark and conclusion" });
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Benchmark and conclusion" });
  s.addText("Scorecard against the 80-alpha pool", { placeholder: "title" });
  const hdr = (t, color = "556070") => ({ text: t, options: { bold: true, color, fontSize: 12 } });
  const cell = (v, r, better) => ({ text: [{ text: v, options: { bold: true, color: "0B1220" } }, { text: better ? "  ▲ " : "  ▼ ", options: { bold: true, color: better ? "0E7C6B" : "B42318" } }, { text: r + " / 80", options: { color: "556070", fontSize: 12 } }] });
  const plain = (t) => ({ text: t, options: { color: "556070" } });
  const rows = [
    [hdr("Metric"), hdr("Pool median"), hdr("Pool top quartile"), hdr("361878 Momentum", HEX.mom), hdr("320000 Reversion", HEX.rev), hdr("336690 Composite", HEX.comp)],
    [{ text: "IR" }, plain("0.102"), plain("0.118"), cell("0.074", 76, false), cell("0.120", 15, true), cell("0.119", 16, true)],
    [{ text: "Return" }, plain("0.065"), plain("0.094"), cell("0.060", 47, false), cell("0.145", 1, true), cell("0.108", 7, true)],
    [{ text: "Turnover (lower is better)" }, plain("0.958"), plain("0.58 or less"), cell("0.368", 14, true), cell("0.987", 44, false), cell("1.369", 74, false)],
    [{ text: "IR/√TVR (platform score)" }, plain("0.118"), plain("0.134"), cell("0.121", 34, true), cell("0.121", 34, true), cell("0.102", 54, false)],
    [{ text: "Drawdown (lower is better)" }, plain("10.8"), plain("7.6 or less"), cell("22.4", 76, false), cell("13.0", 54, false), cell("20.5", 71, false)],
    [{ text: "Ret/DD" }, plain("0.0062"), plain("0.0087"), cell("0.0027", 79, false), cell("0.0112", 9, true), cell("0.0053", 51, false)],
  ];
  s.addTable(rows, { x: M, y: 1.35, w: 12.13, colW: [2.75, 1.45, 1.75, 2.06, 2.06, 2.06], fontSize: 14, fontFace: "Calibri", color: "0B1220", border: { type: "solid", pt: 0.75, color: HEX.hair }, rowH: 0.52, valign: "middle", margin: [0, 0.1, 0, 0.1], objectName: "scorecard" });
  const reads = [
    [HEX.mom, "Momentum", "Cheapest to trade, weakest alone. Earns its place by betting the opposite way to the reversal pool."],
    [HEX.rev, "Reversion", "Strongest standalone: #1 return on the account and top-decile Ret/DD. Turnover is its cost."],
    [HEX.comp, "Composite", "Proof of construction: two bottom-half legs make a top-20 IR. Turnover and drawdown are the price."],
  ];
  reads.forEach(([c, h, b], k) => {
    const x = M + k * 4.12;
    card(s, x, 5.2, 3.85, 1.6);
    txt(s, [{ text: h, options: { bold: true, color: c, breakLine: true } }, { text: b, options: { color: C.text2 } }], { x: x + 0.2, y: 5.3, w: 3.45, h: 1.45, fontSize: 13, paraSpaceAfter: 4 });
  });
  s.addNotes("Rank 1 is best for every metric; for turnover and drawdown lower is better. Top-quartile column is the 75th percentile of the pool (25th for turnover and drawdown). Ties share a rank: 361878 and 320000 have the same score, 0.121. No alpha wins on every axis, which is the point of holding all three.");
}

// 8. Close
{
  const s = pres.addSlide({ masterName: "DARK", sectionTitle: "Benchmark and conclusion" });
  txt(s, "What the three add up to", { x: M, y: 0.6, w: 12, h: 0.8, fontSize: 36, bold: true, color: C.background1 });
  const cols = [
    [HEX.mom, "361878 · Momentum", "Horizon ~20 days · bets continuation", "Fails if post-earnings drift gets arbitraged away, or on earnings seasons where reactions reverse (max DD 22.4)."],
    [HEX.rev, "320000 · Mean reversion", "Horizon 1–5 days · bets reversal", "Fails when a move is news rather than pressure, and in liquidity crises when providers are run over."],
    [HEX.comp, "336690 · Composite", "Horizon 1–5 days · persist + revert", "Fails if late volume and close location start moving together, so ρ rises and the gain over the legs shrinks."],
  ];
  cols.forEach(([c, h, sub, risk], k) => {
    const x = M + k * 4.12;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.7, w: 3.85, h: 2.9, fill: { color: "162033" }, line: { color: "162033" }, rectRadius: 0.08, objectName: nm("darkcard") });
    tag(s, x + 0.25, 1.95, h, c, 3.35);
    txt(s, sub, { x: x + 0.25, y: 2.55, w: 3.35, h: 0.6, fontSize: 14, bold: true, color: C.background1 });
    txt(s, risk, { x: x + 0.25, y: 3.2, w: 3.35, h: 1.6, fontSize: 14, color: HEX.nightText });
  });
  txt(s, [
    { text: "Next checks before trusting this out of sample", options: { bold: true, color: C.background1, breakLine: true } },
    { text: "Measure the pairwise correlation of the three PnL streams  ·  run them on out-of-sample dates  ·  check drawdowns by year", options: { color: HEX.nightText } },
  ], { x: M, y: 5.0, w: 12, h: 1.0, fontSize: 15, paraSpaceAfter: 6 });
  txt(s, "All figures are in-sample backtests on TrexSim; the hypotheses explain why the signals might work, they are not proven causes.", { x: M, y: 6.75, w: 12, h: 0.35, fontSize: 11, color: HEX.nightText });
  s.addNotes("Close on the portfolio view: three different horizons and two opposite signs. The risks named are hypotheses about failure modes, not measured events. The correlation between the three alphas' PnL streams has not been measured here; that and an out-of-sample run are the obvious next checks.");
}

(async () => {
  const out = process.argv[2] || "trexsim-alpha-deck.pptx";
  await pres.writeFile({ fileName: out });
  await applyTheme(out, THEME);
  console.log("wrote", out);
})();
