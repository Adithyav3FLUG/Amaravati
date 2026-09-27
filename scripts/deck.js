const pptxgen = require('pptxgenjs');
const D = require('./deck.json');
const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE'; // 13.333 x 7.5
pres.title = 'Golf Homes Phase-II – Project Execution Plan';
pres.author = 'Project Planning';

const C = { dark: '16302B', green: '2E6B4F', gold: 'C9973A', slate: '5B7B8C', sage: '9DB8A6', ink: '1E2A26',
  muted: '5E6B66', card: 'F1F4F2', line: 'D5DDD8', white: 'FFFFFF', red: 'C0504D', lightInk: 'C9D6CF' };
const HF = 'Cambria', BF = 'Calibri';
const W = 13.333, H = 7.5, MX = 0.6;
let n = 0;
const fmt = (v, d = 1) => Number(v).toLocaleString('en-IN', { minimumFractionDigits: d, maximumFractionDigits: d });
const fmtDate = s => { const d = new Date(s); return d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }); };

function base(dark = false) {
  const s = pres.addSlide(); n += 1;
  s.background = { color: dark ? C.dark : C.white };
  s.addText('Golf Homes Phase-II · Project Execution Plan R1', { x: MX, y: H - 0.42, w: 6, h: 0.3, fontFace: BF, fontSize: 9,
    color: dark ? C.lightInk : C.muted, margin: 0, isTextBox: true });
  s.addText(String(n).padStart(2, '0'), { x: W - MX - 1, y: H - 0.42, w: 1, h: 0.3, fontFace: BF, fontSize: 9, align: 'right',
    color: dark ? C.gold : C.muted, margin: 0, isTextBox: true });
  return s;
}
function title(s, t, sub, dark = false) {
  s.addText(t, { x: MX, y: 0.38, w: W - 2 * MX, h: 0.7, fontFace: HF, fontSize: 32, bold: true, color: dark ? C.white : C.ink, margin: 0, isTextBox: true });
  if (sub) s.addText(sub, { x: MX, y: 1.08, w: W - 2 * MX, h: 0.4, fontFace: BF, fontSize: 14, color: dark ? C.lightInk : C.muted, margin: 0, isTextBox: true });
}
function card(s, x, y, w, h, head, body, opt = {}) {
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: opt.fill || C.card }, line: { color: opt.fill || C.card } });
  if (head) s.addText(head, { x: x + 0.2, y: y + 0.15, w: w - 0.4, h: 0.4, fontFace: BF, fontSize: opt.hs || 15, bold: true,
    color: opt.hc || C.ink, margin: 0, isTextBox: true, valign: 'top' });
  if (body) s.addText(body, { x: x + 0.2, y: y + (head ? 0.6 : 0.2), w: w - 0.4, h: h - (head ? 0.75 : 0.35), fontFace: BF,
    fontSize: opt.bs || 12, color: opt.bc || C.muted, margin: 0, isTextBox: true, valign: 'top', paraSpaceAfter: 4 });
}
function stat(s, x, y, w, h, big, label, opt = {}) {
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: opt.fill || C.card }, line: { color: opt.fill || C.card } });
  s.addText(big, { x: x + 0.2, y: y + 0.12, w: w - 0.4, h: h * 0.55, fontFace: HF, fontSize: opt.fs || 30, bold: true,
    color: opt.color || C.green, margin: 0, isTextBox: true, valign: 'middle', fit: 'shrink' });
  s.addText(label, { x: x + 0.2, y: y + h * 0.62, w: w - 0.4, h: h * 0.32, fontFace: BF, fontSize: opt.ls || 11,
    color: opt.lc || C.muted, margin: 0, isTextBox: true, valign: 'top' });
}
function circleNum(s, x, y, d, txt, color = C.gold) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color }, line: { color } });
  s.addText(txt, { x, y, w: d, h: d, align: 'center', valign: 'middle', fontFace: BF, fontSize: d > 0.6 ? 18 : 12, bold: true, color: C.white, margin: 0, isTextBox: true });
}
function table(s, rows, x, y, w, colW, opt = {}) {
  const hdr = rows[0].map(t => ({ text: t, options: { bold: true, color: C.white, fill: { color: C.dark }, fontFace: BF, fontSize: opt.fs || 11 } }));
  const body = rows.slice(1).map((r, i) => r.map(t => ({ text: String(t), options: { color: C.ink, fontFace: BF, fontSize: opt.fs || 11,
    bold: opt.boldLast && i === rows.length - 2, fill: { color: (opt.boldLast && i === rows.length - 2) ? 'E3EAE6' : C.white } } })));
  s.addTable([hdr, ...body], { x, y, w, colW, border: { type: 'solid', pt: 0.5, color: C.line }, rowH: opt.rowH || 0.36, valign: 'middle', margin: 0.06, autoPage: false });
}
const notes = (s, t) => s.addNotes(t);

// ---------------------------------------------------------------- 1 Title
{
  const s = base(true);
  s.addText('GOLF HOMES', { x: MX + 0.1, y: 1.6, w: 9, h: 0.5, fontFace: BF, fontSize: 16, bold: true, color: C.gold, charSpacing: 6, margin: 0, isTextBox: true });
  s.addText('Phase-II Project Execution Plan', { x: MX + 0.1, y: 2.1, w: 11, h: 1.1, fontFace: HF, fontSize: 48, bold: true, color: C.white, margin: 0, isTextBox: true });
  s.addText('11 residential towers · 2B + Stilt + G/P + 35 floors · 1,507 DUs · GH-02, Sector 4, Greater Noida', { x: MX + 0.1, y: 3.3, w: 12, h: 0.45,
    fontFace: BF, fontSize: 18, color: C.gold, margin: 0, isTextBox: true });
  s.addText('Schedule · Resources · Cost · Baseline · Cash Flow · Earned Value — built on the L3 GRAP-aware schedule and BOQ of 17-Nov-2024',
    { x: MX + 0.1, y: 3.8, w: 12, h: 0.4, fontFace: BF, fontSize: 14, color: C.lightInk, margin: 0, isTextBox: true });
  s.addText('Developer: NBCC (India) Ltd · Architect: Chapman Taylor India LLP', { x: MX + 0.1, y: 5.3, w: 9, h: 0.3, fontFace: BF, fontSize: 12, color: C.lightInk, margin: 0, isTextBox: true });
  s.addText('Student Name · Roll No. · Guide: ____________', { x: MX + 0.1, y: 5.7, w: 9, h: 0.3, fontFace: BF, fontSize: 12, color: C.lightInk, margin: 0, isTextBox: true });
  notes(s, 'Golf Homes Phase-II is a 12-tower-unit (11 sanctioned towers) MIVAN residential development in Greater Noida. This deck follows the 10 mandatory deliverables of the P4 brief.');
}
// ---------------------------------------------------------------- 2 Map
{
  const s = base(); title(s, 'Presentation Map', 'Ten mandatory deliverables → one integrated execution plan');
  const items = [['Project Brief', 'Site, area statement, assumptions'], ['WBS', '7 phases · 78 packages · 2,672 L3 activities'],
    ['Schedule', 'L3 logic-linked, GRAP calendars, MS Project'], ['Resources', `${D.nres.Labour + D.nres.Staff} labour & staff, ${D.nres.Equipment} plant, ${D.nres.Material} materials`],
    ['Cost Plan', 'BOQ-driven activity costs → BAC'], ['Baseline', 'Time-cost Baseline 0 in MSP'], ['Critical Path', 'Zero-float & driving chain'],
    ['Cash Flow', 'Monthly PV & S-curve'], ['EVM', 'PV, EV, AC, CPI, SPI, EAC, ETC, VAC'], ['Presentation', 'Findings, caveats & defence']];
  const cw = 2.3, ch = 2.1, gx = 0.13, gy = 0.3, x0 = MX, y0 = 1.8;
  items.forEach((it, i) => {
    const x = x0 + (i % 5) * (cw + gx), y = y0 + Math.floor(i / 5) * (ch + gy);
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: cw, h: ch, fill: { color: C.card }, line: { color: C.card } });
    circleNum(s, x + 0.2, y + 0.22, 0.5, String(i + 1));
    s.addText(it[0], { x: x + 0.8, y: y + 0.22, w: cw - 0.9, h: 0.5, fontFace: BF, fontSize: 15, bold: true, color: C.ink, valign: 'middle', margin: 0, isTextBox: true });
    s.addText(it[1], { x: x + 0.2, y: y + 0.9, w: cw - 0.35, h: 1.0, fontFace: BF, fontSize: 12, color: C.muted, valign: 'top', margin: 0, isTextBox: true });
  });
}
// ---------------------------------------------------------------- 3 Brief
{
  const s = base(); title(s, 'Project Brief — Site & Scope', 'Real NBCC plot, real BOQ and approved area statement (drawing CT/GH/DOC/SUB/DOC1 R0)');
  const tw = 2.9, g = 0.17;
  stat(s, MX, 1.7, tw, 1.35, '24.76 ha', 'Total plot 2,47,622 sqm · 52,669 sqm under Phase-II development');
  stat(s, MX + (tw + g), 1.7, tw, 1.35, '11 towers', '2B + S + G/P + 35 floors · 1,507 dwelling units');
  stat(s, MX + 2 * (tw + g), 1.7, tw, 1.35, '5,58,450 sqm', 'Built-up area (60.1 lakh sft) incl. basements, stilt, podium');
  stat(s, MX + 3 * (tw + g), 1.7, tw, 1.35, '28.4 months', '01 Jan 2024 – 14 May 2026 · 865 calendar days');
  s.addText('Location & approvals', { x: MX, y: 3.35, w: 5.9, h: 0.4, fontFace: BF, fontSize: 16, bold: true, color: C.ink, margin: 0, isTextBox: true });
  s.addText([
    { text: 'GH-02, Sector 4, Greater Noida (UP 201301) — inside the Delhi-NCR airshed, so CAQM GRAP winter restrictions apply', options: { bullet: true, breakLine: true } },
    { text: 'Developer NBCC (India) Ltd; architect Chapman Taylor India LLP; client/PMC-driven EPC delivery', options: { bullet: true, breakLine: true } },
    { text: 'Permissible FAR 3.5 on remaining plot; utilised 3.445 after Phase-II (8,45,439 sqm)', options: { bullet: true, breakLine: true } },
    { text: 'Ground coverage 19.3 % (47,422 sqm) vs 35 % permitted; 88,000 sqm landscape, 1,785 trees', options: { bullet: true } }],
    { x: MX, y: 3.8, w: 5.9, h: 2.6, fontFace: BF, fontSize: 13, color: C.muted, margin: 0, paraSpaceAfter: 6, valign: 'top', isTextBox: true });
  s.addText('Building configuration', { x: 6.9, y: 3.35, w: 5.8, h: 0.4, fontFace: BF, fontSize: 16, bold: true, color: C.ink, margin: 0, isTextBox: true });
  s.addText([
    { text: 'MIVAN aluminium-formwork monolithic RCC shear-wall towers; 843 sqm FAR per typical floor, 4 flats per floor', options: { bullet: true, breakLine: true } },
    { text: 'Two common basements (47,555 sqm each) + stilt 36,430 sqm → 2,904 + 1,167 car spaces', options: { bullet: true, breakLine: true } },
    { text: 'Amenities: community centre, shopping complex (2B+G+2), nursery & senior secondary school, dispensary, milk booths', options: { bullet: true, breakLine: true } },
    { text: '8 lifts per tower, STP / WTP, substation & DG, 14,123 sqm internal roads, 26,123 sqm green area', options: { bullet: true } }],
    { x: 6.9, y: 3.8, w: 5.8, h: 2.6, fontFace: BF, fontSize: 13, color: C.muted, margin: 0, paraSpaceAfter: 6, valign: 'top', isTextBox: true });
  notes(s, 'All figures come from the master plan area details drawing and the BOQ built-up area table (rows 668–693).');
}
// ---------------------------------------------------------------- 4 Area statement
{
  const s = base(); title(s, 'Area Statement', 'BOQ built-up area table (rows 668–690) — the basis for quantities, resources and cost');
  table(s, [['Component', 'Area (sqm)', 'Area (sft)', 'Use'],
    ['Basement-2 (common)', '47,554.74', '5,11,873', 'Parking, services'],
    ['Basement-1 (common)', '47,554.74', '5,11,873', 'Parking, services'],
    ['Stilt level', '36,430.00', '3,92,129', 'Stilt parking, lobbies'],
    ['Ground / podium level', '11,254.43', '1,21,142', 'Entrance lobbies, podium'],
    ['Residential towers (11 nos.)', '3,96,168.96', '42,64,332', '1,507 DUs, 35 floors each'],
    ['Community centre (G+2 + terrace)', '3,666.06', '39,461', 'Clubhouse, amenities'],
    ['Shopping complex (2B+G+2)', '15,820.62', '1,70,292', 'Commercial, 293 ECS'],
    ['Total built-up area', '5,58,449.55', '60,11,095', '']], MX, 1.65, 7.6, [3.0, 1.45, 1.4, 1.75], { boldLast: true, fs: 11.5, rowH: 0.42 });
  s.addChart(pres.charts.BAR, [{ name: 'Area (\'000 sqm)', labels: ['Towers', 'Basements', 'Stilt', 'Shopping', 'Podium', 'Community'],
    values: [396.2, 95.1, 36.4, 15.8, 11.3, 3.7] }], { x: 8.5, y: 1.55, w: 4.3, h: 3.4, barDir: 'bar', chartColors: [C.green],
    showValue: true, dataLabelPosition: 'outEnd', dataLabelFontSize: 10, dataLabelFormatCode: '0.0', catAxisLabelFontSize: 10, valAxisHidden: true,
    valGridLine: { style: 'none' }, catGridLine: { style: 'none' }, showTitle: true, title: 'Built-up area, \'000 sqm', titleFontSize: 11, titleColor: C.muted, catAxisOrientation: 'maxMin' });
  card(s, 8.5, 5.15, 4.3, 1.55, 'FAR check', 'Residential FAR proposed 3,23,320 sqm + 15 % service 35,243 sqm; total utilised FAR 3.445 ≤ 3.5 permissible.', { bs: 11.5 });
  notes(s, 'The ₹/sqm and quantities are traceable to this table. Tower area includes 35 floors × 843.27 sqm FAR per tower plus service area.');
}
// ---------------------------------------------------------------- 5 Assumptions
{
  const s = base(); title(s, 'Planning Assumptions', 'Stated up-front so every number downstream is auditable');
  const A = [['Contract & delivery', 'Single EPC main contractor under NBCC; MIVAN formwork system, lifts, STP/WTP, substation and façade glazing as specialist subcontracts. Design frozen before NTP.'],
    ['Calendar', '7-day week, 8 h/day. CAQM GRAP window 16 Nov–15 Jan each year: Cat-A outdoor work stops, Cat-B indoor trades at 60 %, Cat-C statutory unaffected.'],
    ['Start / finish', 'NTP 01 Jan 2024 → final handover 14 May 2026 (865 days). A 4-week contingency before the contractual date is recommended for a severe-AQI winter.'],
    ['Structure', 'MIVAN monolithic walls + slab, 7-day floor cycle, 34 levels per tower; 12 tower cranes in 3 staggered batches; substructure in 4 parallel zones.'],
    ['Rates', 'BOQ at DSR 2023 + NDSR (17-Nov-2024). Resource rates NCR 2024: labour ₹650–1,500/day, rebar ₹62/kg, RMC M35 ₹6,500/cum, tower crane ₹25,000/day.'],
    ['Cost basis', 'Construction cost only — no land, fees, approvals or GST. MEP, lifts, services & prelims carried as allowances. 5 % management reserve outside BAC.']];
  const cw = 3.95, ch = 2.35;
  A.forEach((a, i) => card(s, MX + (i % 3) * (cw + 0.12), 1.7 + Math.floor(i / 3) * (ch + 0.2), cw, ch, a[0], a[1], { bs: 12.5 }));
}
// ---------------------------------------------------------------- 6 Methodology
{
  const s = base(); title(s, 'Methodology', 'PMBOK-aligned planning sequence — Excel model mirrored in Microsoft Project');
  const st = ['Scope & WBS', 'Activity definition (L3)', 'Sequencing & calendars', 'Duration estimating', 'Resource loading', 'Cost loading', 'Baseline & critical path', 'Cash flow & EVM'];
  const sp = (W - 2 * MX) / 8;
  s.addShape(pres.shapes.LINE, { x: MX + sp / 2, y: 2.35, w: sp * 7, h: 0, line: { color: C.line, width: 2 } });
  st.forEach((t, i) => {
    circleNum(s, MX + i * sp + sp / 2 - 0.35, 2.0, 0.7, String(i + 1), i < 4 ? C.green : C.gold);
    s.addText(t, { x: MX + i * sp, y: 2.8, w: sp, h: 0.6, align: 'center', fontFace: BF, fontSize: 12, color: C.ink, margin: 0, isTextBox: true, valign: 'top' });
  });
  card(s, MX, 3.8, 5.95, 2.9, 'Tools & outputs', 'Excel model (this submission): L3 schedule + 83-resource sheet, 9,864 assignments, BOQ→WBS cost map, cash flow, histogram and EVM with live formulas. Microsoft Project: tasks, logic, GRAP task calendars, Resource Names and Fixed Cost imported through the Import Wizard; Baseline 0, Earned Value tables and cost reports generated from MSP.', { bs: 12.5 });
  card(s, MX + 6.13, 3.8, 5.95, 2.9, 'Quality checks applied', `CPM re-run in Python reproduces every Excel start/finish exactly (0 of 2,758 rows differ). Only the project start is fixed; no date constraints. ${D.links.toLocaleString('en-IN')} links, ${D.lagged} with lags. Open ends flagged: ${D.no_succ} activities without successors (mainly zone/band ends and external works) — to be tied to handover milestones before baselining. Longest activity ${D.maxdur} d.`, { bs: 12.5 });
}
// ---------------------------------------------------------------- 7 WBS L1-L2
{
  const s = base(); title(s, 'Work Breakdown Structure — Level 1 & 2', '7 phases · 78 work packages · 2,672 Level-3 activities · 2,758 schedule rows');
  const P = [['1 Pre-construction', 'Approvals · Mobilisation · Site office & camp · Survey · Utilities · Batching plant · Tower cranes (3 batches) · MIVAN design & mock-up', '23 activities'],
    ['2 Substructure', 'Excavation & shoring · Dewatering · Raft & PCC · Waterproofing · B2 & B1 walls/slabs · Backfill · Podium — 4 zones', '40 activities'],
    ['3 Superstructure', '12 towers × 34 floors × 5-step MIVAN cycle: rebar → formwork → QA → pour → cure/strike', '2,040 activities'],
    ['4 Finishing', '12 towers × 3 bands: AAC blockwork · putty · façade texture · flooring · painting · doors/windows/railings', '216 activities'],
    ['5 MEP & lifts', 'Electrical · plumbing · HVAC · fire fighting first & second fix · floor T&C · 8 lifts per tower', '240 activities'],
    ['6 External works', 'Compound wall · roads & paving · storm drains · STP · WTP · substation & DG · RWH/UGT · landscape · clubhouse · lighting', '40 activities'],
    ['7 T&C & handover', 'OC filing · lift & fire commissioning · snagging · deep cleaning · handover per tower · completion', '73 activities']];
  const cw = 2.95, ch = 2.45;
  P.forEach((p, i) => {
    const x = MX + (i % 4) * (cw + 0.13), y = 1.65 + Math.floor(i / 4) * (ch + 0.2);
    card(s, x, y, cw, ch, p[0], p[1], { bs: 11.5 });
    s.addText(p[2], { x: x + 0.2, y: y + ch - 0.45, w: cw - 0.4, h: 0.3, fontFace: BF, fontSize: 12, bold: true, color: C.gold, margin: 0, isTextBox: true });
  });
  const x = MX + 3 * (cw + 0.13), y = 1.65 + ch + 0.2;
  s.addShape(pres.shapes.RECTANGLE, { x, y, w: cw, h: ch, fill: { color: C.dark }, line: { color: C.dark } });
  s.addText([{ text: 'Coding', options: { bold: true, fontSize: 15, color: C.gold, breakLine: true, paraSpaceAfter: 8 } },
    { text: 'L1 phase (1–7) · L2 package / tower · L3 activity suffixed by tower, floor, band or zone, e.g. T07.F21.d = Tower 07, floor 21, concreting.', options: { fontSize: 11.5, color: C.white } }],
    { x: x + 0.2, y: y + 0.15, w: cw - 0.4, h: ch - 0.3, fontFace: BF, valign: 'top', margin: 0, isTextBox: true });
}
// ---------------------------------------------------------------- 8 WBS L3 sample
{
  const s = base(); title(s, 'WBS Level 3 & Dictionary — Sample Package', 'Package T01.F10 — Tower 01 (Krishna) typical floor MIVAN cycle, repeated for 34 floors × 12 towers');
  const rows = [['WBS', 'Activity', 'Dur (d)', 'Pred.', 'Resources loaded (units)', 'Cost (₹ L)']];
  D.pkg.forEach(p => rows.push([p[0], p[1].replace(/\s*\(incl\. MEP$/, '').replace(/\(Walls \+ Sla$/, '').replace(/\(Pump-p$/, '').replace(/\(Panel Strike, Clean & Oil$/, ''), p[2], p[3], p[4], fmt(p[5])]));
  table(s, rows, MX, 1.6, W - 2 * MX, [1.2, 3.9, 0.8, 0.9, 4.2, 1.1], { fs: 11, rowH: 0.5 });
  card(s, MX, 4.9, 5.95, 1.85, 'Dictionary — T01.F10.d concreting', 'Scope: pump-place M35 in walls, columns, beams and slab in one monolithic pour (BOQ average 370 cum per floor-unit). Deliverable: pour card, cube register. Responsible: structures engineer. Acceptance: 28-day cube ≥ M35. Calendar: GRAP-A (stops 16 Nov–15 Jan).', { bs: 11.5 });
  card(s, MX + 6.13, 4.9, 5.95, 1.85, 'How cost is built', 'Budget = tower BOQ (rebar ₹360 Cr, M35 ₹155 Cr, MIVAN ₹125 Cr) ÷ 408 floor-units. Labour, plant and material come from the assignments; the balance of the BOQ rate is Fixed Cost (formwork system, subcontract margin, OH&P).', { bs: 11.5 });
}
// ---------------------------------------------------------------- 9 Master Gantt
{
  const s = base(); title(s, 'Master Schedule — Phase Gantt', '29 months · Jan 2024 – May 2026 · 2,672 activities rolled up to L1 · status line 30 Apr 2025');
  const x0 = 3.9, x1 = W - MX, t0 = new Date('2024-01-01'), t1 = new Date('2026-06-01');
  const X = d => x0 + (new Date(d) - t0) / (t1 - t0) * (x1 - x0);
  const months = []; for (let y = 2024; y <= 2026; y++) for (let m = 0; m < 12; m++) { const d = new Date(y, m, 1); if (d < t1) months.push(d); }
  months.forEach((d, i) => {
    s.addShape(pres.shapes.LINE, { x: X(d), y: 1.95, w: 0, h: 4.7, line: { color: 'E6EBE8', width: 0.75 } });
    if (d.getMonth() % 3 === 0) s.addText(d.toLocaleDateString('en-GB', { month: 'short' }) + '-' + String(d.getFullYear()).slice(2), { x: X(d) - 0.1, y: 1.6, w: 0.9, h: 0.3, fontFace: BF, fontSize: 10, color: C.muted, margin: 0, isTextBox: true });
  });
  // GRAP windows
  [['2024-01-01', '2024-01-15'], ['2024-11-16', '2025-01-15'], ['2025-11-16', '2026-01-15']].forEach(g => s.addShape(pres.shapes.RECTANGLE,
    { x: X(g[0]), y: 1.95, w: X(g[1]) - X(g[0]), h: 4.7, fill: { color: C.red, transparency: 88 }, line: { color: C.red, transparency: 100 } }));
  const cols = [C.slate, C.green, C.green, C.green, C.gold, C.green, C.gold];
  D.phases.forEach((p, i) => {
    const y = 2.1 + i * 0.64;
    s.addText(p[0], { x: MX, y, w: 3.2, h: 0.42, fontFace: BF, fontSize: 12, color: C.ink, valign: 'middle', margin: 0, isTextBox: true });
    s.addShape(pres.shapes.RECTANGLE, { x: X(p[1]), y: y + 0.05, w: Math.max(0.05, X(p[2]) - X(p[1])), h: 0.32, fill: { color: cols[i] }, line: { color: cols[i] } });
  });
  const sd = X('2025-04-30');
  s.addShape(pres.shapes.LINE, { x: sd, y: 1.95, w: 0, h: 4.75, line: { color: C.red, width: 1.5, dashType: 'dash' } });
  s.addText('Status 30 Apr 2025', { x: sd + 0.05, y: 6.42, w: 1.8, h: 0.25, fontFace: BF, fontSize: 10, color: C.red, margin: 0, isTextBox: true });
  s.addText('Shaded bands = GRAP winter windows (Cat-A work stops)', { x: MX, y: 6.42, w: 3.3, h: 0.25, fontFace: BF, fontSize: 10, color: C.red, margin: 0, isTextBox: true });
}
// ---------------------------------------------------------------- 10 Calendars & logic
{
  const s = base(); title(s, 'Calendars, Logic & Durations', 'How the network was built — and how MS Project reproduces it');
  const fs = Math.round(D.rel.FS / D.links * 100), ss = 100 - fs;
  card(s, MX, 1.65, 3.95, 5.05, 'Three task calendars', 'Standard — 7-day week, 08:00–17:00 (8 h).\n\nGRAP-A Full Stop — Standard with 16 Nov–15 Jan non-working each winter (2,194 activities: earthwork, RCC shell, façade, external works).\n\nGRAP-B Derated 60 % — same windows at 4.8 h/day, so indoor trades stretch by 1/0.6 (447 activities: MEP, finishes, lifts).\n\nStandard for Cat-C statutory/off-site work (31 activities).', { bs: 12 });
  card(s, MX + 4.07, 1.65, 3.95, 5.05, 'Relationship types used', `FS ${fs} % (${D.rel.FS.toLocaleString('en-IN')} links) — floor-to-floor, trade-to-trade chains.\n\nSS + lag ${ss} % (${D.rel.SS} links) — tower batch staggers (SS+15 d) and parallel site set-up.\n\nFS + lag (${D.lagged} lagged links) — finishing starts 19 d after floor cast, MEP first fix 9 d after, lifts 29 d after top-out.\n\nNo FF/SF used — keeps the L3 network simple to audit.`, { bs: 12 });
  card(s, MX + 8.14, 1.65, 3.95, 5.05, 'Duration estimating', 'Parametric: quantity ÷ (gang output × gangs).\n\nExample: typical floor pour 134.5 cum (BOQ SummaryTower) at 60 cum/h boom pump ≈ 2.3 h + set-up → 1 day.\n\nTypical floor rebar 14.4 t ÷ (2.2 t/gang-day × 2 d) → 3.3 gangs; 5 gangs loaded because tower steel averages 53.8 t per floor-unit incl. stilt & mumty levels.\n\nMIVAN erection 850 sqm floor plate/day → 1 day.\n\nLead-time driven: MIVAN design + fabrication 105 d; lifts 45 d install + 15 d test.', { bs: 12 });
}
// ---------------------------------------------------------------- 11 Milestones
{
  const s = base(); title(s, 'Key Milestones', 'Contractual and internal control points computed by the CPM');
  const rows = [['ID', 'Milestone', 'Planned date', 'Type', 'Driven by']];
  D.milestones.forEach(m => rows.push([m[0], m[1], fmtDate(m[2]), m[3], m[4]]));
  table(s, rows, MX, 1.6, W - 2 * MX, [0.8, 6.0, 1.9, 1.6, 1.83], { fs: 12, rowH: 0.42 });
  s.addText(`First tower handover (T01 Krishna): ${fmtDate(D.first_handover)} — handovers then roll tower-by-tower to May 2026.`, { x: MX, y: 6.75, w: 12, h: 0.3, fontFace: BF, fontSize: 11, italic: true, color: C.muted, margin: 0, isTextBox: true });
}
// ---------------------------------------------------------------- 12 Critical path
{
  const s = base(); title(s, 'Critical Path Analysis', `Driving chain = ${D.driving} activities · ${D.tf.zero} activities at zero total float · finish 14 May 2026`);
  const chain = ['Mobilise & survey', 'Zone-3 excavation', 'Zone-3 raft → B2 → B1', 'Zone-3 podium', 'T09→T10→T11 staggered starts', 'T12 34-floor MIVAN cycle', 'T12 upper-zone finishes', 'OC & handover'];
  const bw = 1.42, gap = 0.1;
  chain.forEach((c, i) => {
    const x = MX + i * (bw + gap), hot = i === 5;
    s.addShape(pres.shapes.RECTANGLE, { x, y: 1.75, w: bw, h: 1.0, fill: { color: hot ? C.gold : C.dark }, line: { color: hot ? C.gold : C.dark } });
    s.addText(c, { x: x + 0.05, y: 1.75, w: bw - 0.1, h: 1.0, align: 'center', valign: 'middle', fontFace: BF, fontSize: 11, color: C.white, margin: 0, isTextBox: true });
  });
  const tw = 2.9, g = 0.17;
  stat(s, MX, 3.05, tw, 1.3, `${D.tf.zero} / 2,672`, 'Zero-float activities — T12 façade restart after the 2025-26 GRAP window');
  stat(s, MX + tw + g, 3.05, tw, 1.3, `${D.driving}`, 'Activities on the driving (longest) path — MSP Task Path');
  stat(s, MX + 2 * (tw + g), 3.05, tw, 1.3, `${D.tf.le30}`, 'Near-critical: total float ≤ 30 days');
  stat(s, MX + 3 * (tw + g), 3.05, tw, 1.3, '238 days', 'T12 structural shell: 34 floors × 7-day MIVAN cycle');
  s.addText([
    { text: 'The GRAP window converts slack: structure on T12 carries 28 d float because façade texture (Cat-A) cannot start until 16 Jan 2026 anyway — the formal critical path starts there.', options: { bullet: true, breakLine: true } },
    { text: 'The last tower batch (T11/T12) is the schedule lever: its start is pushed from Nov-24 to 16 Jan 2025 by the 2024-25 GRAP stop.', options: { bullet: true, breakLine: true } },
    { text: `External works carry ${D.ext_float[0]}–${D.ext_float[1]} d float (finish Mar-25, no successors) — re-sequence later to avoid damage by construction traffic and free cash.`, options: { bullet: true, breakLine: true } },
    { text: `${D.no_succ} open-ended activities must be tied to the tower handovers before Baseline 0, otherwise their float is overstated.`, options: { bullet: true } }],
    { x: MX, y: 4.6, w: W - 2 * MX, h: 2.2, fontFace: BF, fontSize: 13, color: C.muted, margin: 0, paraSpaceAfter: 6, valign: 'top', isTextBox: true });
}
// ---------------------------------------------------------------- 13 MIVAN cycle
{
  const s = base(); title(s, 'Superstructure Strategy — 7-Day MIVAN Floor Cycle', 'One MIVAN set and one tower crane per tower · 12 towers in 3 batches · monolithic walls + slab');
  const steps = [['Rebar fixing – walls, beams, slab, sleeves', 1, 2, C.green], ['MIVAN formwork erection & alignment', 3, 1, C.green], ['Pre-pour checks, MEP embeds, QA/QC', 4, 1, C.slate],
    ['Monolithic concreting (boom pump)', 5, 1, C.gold], ['Curing & panel strike / clean / oil', 6, 2, C.green]];
  const x0 = 4.6, dw = (W - MX - x0) / 7;
  for (let d = 1; d <= 7; d++) s.addText('D' + d, { x: x0 + (d - 1) * dw, y: 1.6, w: dw, h: 0.3, align: 'center', fontFace: BF, fontSize: 11, color: C.muted, margin: 0, isTextBox: true });
  steps.forEach((st, i) => {
    const y = 2.0 + i * 0.52;
    s.addText(st[0], { x: MX, y, w: 3.9, h: 0.4, fontFace: BF, fontSize: 12, color: C.ink, valign: 'middle', margin: 0, isTextBox: true });
    s.addShape(pres.shapes.RECTANGLE, { x: x0 + (st[1] - 1) * dw, y: y + 0.05, w: st[2] * dw, h: 0.32, fill: { color: st[3] }, line: { color: st[3] } });
  });
  card(s, MX, 4.8, 5.95, 1.9, 'Output', '34 floors × 7 days = 238 days per tower before GRAP. Batch starts are staggered SS+15 d so each tower keeps its own crane and MIVAN set. All 12 shells complete 25 Sep 2025 (first: T01 on 10 Jul 2025).', { bs: 12 });
  card(s, MX + 6.13, 4.8, 5.95, 1.9, 'Risk control', 'Every shell activity is GRAP-A. If Stage-IV runs past 15 Jan, a 3rd MIVAN set on batch-3 compresses the cycle to 5 days; pre-stock 45 days of rebar/cement before 16 Nov to beat truck-entry bans.', { bs: 12 });
}
// ---------------------------------------------------------------- 14 Resources
{
  const s = base(); title(s, 'Resource Planning', `${D.nres.Labour + D.nres.Staff + D.nres.Equipment + D.nres.Material} resources · ${D.nassign.toLocaleString('en-IN')} task assignments · loaded on all 2,672 L3 activities`);
  const R = [['Labour (work)', `${D.nres.Labour} trades at ₹650–1,500/day — bar benders, MIVAN fitters, masons, tile-layers, painters, electricians, plumbers. Gangs sized from BOQ labour content (≤5 gangs/task).`],
    ['Staff (work)', `${D.nres.Staff} roles — site, QA/QC, MEP, OEM, liaison engineers, architect, PM & PMO — ₹2,500–25,000/day.`],
    ['Equipment (work)', `${D.nres.Equipment} plant items — 12 tower cranes, boom pumps, excavators, tippers, gondolas, hoists, mobile crane — hire ₹2,500–35,000/day.`],
    ['Material (material)', `${D.nres.Material} materials with BOQ quantities — rebar 33,438 t, RMC 2.76 lakh cum, tiles 5.5 lakh sqm, paint 16 lakh sqm, 1,507 kitchens.`]];
  R.forEach((r, i) => card(s, MX + i * 3.05, 1.65, 2.93, 2.75, r[0], r[1], { bs: 12 }));
  s.addText('Loading & levelling approach', { x: MX, y: 4.65, w: 8, h: 0.4, fontFace: BF, fontSize: 16, bold: true, color: C.ink, margin: 0, isTextBox: true });
  s.addText([
    { text: 'Crew text in the L3 schedule is converted into MS Project "Resource Names" (e.g. Bar Bender[9000%], Rebar Fe500D[53.76]) and imported with the tasks', options: { bullet: true, breakLine: true } },
    { text: 'Subcontract balance of each BOQ rate goes to task Fixed Cost (prorated), so MSP total cost = Excel BAC to the rupee', options: { bullet: true, breakLine: true } },
    { text: `Peak ${D.peak_head.toLocaleString('en-IN')} workers (monthly average, ${D.peak_head_month}); levelling by float on finishes & external works — no automatic levelling of the tower cycle`, options: { bullet: true } }],
    { x: MX, y: 5.05, w: W - 2 * MX, h: 1.7, fontFace: BF, fontSize: 13, color: C.muted, margin: 0, paraSpaceAfter: 5, valign: 'top', isTextBox: true });
}
// ---------------------------------------------------------------- 15 Histogram
{
  const s = base(); title(s, 'Labour Histogram', 'Average daily site headcount by month (labour + staff) — reproduced in MSP Resource Graph');
  s.addChart(pres.charts.BAR, [{ name: 'Avg daily headcount', labels: D.head.map(h => h[0]), values: D.head.map(h => h[1]) }],
    { x: MX, y: 1.55, w: W - 2 * MX, h: 4.75, chartColors: [C.green], showValue: true, dataLabelPosition: 'outEnd', dataLabelFontSize: 8, dataLabelFormatCode: '#,##0',
      catAxisLabelFontSize: 9, valAxisLabelFontSize: 9, valGridLine: { color: 'E6EBE8', size: 0.5 }, catGridLine: { style: 'none' }, barGapWidthPct: 40, valAxisLabelFormatCode: '#,##0' });
  s.addText(`Peak ${D.peak_head.toLocaleString('en-IN')} (monthly average, ${D.peak_head_month}) when all 12 shells, finishing and MEP overlap · peak day ${D.peak_day.toLocaleString('en-IN')} · ≈ ${D.worker_months.toLocaleString('en-IN')} worker-months · dips = GRAP winters`,
    { x: MX, y: 6.4, w: W - 2 * MX, h: 0.35, fontFace: BF, fontSize: 11.5, color: C.muted, margin: 0, isTextBox: true });
}
// ---------------------------------------------------------------- 16 Cost plan
{
  const s = base(); title(s, 'Cost Plan — Budget at Completion', `Bottom-up from 2,672 activity costs · BAC ₹ ${fmt(D.bac, 1)} Cr`);
  s.addChart(pres.charts.BAR, [{ name: '₹ Cr', labels: D.phase_cost.map(p => p[0]), values: D.phase_cost.map(p => p[1]) }],
    { x: MX, y: 1.55, w: 7.6, h: 5.2, barDir: 'bar', chartColors: [C.gold], showValue: true, dataLabelPosition: 'outEnd', dataLabelFormatCode: '#,##0.0', dataLabelFontSize: 11,
      catAxisOrientation: 'maxMin', catAxisLabelFontSize: 12, valAxisLabelFontSize: 10, valGridLine: { color: 'E6EBE8', size: 0.5 }, catGridLine: { style: 'none' } });
  stat(s, 8.5, 1.6, 4.25, 1.55, `₹ ${fmt(D.bac, 0)} Cr`, 'Budget at Completion (BAC) — excl. GST');
  stat(s, 8.5, 3.3, 4.25, 1.55, `₹ ${D.psf.toLocaleString('en-IN')} / sft`, 'Construction cost per sft BUA (₹29,290/sqm) · ₹1.09 Cr per DU');
  stat(s, 8.5, 5.0, 4.25, 1.55, `₹ ${fmt(D.bac * 0.05, 1)} Cr`, 'Management reserve (5 %) — outside BAC');
}
// ---------------------------------------------------------------- 17 Cost by source & type
{
  const s = base(); title(s, 'Cost Build-up by Source & Resource Type', 'How the BOQ becomes a resource-loaded budget');
  const T = D.types, lab = Object.keys(T), val = Object.values(T);
  s.addChart(pres.charts.DOUGHNUT, [{ name: 'Type', labels: lab, values: val }], { x: MX, y: 1.5, w: 5.0, h: 4.9, holeSize: 55,
    chartColors: [C.green, C.gold, C.slate, C.dark, C.sage], showPercent: true, showValue: false, dataLabelColor: C.white, dataLabelFontSize: 11,
    showLegend: true, legendPos: 'b', legendFontSize: 10 });
  table(s, [['Source', '₹ Cr', 'Loaded in MSP as'],
    ['Civil BOQ – structure + finishes (118 lines)', fmt(D.boq, 1), 'Material × qty + crew + Fixed Cost'],
    ['External development BOQ', fmt(D.ext, 1), 'Crew + Fixed Cost'],
    ['Allowances – MEP, lifts, services, prelims', fmt(D.allow, 1), 'Crew + Fixed Cost'],
    ['Crew cost with no BOQ item (QA, T&C, statutory) + uplift', fmt(D.bac - D.boq - D.ext - D.allow, 1), 'Work resources'],
    ['Total BAC', fmt(D.bac, 1), '']], 6.0, 1.6, 6.73, [3.4, 1.0, 2.33], { boldLast: true, fs: 11, rowH: 0.44 });
  table(s, [['Resource class', '₹ Cr', 'Share'], ...lab.map((l, i) => [l, fmt(val[i], 1), fmt(val[i] / D.bac * 100, 1) + ' %'])], 6.0, 4.4, 6.73, [3.9, 1.3, 1.53], { fs: 11, rowH: 0.36 });
}
// ---------------------------------------------------------------- 18 Baseline
{
  const s = base(); title(s, 'Time-Cost Baseline', 'Baseline 0 saved in MS Project before NTP (31 Dec 2023)');
  card(s, MX, 1.6, 3.95, 2.1, 'Baseline schedule', 'Start 01 Jan 2024 · finish 14 May 2026 · 865 days · 2,758 rows (2,672 L3) · 11 milestones · driving path 197 activities.', { bs: 12.5 });
  card(s, MX + 4.07, 1.6, 3.95, 2.1, 'Baseline cost', `BAC ₹ ${fmt(D.bac, 1)} Cr time-phased by working day on the GRAP calendars (PV curve). Fixed costs prorated; labour, plant and material rate-based.`, { bs: 12.5 });
  card(s, MX + 8.14, 1.6, 3.95, 2.1, 'Baseline resources', `${D.nres.Labour} trades, ${D.nres.Staff} staff roles, ${D.nres.Equipment} plant, ${D.nres.Material} materials; ${D.nassign.toLocaleString('en-IN')} assignments with baseline work and cost.`, { bs: 12.5 });
  const q = D.qpv; let cum = 0; const cumv = q.map(r => (cum += r[1]));
  table(s, [['Quarter', ...q.map(r => r[0].replace('20', '').replace('-Q', ' Q'))], ['Planned value (₹ Cr)', ...q.map(r => fmt(r[1], 1))],
    ['Cumulative PV (₹ Cr)', ...cumv.map(v => fmt(v, 1))], ['Cumulative % of BAC', ...cumv.map(v => Math.round(v / D.bac * 100) + ' %')]],
    MX, 4.1, W - 2 * MX, [2.33, ...q.map(() => 0.975)], { fs: 11, rowH: 0.45 });
  s.addText('Change control: approved variations are re-baselined into Baseline 1 only; Baseline 0 stays the contract reference.', { x: MX, y: 6.2, w: 12, h: 0.3, fontFace: BF, fontSize: 11, italic: true, color: C.muted, margin: 0, isTextBox: true });
}
// ---------------------------------------------------------------- 19 Cash flow
{
  const s = base(); title(s, 'Monthly Cash Flow & S-Curve', 'Planned value by month (bars) and cumulative (line)');
  let cum = 0; const cumv = D.monthly.map(m => +(cum += m[1]).toFixed(1));
  const labels = D.monthly.map(m => m[0]);
  s.addChart([{ type: pres.charts.BAR, data: [{ name: 'Monthly PV (₹ Cr)', labels, values: D.monthly.map(m => m[1]) }], options: { chartColors: [C.green], barGapWidthPct: 40 } },
    { type: pres.charts.LINE, data: [{ name: 'Cumulative PV (₹ Cr)', labels, values: cumv }], options: { chartColors: [C.gold], lineSize: 3, lineDataSymbol: 'none', secondaryValAxis: true, secondaryCatAxis: true } }],
    { x: MX, y: 1.5, w: W - 2 * MX, h: 4.85, valAxes: [{ showValAxisTitle: true, valAxisTitle: 'Monthly ₹ Cr', valAxisLabelFontSize: 9, valGridLine: { color: 'E6EBE8', size: 0.5 } },
      { showValAxisTitle: true, valAxisTitle: 'Cumulative ₹ Cr', valAxisLabelFontSize: 9, valGridLine: { style: 'none' } }],
      catAxes: [{ catAxisLabelFontSize: 9 }, { catAxisHidden: true }], showLegend: true, legendPos: 'b', legendFontSize: 10 });
  s.addText(`Peak month ₹ ${fmt(D.monthly_peak[1], 1)} Cr (${D.monthly_peak[0]}) · ${D.years[0]} % of BAC in 2024, ${D.years[1]} % in 2025, ${D.years[2]} % in 2026 · spend collapses in each GRAP window (Dec-24, Dec-25)`,
    { x: MX, y: 6.45, w: W - 2 * MX, h: 0.3, fontFace: BF, fontSize: 11.5, color: C.muted, margin: 0, isTextBox: true });
}
// ---------------------------------------------------------------- 20 EVM framework
{
  const s = base(); title(s, 'Earned Value Management — Framework', 'Reported monthly from MS Project Earned Value tables and the Excel EVM sheet');
  const E = [['PV', 'Planned value = baseline cost of work scheduled to the status date (BCWS)'], ['EV', 'Earned value = BAC × physical % complete (BCWP)'],
    ['AC', 'Actual cost of work performed to date (ACWP)'], ['CPI', 'Cost performance index = EV ÷ AC'], ['SPI', 'Schedule performance index = EV ÷ PV'],
    ['EAC', 'Estimate at completion = BAC ÷ CPI'], ['ETC', 'Estimate to complete = EAC − AC'], ['VAC', 'Variance at completion = BAC − EAC']];
  E.forEach((e, i) => {
    const x = MX + (i % 2) * 6.1, y = 1.6 + Math.floor(i / 2) * 1.2;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 5.98, h: 1.05, fill: { color: C.card }, line: { color: C.card } });
    s.addText(e[0], { x: x + 0.2, y, w: 1.1, h: 1.05, fontFace: HF, fontSize: 26, bold: true, color: C.gold, valign: 'middle', margin: 0, isTextBox: true });
    s.addText(e[1], { x: x + 1.35, y, w: 4.5, h: 1.05, fontFace: BF, fontSize: 13, color: C.ink, valign: 'middle', margin: 0, isTextBox: true });
  });
  s.addText('% complete rule: 0/100 for floor-cycle steps (1–2 days), 50/50 for 6–15-day activities, physical measurement for finishing bands, MEP and external works.', { x: MX, y: 6.45, w: 12, h: 0.3, fontFace: BF, fontSize: 11, color: C.muted, margin: 0, isTextBox: true });
}
// ---------------------------------------------------------------- 21 EVM status (dark)
{
  const s = base(true); title(s, 'EVM Status — 30 April 2025', 'Month 16 of 29 · substructure complete · 12 shells rising · finishing & MEP started (illustrative progress)', true);
  const e = D.evm, sv = e.ev - e.pv, cv = e.ev - e.ac, vac = e.bac - e.eac;
  const tiles = [[`₹ ${fmt(e.pv, 1)} Cr`, 'PV', 0], [`₹ ${fmt(e.ev, 1)} Cr`, 'EV', 0], [`₹ ${fmt(e.ac, 1)} Cr`, 'AC', 0], [`₹ ${fmt(e.bac, 1)} Cr`, 'BAC', 0],
    [fmt(e.cpi, 2), 'CPI', 1], [fmt(e.spi, 2), 'SPI', 1], [`₹ ${fmt(e.eac, 1)} Cr`, 'EAC', 0], [`₹ ${fmt(e.eac - e.ac, 1)} Cr`, 'ETC', 0],
    [`₹ ${fmt(vac, 1)} Cr`, 'VAC', 1], [`₹ ${fmt(sv, 1)} Cr`, 'SV', 1], [`₹ ${fmt(cv, 1)} Cr`, 'CV', 1], [fmt(e.ev / e.bac * 100, 1) + ' %', '% complete', 0]];
  tiles.forEach((t, i) => stat(s, MX + (i % 4) * 3.07, 1.65 + Math.floor(i / 4) * 1.5, 2.95, 1.35, t[0], t[1], { fill: '24463D', color: t[2] ? 'F08C7A' : C.gold, lc: C.lightInk, fs: 28 }));
  s.addText(`Interpretation: ${Math.round((1 - e.spi) * 100)} % behind schedule (batch-3 towers 3–4 floors late after an extended GRAP Stage-IV) and ${Math.round((1 / e.cpi - 1) * 100)} % over cost (rebar +5 %). At current CPI the project lands ₹ ${fmt(-vac, 1)} Cr over BAC — inside the ₹ ${fmt(e.bac * 0.05, 1)} Cr management reserve, but recovery is required.`,
    { x: MX, y: 6.2, w: W - 2 * MX, h: 0.6, fontFace: BF, fontSize: 12, color: C.white, margin: 0, isTextBox: true });
}
// ---------------------------------------------------------------- 22 EVM curves
{
  const s = base(); title(s, 'EVM Dashboard — PV / EV / AC Curves', 'Cumulative ₹ Cr · status 30 Apr 2025 · forecast to EAC');
  const R = D.evm_rows, labels = R.map(r => r[0]);
  const ser = (k, nm) => ({ name: nm, labels, values: R.map(r => r[k] === null ? null : r[k]) });
  s.addChart(pres.charts.LINE, [ser(1, 'PV'), ser(2, 'EV'), ser(3, 'AC'), ser(4, 'Forecast (EAC)')], { x: MX, y: 1.5, w: 8.3, h: 5.2,
    chartColors: [C.slate, C.green, C.red, C.gold], lineSize: 2.5, lineDataSymbol: 'none', catAxisLabelFontSize: 9, valAxisLabelFontSize: 9,
    valGridLine: { color: 'E6EBE8', size: 0.5 }, catGridLine: { style: 'none' }, showLegend: true, legendPos: 'b', legendFontSize: 10, displayBlanksAs: 'gap' });
  const e = D.evm;
  card(s, 9.1, 1.6, 3.63, 1.6, 'Cost variance', `CV −₹${fmt(e.ac - e.ev, 1)} Cr · CPI ${fmt(e.cpi, 2)} · rebar escalation + RMC diesel surcharge`, { bs: 12 });
  card(s, 9.1, 3.35, 3.63, 1.6, 'Schedule variance', `SV −₹${fmt(e.pv - e.ev, 1)} Cr · SPI ${fmt(e.spi, 2)} · ≈ 3–4 floors late on T11/T12`, { bs: 12 });
  card(s, 9.1, 5.1, 3.63, 1.6, 'Forecast', `EAC ₹${fmt(e.eac, 1)} Cr · TCPI to hit BAC = ${fmt((e.bac - e.ev) / (e.bac - e.ac), 2)} · recovery plan next`, { bs: 12 });
}
// ---------------------------------------------------------------- 23 Risks
{
  const s = base(); title(s, 'Risks & Recovery Plan', 'Turning the EVM signal into action');
  table(s, [['Risk / issue', 'Impact', 'Response', 'Owner', 'Effect'],
    ['Batch-3 towers late (SPI 0.95)', 'Top-out of T12 slips → finishes & OC late', '3rd MIVAN set on T11/T12; Sunday pours outside GRAP', 'Structures Mgr', 'Cycle 7 → 5 d, recovers ~25 d'],
    ['Rebar & cement escalation (CPI 0.97)', '+₹23.5 Cr to date, +₹56 Cr at EAC', 'Lock balance ~17,000 t rebar price; value-engineer finishes', 'Commercial Mgr', '−₹15 to 20 Cr'],
    ['GRAP Stage-IV beyond 62-day window', 'Cat-A work stops; truck-entry ban', '45-day material stockpile before 16 Nov; indoor trades re-sequenced', 'Planner', 'Protects M6/M9'],
    ['External works scheduled too early', 'Damage by site traffic; cash tied up', `Use ${D.ext_float[0]}+ d float: move roads/landscape to Oct-25–Feb-26`, 'Planner', 'Nil on CP; cash deferred'],
    ['MEP / lift scope not in BOQ', 'Allowance ₹263 Cr unverified', 'Issue MEP & lift BOQs; replace allowances before Baseline 1', 'MEP Mgr', 'BAC certainty'],
    ['OC approvals tower-wise', 'Handover delay', 'File fire NOC & lift licences 90 d before PC', 'Liaison', 'Nil']], MX, 1.6, W - 2 * MX, [2.75, 2.45, 3.55, 1.45, 1.93], { fs: 11, rowH: 0.62 });
  s.addText('Post-recovery forecast: finish 14 May 2026 inside a 4-week buffer; EAC ≈ ₹1,670 Cr (VAC ≈ −₹34 Cr, absorbed by the ₹81.8 Cr management reserve).',
    { x: MX, y: 6.3, w: W - 2 * MX, h: 0.4, fontFace: BF, fontSize: 12.5, bold: true, color: C.ink, margin: 0, isTextBox: true });
}
// ---------------------------------------------------------------- 24 MSP integration
{
  const s = base(); title(s, 'Linking Excel to MS Project', 'Import Wizard route — full step list in workbook sheet MSP_Link_Guide');
  const st = [['Options & calendars', 'Start 01-Jan-24 · 8 h/day, 7-day Standard · GRAP-A and GRAP-B (4.8 h) calendars'], ['Import resources', 'MSP_Import_Resources → Resource Sheet: type, label, max units, rate/d'],
    ['Import tasks', 'MSP_Import_Tasks: outline, duration, predecessors, Resource Names, Fixed Cost, Task Calendar'], ['Verify', 'Finish 14-May-26 · project cost = BAC ₹1,635.8 Cr · histogram matches'],
    ['Baseline 0', 'Project > Set Baseline — freezes dates, work, cost and PV'], ['Status & EVM', 'Status date → % complete → Earned Value tables & reports']];
  const bw = 1.93;
  st.forEach((t, i) => {
    const x = MX + i * (bw + 0.1);
    circleNum(s, x + bw / 2 - 0.35, 1.7, 0.7, String(i + 1), i % 2 ? C.gold : C.green);
    s.addText(t[0], { x, y: 2.55, w: bw, h: 0.4, align: 'center', fontFace: BF, fontSize: 13, bold: true, color: C.ink, margin: 0, isTextBox: true });
    s.addText(t[1], { x, y: 2.95, w: bw, h: 1.3, align: 'center', fontFace: BF, fontSize: 11, color: C.muted, margin: 0, isTextBox: true, valign: 'top' });
  });
  card(s, MX, 4.55, 5.95, 2.15, 'Why it reconciles exactly', 'Resource Names carry every crew and material quantity; Fixed Cost carries the balance of each BOQ-based budget; the GRAP task calendars reproduce the Excel dates (checked by an independent CPM run).', { bs: 12 });
  card(s, MX + 6.13, 4.55, 5.95, 2.15, 'Watch-outs', 'Do not map Start/Finish (creates constraints). Keep row order so IDs 1–2,758 match predecessors. Use your Windows list separator in Resource Names. Save the workbook after recalculation — MSP reads stored values.', { bs: 12 });
}
// ---------------------------------------------------------------- 25 Data reconciliation
{
  const s = base(); title(s, 'Data Reconciliation & Caveats', 'What the inputs disagree on, and how the plan treats it');
  const K = [['11 vs 12 towers', 'Area statement & BOQ sanction 11 towers (2B+S+G/P+35); the L3 schedule carries 12 tower-units (T08 "Godavari-2"). BOQ tower quantities are spread over the 12 units — confirm the tower list before Baseline 0.'],
    ['BOQ subtotal slip', 'BOQ row 647 (polycarbonate sheets, ₹0.94 Cr) sits outside the NDSR SUM range, so the printed total ₹1,353.98 Cr is short; the plan carries all lines = ₹1,354.91 Cr.'],
    ['Scope missing from BOQ', 'No MEP, lifts, substation/DG, STP/WTP or preliminaries in the BOQ → ₹262.7 Cr of allowances at benchmark rates, clearly separated for replacement.'],
    ['Crew vs quantity', 'Single-gang crews in the L3 text cannot place the BOQ quantities in the stated durations; gangs were scaled from the labour content of the DSR rates (cap 5 gangs).'],
    ['Open-ended logic', `${D.no_succ} activities have no successor (zone ends, external works) — tie them to handover before baselining so float is not overstated.`],
    ['EVM data', 'Status at 30-Apr-2025 is an illustrative scenario to demonstrate the method; the EVM sheet takes real % complete and actual cost as blue inputs.']];
  K.forEach((k, i) => card(s, MX + (i % 3) * 4.07, 1.6 + Math.floor(i / 3) * 2.6, 3.95, 2.45, k[0], k[1], { bs: 11.5 }));
}
// ---------------------------------------------------------------- 26 Conclusion
{
  const s = base(true); title(s, 'Conclusion & Deliverables', 'What this execution plan delivers', true);
  const L = ['Excel workbook — L3 schedule with resources & costs on all 2,672 activities', 'MSP import sheets + step-by-step linking guide (calendars, import, baseline, EVM)',
    `Resource sheet — ${D.nres.Labour + D.nres.Staff + D.nres.Equipment + D.nres.Material} resources, ${D.nassign.toLocaleString('en-IN')} assignments`, `Cost plan — BAC ₹${fmt(D.bac, 1)} Cr, ₹${D.psf.toLocaleString('en-IN')}/sft, reconciled to the BOQ`,
    'Cash-flow S-curve and labour histogram by month', 'Critical path: 7 zero-float, 197-activity driving chain, GRAP insight',
    'EVM dashboard — PV/EV/AC, CPI, SPI, EAC, ETC, VAC at month 16', 'Risk & recovery plan and data-reconciliation log'];
  L.forEach((t, i) => {
    const x = MX + (i % 2) * 6.1, y = 1.75 + Math.floor(i / 2) * 1.0;
    s.addShape(pres.shapes.OVAL, { x, y: y + 0.05, w: 0.42, h: 0.42, fill: { color: C.gold }, line: { color: C.gold } });
    s.addText('✓', { x, y: y + 0.05, w: 0.42, h: 0.42, align: 'center', valign: 'middle', fontFace: BF, fontSize: 14, bold: true, color: C.dark, margin: 0, isTextBox: true });
    s.addText(t, { x: x + 0.6, y, w: 5.3, h: 0.55, fontFace: BF, fontSize: 13.5, color: C.white, valign: 'middle', margin: 0, isTextBox: true });
  });
  s.addText('Thank you — questions & viva', { x: MX, y: 6.05, w: 8, h: 0.5, fontFace: HF, fontSize: 24, bold: true, color: C.gold, margin: 0, isTextBox: true });
}

pres.writeFile({ fileName: '../deliverables/Golf_Homes_PhII_Project_Execution_Plan.pptx' }).then(f => console.log('written', f));
