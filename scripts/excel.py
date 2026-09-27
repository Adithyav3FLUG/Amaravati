import pickle, datetime as dt, re
import numpy as np, pandas as pd, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as CL
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.comments import Comment
from openpyxl.worksheet.datavalidation import DataValidation

SRC = '../inputs/Golf_Homes_PhII_L3_Schedule_GRAP.xlsx'
OUT = 'Golf_Homes_PhII_L3_Schedule_Resource_Cost_Loaded.xlsx'
M = pickle.load(open('model.pkl', 'rb'))
L = M['L']; bm = M['bm']; RES = M['RES']; ALLOW = M['ALLOW']
df, preds, parent, children = pickle.load(open('sched.pkl', 'rb'))

NAVY = '1F4E78'
thin = Side(style='thin', color='BFBFBF'); BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
F = lambda **k: Font(name='Arial', size=k.pop('size', 10), **k)
HFILL = PatternFill('solid', fgColor=NAVY); INFILL = PatternFill('solid', fgColor='FFF2CC')
SUBFILL = PatternFill('solid', fgColor='D9E1F2'); TOTFILL = PatternFill('solid', fgColor='BDD7EE')
BLUE = F(color='0000FF'); GREEN = F(color='008000'); BOLD = F(bold=True)
INR = '#,##0;(#,##0);"-"'; CR = '#,##0.00;(#,##0.00);"-"'; PCT = '0.0%;(0.0%);"-"'

def header(ws, row, names, widths=None, height=30):
    for j, n in enumerate(names, 1):
        c = ws.cell(row, j, n); c.font = F(bold=True, color='FFFFFF'); c.fill = HFILL
        c.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center'); c.border = BOX
        if widths: ws.column_dimensions[CL(j)].width = widths[j - 1]
    ws.row_dimensions[row].height = height

def title(ws, text, sub=None):
    ws['A1'] = text; ws['A1'].font = F(bold=True, size=14, color=NAVY)
    if sub: ws['A2'] = sub; ws['A2'].font = F(italic=True, color='595959')

wb = openpyxl.load_workbook(SRC)
ws = wb['L3_Schedule']
N = ws.max_row  # 2759
ID2ROW = {int(ws.cell(r, 1).value): r for r in range(2, N + 1)}
Lr = L.set_index('ID')
ol = {r.ID: r.OL for _, r in df.iterrows()}
PH = {'1': '1 Pre-construction', '2': '2 Substructure', '3': '3 Superstructure', '4': '4 Finishing', '5': '5 MEP & Lifts',
      '6': '6 External Development', '7': '7 T&C & Handover'}

# ------------------------------------------------------------------ Resources sheet
wr = wb.create_sheet('Resources')
title(wr, 'Resource Sheet – labour, staff, equipment & material (MS Project Resource Sheet order)',
      'Blue = editable input rate. Work resources are priced per 8-h day (MSP "Std. Rate" = rate/d); material per material label. '
      'Max Units = peak concurrent requirement from the loaded schedule.')
RH = ['Resource Name', 'Type', 'Material Label', 'Initials', 'Group', 'Max Units (nos.)', 'Std. Rate (₹)', 'Rate basis',
      'Std. Rate (MSP text)', 'Base Calendar', 'Code', 'Total units x days / qty', 'Total Cost (₹)', 'Total Cost (₹ Cr)', 'Basis / source']
header(wr, 4, RH, [30, 10, 10, 9, 12, 12, 13, 14, 14, 12, 8, 16, 17, 11, 60])
RSRC = {'Labour': 'NCR 2024 all-in daily wage incl. PF/ESI/BOCW welfare & contractor burden (planning assumption)',
        'Staff': 'Cost-to-company per working day, NCR 2024 (planning assumption)',
        'Equipment': 'Hire rate per day incl. fuel & maintenance; operator loaded separately (planning assumption)',
        'Material': 'Basic supply rate ex-site, Noida/NCR 2024 market (planning assumption; installation in BOQ rate balance)'}
prefix = {'Labour': 'L', 'Staff': 'S', 'Equipment': 'E', 'Material': 'M'}
cnt = {}
RES_ROW = {}
r0 = 5
for k, (name, typ, lab, grp, rate) in enumerate(RES):
    r = r0 + k; RES_ROW[name] = r
    cnt[grp] = cnt.get(grp, 0) + 1
    ini = ''.join(w[0] for w in re.findall(r'[A-Za-z0-9]+', name))[:4].upper()
    pk = M['peak'].get(name, 1 if typ == 'Work' else None)
    vals = [name, typ, lab or None, ini, grp, (int(np.ceil(pk)) if typ == 'Work' else None), rate,
            'per day (8 h)' if typ == 'Work' else 'per ' + lab, None, 'Standard', '%s%02d' % (prefix[grp], cnt[grp]), None, None, None, RSRC[grp]]
    for j, v in enumerate(vals, 1):
        c = wr.cell(r, j, v); c.font = F(); c.border = BOX
    wr.cell(r, 7).font = BLUE; wr.cell(r, 7).number_format = INR; wr.cell(r, 6).number_format = '#,##0'
    wr.cell(r, 9, '=IF(B%d="Work",TEXT(G%d,"0")&"/d",TEXT(G%d,"0"))' % (r, r, r))
    wr.cell(r, 12).number_format = '#,##0.0'; wr.cell(r, 13).number_format = INR; wr.cell(r, 14).number_format = CR
RLAST = r0 + len(RES) - 1
wr.freeze_panes = 'B5'

# ------------------------------------------------------------------ Assignments sheet
wa = wb.create_sheet('Assignments')
title(wa, 'Resource Assignments – one row per task x resource (drives labour / equipment / material cost in L3_Schedule)',
      'Units = persons (work) or quantity in material label (material). Cost = Units x Duration x Rate for work; Units x Rate for material.')
header(wa, 4, ['Task ID', 'WBS', 'Resource', 'Class', 'Units / Qty', 'Duration (d)', 'Rate (₹)', 'Cost (₹)', 'Qty basis'],
       [8, 16, 30, 11, 12, 11, 12, 15, 40])
RGRP = {x[0]: x[3] for x in RES}
ABLK = {}
for _, t in L.iterrows():
    a0 = wa.max_row + 1 if wa.max_row >= 5 else 5
    for n, u in t.crew:
        wa.append([t.ID, t.WBS, n, None, u, t.dur, None, None, ('Crew composition (L3 col I) x gang multiplier %.2f' % t.kfac) if (t.kfac > 1 and RGRP[n] == 'Labour' and n != 'Equipment Operator') else 'Crew composition (L3 col I)'])
    for n, q in t.mats:
        wa.append([t.ID, t.WBS, n, None, round(q, 3), t.dur, None, None, 'BOQ quantity allocated to task (see BOQ_Map)'])
    ABLK[t.ID] = (a0, wa.max_row)
ALAST = wa.max_row
for r in range(5, ALAST + 1):
    wa.cell(r, 4, '=INDEX(Resources!$E$5:$E$%d,MATCH(C%d,Resources!$A$5:$A$%d,0))' % (RLAST, r, RLAST))
    wa.cell(r, 7, '=INDEX(Resources!$G$5:$G$%d,MATCH(C%d,Resources!$A$5:$A$%d,0))' % (RLAST, r, RLAST))
    wa.cell(r, 8, '=IF(D%d="Material",E%d*G%d,E%d*F%d*G%d)' % (r, r, r, r, r, r))
    wa.cell(r, 7).number_format = INR; wa.cell(r, 8).number_format = INR; wa.cell(r, 5).number_format = '#,##0.###'
    for j in range(1, 10): wa.cell(r, j).font = F()
wa.freeze_panes = 'A5'; wa.auto_filter.ref = 'A4:I%d' % ALAST
AR = lambda col: 'Assignments!$%s$5:$%s$%d' % (col, col, ALAST)
for name, r in RES_ROW.items():
    wr.cell(r, 12, '=IF(B{r}="Work",SUMPRODUCT(({c}=A{r})*{e}*{f}),SUMIFS({e},{c},A{r}))'.format(r=r, c=AR('C'), e=AR('E'), f=AR('F')))
    wr.cell(r, 13, '=SUMIFS(%s,%s,A%d)' % (AR('H'), AR('C'), r))
    wr.cell(r, 14, '=M%d/10^7' % r)
tr = RLAST + 1
wr.cell(tr, 1, 'TOTAL').font = BOLD
wr.cell(tr, 13, '=SUM(M5:M%d)' % RLAST).number_format = INR; wr.cell(tr, 14, '=M%d/10^7' % tr).number_format = CR
for j in (1, 13, 14): wr.cell(tr, j).fill = TOTFILL; wr.cell(tr, j).font = BOLD

# ------------------------------------------------------------------ L3_Schedule: new columns
NEW = ['Outline Level', 'Phase', 'Activity Code', 'Task Calendar (MSP)', 'Cost Source', 'Labour + Staff (₹)', 'Equipment (₹)',
       'Material (₹)', 'Budget (₹)', 'Fixed Cost – Subcontract / Other (₹)', 'Total Cost (₹)', 'Resource Names (MSP syntax)',
       'Total Float (d)', 'Critical (TF≤0)', 'Gang multiplier']
W = [8, 18, 11, 18, 18, 15, 13, 15, 16, 17, 16, 70, 9, 9, 9]
for j, (n, w) in enumerate(zip(NEW, W)):
    c = ws.cell(1, 12 + j, n); c.font = F(bold=True, color='FFFFFF'); c.fill = HFILL; c.border = BOX
    c.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
    ws.column_dimensions[CL(12 + j)].width = w
# column letters
cL, cM, cN, cO, cP, cQ, cR, cS, cT, cU, cV, cW, cX, cY, cZ = [CL(12 + j) for j in range(15)]
ids = [int(ws.cell(r, 1).value) for r in range(2, N + 1)]
OLs = [ol[i] for i in ids]
for idx, r in enumerate(range(2, N + 1)):
    i = ids[idx]; o = OLs[idx]
    ref = ws.cell(r, 3)
    ws.cell(r, 12, o)
    if o == 4:
        t = Lr.loc[i]
        ws.cell(r, 13, PH[t.phase]); ws.cell(r, 14, t.acode); ws.cell(r, 15, t.cal); ws.cell(r, 16, t.src)
        a0, a1 = ABLK[i]
        hh = 'Assignments!$H$%d:$H$%d' % (a0, a1); dd = 'Assignments!$D$%d:$D$%d' % (a0, a1)
        ws.cell(r, 17, '=SUMIFS({h},{d},"Labour")+SUMIFS({h},{d},"Staff")'.format(h=hh, d=dd))
        ws.cell(r, 18, '=SUMIFS({h},{d},"Equipment")'.format(h=hh, d=dd))
        ws.cell(r, 19, '=SUMIFS({h},{d},"Material")'.format(h=hh, d=dd))
        ws.cell(r, 20, round(float(t.budget), 0)).font = BLUE
        ws.cell(r, 21, '=%s%d-%s%d-%s%d-%s%d' % (cT, r, cQ, r, cR, r, cS, r))
        ws.cell(r, 22, '=%s%d+%s%d+%s%d+%s%d' % (cQ, r, cR, r, cS, r, cU, r))
        ws.cell(r, 23, t.resnames)
        ws.cell(r, 24, int(t.TF)); ws.cell(r, 25, 'Yes' if t.TF <= 0 else 'No'); ws.cell(r, 26, round(float(t.kfac), 2))
    else:
        # descendants are contiguous until next row with outline <= o
        end = idx + 1
        while end < len(ids) and OLs[end] > o: end += 1
        a, b = r + 1, end + 1
        for col in (17, 18, 19, 20, 21, 22):
            X = CL(col)
            ws.cell(r, col, '=SUMIFS(%s%d:%s%d,%s%d:%s%d,4)' % (X, a, X, b, cL, a, cL, b))
        if o == 2: ws.cell(r, 13, PH[ws.cell(r, 2).value])
    for col in range(12, 27):
        c = ws.cell(r, col); c.font = Font(name='Arial', size=10, bold=ref.font.b, color=ref.font.color.rgb if ref.font.color else None)
        c.fill = PatternFill('solid', fgColor=ref.fill.fgColor.rgb) if ref.fill and ref.fill.fgColor.rgb not in (None, '00000000') else PatternFill()
        c.border = BOX
    if o == 4: ws.cell(r, 20).font = Font(name='Arial', size=10, color='0000FF')
    for col in range(17, 23): ws.cell(r, col).number_format = INR
ws.freeze_panes = 'D2'
ws.auto_filter.ref = 'A1:%s%d' % (cZ, N)
ws.cell(1, 20).comment = Comment('Budget (blue input) = BOQ amount allocated to this activity via BOQ_Map bucket (equal split across '
                                 'the repeated floors / zones / towers) + Allowances + own resource cost where no BOQ item exists. '
                                 'Fixed Cost = Budget - Labour - Equipment - Material and is imported to MS Project "Fixed Cost".', 'Planner')
S = lambda col: "L3_Schedule!$%s$2:$%s$%d" % (col, col, N)

# ------------------------------------------------------------------ BOQ_Map
wbm = wb.create_sheet('BOQ_Map')
title(wbm, 'BOQ → WBS cost-bucket map (BOQ Golf Homes Phase-II 17.11.2024, DSR 2023 + NDSR; excl. GST)',
      'Each priced BOQ line is split by building column and mapped to a schedule activity code. SUB_RCC / SUB_WP spread over several '
      'substructure packages (weights at right). Amount = Qty x Rate.')
header(wbm, 4, ['BOQ row', 'Part', 'Section', 'Description', 'Unit', 'Building column', 'Qty', 'BOQ Rate (₹)', 'Amount (₹)',
                'WBS bucket (activity code)', 'Material resource', 'Qty→label factor', 'Amount (₹ Cr)'],
       [8, 6, 22, 60, 7, 15, 14, 13, 16, 14, 26, 9, 11])
for k, x in enumerate(bm.itertuples(), 5):
    vals = [x.row if x.row == x.row else None, x.part, x.sec, x.desc[:180], x.unit, x.col, round(x.qty, 3), round(x.rate, 2), None,
            x.bucket, x.mat if isinstance(x.mat, str) else None, x.matf if isinstance(x.mat, str) else None, None]
    for j, v in enumerate(vals, 1):
        c = wbm.cell(k, j, v); c.font = F(); c.border = BOX
    wbm.cell(k, 7).font = BLUE; wbm.cell(k, 8).font = BLUE
    wbm.cell(k, 9, '=G%d*H%d' % (k, k)).number_format = INR; wbm.cell(k, 13, '=I%d/10^7' % k).number_format = CR
    wbm.cell(k, 7).number_format = '#,##0.00'; wbm.cell(k, 8).number_format = '#,##0.00'
BLAST = 4 + len(bm)
t = BLAST + 2
rows_ = [('Total – civil BOQ (sheet "BOQ", structure + finishing line items)', '=SUMIFS(I5:I%d,B5:B%d,"STR")+SUMIFS(I5:I%d,B5:B%d,"FIN")' % (BLAST, BLAST, BLAST, BLAST)),
         ('Total – external development BOQ (sheet "BOQ EXTERNAL")', '=SUMIFS(I5:I%d,B5:B%d,"EXT")' % (BLAST, BLAST)),
         ('BOQ grand total as printed in source (G665)', 13539780035.19),
         ('Difference (line items vs printed total)', '=I%d-I%d' % (t, t + 2))]
for k, (lab, f) in enumerate(rows_):
    wbm.cell(t + k, 4, lab).font = BOLD; c = wbm.cell(t + k, 9, f); c.number_format = INR; c.font = BOLD
    wbm.cell(t + k, 13, '=I%d/10^7' % (t + k)).number_format = CR
wbm.cell(t + 2, 9).font = BLUE
wbm.cell(t + 4, 4, 'Note: the difference is BOQ row 647 "Providing and fixing polycarbonate sheets" (₹0.94 Cr), which lies outside the '
         'source NDSR subtotal range =SUM(G457:G645) and so is missing from the printed total. The plan carries every priced line (₹1,354.91 Cr).').font = F(italic=True, color='C00000')
wbm.cell(4, 15, 'Bucket').font = BOLD; wbm.cell(4, 16, 'Package').font = BOLD; wbm.cell(4, 17, 'Weight').font = BOLD
k = 5
for b, wts in (('SUB_RCC', M['SUB_RCC_W']), ('SUB_WP', M['SUB_WP_W'])):
    for p, w in wts.items():
        wbm.cell(k, 15, b); wbm.cell(k, 16, p); c = wbm.cell(k, 17, w); c.font = BLUE; c.number_format = '0%'; k += 1
wbm.freeze_panes = 'E5'; wbm.auto_filter.ref = 'A4:M%d' % BLAST

# ------------------------------------------------------------------ Allowances
wal = wb.create_sheet('Allowances')
title(wal, 'Allowances for scope NOT priced in the supplied BOQ (MEP, lifts, external services, preliminaries)',
      'Planning allowances – replace with the MEP / specialist BOQ when issued. Blue = input.')
header(wal, 4, ['Code', 'Description', 'Basis', 'Quantity', 'Rate (₹)', 'Amount (₹)', 'Amount (₹ Cr)', 'Loaded on activity codes (share)', 'Rationale'],
       [7, 58, 12, 13, 13, 16, 11, 34, 60])
RAT = {'AL-E': 'High-rise residential internal electrical ≈ ₹130/sqft BUA', 'AL-P': 'Plumbing + sanitaryware ≈ ₹85/sqft',
       'AL-H': 'Ventilation only (no central AC) ≈ ₹28/sqft', 'AL-F': 'Wet-riser, sprinklers in basements, FA/PA ≈ ₹42/sqft',
       'AL-L': '8 cars per tower (schedule col. C) x 11 sanctioned towers @ ₹40 L per 2.5 m/s car',
       'AL-S': '~12 MVA load for 1,507 DUs + amenities; 33/11 kV substation, 6 x 2,000 kVA DG',
       'AL-T': 'MBBR STP sized for ~6,800 added residents (assumed)', 'AL-W': 'Softening + domestic/flushing pump houses', 'AL-R': 'RWH pits + fire/domestic UGTs',
       'AL-C': 'Lump-sum allowance – boundary wall, gates, cabins', 'AL-D': 'Lump-sum allowance – SW drains to outfall', 'AL-X': 'Street, landscape & facade lighting',
       'AL-1': 'Site offices, stores, BOCW-compliant camp for ~4,000 workers', 'AL-2': 'Temporary DG/grid connection, borewells, site roads',
       'AL-3': '2 x 60 cum/h batching plants + yard', 'AL-4': '12 tower-crane rafts, erection & dismantling',
       'AL-5': 'Hoarding, fencing, grubbing'}
for k, (code, desc, basis, q, rate, mp) in enumerate(ALLOW, 5):
    vals = [code, desc, basis, q, rate, '=D%d*E%d' % (k, k), '=F%d/10^7' % k, ', '.join('%s (%d%%)' % (a, s * 100) for a, s in mp.items()), RAT.get(code, '')]
    for j, v in enumerate(vals, 1):
        c = wal.cell(k, j, v); c.font = F(); c.border = BOX
    wal.cell(k, 4).font = BLUE; wal.cell(k, 5).font = BLUE
    wal.cell(k, 4).number_format = '#,##0.00'; wal.cell(k, 5).number_format = INR; wal.cell(k, 6).number_format = INR; wal.cell(k, 7).number_format = CR
ALL_LAST = 4 + len(ALLOW)
wal.cell(ALL_LAST + 1, 2, 'TOTAL ALLOWANCES').font = BOLD
wal.cell(ALL_LAST + 1, 6, '=SUM(F5:F%d)' % ALL_LAST).number_format = INR; wal.cell(ALL_LAST + 1, 7, '=F%d/10^7' % (ALL_LAST + 1)).number_format = CR
wal.cell(ALL_LAST + 3, 2, 'BUA basis for ₹/sqm allowances = 558,449.55 sqm (BOQ "Total Built up Area", row 690).').font = F(italic=True)

# ------------------------------------------------------------------ Rate_Norms
wn = wb.create_sheet('Rate_Norms')
title(wn, 'Resource-loading norms used to convert BOQ value into crew size (labour content of DSR rates)',
      'Gang multiplier = min(5, Budget x Labour share ÷ crew labour cost at 1 gang). Labour content beyond 5 gangs remains in Fixed Cost '
      '(piece-rate / labour-supply subcontract).')
header(wn, 4, ['Activity code', 'Labour share of BOQ rate', 'Typical activity', 'Tasks', 'Avg gang multiplier'], [14, 14, 60, 8, 12])
LS = {}
for a, g in L.groupby('acode'):
    if g.LSHARE.notna().any(): LS[a] = (g.LSHARE.iloc[0], g['Task Name'].iloc[0][:80], len(g), g.kfac.mean())
for k, (a, (sh, nm, n, kf)) in enumerate(sorted(LS.items()), 5):
    for j, v in enumerate([a, sh, nm, n, round(kf, 2)], 1):
        c = wn.cell(k, j, v); c.font = F(); c.border = BOX
    wn.cell(k, 2).number_format = '0%'; wn.cell(k, 2).font = BLUE
k += 2
wn.cell(k, 1, 'Gang multiplier cap').font = BOLD; wn.cell(k, 2, 5).font = BLUE
wn.cell(k + 1, 1, 'Time-related plant added: tower crane + operator on floor-cycle steps (a) rebar and (e) deshuttering; material hoist on finishing trades t1, t2, t4, t6.').font = F(italic=True)
wn.cell(k + 2, 1, 'Secondary floor-cycle steps (c) QA/embeds and (e) curing/deshuttering carry only their own resource cost; it is netted out of the '
        'concrete (d) and MIVAN formwork (b) BOQ values so the BOQ is not double-counted.').font = F(italic=True)

# ------------------------------------------------------------------ Cost_Summary
wc = wb.create_sheet('Cost_Summary')
title(wc, 'Cost Plan – Budget at Completion (BAC) by phase and by resource class', 'All figures ₹ Crore unless noted; formulas read L3_Schedule leaf rows.')
header(wc, 4, ['Phase', 'Labour + Staff', 'Equipment', 'Material', 'Subcontract / Other (Fixed)', 'Total (BAC)', '% of BAC', 'Leaf activities'],
       [26, 14, 12, 12, 16, 14, 10, 12])
for k, p in enumerate(PH.values(), 5):
    wc.cell(k, 1, p)
    for j, col in enumerate([cQ, cR, cS, cU, cV], 2):
        wc.cell(k, j, '=SUMIFS(%s,%s,$A%d,%s,4)/10^7' % (S(col), S(cM), k, S(cL))).number_format = CR
    wc.cell(k, 7, '=F%d/$F$12' % k).number_format = PCT
    wc.cell(k, 8, '=COUNTIFS(%s,$A%d,%s,4)' % (S(cM), k, S(cL)))
wc.cell(12, 1, 'TOTAL – BAC')
for j in range(2, 7): wc.cell(12, j, '=SUM(%s5:%s11)' % (CL(j), CL(j))).number_format = CR
wc.cell(12, 7, '=F12/F12').number_format = PCT; wc.cell(12, 8, '=SUM(H5:H11)')
wc.cell(13, 1, 'Share of BAC'); [wc.cell(13, j, '=%s12/$F$12' % CL(j)) for j in range(2, 7)]
for j in range(2, 7): wc.cell(13, j).number_format = PCT
for r in (12, 13):
    for j in range(1, 9): wc.cell(r, j).font = BOLD; wc.cell(r, j).fill = TOTFILL
wc['A15'] = 'Key indicators'; wc['A15'].font = F(bold=True, size=12, color=NAVY)
KI = [('Built-up area (sqm) – BOQ row 690', 558449.55, '#,##0'), ('Built-up area (sqft)', '=B16*10.7639', '#,##0'),
      ('BAC (₹ Cr)', '=F12', CR), ('Cost per sqft BUA (₹)', '=B18*10^7/B17', INR), ('Cost per sqm BUA (₹)', '=B18*10^7/B16', INR),
      ('Management reserve % (outside BAC)', 0.05, '0%'), ('Management reserve (₹ Cr)', '=B18*B21', CR), ('Dwelling units (area statement)', 1507, '#,##0'),
      ('Cost per DU (₹ lakh)', '=B18*100/B23', CR)]
for k, (lab, v, fmt) in enumerate(KI, 16):
    wc.cell(k, 1, lab); c = wc.cell(k, 2, v); c.number_format = fmt
    if not str(v).startswith('='): c.font = BLUE
wc['D15'] = 'BAC reconciliation by source'; wc['D15'].font = F(bold=True, size=12, color=NAVY)
REC = [('Civil BOQ line items', '=(SUMIFS(BOQ_Map!I:I,BOQ_Map!B:B,"STR")+SUMIFS(BOQ_Map!I:I,BOQ_Map!B:B,"FIN"))/10^7'),
       ('External development BOQ', '=SUMIFS(BOQ_Map!I:I,BOQ_Map!B:B,"EXT")/10^7'),
       ('Allowances (MEP, lifts, services, prelims)', '=Allowances!F%d/10^7' % (ALL_LAST + 1)),
       ('Resource cost of activities with no BOQ item (T&C, QA, dewatering, statutory, survey)', None),
       ('Crew cost above allocated BOQ value (resource-driven uplift)', None),
       ('Total', '=SUM(E16:E20)'), ('Check vs BAC (should be 0)', '=E21-F12')]
unb = L[L.src == 'Resource cost'].total.sum() / 1e7
upl = L.uplift.sum() / 1e7
for k, (lab, v) in enumerate(REC, 16):
    wc.cell(k, 4, lab)
    if v is None:
        v = round(unb if 'no BOQ' in lab else upl, 4); c = wc.cell(k, 5, v); c.font = BLUE
        c.comment = Comment('Computed by the allocation model (sum over the relevant L3 rows).', 'Planner')
    else: c = wc.cell(k, 5, v)
    c.number_format = CR
wc.cell(21, 4).font = BOLD; wc.cell(21, 5).font = BOLD
wc.column_dimensions['D'].width = 58
wc['A27'] = 'Presentation split of Subcontract / Other (Fixed Cost)'; wc['A27'].font = F(bold=True, size=12, color=NAVY)
wc['A28'] = 'Contractor OH&P embedded in DSR/NDSR rates (15/115 of BOQ value)'; wc['B28'] = '=(E16+E17)*15/115'
wc['A29'] = 'Specialist subcontract packages, formwork system, balance of rates'; wc['B29'] = '=E12-B28'
for c in ('B28', 'B29'): wc[c].number_format = CR
bar = BarChart(); bar.type = 'bar'; bar.style = 10; bar.title = 'BAC by phase (₹ Cr)'; bar.y_axis.title = '₹ Cr'
bar.add_data(Reference(wc, min_col=6, min_row=4, max_row=11), titles_from_data=True)
bar.set_categories(Reference(wc, min_col=1, min_row=5, max_row=11)); bar.height = 8; bar.width = 16; bar.legend = None
wc.add_chart(bar, 'J4')

# ------------------------------------------------------------------ CashFlow
wf = wb.create_sheet('CashFlow_S-Curve')
title(wf, 'Monthly cash flow (Planned Value) & cumulative S-curve – baseline time-phased at task level',
      'Task cost spread over its working days using the GRAP calendars (cost stops in GRAP-A windows, 60% rate in GRAP-B) – the same '
      'prorated accrual MS Project applies. Phase columns are exported values; totals are formulas.')
cols = ['Month'] + list(PH.values()) + ['Monthly PV (₹ Cr)', 'Cumulative PV (₹ Cr)', 'Cumulative % of BAC']
header(wf, 4, cols, [10] + [13] * 7 + [13, 14, 12], 44)
mph = M['m_ph']
for k, per in enumerate(M['months'], 5):
    c = wf.cell(k, 1, per.to_timestamp().to_pydatetime()); c.number_format = 'mmm-yy'
    for j, p in enumerate(PH, 2):
        v = float(mph[p].get(per, 0)) / 1e7 if p in mph else 0
        cc = wf.cell(k, j, round(v, 4)); cc.number_format = CR
    wf.cell(k, 9, '=SUM(B%d:H%d)' % (k, k)).number_format = CR
    wf.cell(k, 10, ('=J%d+I%d' % (k - 1, k)) if k > 5 else '=I5').number_format = CR
    wf.cell(k, 11, '=J%d/Cost_Summary!$F$12' % k).number_format = PCT
FL = 4 + len(M['months'])
wf.cell(FL + 1, 1, 'Total').font = BOLD
for j in range(2, 10): wf.cell(FL + 1, j, '=SUM(%s5:%s%d)' % (CL(j), CL(j), FL)).number_format = CR
b1 = BarChart(); b1.title = 'Monthly PV & cumulative S-curve'; b1.y_axis.title = 'Monthly ₹ Cr'
b1.add_data(Reference(wf, min_col=9, min_row=4, max_row=FL), titles_from_data=True)
b1.set_categories(Reference(wf, min_col=1, min_row=5, max_row=FL))
l1 = LineChart(); l1.add_data(Reference(wf, min_col=10, min_row=4, max_row=FL), titles_from_data=True)
l1.y_axis.axId = 200; l1.y_axis.title = 'Cumulative ₹ Cr'; l1.y_axis.crosses = 'max'
b1 += l1; b1.height = 9; b1.width = 22
wf.add_chart(b1, 'M4')

# ------------------------------------------------------------------ Labour histogram
wh = wb.create_sheet('Labour_Histogram')
title(wh, 'Site labour histogram (labour + staff headcount) – from the loaded schedule',
      'Average and peak daily headcount per month = Σ assigned units on working days. Regenerate in MSP: Resource Usage / Resource Graph.')
header(wh, 4, ['Month', 'Avg daily headcount', 'Peak daily headcount', 'Worker-days'], [10, 14, 14, 14])
days = M['days']; head = M['head']
didx = pd.PeriodIndex([pd.Period(d, 'M') for d in days])
wd = pd.Series(head, index=didx).groupby(level=0).sum()
for k, per in enumerate(M['months'], 5):
    c = wh.cell(k, 1, per.to_timestamp().to_pydatetime()); c.number_format = 'mmm-yy'
    wh.cell(k, 2, int(round(M['m_head'][per]))).number_format = '#,##0'
    wh.cell(k, 3, int(round(M['m_headpeak'][per]))).number_format = '#,##0'
    wh.cell(k, 4, int(round(wd.get(per, 0)))).number_format = '#,##0'
HL = 4 + len(M['months'])
wh.cell(HL + 1, 1, 'Peak / total').font = BOLD
wh.cell(HL + 1, 2, '=MAX(B5:B%d)' % HL); wh.cell(HL + 1, 3, '=MAX(C5:C%d)' % HL); wh.cell(HL + 1, 4, '=SUM(D5:D%d)' % HL)
wh.cell(HL + 2, 1, 'Worker-months (÷26)'); wh.cell(HL + 2, 4, '=D%d/26' % (HL + 1))
for j in (2, 3, 4): wh.cell(HL + 1, j).number_format = '#,##0'
wh.cell(HL + 2, 4).number_format = '#,##0'
b2 = BarChart(); b2.title = 'Average daily site headcount by month'; b2.add_data(Reference(wh, min_col=2, min_row=4, max_row=HL), titles_from_data=True)
b2.set_categories(Reference(wh, min_col=1, min_row=5, max_row=HL)); b2.legend = None; b2.height = 9; b2.width = 22
wh.add_chart(b2, 'G4')

# ------------------------------------------------------------------ EVM
we = wb.create_sheet('EVM')
title(we, 'Earned Value Management – status 30-Apr-2025 (month 16 of 29) – ILLUSTRATIVE progress data',
      'Blue cells are status inputs (physical % complete and actual cost to date per phase) – replace with site data each month. '
      'PV reads the baseline cash flow up to the status month.')
we['A3'] = 'Status month'; we['B3'] = dt.datetime(2025, 4, 1); we['B3'].number_format = 'mmm-yy'; we['B3'].font = BLUE
header(we, 5, ['Phase', 'BAC (₹ Cr)', 'PV (₹ Cr)', '% complete (physical)', 'EV (₹ Cr)', 'AC (₹ Cr)', 'SV = EV-PV', 'CV = EV-AC',
               'SPI = EV/PV', 'CPI = EV/AC', 'EAC = BAC/CPI', 'ETC = EAC-AC', 'VAC = BAC-EAC'], [24, 11, 11, 12, 11, 11, 11, 11, 10, 10, 12, 12, 12])
# illustrative status: PV at status per phase from model
st = pd.Period('2025-04', 'M')
pv_ph = {p: float(mph[p].loc[:st].sum()) / 1e7 for p in PH}
bac_ph = {p: float(L[L.phase == p].total.sum()) / 1e7 for p in PH}
SPIx = {'1': 1.0, '2': 1.0, '3': 0.90, '4': 0.84, '5': 0.88, '6': 0.95, '7': 1.0}
CPIx = {'1': 0.98, '2': 0.97, '3': 0.95, '4': 0.98, '5': 1.00, '6': 0.99, '7': 1.0}
EVPH = {}
for k, p in enumerate(PH, 6):
    ev = pv_ph[p] * SPIx[p]; pc = ev / bac_ph[p] if bac_ph[p] else 0
    ac = ev / CPIx[p]; EVPH[p] = (pc, ac)
    we.cell(k, 1, PH[p])
    we.cell(k, 2, '=SUMIFS(%s,%s,A%d,%s,4)/10^7' % (S(cV), S(cM), k, S(cL)))
    colL = CL(2 + list(PH).index(p))
    we.cell(k, 3, "=SUMIFS('CashFlow_S-Curve'!%s$5:%s$%d,'CashFlow_S-Curve'!$A$5:$A$%d,\"<=\"&$B$3)" % (colL, colL, FL, FL))
    c = we.cell(k, 4, round(pc, 4)); c.font = BLUE; c.number_format = PCT
    we.cell(k, 5, '=B%d*D%d' % (k, k))
    c = we.cell(k, 6, round(ac, 2)); c.font = BLUE
    we.cell(k, 7, '=E%d-C%d' % (k, k)); we.cell(k, 8, '=E%d-F%d' % (k, k))
    we.cell(k, 9, '=IFERROR(E%d/C%d,0)' % (k, k)); we.cell(k, 10, '=IFERROR(E%d/F%d,0)' % (k, k))
    we.cell(k, 11, '=IF(J%d>0,B%d/J%d,B%d)' % (k, k, k, k)); we.cell(k, 12, '=K%d-F%d' % (k, k)); we.cell(k, 13, '=B%d-K%d' % (k, k))
    for j in (2, 3, 5, 6, 7, 8, 11, 12, 13): we.cell(k, j).number_format = CR
    for j in (9, 10): we.cell(k, j).number_format = '0.00'
tr = 13
we.cell(tr, 1, 'PROJECT TOTAL')
for j in (2, 3, 5, 6, 7, 8): we.cell(tr, j, '=SUM(%s6:%s12)' % (CL(j), CL(j)))
we.cell(tr, 4, '=E13/B13'); we.cell(tr, 9, '=E13/C13'); we.cell(tr, 10, '=E13/F13'); we.cell(tr, 11, '=B13/J13')
we.cell(tr, 12, '=K13-F13'); we.cell(tr, 13, '=B13-K13')
for j in range(1, 14): we.cell(tr, j).font = BOLD; we.cell(tr, j).fill = TOTFILL
for j in (2, 3, 5, 6, 7, 8, 11, 12, 13): we.cell(tr, j).number_format = CR
we.cell(tr, 4).number_format = PCT; we.cell(tr, 9).number_format = '0.00'; we.cell(tr, 10).number_format = '0.00'
EXTRA = [('TCPI to achieve BAC = (BAC-EV)/(BAC-AC)', '=(B13-E13)/(B13-F13)', '0.00'),
         ('EAC (CPI x SPI method) = AC + (BAC-EV)/(CPI x SPI)', '=F13+(B13-E13)/(J13*I13)', CR),
         ('Schedule slip indicator ≈ (1-SPI) x elapsed months', '=(1-I13)*16', '0.0'),
         ('Management reserve (5% of BAC)', '=Cost_Summary!B22', CR)]
for k, (lab, f, fmt) in enumerate(EXTRA, 15):
    we.cell(k, 1, lab); c = we.cell(k, 5, f); c.number_format = fmt
we.column_dimensions['A'].width = 46
we['A20'] = 'Illustrative status narrative (for the EVM demonstration – replace with actuals):'; we['A20'].font = BOLD
we['A21'] = ('• GRAP Stage-III/IV in winter 2024-25 ran ~12 days beyond the modelled 15-Jan window → tower batch-3 (T11/T12) 3–4 floors behind; '
             'finishing & MEP starts on batch-1 towers slipped with it.')
we['A22'] = '• Fe-500D rebar landed ~5% above the ₹62,000/t planning rate from Nov-24; RMC diesel surcharge on substructure pours.'
we['A23'] = '• Pre-construction and substructure are physically complete; external development is marginally behind (drainage Zone-3).'
# monthly EVM curve (values, illustrative)
we['A26'] = 'Monthly cumulative curves (₹ Cr) – PV baseline; EV/AC illustrative to status; forecast to EAC'; we['A26'].font = BOLD
header(we, 27, ['Month', 'PV cum', 'EV cum', 'AC cum', 'Forecast (EAC)'], None)
cum = (M['m_cost'].cumsum() / 1e7)
spi_end = sum(pv_ph[p] * SPIx[p] for p in PH) / sum(pv_ph.values())
ac_end = sum(EVPH[p][1] for p in PH); ev_end = sum(pv_ph[p] * SPIx[p] for p in PH)
cpi_end = ev_end / ac_end
bac = L.total.sum() / 1e7; eac = bac / cpi_end
months = list(M['months']); si = months.index(st)
EVM_ROWS = []
for k, per in enumerate(months, 28):
    pv = float(cum[per]); c = we.cell(k, 1, per.to_timestamp().to_pydatetime()); c.number_format = 'mmm-yy'
    we.cell(k, 2, round(pv, 2))
    if k - 28 <= si:
        f = (k - 28) / si
        s_ = 1 - (1 - spi_end) * f ** 1.5; c_ = 1 - (1 - cpi_end) * f ** 1.5
        ev = pv * s_; ac = ev / c_
        we.cell(k, 3, round(ev, 2)); we.cell(k, 4, round(ac, 2)); EVM_ROWS.append((per, pv, ev, ac, None))
        if k - 28 == si: we.cell(k, 5, round(ac, 2)); last_ac = ac; last_pv = pv
    else:
        # remaining work re-spread on the baseline shape, shifted by the slip and scaled to EAC
        frac = (pv - last_pv) / (bac - last_pv)
        fc = last_ac + frac * (eac - last_ac)
        we.cell(k, 5, round(fc, 2)); EVM_ROWS.append((per, pv, None, None, fc))
EL = 27 + len(months)
lc = LineChart(); lc.title = 'PV / EV / AC & forecast (₹ Cr, cumulative)'
lc.add_data(Reference(we, min_col=2, min_row=27, max_col=5, max_row=EL), titles_from_data=True)
lc.set_categories(Reference(we, min_col=1, min_row=28, max_row=EL)); lc.height = 9; lc.width = 20
we.add_chart(lc, 'H20')

# ------------------------------------------------------------------ MSP import sheets
wi = wb.create_sheet('MSP_Import_Tasks')
TH = ['ID (check only)', 'Outline Level', 'Name', 'Duration', 'Predecessors', 'Resource Names', 'Fixed Cost', 'Fixed Cost Accrual',
      'Task Calendar', 'Text1 (WBS)', 'Text2 (GRAP category)', 'Text3 (Crew composition)', 'Text4 (Activity code)', 'Number1 (Budget ₹)']
header(wi, 1, TH, [8, 8, 60, 10, 18, 70, 15, 10, 18, 14, 22, 30, 10, 15])
for k, r in enumerate(range(2, N + 1), 2):
    o = OLs[k - 2]
    wi.cell(k, 1, '=L3_Schedule!A%d' % r); wi.cell(k, 2, '=L3_Schedule!%s%d' % (cL, r)); wi.cell(k, 3, '=L3_Schedule!C%d' % r)
    if o == 4:
        wi.cell(k, 4, '=L3_Schedule!D%d' % r)
        wi.cell(k, 5, '=IF(L3_Schedule!G%d="","",L3_Schedule!G%d)' % (r, r))
        wi.cell(k, 6, '=L3_Schedule!%s%d' % (cW, r))
        wi.cell(k, 7, '=ROUND(L3_Schedule!%s%d,0)' % (cU, r)); wi.cell(k, 7).number_format = '0'
        wi.cell(k, 8, 'Prorated'); wi.cell(k, 9, '=L3_Schedule!%s%d' % (cO, r))
        wi.cell(k, 11, '=L3_Schedule!K%d' % r); wi.cell(k, 12, '=L3_Schedule!I%d' % r); wi.cell(k, 13, '=L3_Schedule!%s%d' % (cN, r))
        wi.cell(k, 14, '=ROUND(L3_Schedule!%s%d,0)' % (cT, r)); wi.cell(k, 14).number_format = '0'
    wi.cell(k, 10, '=L3_Schedule!B%d' % r)
    for j in range(1, 15): wi.cell(k, j).font = F()
wi.freeze_panes = 'D2'

wir = wb.create_sheet('MSP_Import_Resources')
header(wir, 1, ['Resource Name', 'Type', 'Material Label', 'Initials', 'Group', 'Max. Units', 'Std. Rate', 'Ovt. Rate', 'Cost/Use', 'Accrue At',
                'Base Calendar', 'Code'], [30, 9, 10, 9, 11, 10, 12, 10, 9, 10, 12, 8])
for k in range(len(RES)):
    r = 5 + k; q = 2 + k
    wir.cell(q, 1, '=Resources!A%d' % r); wir.cell(q, 2, '=Resources!B%d' % r); wir.cell(q, 3, '=IF(Resources!C%d="","",Resources!C%d)' % (r, r))
    wir.cell(q, 4, '=Resources!D%d' % r); wir.cell(q, 5, '=Resources!E%d' % r)
    wir.cell(q, 6, '=IF(Resources!B%d="Work",TEXT(Resources!F%d*100,"0")&"%%","")' % (r, r))
    wir.cell(q, 7, '=Resources!I%d' % r); wir.cell(q, 8, '0'); wir.cell(q, 9, '0'); wir.cell(q, 10, 'Prorated')
    wir.cell(q, 11, '=IF(Resources!B%d="Work","Standard","")' % r); wir.cell(q, 12, '=Resources!K%d' % r)
    for j in range(1, 13): wir.cell(q, j).font = F()

# ------------------------------------------------------------------ MSP guide sheet
wg = wb.create_sheet('MSP_Link_Guide')
GUIDE = open('msp_guide.txt', encoding='utf-8').read().replace('{NRES}', str(len(RES))).replace('{RL}', str(len(RES) + 1)).splitlines()
open('msp_guide_final.txt', 'w', encoding='utf-8').write('\n'.join(GUIDE))
wg.column_dimensions['A'].width = 6; wg.column_dimensions['B'].width = 150
for k, line in enumerate(GUIDE, 1):
    if line.startswith('# '):
        c = wg.cell(k, 1, line[2:]); c.font = F(bold=True, size=14, color=NAVY)
    elif line.startswith('## '):
        c = wg.cell(k, 1, line[3:]); c.font = F(bold=True, size=12, color='FFFFFF')
        for j in (1, 2): wg.cell(k, j).fill = HFILL
    elif re.match(r'^\d+[a-z]?\.', line):
        n, rest = line.split(' ', 1); wg.cell(k, 1, n).font = F(bold=True); c = wg.cell(k, 2, rest); c.font = F()
        c.alignment = Alignment(wrap_text=True, vertical='top')
    elif line.startswith('   '):
        c = wg.cell(k, 2, line.strip()); c.font = F(color='404040'); c.alignment = Alignment(wrap_text=True, indent=2)
    else:
        c = wg.cell(k, 2, line); c.font = F(italic=bool(line.startswith('Note') or line.startswith('Tip'))); c.alignment = Alignment(wrap_text=True)

# ------------------------------------------------------------------ READ_ME
wrm = wb.create_sheet('READ_ME', 0)
title(wrm, 'Golf Homes Phase-II (GH-02, Sector 4, Greater Noida) – L3 schedule, RESOURCE & COST LOADED (R1)',
      'Adds resources, BOQ-based costing, cash flow, labour histogram, EVM and MS Project import sheets to the L3 GRAP schedule. Prepared 27-Sep-2026.')
IDX = [('L3_Schedule', 'Original 2,758-row L3 schedule (cols A–K untouched) + cols L–Z: outline level, phase, activity code, MSP task calendar, '
                       'labour/equipment/material/fixed/total cost, MSP Resource Names string, total float, critical flag.'),
       ('Assumptions_Basis', 'Original basis-of-schedule notes (unchanged).'),
       ('Resources', '%d resources: %d labour trades, %d staff, %d equipment, %d materials – rates (blue) drive every cost.' % (
           len(RES), sum(r[3] == 'Labour' for r in RES), sum(r[3] == 'Staff' for r in RES), sum(r[3] == 'Equipment' for r in RES), sum(r[3] == 'Material' for r in RES))),
       ('Assignments', '%d task-resource assignments (units x duration x rate).' % (ALAST - 4)),
       ('BOQ_Map', 'Every priced BOQ line split by building column and mapped to a WBS activity code (the cost buckets).'),
       ('Allowances', 'MEP, lifts, external services & preliminaries not priced in the BOQ – planning allowances (blue).'),
       ('Rate_Norms', 'Labour share of BOQ rates used to size gangs; secondary-step and plant rules.'),
       ('Cost_Summary', 'BAC by phase x resource class, ₹/sqft, reconciliation BOQ → BAC.'),
       ('CashFlow_S-Curve', 'Monthly PV by phase and cumulative S-curve.'),
       ('Labour_Histogram', 'Average / peak daily headcount by month.'),
       ('EVM', 'PV/EV/AC, SPI, CPI, EAC, ETC, VAC, TCPI at 30-Apr-2025 (illustrative status inputs in blue).'),
       ('MSP_Import_Tasks', 'Clean task table for the MS Project Import Wizard.'),
       ('MSP_Import_Resources', 'Clean resource table for the Import Wizard / paste into Resource Sheet.'),
       ('MSP_Link_Guide', 'Step-by-step: calendars, import, verification, baseline, critical path, cash flow, EVM in MS Project.')]
header(wrm, 4, ['Sheet', 'Contents'], [24, 140])
for k, (a, b) in enumerate(IDX, 5):
    wrm.cell(k, 1, a).font = BOLD; wrm.cell(k, 2, b).font = F()
k += 2
wrm.cell(k, 1, 'Headline figures').font = F(bold=True, size=12, color=NAVY)
HL_ = [('BAC (₹ Cr)', '=Cost_Summary!F12', CR), ('Cost / sqft BUA (₹)', '=Cost_Summary!B19', INR), ('Peak avg daily headcount', '=Labour_Histogram!B%d' % (HL + 1), '#,##0'),
       ('Leaf activities / assignments', '=COUNTIFS(L3_Schedule!L:L,4)&" / "&COUNTA(Assignments!A5:A%d)' % ALAST, '@'),
       ('Schedule', '01-Jan-2024 → 14-May-2026 (865 days), 7 zero-float activities; driving path 197 activities', '@')]
for j, (a, f, fmt) in enumerate(HL_, k + 1):
    wrm.cell(j, 1, a).font = F(); c = wrm.cell(j, 2, f); c.number_format = fmt; c.font = BOLD; c.alignment = Alignment(horizontal='left')
k = j + 2
wrm.cell(k, 1, 'Colour code').font = F(bold=True, size=12, color=NAVY)
wrm.cell(k + 1, 1, 'Blue text').font = BLUE; wrm.cell(k + 1, 2, 'Hard-coded input / assumption – edit these').font = F()
wrm.cell(k + 2, 1, 'Black text').font = F(); wrm.cell(k + 2, 2, 'Formula – do not overwrite').font = F()
wrm.cell(k + 4, 1, 'Data caveats').font = F(bold=True, size=12, color='C00000')
CAV = ['The L3 schedule models 12 tower-units (T01–T12, incl. "GODAVARI-2"); the BOQ and approved area statement sanction 11 towers (2B+S+G/P+35). '
       'BOQ tower quantities are spread over the 12 schedule units – reconcile the tower list before freezing the baseline.',
       'The supplied BOQ covers civil structure, finishes and external development only. MEP, lifts, substation/DG, STP/WTP and preliminaries are '
       'carried as allowances (sheet Allowances) and must be replaced by the specialist BOQs.',
       'All BOQ amounts are at DSR 2023 / NDSR rates excluding GST; the plan is a construction-cost baseline (no land, fees, approvals or GST).',
       'EVM progress (30-Apr-2025) is illustrative, to demonstrate the method – overwrite the blue inputs with actual measurements.']
for j, t_ in enumerate(CAV, k + 5):
    c = wrm.cell(j, 2, '• ' + t_); c.font = F(); c.alignment = Alignment(wrap_text=True); wrm.row_dimensions[j].height = 28

# order sheets
order = ['READ_ME', 'L3_Schedule', 'Assumptions_Basis', 'Resources', 'Assignments', 'BOQ_Map', 'Allowances', 'Rate_Norms', 'Cost_Summary',
         'CashFlow_S-Curve', 'Labour_Histogram', 'EVM', 'MSP_Import_Tasks', 'MSP_Import_Resources', 'MSP_Link_Guide']
wb._sheets = [wb[n] for n in order]
wb.active = 0
wb.save(OUT)
pickle.dump(dict(pv_ph=pv_ph, bac_ph=bac_ph, SPIx=SPIx, CPIx=CPIx, EVPH=EVPH, EVM_ROWS=EVM_ROWS, eac=eac, cpi=cpi_end, spi=spi_end,
                 ev=ev_end, ac=ac_end, bac=bac), open('evm.pkl', 'wb'))
print('saved', OUT, 'assignments', ALAST - 4, 'EAC', eac, 'CPI', cpi_end, 'SPI', spi_end)
