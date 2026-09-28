import xlrd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as CL
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.comments import Comment
from openpyxl.chart import BarChart, Reference

SRC_FILE = "Spire-L.A-_250926.xls"
src = xlrd.open_workbook("src.xls")
sA = src.sheet_by_name("ABS-LA bldg")
sM = src.sheet_by_name("Measurements")

F = "Arial"
f_base = Font(name=F, size=10)
f_b = Font(name=F, size=10, bold=True)
f_title = Font(name=F, size=14, bold=True)
f_blue = Font(name=F, size=10, color="0000FF")
f_green = Font(name=F, size=10, color="008000")
f_link = Font(name=F, size=10, color="0563C1", underline="single")
f_hdr = Font(name=F, size=10, bold=True, color="FFFFFF")
fill_hdr = PatternFill("solid", fgColor="1F4E78")
fill_sub = PatternFill("solid", fgColor="DDEBF7")
fill_yel = PatternFill("solid", fgColor="FFFF00")
fill_red = PatternFill("solid", fgColor="FCE4D6")
fill_tot = PatternFill("solid", fgColor="E2EFDA")
thin = Side(style="thin", color="A6A6A6")
box = Border(left=thin, right=thin, top=thin, bottom=thin)
wrap = Alignment(wrap_text=True, vertical="top")
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
NUM = '#,##0;(#,##0);"-"'
NUM2 = '#,##0.00;(#,##0.00);"-"'
PCT = '0.0%;(0.0%);"-"'

wb = Workbook()

# ---------------------------------------------------------------- bands
BANDS = [("Basement", -10.0, 0.0)] + [(f"{lo:.0f} - {lo+25:.0f} m", float(lo), float(lo + 25)) for lo in range(0, 250, 25)]
BAND_SHEETS = ["Band_Basement"] + [f"Band_{lo:03d}-{lo+25:03d}m" for lo in range(0, 250, 25)]
ELEMENTS = ["Column", "Shear Wall", "Slab", "Beam"]

# ---------------------------------------------------------------- items
# (id, element, abstract row, floor label, bottom elev, top elev, measurement rows, elevation note)
L0, L1, L2, L3 = 0.00, 7.32, 12.81, 18.30
ITEMS = [
    ("BM01", "Beam", 99, "Basement roof (Level 0)", L0, L0, [506], "Level 0 = +/-0.00 (basement top)"),
    ("BM02", "Beam", 100, "Ground floor (Level 1) incl. PT beams", L1, L1, [507, 527], "L0 + 7.32 m unsupported height"),
    ("BM03", "Beam", 101, "First floor (Level 2) incl. PT beams", L2, L2, [508, 528], "L1 + 5.49 m"),
    ("BM04", "Beam", 102, "Second floor (Level 3) incl. PT beams", L3, L3, [509, 529], "L2 + 5.49 m"),
    ("BM05", "Beam", 103, "Level 2 roof beams (3.66 m)", 16.47, 16.47, [540, 541, 553, 554, 555, 556, 557], "L2 + 3.66 m"),
    ("BM06", "Beam", 104, "Level 2 roof beams (4.88 m)", 17.69, 17.69, [542, 558, 559], "L2 + 4.88 m"),
    ("BM07", "Beam", 105, "Level 2 roof beams (6.10 m)", 18.91, 18.91, [543], "L2 + 6.10 m"),
    ("BM08", "Beam", 106, "Level 2 edge beam (12.81 m)", 12.81, 12.81, [617], "Unsupported 12.81 m = L2"),
    ("BM09", "Beam", 107, "Level 3 roof beams (3.66 m)", 21.96, 21.96, list(range(544, 551)) + list(range(560, 565)), "L3 + 3.66 m"),
]
for i, (lvl, ar, mr) in enumerate([(54.40, 108, 510), (60.85, 109, 511), (70.0, 110, 512), (75.0, 111, 513), (80.0, 112, 514),
                                   (85.0, 113, 515), (223.0, 114, 516), (229.0, 115, 517), (235.0, 116, 518), (241.0, 117, 519),
                                   (247.0, 118, 520), (250.0, 119, 521)]):
    ITEMS.append((f"BM{10+i:02d}", "Beam", ar, f"Level +{lvl:.2f} m", lvl, lvl, [mr], "Level stated in source"))
ITEMS += [
    ("SL01", "Slab", 121, "Basement roof (Level 0)", L0, L0, [569], "Level 0 = +/-0.00 (basement top)"),
    ("SL02", "Slab", 122, "Ground floor (Level 1) incl. PT slabs", L1, L1, [570, 590, 603], "L0 + 7.32 m"),
    ("SL03", "Slab", 123, "First floor (Level 2) incl. PT slab", L2, L2, [571, 591], "L1 + 5.49 m"),
    ("SL04", "Slab", 124, "Second floor (Level 3) incl. PT slab", L3, L3, [572, 592], "L2 + 5.49 m"),
    ("SL05", "Slab", 125, "Level 2 roof slab (3.66 m)", 16.47, 16.47, [628], "L2 + 3.66 m"),
    ("SL06", "Slab", 126, "Level 2 roof slab (5.49 m)", 18.30, 18.30, [633, 635, 637, 639, 641], "L2 + 5.49 m"),
    ("SL07", "Slab", 127, "Level 2 roof slab (6.71 m)", 19.53, 19.53, [629, 631, 643, 645, 647], "L2 + 6.72 m"),
    ("SL08", "Slab", 128, "Level 3 roof slab (3.66 m)", 21.96, 21.96, [630, 632, 634, 636, 638, 640, 642, 644, 646, 648], "L3 + 3.66 m"),
]
for i, (lvl, ar, mr) in enumerate([(54.40, 129, 573), (60.85, 130, 574), (70.0, 131, 575), (75.0, 132, 576), (80.0, 133, 577), (85.0, 134, 578)]):
    ITEMS.append((f"SL{9+i:02d}", "Slab", ar, f"Level +{lvl:.2f} m", lvl, lvl, [mr], "Level stated in source"))
ITEMS += [
    ("CO01", "Column", 136, "Basement columns (600 mm)", -6.10, 0.00, [683], "Basement height 6.10 m below +/-0.00"),
    ("SW01", "Shear Wall", 138, "Basement shear walls (1000 + 600 mm)", -6.10, 0.00, [653, 668], "Basement height 6.10 m below +/-0.00"),
    ("SW02", "Shear Wall", 139, "Petal shear wall 600 mm (0.0 - 58.6 m)", 0.0, 58.6, [697], "Range stated in source"),
    ("SW03", "Shear Wall", 140, "Petal shear wall 600 mm (0.0 - 84.2 m)", 0.0, 84.2, [698], "Range stated in source"),
    ("SW04", "Shear Wall", 141, "Petal shear wall 600 mm (0.0 - 168.4 m)", 0.0, 168.4, [699], "Range stated in source"),
    ("SW05", "Shear Wall", 142, "Petal shear wall 1000 mm (91.5 - 250.0 m)", 91.5, 250.0, [700],
     "Abstract & steel summary say 91.5 m; Measurements row 700 text says 80.0 m"),
]
ELEM_ORDER = {e: i for i, e in enumerate(ELEMENTS)}
ITEMS.sort(key=lambda t: (ELEM_ORDER[t[1]], t[0]))


def hdr(ws, row, headers, widths=None):
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row, c, h)
        cell.font, cell.fill, cell.alignment, cell.border = f_hdr, fill_hdr, center, box
    if widths:
        for c, w in enumerate(widths, 1):
            ws.column_dimensions[CL(c)].width = w


def link(cell, sheet, ref="A1", text=None):
    cell.hyperlink = f"#'{sheet}'!{ref}"
    if text is not None:
        cell.value = text
    cell.font = f_link


# ================================================================= Index
ix = wb.active
ix.title = "Index"
# ================================================================= Summary
sm = wb.create_sheet("Summary_25m")
# ================================================================= Register
rg = wb.create_sheet("Item_Register")
# ================================================================= Matrix
mx = wb.create_sheet("Band_Matrix")
band_ws = [wb.create_sheet(n) for n in BAND_SHEETS]
sa = wb.create_sheet("SRC_Abstract")
sme = wb.create_sheet("SRC_Measurements")
nt = wb.create_sheet("Notes_Checks")

# ---------------------------------------------------------------- source copies (same cell addresses as original)
def copy_rows(ws, sheet, rows, title):
    for r in rows:
        for c in range(sheet.ncols):
            v = sheet.cell_value(r - 1, c)
            if v in ("", None):
                continue
            cell = ws.cell(r, c + 1, v.strip() if isinstance(v, str) else v)
            cell.font = f_base
            if isinstance(v, float):
                cell.number_format = "#,##0.000"
    ws["K1" if ws.title == "SRC_Abstract" else "L1"].value = title
    ws["K1" if ws.title == "SRC_Abstract" else "L1"].font = Font(name=F, size=9, italic=True, color="7F7F7F")


copy_rows(sa, sA, [1, 2] + list(range(93, 143)),
          f"Values copied from '{SRC_FILE}' sheet 'ABS-LA bldg' - every cell sits at the SAME address as in the original (e.g. C99 here = C99 there).")
copy_rows(sme, sM, list(range(1, 11)) + list(range(494, 703)),
          f"Values copied from '{SRC_FILE}' sheet 'Measurements' - every cell sits at the SAME address as in the original (H = qty with contingency).")
for ws, widths in ((sa, [8, 70, 10, 7, 8, 12, 7, 7, 16]), (sme, [8, 62, 7, 12, 8, 4, 12, 12, 8])):
    for c, w in enumerate(widths, 1):
        ws.column_dimensions[CL(c)].width = w
    ws["B"+str(1)].alignment = Alignment(wrap_text=False)
sa.freeze_panes = "A3"
sme.freeze_panes = "A9"

# ---------------------------------------------------------------- Item register
R0 = 5
rg["A1"] = "ITEM REGISTER - Formwork / shuttering items (Columns, Shear Walls, Slabs, Beams)"
rg["A1"].font = f_title
rg["A2"] = ("Green = link to source sheet  |  Blue = input / assumption (edit if drawings differ)  |  "
            "Qty Used follows the basis selected on the Index sheet (cell C6).")
rg["A2"].font = Font(name=F, size=9, italic=True)
rg["A3"] = "Back to Index"
link(rg["A3"], "Index")
heads = ["Item ID", "Element", "Abstract Ref", "Description (from Abstract)", "Floor / Level", "Bottom Elev (m)",
         "Top Elev (m)", "Elevation basis", "Qty - Abstract (Sqm)", "Measurements rows (col H)", "Qty - Measurements (Sqm)",
         "Difference (Sqm)", "Check", "Qty Used (Sqm)", "Distribution", "Go to Matrix"]
hdr(rg, R0 - 1, heads, [8, 11, 11, 52, 30, 10, 10, 30, 12, 26, 13, 12, 10, 12, 14, 10])
rg.row_dimensions[R0 - 1].height = 42
reg_row = {}
for i, (iid, el, ar, lab, bot, top, mrows, enote) in enumerate(ITEMS):
    r = R0 + i
    reg_row[iid] = r
    rg.cell(r, 1, iid).font = f_b
    rg.cell(r, 2, el)
    c = rg.cell(r, 3, f"ABS C{ar}")
    link(c, "SRC_Abstract", f"B{ar}")
    rg.cell(r, 4, f"=SRC_Abstract!B{ar}").font = f_green
    rg.cell(r, 5, lab)
    rg.cell(r, 6, bot).font = f_blue
    rg.cell(r, 7, top).font = f_blue
    rg.cell(r, 8, enote)
    rg.cell(r, 9, f"=SRC_Abstract!C{ar}").font = f_green
    c = rg.cell(r, 10, ", ".join(f"H{m}" for m in mrows))
    link(c, "SRC_Measurements", f"B{mrows[0]}")
    rg.cell(r, 11, "=" + "+".join(f"SRC_Measurements!H{m}" for m in mrows)).font = f_green
    rg.cell(r, 12, f"=I{r}-K{r}")
    rg.cell(r, 13, f'=IF(ABS(L{r})<0.5,"OK","MISMATCH")')
    rg.cell(r, 14, f'=IF(QtyBasis="ABSTRACT",I{r},K{r})').font = f_b
    rg.cell(r, 15, f'=IF(F{r}=G{r},"At level","Pro-rata height")')
    c = rg.cell(r, 16, "Matrix")
    link(c, "Band_Matrix", f"A{r}")
    for cc in range(1, 17):
        cell = rg.cell(r, cc)
        cell.border = box
        if cell.font == f_base or cell.font.name != F:
            pass
        if cc in (6, 7):
            cell.number_format = "0.00"
        if cc in (9, 11, 12, 14):
            cell.number_format = NUM
        if cc in (4, 5, 8, 10):
            cell.alignment = wrap
    for cc in range(1, 17):
        if rg.cell(r, cc).font.name != F:
            rg.cell(r, cc).font = f_base
RL = R0 + len(ITEMS) - 1
rt = RL + 1
rg.cell(rt, 1, "TOTAL").font = f_b
for cc in (9, 11, 12, 14):
    c = rg.cell(rt, cc, f"=SUM({CL(cc)}{R0}:{CL(cc)}{RL})")
    c.font, c.number_format, c.fill, c.border = f_b, NUM, fill_tot, box
rg.cell(rt, 13, f'=COUNTIF(M{R0}:M{RL},"MISMATCH")&" mismatch(es)"').font = f_b
from openpyxl.formatting.rule import FormulaRule
rg.conditional_formatting.add(f"A{R0}:P{RL}", FormulaRule(formula=[f'$M{R0}="MISMATCH"'], fill=fill_red))
rg.freeze_panes = f"E{R0}"
rg.auto_filter.ref = f"A{R0-1}:P{RL}"
rg[f"F{R0}"].comment = Comment("Blue elevations are assumptions. Levels 1-3 are derived from the unsupported heights "
                               "stated in the Abstract (7.32 m, 5.49 m, 5.49 m). All of them fall inside the 0-25 m band, "
                               "so small corrections will not move quantities between bands.", "Estimator")

# ---------------------------------------------------------------- Band matrix
mx["A1"] = "BAND MATRIX - Qty Used (Sqm) distributed into 25 m height bands"
mx["A1"].font = f_title
mx["A2"] = ("At-level items go wholly to the band containing their level (lower bound exclusive, upper bound inclusive). "
            "Vertical elements spanning a height range (petal / basement walls, columns) are split pro-rata to the height inside each band.")
mx["A2"].font = Font(name=F, size=9, italic=True)
mx["A3"] = "Back to Index"
link(mx["A3"], "Index")
BC0 = 7  # first band column
mh = ["Item ID", "Element", "Qty Used (Sqm)", "Bottom (m)", "Top (m)", "Register"] + [b[0] for b in BANDS] + ["Allocated", "Check"]
hdr(mx, R0 - 1, mh, [8, 11, 12, 9, 9, 9] + [11] * len(BANDS) + [12, 9])
mx.row_dimensions[R0 - 1].height = 32
# band limit rows (linked to Summary)
mx.cell(R0 - 3, BC0 - 1, "Band from (m) >").font = f_b
mx.cell(R0 - 2, BC0 - 1, "Band to (m) <=").font = f_b
SUM_R0 = 8  # summary first band row
for j in range(len(BANDS)):
    col = BC0 + j
    a = mx.cell(R0 - 3, col, f"=Summary_25m!C{SUM_R0+j}")
    b = mx.cell(R0 - 2, col, f"=Summary_25m!D{SUM_R0+j}")
    for x in (a, b):
        x.font, x.number_format, x.alignment = f_green, "0.00", center
    link(mx.cell(R0 - 1, col), BAND_SHEETS[j], "A1")
    mx.cell(R0 - 1, col).font = Font(name=F, size=10, bold=True, color="FFFFFF", underline="single")
for i, it in enumerate(ITEMS):
    r = R0 + i
    iid = it[0]
    mx.cell(r, 1, f"=Item_Register!A{r}").font = f_green
    mx.cell(r, 2, f"=Item_Register!B{r}").font = f_green
    mx.cell(r, 3, f"=Item_Register!N{r}").font = f_green
    mx.cell(r, 4, f"=Item_Register!F{r}").font = f_green
    mx.cell(r, 5, f"=Item_Register!G{r}").font = f_green
    link(mx.cell(r, 6), "Item_Register", f"A{r}", "Register")
    for j in range(len(BANDS)):
        col = CL(BC0 + j)
        lo, hi = f"{col}${R0-3}", f"{col}${R0-2}"
        fml = (f"=IF($D{r}=$E{r},IF(AND($D{r}>{lo},$D{r}<={hi}),$C{r},0),"
               f"$C{r}*MAX(0,MIN($E{r},{hi})-MAX($D{r},{lo}))/($E{r}-$D{r}))")
        c = mx.cell(r, BC0 + j, fml)
        c.number_format, c.font, c.border = NUM, f_base, box
    ca = CL(BC0 + len(BANDS))
    c = mx.cell(r, BC0 + len(BANDS), f"=SUM({CL(BC0)}{r}:{CL(BC0+len(BANDS)-1)}{r})")
    c.number_format, c.font, c.border = NUM, f_b, box
    mx.cell(r, BC0 + len(BANDS) + 1, f'=IF(ABS({ca}{r}-C{r})<0.01,"OK","CHECK")').font = f_base
    for cc in range(1, 6):
        mx.cell(r, cc).border = box
        if cc == 3:
            mx.cell(r, cc).number_format = NUM
        if cc in (4, 5):
            mx.cell(r, cc).number_format = "0.00"
mt = RL + 1
mx.cell(mt, 1, "TOTAL").font = f_b
for cc in [3] + list(range(BC0, BC0 + len(BANDS) + 1)):
    c = mx.cell(mt, cc, f"=SUM({CL(cc)}{R0}:{CL(cc)}{RL})")
    c.font, c.number_format, c.fill, c.border = f_b, NUM, fill_tot, box
mx.freeze_panes = mx.cell(R0, BC0)

# ---------------------------------------------------------------- Summary
sm["A1"] = "SPIRE - LEGISLATIVE ASSEMBLY BUILDING : SHUTTERING AREA PER 25 m HEIGHT"
sm["A1"].font = f_title
sm["A2"] = f"Source: '{SRC_FILE}' - sheets 'ABS-LA bldg' (Formwork item 8) and 'Measurements' (FORMWORK summary rows 494-702)."
sm["A2"].font = Font(name=F, size=9, italic=True)
sm["A3"] = '="Quantity basis in use: "&QtyBasis&"   (change on Index!C6)"'
sm["A3"].font = f_b
link(sm["A4"], "Index", "A1", "Back to Index")
sh = ["#", "Height band", "From (m) >", "To (m) <=", "Column (Sqm)", "Shear Wall (Sqm)", "Slab (Sqm)", "Beam (Sqm)",
      "Total (Sqm)", "% of Total", "Cumulative (Sqm)", "Band sheet"]
hdr(sm, SUM_R0 - 1, sh, [5, 16, 10, 10, 13, 15, 13, 13, 14, 10, 15, 20])
sm.row_dimensions[SUM_R0 - 1].height = 32
mxRange = lambda col: f"Band_Matrix!${col}${R0}:${col}${RL}"
for j, (name, lo, hi) in enumerate(BANDS):
    r = SUM_R0 + j
    sm.cell(r, 1, j)
    sm.cell(r, 2, name).font = f_b
    sm.cell(r, 3, lo).font = f_blue
    sm.cell(r, 4, hi).font = f_blue
    bcol = CL(BC0 + j)
    for k, el in enumerate(ELEMENTS):
        c = sm.cell(r, 5 + k, f'=SUMIFS({mxRange(bcol)},{mxRange("B")},"{el}")')
        c.font = f_green
    sm.cell(r, 9, f"=SUM(E{r}:H{r})").font = f_b
    sm.cell(r, 10, f"=IF($I${SUM_R0+len(BANDS)}=0,0,I{r}/$I${SUM_R0+len(BANDS)})")
    sm.cell(r, 11, f"=SUM($I${SUM_R0}:I{r})")
    link(sm.cell(r, 12), BAND_SHEETS[j], "A1", BAND_SHEETS[j])
    for cc in range(1, 13):
        cell = sm.cell(r, cc)
        cell.border = box
        if cell.font.name != F:
            cell.font = f_base
        if 5 <= cc <= 9 or cc == 11:
            cell.number_format = NUM
        if cc in (3, 4):
            cell.number_format = "0.00"
        if cc == 10:
            cell.number_format = PCT
ST = SUM_R0 + len(BANDS)
sm.cell(ST, 2, "TOTAL").font = f_b
for cc in range(5, 11):
    c = sm.cell(ST, cc, f"=SUM({CL(cc)}{SUM_R0}:{CL(cc)}{ST-1})")
    c.font, c.fill, c.border = f_b, fill_tot, box
    c.number_format = PCT if cc == 10 else NUM
sm.cell(ST + 1, 2, "Check vs Register").font = f_base
c = sm.cell(ST + 1, 9, f"=Item_Register!N{rt}")
c.font, c.number_format = f_green, NUM
c = sm.cell(ST + 1, 10, f'=IF(ABS(I{ST}-I{ST+1})<0.5,"OK","CHECK")')
c.font = f_b
sm.cell(ST + 3, 2, "Band limits (blue) are inputs: lower bound exclusive, upper bound inclusive. "
        "Basement = everything at or below +/-0.00 (basement walls/columns and the Level 0 slab & beams).").font = Font(name=F, size=9, italic=True)
sm.freeze_panes = f"C{SUM_R0}"
ch = BarChart()
ch.type = "bar"
ch.grouping = "stacked"
ch.overlap = 100
ch.title = "Shuttering area per 25 m band (Sqm)"
ch.y_axis.title = "Sqm"
ch.x_axis.title = "Height band"
data = Reference(sm, min_col=5, max_col=8, min_row=SUM_R0 - 1, max_row=ST - 1)
cats = Reference(sm, min_col=2, min_row=SUM_R0, max_row=ST - 1)
ch.add_data(data, titles_from_data=True)
ch.set_categories(cats)
ch.height, ch.width = 10, 20
ch.x_axis.scaling.orientation = "minMax"
sm.add_chart(ch, f"B{ST+5}")

# ---------------------------------------------------------------- Band sheets
for j, ws in enumerate(band_ws):
    name, lo, hi = BANDS[j]
    bcol = CL(BC0 + j)
    ws["A1"] = f"SHUTTERING AREA - HEIGHT BAND: {name}"
    ws["A1"].font = f_title
    ws["A2"] = f'="From > "&TEXT(Summary_25m!C{SUM_R0+j},"0.00")&" m  to <= "&TEXT(Summary_25m!D{SUM_R0+j},"0.00")&" m   |   Basis: "&QtyBasis'
    ws["A2"].font = f_b
    link(ws["A3"], "Summary_25m", f"B{SUM_R0+j}", "<< Summary")
    link(ws["C3"], "Index", "A1", "Index")
    if j > 0:
        link(ws["D3"], BAND_SHEETS[j - 1], "A1", f"< Prev: {BAND_SHEETS[j-1]}")
    if j < len(BANDS) - 1:
        link(ws["E3"], BAND_SHEETS[j + 1], "A1", f"Next: {BAND_SHEETS[j+1]} >")
    # element subtotal block
    hdr(ws, 5, ["Element", "Area (Sqm)", "% of band"], None)
    for k, el in enumerate(ELEMENTS):
        r = 6 + k
        ws.cell(r, 1, el).font = f_b
        c = ws.cell(r, 2, f"=Summary_25m!{CL(5+k)}{SUM_R0+j}")
        c.font, c.number_format = f_green, NUM
        c = ws.cell(r, 3, f"=IF($B$10=0,0,B{r}/$B$10)")
        c.number_format, c.font = PCT, f_base
        for cc in (1, 2, 3):
            ws.cell(r, cc).border = box
    ws.cell(10, 1, "Band total").font = f_b
    c = ws.cell(10, 2, "=SUM(B6:B9)")
    c.font, c.number_format, c.fill = f_b, NUM, fill_tot
    c = ws.cell(10, 3, "=SUM(C6:C9)")
    c.number_format, c.font, c.fill = PCT, f_b, fill_tot
    # item detail
    hdr(ws, 12, ["Item ID", "Element", "Description (from Abstract)", "Floor / Level", "Bottom (m)", "Top (m)",
                 "Item Qty Used (Sqm)", "Share in band (Sqm)", "% of item", "Abstract ref", "Measurements ref", "Matrix cell"], None)
    ws.row_dimensions[12].height = 32
    r = 13
    for i, it in enumerate(ITEMS):
        iid, el, ar, lab, bot, top, mrows, _ = it
        if bot == top:
            inband = (bot > lo and bot <= hi)
        else:
            inband = max(0.0, min(top, hi) - max(bot, lo)) > 0
        if not inband:
            continue
        rr = R0 + i
        ws.cell(r, 1, f"=Item_Register!A{rr}")
        ws.cell(r, 2, f"=Item_Register!B{rr}")
        ws.cell(r, 3, f"=Item_Register!D{rr}")
        ws.cell(r, 4, f"=Item_Register!E{rr}")
        ws.cell(r, 5, f"=Item_Register!F{rr}")
        ws.cell(r, 6, f"=Item_Register!G{rr}")
        ws.cell(r, 7, f"=Item_Register!N{rr}")
        ws.cell(r, 8, f"=Band_Matrix!{bcol}{rr}")
        ws.cell(r, 9, f"=IF(G{r}=0,0,H{r}/G{r})")
        for cc in range(1, 10):
            ws.cell(r, cc).font = f_green if cc != 9 else f_base
        ws.cell(r, 8).font = Font(name=F, size=10, bold=True, color="008000")
        link(ws.cell(r, 10), "SRC_Abstract", f"B{ar}", f"ABS C{ar}")
        link(ws.cell(r, 11), "SRC_Measurements", f"B{mrows[0]}", ", ".join(f"H{m}" for m in mrows))
        link(ws.cell(r, 12), "Band_Matrix", f"{bcol}{rr}", f"{bcol}{rr}")
        for cc in range(1, 13):
            ws.cell(r, cc).border = box
            if cc in (5, 6):
                ws.cell(r, cc).number_format = "0.00"
            if cc in (7, 8):
                ws.cell(r, cc).number_format = NUM
            if cc == 9:
                ws.cell(r, cc).number_format = PCT
            if cc in (3, 4):
                ws.cell(r, cc).alignment = wrap
        r += 1
    if r == 13:
        ws.cell(r, 1, "No Column / Shear Wall / Slab / Beam formwork in this band per the source estimate.").font = Font(name=F, size=10, italic=True)
        r += 1
    else:
        ws.cell(r, 7, "Total").font = f_b
        c = ws.cell(r, 8, f"=SUM(H13:H{r-1})")
        c.font, c.number_format, c.fill, c.border = f_b, NUM, fill_tot, box
        c = ws.cell(r, 9, f'=IF(ABS(H{r}-B10)<0.5,"= band total OK","CHECK")')
        c.font = f_b
    for c, w in enumerate([9, 11, 50, 30, 9, 9, 12, 12, 9, 11, 26, 11], 1):
        ws.column_dimensions[CL(c)].width = w
    ws.freeze_panes = "A13"

# ---------------------------------------------------------------- Notes & checks
nt["A1"] = "NOTES, ASSUMPTIONS & DISCREPANCIES FOUND IN SOURCE"
nt["A1"].font = f_title
link(nt["A2"], "Index", "A1", "Back to Index")
notes = [
    ("Scope", "Formwork for Columns, Shear Walls (incl. petal walls), Slabs (incl. PT & roof slabs) and Beams (incl. PT, roof & edge beams) "
              "from Abstract item 8 (rows 99-142). Excluded: pile caps (row 96, foundation) and the lump-sum spire formwork material item (row 97).", None),
    ("Quantity basis", "Default = MEASUREMENTS (column H 'with contingency' of the Measurements FORMWORK summary). Switch to ABSTRACT on Index!C6 "
                       "to reproduce the priced BOQ quantities. The two agree on every item except the two below.", None),
    ("DISCREPANCY 1", "Abstract C127 'Second floor (6.71 m)' slab formwork = 1,497 Sqm, which equals the CONCRETE volume in C67 (1,497 Cum). "
                      "Measurements roof-slab formwork (6.72 m) = H629+H631+H643+H645+H647 = 4,980 Sqm.", "SL07"),
    ("DISCREPANCY 2", "Abstract C128 'Third floor (3.66 m)' slab formwork = 2,880 Sqm, which equals the CONCRETE volume in C68 (2,880 Cum). "
                      "Measurements Level 3 roof-slab formwork = 9,586 Sqm (10 rows, H630...H648).", "SL08"),
    ("DISCREPANCY 3", "Top petal wall (1000 mm) start level: Abstract C142 and steel summary row 482 say 91.5 m; Measurements row 700 text says 80.0 m. "
                      "91.5 m is used (Item_Register G/F cells are editable).", "SW05"),
    ("Not carried to Abstract", "Measurements rows 654-662 (SW 1000 mm L0-L80), 669-677 (SW 600 mm L0-L80) and 684-687 (Columns L0-L3) have no "
                                "contingency qty and are NOT in the Abstract (petal walls are measured separately in rows 697-700). They are excluded here too - confirm with the QS.", None),
    ("Pro-rata split", "Walls/columns spanning a height range are split in proportion to the height inside each band, i.e. uniform formwork area per metre of height. "
                       "Petal walls taper/curve, so for construction planning replace with lift-wise areas from the drawings if available.", None),
    ("Elevations", "Only levels +54.4 m and above are stated in the source. Level 0 = +/-0.00; Levels 1-3 (7.32 / 12.81 / 18.30 m) and the Level 2/3 roof items are "
                   "estimated from the unsupported heights in the Abstract. All of them are below 25 m, so the band results do not depend on the exact values.", None),
    ("Cross references", "Every quantity is a live formula. SRC_Abstract and SRC_Measurements hold the source values at the SAME cell addresses as the original workbook, "
                         "so e.g. SRC_Measurements!H506 = Measurements!H506 of the original. Blue underlined text = click-through hyperlink.", None),
]
hdr(nt, 4, ["Topic", "Note", "Item"], [22, 120, 10])
for i, (t, n, iid) in enumerate(notes):
    r = 5 + i
    nt.cell(r, 1, t).font = f_b
    nt.cell(r, 2, n).font = f_base
    nt.cell(r, 2).alignment = wrap
    nt.row_dimensions[r].height = 42
    if iid:
        link(nt.cell(r, 3), "Item_Register", f"A{reg_row[iid]}", iid)
        nt.cell(r, 1).fill = fill_red
    for cc in (1, 2, 3):
        nt.cell(r, cc).border = box
        nt.cell(r, cc).alignment = wrap

# ---------------------------------------------------------------- Index
ix["A1"] = "SHUTTERING AREA PER 25 m HEIGHT - SPIRE, LEGISLATIVE ASSEMBLY BUILDING, AMARAVATI"
ix["A1"].font = f_title
ix["A2"] = f"Prepared from '{SRC_FILE}' (Abstract 'ABS-LA bldg' + 'Measurements'). Units: Sqm of formwork contact area."
ix["A2"].font = Font(name=F, size=9, italic=True)
ix["A4"] = "SETTINGS"
ix["A4"].font = f_b
ix["A6"] = "Quantity basis"
ix["A6"].font = f_b
ix["C6"] = "MEASUREMENTS"
ix["C6"].font, ix["C6"].fill, ix["C6"].border = Font(name=F, size=10, bold=True, color="0000FF"), fill_yel, box
ix["D6"] = "<- choose MEASUREMENTS (true formwork area, default) or ABSTRACT (priced BOQ qty). See Notes_Checks."
ix["D6"].font = Font(name=F, size=9, italic=True)
dv = DataValidation(type="list", formula1='"MEASUREMENTS,ABSTRACT"', allow_blank=False)
ix.add_data_validation(dv)
dv.add("C6")
wb.defined_names["QtyBasis"] = DefinedName("QtyBasis", attr_text="Index!$C$6")

ix["A8"] = "HEADLINE RESULTS"
ix["A8"].font = f_b
hdr(ix, 9, ["Height band", "", "Column", "Shear Wall", "Slab", "Beam", "Total (Sqm)", "Open"], None)
for j in range(len(BANDS)):
    r = 10 + j
    c = ix.cell(r, 1, f"=Summary_25m!B{SUM_R0+j}")
    c.font = f_green
    for k in range(5):
        c = ix.cell(r, 3 + k, f"=Summary_25m!{CL(5+k)}{SUM_R0+j}")
        c.font, c.number_format, c.border = f_green, NUM, box
    link(ix.cell(r, 8), BAND_SHEETS[j], "A1", BAND_SHEETS[j])
    ix.cell(r, 1).border = box
r = 10 + len(BANDS)
ix.cell(r, 1, "TOTAL").font = f_b
for k in range(5):
    c = ix.cell(r, 3 + k, f"=Summary_25m!{CL(5+k)}{ST}")
    c.font, c.number_format, c.fill, c.border = Font(name=F, size=10, bold=True, color="008000"), NUM, fill_tot, box

r += 2
ix.cell(r, 1, "SHEETS (click to open)").font = f_b
r += 1
toc = [("Summary_25m", "Band-wise summary (Column / Shear Wall / Slab / Beam) + chart"),
       ("Item_Register", "All 41 formwork items with source links, elevations, Abstract vs Measurements check"),
       ("Band_Matrix", "Each item distributed across the 25 m bands (formulas)")] + \
      [(n, f"Detail of items in band {BANDS[j][0]}") for j, n in enumerate(BAND_SHEETS)] + \
      [("SRC_Abstract", "Copy of source 'ABS-LA bldg' formwork rows (same cell addresses)"),
       ("SRC_Measurements", "Copy of source 'Measurements' FORMWORK summary rows 494-702 (same cell addresses)"),
       ("Notes_Checks", "Assumptions and 3 discrepancies found in the source file")]
for n, d in toc:
    link(ix.cell(r, 1), n, "A1", n)
    ix.cell(r, 3, d).font = f_base
    r += 1
r += 1
ix.cell(r, 1, "LEGEND").font = f_b
for t, fnt, fl in [("Blue text = input / assumption you may edit", f_blue, None),
                   ("Green text = formula linking to another sheet", f_green, None),
                   ("Black text = calculation on the same sheet", f_base, None),
                   ("Yellow fill = key setting", f_base, fill_yel),
                   ("Blue underlined = hyperlink (cross reference)", f_link, None)]:
    r += 1
    c = ix.cell(r, 1, t)
    c.font = fnt
    if fl:
        c.fill = fl
for c, w in enumerate([22, 2, 13, 13, 13, 13, 14, 22], 1):
    ix.column_dimensions[CL(c)].width = w

# default font for any cell left without Arial
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None and cell.font.name != F:
                cell.font = Font(name=F, size=cell.font.size or 10, bold=cell.font.bold, italic=cell.font.italic, color=cell.font.color)
    ws.sheet_view.showGridLines = ws.title.startswith("SRC_")

tabc = {"Index": "1F4E78", "Summary_25m": "00B050", "Item_Register": "FFC000", "Band_Matrix": "FFC000",
        "SRC_Abstract": "7F7F7F", "SRC_Measurements": "7F7F7F", "Notes_Checks": "C00000"}
for ws in wb.worksheets:
    ws.sheet_properties.tabColor = tabc.get(ws.title, "9BC2E6")
from openpyxl.workbook.properties import CalcProperties
wb.calculation = CalcProperties(fullCalcOnLoad=True)
wb.active = 0
wb.save("Spire_Shuttering_per_25m.xlsx")
print("items", len(ITEMS))
