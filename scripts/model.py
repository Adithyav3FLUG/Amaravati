"""Resource + cost model for Golf Homes Phase-II L3 schedule.
Produces model.pkl consumed by the Excel and PPT builders."""
import pickle, re, datetime as dt, math
import numpy as np, pandas as pd, openpyxl

BOQ = '../inputs/BOQ_Golf_Homes_Phase-II_17.11.2024.xlsx'
df, preds, parent, children = pickle.load(open('sched.pkl', 'rb'))
START = dt.date(2024, 1, 1)
rows = {r.ID: r for _, r in df.iterrows()}

# ------------------------------------------------------------------ resources
# (name, type, label, group, std rate, unit) ; work rates are Rs/day (8 h), material Rs/label
RES = [
 # --- labour (trades)
 ('Bar Bender','Work','','Labour',950),('Helper (Unskilled)','Work','','Labour',700),
 ('MIVAN Formwork Fitter','Work','','Labour',1000),('Formwork Striker','Work','','Labour',800),
 ('Curing Attendant','Work','','Labour',700),('Mason','Work','','Labour',950),
 ('Carpenter','Work','','Labour',950),('Electrician','Work','','Labour',1000),
 ('Plumber','Work','','Labour',1000),('Pipe-fitter','Work','','Labour',1000),
 ('Welder','Work','','Labour',1050),('Ducting Technician','Work','','Labour',1000),
 ('Tile-layer','Work','','Labour',1000),('Painter','Work','','Labour',900),
 ('Putty Applicator','Work','','Labour',850),('Facade Applicator','Work','','Labour',950),
 ('Waterproofing Applicator','Work','','Labour',950),('Fabricator','Work','','Labour',1000),
 ('Rigger','Work','','Labour',950),('Equipment Operator','Work','','Labour',1100),
 ('MEP / OEM Technician','Work','','Labour',1200),('Gardener','Work','','Labour',750),
 ('Housekeeping Staff','Work','','Labour',650),('Trade Supervisor','Work','','Labour',1500),
 # --- staff
 ('Site Engineer','Work','','Staff',3500),('QA/QC Engineer','Work','','Staff',3500),
 ('MEP Engineer','Work','','Staff',4000),('OEM Engineer','Work','','Staff',6000),
 ('Electrical Engineer','Work','','Staff',4000),('Sr. Liaison Engineer','Work','','Staff',5000),
 ('Liaison Engineer','Work','','Staff',3500),('Architect','Work','','Staff',5000),
 ('HR/Admin Executive','Work','','Staff',2500),('Design / CAD Engineer','Work','','Staff',4000),
 ('Licensed Surveyor','Work','','Staff',4000),('Hydrogeologist','Work','','Staff',6000),
 ('Landscape Architect','Work','','Staff',5000),('Project Manager','Work','','Staff',10000),
 ('CRM Executive','Work','','Staff',3000),('PMO Core Team','Work','','Staff',25000),
 # --- equipment (hire incl. fuel; operator where noted separately)
 ('Tower Crane','Work','','Equipment',25000),('Concrete Boom Pump 60 cum/h','Work','','Equipment',30000),
 ('Excavator 0.9 cum','Work','','Equipment',12000),('Excavator 1.2 cum','Work','','Equipment',15000),
 ('Dozer','Work','','Equipment',14000),('Tipper','Work','','Equipment',6000),
 ('Mobile Crane','Work','','Equipment',35000),('Gondola','Work','','Equipment',3000),
 ('Wellpoint Dewatering Pump','Work','','Equipment',2500),('Paver Finisher','Work','','Equipment',25000),
 ('Drilling Rig','Work','','Equipment',15000),('Material Hoist','Work','','Equipment',4000),
]
# materials: (name, label, rate) market supply rates, Noida/NCR 2024 (planning assumptions)
MAT = {
 'RMC M10':('cum',4600),'RMC M25':('cum',6000),'RMC M30':('cum',6200),'RMC M35':('cum',6500),'RMC M40':('cum',6800),
 'Rebar Fe500D':('t',62000),'AAC Block':('cum',4500),'POP Putty':('sqm',60),
 'Cement Plaster Material':('sqm',140),'Vitrified Tile':('sqm',900),'Ceramic Wall Tile':('sqm',700),
 'Granite Stone':('sqm',2400),'Italian Marble':('sqm',4500),'Engineered Wood Floor':('sqm',4500),
 'Vinyl Flooring':('sqm',2200),'Kota Stone':('sqm',900),'Screed Concrete':('cum',5200),
 'Gypsum / Grid Ceiling':('sqm',800),'Exterior Texture Paint':('sqm',90),'Interior Paint':('sqm',60),
 'Waterproofing Membrane':('sqm',450),'IG Glazing Unit':('sqm',7200),'Aluminium Window':('sqm',6000),
 'Glass Railing':('sqm',8000),'Door Frame Timber':('cum',95000),'Door Shutter (Laminated/uPVC)':('sqm',2600),
 'Fire Door Set':('sqm',5500),'Modular Kitchen':('each',320000),'Wardrobe':('sqm',6800),
 'Structural Steel':('kg',75),'Anti-termite Chemical':('sqm',120),
}
for m,(lab,rate) in MAT.items(): RES.append((m,'Material',lab,'Material',rate))
RATE = {r[0]: r[4] for r in RES}; RTYPE = {r[0]: r[3] for r in RES}

def H(a,b):  # split "Masons/Helpers" etc
    return a, b
CREW = {
 '15 Masons + 15 Helpers':[('Mason',15),('Helper (Unskilled)',15)],
 '1 Engineer + 20 Masons/Helpers':[('Site Engineer',1),('Mason',10),('Helper (Unskilled)',10)],
 '1 Engineer + 30 Masons/Finishing Trades':[('Site Engineer',1),('Mason',30)],
 '1 Boom Pump (60 cum/hr) + 10 Masons + RMC Fleet':[('Concrete Boom Pump 60 cum/h',1),('Equipment Operator',1),('Mason',10)],
 '1 Rigger-in-charge + 10 Riggers + Mobile Crane':[('Trade Supervisor',1),('Rigger',10),('Mobile Crane',1)],
 '1 Engineer + 12 Masons/Bar-benders (4 foundations)':[('Site Engineer',1),('Mason',6),('Bar Bender',6)],
 '8 Curing Attendants + 15 Formwork Strikers':[('Curing Attendant',8),('Formwork Striker',15)],
 '1 Hydrogeologist + 12 Wellpoint Pumps + 6 Operators':[('Hydrogeologist',1),('Wellpoint Dewatering Pump',12),('Equipment Operator',6)],
 '1 Excavator + 15 Masons/Pipe-layers':[('Excavator 0.9 cum',1),('Mason',8),('Plumber',7)],
 '2 Excavators (0.9 cum) + 2 Dozers + 6 Tippers':[('Excavator 0.9 cum',2),('Dozer',2),('Tipper',6)],
 '1 Electrical Supervisor + 4 Wiremen':[('Trade Supervisor',1),('Electrician',4)],
 '20 Electricians + 10 Helpers':[('Electrician',20),('Helper (Unskilled)',10)],
 '1 OEM Commissioning Engineer + 2 Technicians':[('OEM Engineer',1),('MEP / OEM Technician',2)],
 '1 OEM Commissioning Engineer + 4 Technicians':[('OEM Engineer',1),('MEP / OEM Technician',4)],
 '1 OEM Engineer + 8 Installation Technicians':[('OEM Engineer',1),('MEP / OEM Technician',8)],
 '3 Excavators (1.2 cum) + 15 Tippers + Shoring Crew (10)':[('Excavator 1.2 cum',3),('Tipper',15),('Helper (Unskilled)',10)],
 '1 Electrical Supervisor + 10 Electricians':[('Trade Supervisor',1),('Electrician',10)],
 '2 Gondolas + 16 Applicators':[('Gondola',2),('Facade Applicator',16)],
 '1 Supervisor + 10 Helpers':[('Trade Supervisor',1),('Helper (Unskilled)',10)],
 '10 Pipe-fitters + 6 Welders':[('Pipe-fitter',10),('Welder',6)],
 '25 Tile-layers + Helpers':[('Tile-layer',25),('Helper (Unskilled)',10)],
 '1 Design Coordinator + 3 CAD Engineers (Vendor)':[('Design / CAD Engineer',4)],
 'Vendor factory - 3 fabrication lines':[('Site Engineer',1)],
 '1 Engineer + 40 Bar-benders + 25 Carpenters + RMC fleet':[('Site Engineer',1),('Bar Bender',40),('Carpenter',25),('Concrete Boom Pump 60 cum/h',1)],
 '1 Rig + 4 Technicians':[('Drilling Rig',1),('MEP / OEM Technician',4)],
 '12 Ducting Technicians + 6 Pipe-fitters':[('Ducting Technician',12),('Pipe-fitter',6)],
 '1 Project Manager + 1 CRM Executive':[('Project Manager',1),('CRM Executive',1)],
 '1 Supervisor + 20 Housekeeping Staff':[('Trade Supervisor',1),('Housekeeping Staff',20)],
 '12 Carpenters + 8 Fabricators':[('Carpenter',12),('Fabricator',8)],
 '1 Supervisor + 12 Masons/Helpers':[('Trade Supervisor',1),('Mason',6),('Helper (Unskilled)',6)],
 '1 Landscape Architect + 20 Gardeners/Masons':[('Landscape Architect',1),('Gardener',15),('Mason',5)],
 '1 HR/Admin Executive':[('HR/Admin Executive',1)],
 '1 Liaison Engineer':[('Liaison Engineer',1)],
 '1 Sr. Liaison Engineer':[('Sr. Liaison Engineer',1)],
 '1 Sr. Liaison Engineer + 1 Architect':[('Sr. Liaison Engineer',1),('Architect',1)],
 '1 Sr. Liaison Engineer + 1 Architect (part-time)':[('Sr. Liaison Engineer',1),('Architect',0.5)],
 '4 MEP Engineers + 6 Technicians':[('MEP Engineer',4),('MEP / OEM Technician',6)],
 '22 Formwork Fitters + 1 Tower Crane + Operator':[('MIVAN Formwork Fitter',22),('Tower Crane',1),('Equipment Operator',1)],
 '1 Engineer + 15 Carpenters/Fitters + 10 Bar-benders':[('Site Engineer',1),('Carpenter',15),('Bar Bender',10)],
 '18 Painters (Spray + Roller)':[('Painter',18)],
 '1 Plant Engineer + 6 Riggers/Electricians':[('Site Engineer',1),('Rigger',3),('Electrician',3)],
 '18 Plumbers + 10 Helpers':[('Plumber',18),('Helper (Unskilled)',10)],
 '1 Rig + 3 Plumbers':[('Drilling Rig',1),('Plumber',3)],
 '1 Supervisor + 8 Fitters':[('Trade Supervisor',1),('Fabricator',8)],
 'Core PMO Team':[('PMO Core Team',1)],
 '20 Putty Applicators':[('Putty Applicator',20)],
 '2 QA/QC Engineers + 4 MEP Fitters':[('QA/QC Engineer',2),('Electrician',2),('Plumber',2)],
 '40 Bar-benders + 35 Carpenters + Tower Crane':[('Bar Bender',40),('Carpenter',35),('Tower Crane',1),('Equipment Operator',1)],
 '40 Bar-benders + 35 Carpenters + MEP Coordination (6)':[('Bar Bender',40),('Carpenter',35),('MEP / OEM Technician',6)],
 '35 Bar-benders + 30 Carpenters + Tower Crane':[('Bar Bender',35),('Carpenter',30),('Tower Crane',1),('Equipment Operator',1)],
 '18 Bar-benders + 6 Helpers':[('Bar Bender',18),('Helper (Unskilled)',6)],
 '1 Paver Machine + 25 Masons/Helpers':[('Paver Finisher',1),('Mason',13),('Helper (Unskilled)',12)],
 '1 Engineer + 25 Masons/Fitters + OEM Team':[('Site Engineer',1),('Mason',13),('Pipe-fitter',12),('MEP / OEM Technician',4)],
 '16 Electricians + 14 Plumbers':[('Electrician',16),('Plumber',14)],
 '1 QA Engineer + 12 Multi-trade Technicians':[('QA/QC Engineer',1),('MEP / OEM Technician',12)],
 '1 Electrical Engineer + 12 Technicians':[('Electrical Engineer',1),('MEP / OEM Technician',12)],
 '1 Licensed Surveyor + 2 Chainmen (Total Station)':[('Licensed Surveyor',1),('Helper (Unskilled)',2)],
 '1 Engineer + 15 Masons/Bar-benders':[('Site Engineer',1),('Mason',8),('Bar Bender',7)],
 '1 Engineer + 18 Masons/Fitters + OEM Team':[('Site Engineer',1),('Mason',9),('Pipe-fitter',9),('MEP / OEM Technician',4)],
 '1 Applicator Team (8) + 2 Excavators + 10 Tippers':[('Waterproofing Applicator',8),('Excavator 0.9 cum',2),('Tipper',10)],
 '1 Specialist Applicator Team (8 no.) + Anti-termite Agency':[('Waterproofing Applicator',8),('Helper (Unskilled)',4)],
 '1 Store Supervisor + 8 Helpers':[('Trade Supervisor',1),('Helper (Unskilled)',8)],
}
L = df[df.leaf].copy()
missing = set(L['Crew Composition']) - set(CREW)
assert not missing, missing
for lst in CREW.values():
    for n, _ in lst: assert n in RATE, n

# ------------------------------------------------------------------ activity type codes
def acode(r):
    w = r.WBS
    m = re.match(r'T\d\d\.F\d\d\.([a-e])$', w)
    if m: return '3.' + m.group(1)
    m = re.match(r'4\.T\d\d\.[LMU]B\.(t\d)$', w)
    if m: return '4.' + m.group(1)
    m = re.match(r'5\.T\d\d\.[LMU]B\.(m\d)$', w)
    if m: return '5.' + m.group(1)
    m = re.match(r'5\.13\.T\d\d\.(inst|test)$', w)
    if m: return '5.13.' + m.group(1)
    m = re.match(r'7\.T\d\d\.(\d)$', w)
    if m: return '7.' + m.group(1)
    m = re.match(r'(\d\.\d+)\.Z\d$', w)
    if m: return m.group(1)
    return w  # phase-1 and 7.13.1 individually
L['acode'] = L.apply(acode, axis=1)

# ------------------------------------------------------------------ BOQ items
wb = openpyxl.load_workbook(BOQ, data_only=True)
ws = wb['BOQ']
COLS = ['Tower', 'Basement/Podium', 'Misc Bldg', 'Community Ctr', 'Guard/Elec Rm', 'Commercial']
items = []
sec = None; part = 'STR'
for i, r in enumerate(ws.iter_rows(min_row=1, max_row=652, values_only=True), 1):
    if i > 172: part = 'FIN'
    if isinstance(r[0], int) and r[2]: sec = str(r[2]).strip()
    a = r[6]
    if i == 643:  # swimming-pool equipment, LS item without qty
        items.append(dict(row=i, part=part, sec=sec, desc=' '.join(str(r[2]).split()), unit='LS', qty=1.0, rate=float(a),
                          amt=float(a), split=[0, 0, 0, 1.0, 0, 0]))
        continue
    if isinstance(a, (int, float)) and a > 0 and isinstance(r[4], (int, float)) and r[4] > 0:
        q = float(r[4]); split = [float(x) if isinstance(x, (int, float)) else 0.0 for x in r[7:13]]
        rem = q - sum(split)
        if rem > 0.01 * q:  # unsplit balance
            k = int(np.argmax(split)) if sum(split) > 0 else (1 if part == 'STR' else 0)
            split[k] += rem
        elif sum(split) > q * 1.0001:  # over-split: scale down
            split = [s * q / sum(split) for s in split]
        items.append(dict(row=i, part=part, sec=sec, desc=' '.join(str(r[2]).split()), unit=str(r[3]).strip(),
                          qty=q, rate=float(r[5]), amt=float(a), split=split))
boq_total = sum(x['amt'] for x in items)
print('BOQ items', len(items), 'total Cr', boq_total / 1e7)

def classify(it, col):
    """return (bucket, material, mat_factor) for BOQ item portion in column col"""
    d = it['desc'].lower(); s = (it['sec'] or '').upper(); row = it['row']
    if col >= 2: bucket = '6.9'
    else: bucket = None
    mat = None; f = 1.0
    if it['part'] == 'STR':
        if s.startswith('EARTH'):
            if 'filling' in d: b = '2.9'
            elif 'anti' in d or row == 34: b = '2.4'; mat = 'Anti-termite Chemical'
            else: b = '2.1'
        elif s.startswith('CONCRETE') or s.startswith('STEEL'):
            if row in (48, 51): b = '2.3'; mat = 'RMC M10'
            elif row == 55: b = '6.2'
            elif row == 63: b = 'SUB_RCC' if col == 1 else '2.3'; mat = 'RMC M30'
            elif row == 65: b = 'SUB_RCC'; mat = 'RMC M40'
            elif row == 72: b = '6.9'; mat = 'RMC M30'
            elif row == 73: b = '3.d'; mat = 'RMC M35'
            elif row == 84: b = '3.b'
            elif row == 88: b = '3.a' if col == 0 else 'SUB_RCC'; mat = 'Rebar Fe500D'; f = 1 / 1000
            elif row == 80 and col == 0: b = '2.10'
            else: b = 'SUB_RCC'
        elif 'WATER' in s:
            b = '4.t4' if row == 134 else 'SUB_WP'; mat = 'Waterproofing Membrane'
        else: b = 'SUB_RCC'
    else:
        if s.startswith('INTERNAL WALLS'):
            b = '4.t1'
            if row in (177, 180): mat = 'AAC Block'
            elif row == 200: mat = 'Rebar Fe500D'; f = 1 / 1000
            elif row == 204: mat = 'RMC M25'
        elif s.startswith('PANTING'):
            b = '4.t2'; mat = 'POP Putty' if row == 207 else 'Cement Plaster Material'
        elif s.startswith('FLOORING'):
            if row in (328, 338): b = '4.t5'; mat = 'Gypsum / Grid Ceiling'
            else:
                b = '4.t4'
                if 'vitrified' in d or 'skirting' in d and row in (240, 245, 248): mat = 'Vitrified Tile'
                if 'ceramic' in d: mat = 'Ceramic Wall Tile'
                if 'granite' in d: mat = 'Granite Stone'
                if 'marble' in d: mat = 'Italian Marble'
                if row in (295, 302): mat = 'Kota Stone'
                if row == 306: mat = 'Screed Concrete'
        elif s.startswith('WALL FINISHES'):
            if row == 346: b = '4.t3'; mat = 'Exterior Texture Paint'
            else: b = '4.t5'; mat = 'Interior Paint' if row in (354, 358, 362, 366, 371) else None
        elif s.startswith('DOOR'):
            b = '4.t6'
            if row == 388: mat = 'Door Frame Timber'
            elif row in (394, 397, 399): mat = 'Door Shutter (Laminated/uPVC)'
            elif row in (417, 419, 421): mat = 'Fire Door Set' if row != 417 else None
            elif row in (426, 429): mat = 'Aluminium Window'
            elif row in (441, 443, 451): mat = 'Structural Steel'
            if row == 443 and col == 1: b = '2.10'
        else:  # NDSR finishing items
            if row == 470: b = '4.t3'; mat = 'IG Glazing Unit'
            elif row == 590: b = '4.t3'
            elif row in (490, 491): b = '4.t4'; mat = 'Engineered Wood Floor'
            elif row == 629: b = '4.t4'; mat = 'Vinyl Flooring'
            elif row in (493, 577, 579, 608, 635, 552): b = '4.t4'; mat = 'Vitrified Tile' if row == 635 else None
            elif row in (581, 583, 585): b = '4.t5'
            elif row == 568: b = '4.t6'; mat = 'Modular Kitchen'
            elif row == 572: b = '4.t6'; mat = 'Glass Railing'
            elif row == 637: b = '4.t6'; mat = 'Door Shutter (Laminated/uPVC)'
            elif row == 645: b = '4.t6'; mat = 'Wardrobe'
            else: b = '4.t6'
        if col == 1 and b.startswith('4.'): b = '2.10'
    if bucket is None: bucket = b
    if bucket == '6.9' and mat and mat not in ('RMC M30', 'RMC M35', 'Rebar Fe500D'):
        pass
    return bucket, mat, f

boqmap = []
for it in items:
    for c, qc in enumerate(it['split']):
        if qc <= 0: continue
        b, mat, f = classify(it, c)
        boqmap.append(dict(row=it['row'], part=it['part'], sec=it['sec'], desc=it['desc'], unit=it['unit'], rate=it['rate'],
                           col=COLS[c], qty=qc, amt=qc * it['rate'], bucket=b, mat=mat, matf=f))
bm = pd.DataFrame(boqmap)
print('BOQ map total Cr', bm.amt.sum() / 1e7)

# external BOQ
EXT = [('External BOQ – Road work (earthwork, sub-grade, WBM/DBM, bitumen)', 70820357, '6.2'),
       ('External BOQ – Concrete paver blocks', 2629312, '6.2'),
       ('External BOQ – Kerb stone', 4340886, '6.2'),
       ('External BOQ – Horticulture (grassing, trees, tree guards, benches)', 17168859, '6.8'),
       ('External BOQ – Hardscape (Kota cobble, granite, pavers, VDF, courts)', 43487224, '6.8')]
for d, a, b in EXT:
    bm = pd.concat([bm, pd.DataFrame([dict(row=None, part='EXT', sec='BOQ EXTERNAL', desc=d, unit='LS', rate=a, col='External',
                                            qty=1, amt=a, bucket=b, mat=None, matf=1)])], ignore_index=True)

# ------------------------------------------------------------------ allowances (not in BOQ) – planning assumptions
BUA = 558449.55  # sqm, BOQ built-up area
ALLOW = [  # (code, description, basis, qty, rate, bucket-map {acode: share})
 ('AL-E', 'Electrical installation (internal wiring, DBs, light & power points)', 'Rs/sqm BUA', BUA, 1400, {'5.m1': 0.6, '5.m5': 0.4}),
 ('AL-P', 'Plumbing & sanitary (CPVC/UPVC, sanitaryware, CP fittings)', 'Rs/sqm BUA', BUA, 900, {'5.m2': 0.55, '5.m5': 0.45}),
 ('AL-H', 'HVAC / ventilation (basement & toilet exhaust, VRF provision)', 'Rs/sqm BUA', BUA, 300, {'5.m3': 1.0}),
 ('AL-F', 'Fire fighting & fire alarm (sprinkler, hydrant, FA/PA)', 'Rs/sqm BUA', BUA, 450, {'5.m4': 1.0}),
 ('AL-L', 'Passenger/service lifts – 11 towers x 8 cars (supply, install, test)', 'Rs/car', 88, 4000000, {'5.13.inst': 0.9, '5.13.test': 0.1}),
 ('AL-S', 'Substation, transformers, HT/LT panels & DG sets', 'LS', 1, 180000000, {'6.6': 1.0}),
 ('AL-T', 'Sewage treatment plant (MBBR) – civil + E&M', 'LS', 1, 60000000, {'6.4': 1.0}),
 ('AL-W', 'Water treatment / softening plant & pump house E&M', 'LS', 1, 25000000, {'6.5': 1.0}),
 ('AL-R', 'Rainwater harvesting pits & UG tanks (civil)', 'LS', 1, 40000000, {'6.7': 1.0}),
 ('AL-C', 'Compound wall, gates & security cabins', 'LS', 1, 45000000, {'6.1': 1.0}),
 ('AL-D', 'Storm-water drainage network', 'LS', 1, 35000000, {'6.3': 1.0}),
 ('AL-X', 'External lighting & signage', 'LS', 1, 15000000, {'6.10': 1.0}),
 ('AL-1', 'Site establishment – offices, stores, BOCW labour camp', 'LS', 1, 60000000, {'1.3.1': 0.4, '1.3.2': 0.6}),
 ('AL-2', 'Temporary power, water & site roads', 'LS', 1, 20000000, {'1.5.1': 0.6, '1.5.2': 0.4}),
 ('AL-3', 'Batching plant & material yard set-up', 'LS', 1, 35000000, {'1.6.1': 0.7, '1.6.2': 0.3}),
 ('AL-4', 'Tower crane foundations, erection & dismantling (12 nos.)', 'LS', 1, 45000000, {'1.7.1': 0.4, '1.7.2': 0.6}),
 ('AL-5', 'Hoarding, fencing & site clearance', 'LS', 1, 12000000, {'1.2.1': 0.6, '1.2.2': 0.4}),
]

# ------------------------------------------------------------------ labour/equipment per task
L['crew'] = L['Crew Composition'].map(CREW)
# time-related plant not named in the crew text: crane on every lifting day of the floor cycle, hoists for finishing trades
EXTRA = {'3.a': [('Tower Crane', 1), ('Equipment Operator', 1)], '3.e': [('Tower Crane', 1), ('Equipment Operator', 1)],
         '4.t1': [('Material Hoist', 1)], '4.t2': [('Material Hoist', 1)], '4.t4': [('Material Hoist', 1)],
         '4.t6': [('Material Hoist', 1)]}
L['crew'] = [c + EXTRA.get(a, []) for c, a in zip(L.crew, L.acode)]
def le_cost(r):
    lab = eq = 0.0
    for n, u in r.crew:
        c = u * r.dur * RATE[n]
        if RTYPE[n] == 'Equipment': eq += c
        else: lab += c
    return pd.Series({'lab': lab, 'equip': eq})
L[['lab', 'equip']] = L.apply(le_cost, axis=1)

# ------------------------------------------------------------------ bucket -> tasks
# multi-package buckets
SUB_RCC_W = {'2.3': 0.30, '2.5': 0.15, '2.6': 0.15, '2.7': 0.15, '2.8': 0.15, '2.10': 0.10}
SUB_WP_W = {'2.4': 0.6, '2.9': 0.4}
def expand_bucket(b):
    if b == 'SUB_RCC': return SUB_RCC_W
    if b == 'SUB_WP': return SUB_WP_W
    return {b: 1.0}
# secondary activities: budget = own L+E (their cost sits inside the BOQ rate of the primary)
SECONDARY = {'3.c': '3.d', '3.e': '3.b'}

L['budget'] = 0.0
L['src'] = ''
matq = {}  # (ID, material) -> qty
# BOQ buckets
alloc = {}
for _, x in bm.iterrows():
    for ac, w in expand_bucket(x.bucket).items():
        alloc.setdefault(ac, []).append((x, w))
by_ac = {ac: g.index.tolist() for ac, g in L.groupby('acode')}
for ac, lst in alloc.items():
    idx = by_ac[ac]
    total = sum(x.amt * w for x, w in lst)
    sec_ac = [k for k, v in SECONDARY.items() if v == ac]
    sec_cost = sum(L.loc[by_ac[s], 'lab'].sum() + L.loc[by_ac[s], 'equip'].sum() for s in sec_ac)
    for s in sec_ac:
        for i in by_ac[s]:
            L.at[i, 'budget'] += L.at[i, 'lab'] + L.at[i, 'equip']; L.at[i, 'src'] = 'BOQ (in %s rate)' % ac
    prim = total - sec_cost
    n = len(idx)
    for i in idx:
        L.at[i, 'budget'] += prim / n; L.at[i, 'src'] = 'BOQ'
    for x, w in lst:
        if isinstance(x.mat, str):
            q = x.qty * w * x.matf / n
            for i in idx: matq[(L.at[i, 'ID'], x.mat)] = matq.get((L.at[i, 'ID'], x.mat), 0) + q
# allowances
for code, desc, basis, q, rate, mp in ALLOW:
    amt = q * rate
    for ac, sh in mp.items():
        idx = by_ac[ac]
        for i in idx:
            L.at[i, 'budget'] += amt * sh / len(idx)
            L.at[i, 'src'] = (L.at[i, 'src'] + '+' if L.at[i, 'src'] else '') + code
# unbudgeted tasks -> resource cost only
for i in L.index:
    if L.at[i, 'budget'] == 0:
        L.at[i, 'budget'] = L.at[i, 'lab'] + L.at[i, 'equip']; L.at[i, 'src'] = 'Resource cost'
    elif L.at[i, 'src'].startswith('AL') or L.at[i, 'src'] in ('BOQ',) or '+' in L.at[i, 'src']:
        pass

# ---- labour scaling: BOQ labour content (DSR analysis-of-rates shares) -> gang multiplier
LSHARE = {'2.1': 0.10, '2.3': 0.14, '2.4': 0.25, '2.5': 0.18, '2.6': 0.18, '2.7': 0.18, '2.8': 0.18, '2.9': 0.20,
          '2.10': 0.18, '3.a': 0.10, '3.b': 0.25, '3.d': 0.06, '4.t1': 0.32, '4.t2': 0.45, '4.t3': 0.12, '4.t4': 0.22,
          '4.t5': 0.35, '4.t6': 0.10, '5.m1': 0.25, '5.m2': 0.22, '5.m3': 0.25, '5.m4': 0.25, '5.m5': 0.18,
          '5.13.inst': 0.10, '5.13.test': 0.10, '6.1': 0.25, '6.2': 0.12, '6.3': 0.22, '6.4': 0.15, '6.5': 0.15,
          '6.6': 0.08, '6.7': 0.20, '6.8': 0.25, '6.9': 0.20, '6.10': 0.20}
KMAX = 5  # cap on gang multiplier; labour content beyond this sits in piece-rate/labour-supply subcontract (fixed cost)
L['kfac'] = 1.0
for i, r in L.iterrows():
    sh = LSHARE.get(r.acode)
    if not sh: continue
    trade_cost = sum(u * r.dur * RATE[n] for n, u in r.crew if RTYPE[n] == 'Labour' and n != 'Equipment Operator')
    target = r.budget * sh
    if trade_cost > 0 and target > trade_cost:
        k = min(KMAX, target / trade_cost)
        L.at[i, 'kfac'] = k
        L.at[i, 'crew'] = [(n, (max(1, round(u * k)) if (RTYPE[n] == 'Labour' and n != 'Equipment Operator') else u)) for n, u in r.crew]
L[['lab', 'equip']] = L.apply(le_cost, axis=1)
L['LSHARE'] = L.acode.map(LSHARE)

# materials per task (limit 4 largest by value, rest stays in fixed cost)
L['mats'] = [[] for _ in range(len(L))]
for (tid, m), q in matq.items():
    i = L.index[L.ID == tid][0]
    L.at[i, 'mats'].append((m, q))
for i in L.index:
    ms = sorted(L.at[i, 'mats'], key=lambda t: -t[1] * RATE[t[0]])[:4]
    L.at[i, 'mats'] = ms
L['mat'] = L.mats.apply(lambda ms: sum(q * RATE[m] for m, q in ms))
L['uplift'] = (L.lab + L.equip + L.mat - L.budget).clip(lower=0)
L['budget'] = L.budget + L.uplift
L['fixed'] = L.budget - L.lab - L.equip - L.mat
print('resource-driven uplift Cr', L.uplift.sum() / 1e7, (L.uplift > 0).sum())
neg = L[L.fixed < -1]
print('negative fixed tasks', len(neg))
if len(neg): print(neg[['ID', 'WBS', 'acode', 'budget', 'lab', 'equip', 'mat', 'fixed']].head(20))
L['total'] = L.lab + L.equip + L.mat + L.fixed
print('BAC Cr', L.total.sum() / 1e7, 'lab', L.lab.sum() / 1e7, 'equip', L.equip.sum() / 1e7, 'mat', L.mat.sum() / 1e7, 'fixed', L.fixed.sum() / 1e7)

# ------------------------------------------------------------------ MSP resource-name strings
def rn(r):
    parts = []
    for n, u in r.crew:
        parts.append('%s[%s%%]' % (n, ('%g' % (u * 100))))
    for m, q in r.mats:
        parts.append('%s[%s]' % (m, ('%.2f' % q).rstrip('0').rstrip('.')))
    return ','.join(parts)
L['resnames'] = L.apply(rn, axis=1)
print('max resname len', L.resnames.str.len().max())
L['cal'] = L['GRAP Calendar Category'].str[0].map({'A': 'GRAP-A Full Stop', 'B': 'GRAP-B Derated 60%', 'C': 'Standard'})

# ------------------------------------------------------------------ daily time-phasing
def grap(d): return (d.month == 11 and d.day >= 16) or d.month == 12 or (d.month == 1 and d.day <= 15)
NDAY = 900
days = [START + dt.timedelta(i) for i in range(NDAY)]
cap = {'A': np.array([0 if grap(d) else 1 for d in days], float),
       'B': np.array([0.6 if grap(d) else 1 for d in days], float), 'C': np.ones(NDAY)}
L['c'] = L['GRAP Calendar Category'].str[0]
L['s'] = (pd.to_datetime(L.Start) - pd.Timestamp(START)).dt.days
L['f'] = (pd.to_datetime(L.Finish) - pd.Timestamp(START)).dt.days
cost_day = np.zeros(NDAY); ph_cost = {}; head = np.zeros(NDAY); head_trade = {}
typ_day = {k: np.zeros(NDAY) for k in ['Labour', 'Staff', 'Equipment', 'Material', 'Subcontract/Other']}
L['phase'] = L.WBS.str.extract(r'^(\d)')[0]
L.loc[L.WBS.str.match(r'^T\d'), 'phase'] = '3'
for i, r in L.iterrows():
    prof = np.zeros(NDAY); prof[r.s:r.f + 1] = cap[r.c][r.s:r.f + 1]
    prof = prof / prof.sum()
    cost_day += prof * r.total
    ph_cost.setdefault(r.phase, np.zeros(NDAY)); ph_cost[r.phase] += prof * r.total
    typ_day['Material'] += prof * r.mat; typ_day['Subcontract/Other'] += prof * r.fixed
    # headcount: resources present on working days (units)
    work = np.zeros(NDAY); work[r.s:r.f + 1] = (cap[r.c][r.s:r.f + 1] > 0)
    for n, u in r.crew:
        t = RTYPE[n]
        c = u * r.dur * RATE[n]
        typ_day['Equipment' if t == 'Equipment' else ('Staff' if t == 'Staff' else 'Labour')] += prof * c
        if t in ('Labour', 'Staff'):
            head += work * u
            head_trade.setdefault(n, np.zeros(NDAY)); head_trade[n] += work * u
# peak units per resource (max units)
peak = {n: float(v.max()) for n, v in head_trade.items()}
eq_peak = {}
for i, r in L.iterrows():
    pass
eqd = {}
for i, r in L.iterrows():
    w = np.zeros(NDAY); w[r.s:r.f + 1] = (cap[r.c][r.s:r.f + 1] > 0)
    for n, u in r.crew:
        if RTYPE[n] == 'Equipment':
            eqd.setdefault(n, np.zeros(NDAY)); eqd[n] += w * u
peak.update({n: float(v.max()) for n, v in eqd.items()})

months = pd.period_range('2024-01', '2026-05', freq='M')
didx = pd.PeriodIndex([pd.Period(d, 'M') for d in days])
def monthly(arr, how='sum'):
    s = pd.Series(arr, index=didx)
    return (s.groupby(level=0).sum() if how == 'sum' else s.groupby(level=0).mean()).reindex(months).fillna(0)
m_cost = monthly(cost_day); m_head = monthly(head, 'mean'); m_headpeak = pd.Series(head, index=didx).groupby(level=0).max().reindex(months).fillna(0)
m_ph = pd.DataFrame({p: monthly(v) for p, v in sorted(ph_cost.items())})
m_typ = pd.DataFrame({k: monthly(v) for k, v in typ_day.items()})
print('monthly cost sum Cr', m_cost.sum() / 1e7, 'peak head', head.max(), 'avg', m_head[m_head > 0].mean())
print(m_cost.round(-5) / 1e7)

pickle.dump(dict(L=L, bm=bm, items=items, RES=RES, MAT=MAT, CREW=CREW, ALLOW=ALLOW, EXT=EXT, peak=peak,
                 m_cost=m_cost, m_head=m_head, m_headpeak=m_headpeak, m_ph=m_ph, m_typ=m_typ, months=months,
                 head=head, days=days, SUB_RCC_W=SUB_RCC_W, SUB_WP_W=SUB_WP_W, BUA=BUA, cost_day=cost_day,
                 ph_cost=ph_cost), open('model.pkl', 'wb'))
