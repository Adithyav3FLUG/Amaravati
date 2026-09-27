# Golf Homes Phase-II – Project Execution Plan (P4)

This is the resource- and cost-loaded execution plan for Golf Homes Phase-II (GH-02, Sector 4, Greater Noida; NBCC). It follows the ten deliverables of the *Sample P4 Project Plan*. It is built from four inputs: the L3 GRAP-aware schedule, the BOQ dated 17-Nov-2024, the master-plan area details and the sample deck.

## Deliverables (`deliverables/`)

| File | What it is |
|---|---|
| `Golf_Homes_PhII_Project_Execution_Plan.pptx` | 26-slide deck following the Sample P4 storyline, from project brief to conclusion (full slide list below). |
| `Golf_Homes_PhII_L3_Schedule_Resource_Cost_Loaded.xlsx` | The original L3 schedule, unchanged in columns A–K. Adds resources and costs on all 2,672 L3 activities, plus the sheets Resources, Assignments, BOQ_Map, Allowances, Rate_Norms, Cost_Summary, CashFlow_S-Curve, Labour_Histogram, EVM, MSP_Import_Tasks, MSP_Import_Resources and MSP_Link_Guide. |
| `MS_Project_Linking_Guide.md` | Step-by-step guide to linking the workbook to MS Project: calendars, Import Wizard mapping, checks, Baseline 0, critical path, cash flow and EVM. |

Deck slides, in order: brief, area statement, assumptions, methodology, WBS L1–L3, Gantt, calendars and logic, milestones, critical path, MIVAN cycle, resources, labour histogram, cost plan, cost build-up, baseline, S-curve, EVM framework, EVM status, EVM dashboard, risks, MS Project linking, data reconciliation, conclusion.

## Headline results

- **Schedule:** 01-Jan-2024 to 14-May-2026 (865 days). An independent CPM run reproduces every date in the Excel schedule. 7 activities have zero float. The driving path is 197 activities long, running through the T12 MIVAN cycle and the restart after the GRAP window.
- **BAC:** ₹1,635.8 Cr (₹2,721/sft on 5,58,450 sqm BUA). It is made up of:
  - the civil BOQ: ₹1,354.9 Cr
  - the external BOQ: ₹13.8 Cr
  - allowances for MEP, lifts, services and preliminaries: ₹262.7 Cr
  - resource-only activities: ₹4.4 Cr
- **Management reserve:** ₹81.8 Cr (5 %), held outside the BAC.
- **Resources:** 83 resources (24 trades, 16 staff roles, 12 plant items, 31 materials) and 9,864 assignments. Headcount peaks at about 3,900 (monthly average, Jun-2025).
- **EVM:** an illustrative status at 30-Apr-2025 gives CPI 0.97, SPI 0.95, EAC ₹1,692 Cr and VAC −₹56 Cr, which is within the management reserve.

## Data caveats (read before baselining)

1. The area statement and BOQ sanction **11 towers**. The L3 schedule carries **12 tower-units** (T08 is "Godavari-2"). The BOQ tower quantities are spread across the 12 units.
2. The BOQ's printed total leaves out row 647 (polycarbonate sheets, ₹0.94 Cr). This plan carries every priced line.
3. The BOQ has no MEP, lifts, substation/DG, STP/WTP or preliminaries, so these are carried as allowances. Replace them with the specialist BOQs.
4. 107 L3 activities have no successor. Tie them to the handover milestones before saving Baseline 0.

## Rebuilding (`scripts/`)

Put the two source workbooks in `inputs/`, named as follows:

- `Golf_Homes_PhII_L3_Schedule_GRAP.xlsx`
- `BOQ_Golf_Homes_Phase-II_17.11.2024.xlsx`

Then run the following from `scripts/`:

```bash
python3 cpm.py && python3 model.py && python3 excel.py   # workbook (recalculate in Excel/LibreOffice after)
python3 deckdata.py && node deck.js                     # deck (needs: npm install pptxgenjs)
```
