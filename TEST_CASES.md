# Validation Notes

## Source alignment

Validation references from finance kit:

- `c:\Users\mwade\Downloads\finance_plugin_kit\finance_plugin_kit\lectures_markdown\Lecture_6.1_CAPM.md`
- `c:\Users\mwade\Downloads\finance_plugin_kit\finance_plugin_kit\workbooks_export\Lecture_2.3_Buy_vs_Make_EAC_solved\Equivalent Annual Cost.formulas.json`

## Numeric checks executed

### CAPM check (Lecture 6.1 example)

- Inputs: `rf=4%`, `beta=1.89`, `market risk premium=5.5%`
- Expected from slide: Tesla ≈ `14.40%`
- Quick check output: `14.39%` (`0.1439`), rounding-consistent.

### EAC check (Buy vs Make workbook)

Workbook formulas indicate:

- `NPV_A = NPV(10%, C7:E7) + B7`
- `NPV_B = NPV(10%, C8:F8) + B8`
- `EAC_A = PMT(10%, 3, NPV_A)`
- `EAC_B = PMT(10%, 4, NPV_B)`

Using values from workbook sample:

- A cash flows: `500,120,120,120`
- B cash flows: `600,100,100,100,100`
- Results: `EAC_A=321.0574`, `EAC_B=289.2825` -> Machine B lower EAC.

This matches workbook conclusion text: machine B preferred by EAC.

## Build/package checks executed

- `npm run build` succeeds and emits `dist/taskpane.html`, `dist/taskpane.bundle.js`.
- `npm run package` succeeds and creates release zip under `release/`.

## Manual platform check steps (to run on both Windows and Mac Excel)

1. Upload manifest (dev: `manifest.xml`; hosted: `manifest.github.xml` with real URL).
2. Open add-in task pane.
3. Test at least: CAPM, NPV/IRR, EAC, WACC.
4. Verify:
   - Topic switch updates fields.
   - Formula text and result are shown.
   - "Insert result in active cell" writes output to selected cell.
   - Cash-flow selection button populates cash-flow field.
