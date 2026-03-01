# Managerial Finance Excel Add-in (Office Add-in)

Cross-platform Excel add-in (Windows + Mac + Web) for managerial finance exam problems.
The task pane provides topic-driven calculators and always shows:

- Formula
- Formula with your numbers
- Final result

## Topics implemented

- TVM (lump sum)
- NPV / IRR
- MIRR
- Payback / Discounted Payback
- EAC
- WACC
- DCF (FCF + terminal growth)
- Relative valuation (multiple x metric)
- Risk/return (one asset)
- Two-asset portfolio
- N-asset portfolio
- Efficient frontier (2-asset grid)
- CAPM
- Bond price / YTM
- Gordon growth stock value
- Core financial ratios

## Project layout

- `src/taskpane.ts` - Task pane UI + topic wiring
- `src/finance.ts` - Core finance calculations
- `src/functions.ts` - Optional custom functions
- `manifest.xml` - Local dev manifest (`https://localhost:3000`)
- `manifest.github.xml` - Hosted manifest template for GitHub Pages
- `.github/workflows/deploy-pages.yml` - CI build/deploy for hosted artifact

## Run locally (developer)

1. Install dependencies:
   - `npm install`
2. Run dev server:
   - `npm run start`
3. Sideload `manifest.xml` in Excel:
   - Excel -> Insert -> Get Add-ins -> Upload My Add-in

## Share with others (no terminal for users)

Users do **not** need npm or terminal commands.

1. Host built files on GitHub Pages (CI workflow included).
2. Set correct hosted URLs in `manifest.github.xml`.
3. Share that manifest file (or link) with users.
4. Users install via:
   - Excel -> Insert -> Get Add-ins -> Upload My Add-in

Once installed, Excel loads the add-in from your hosted URL.

## Packaging for release

For maintainers, packaging is available:

- Build: `npm run build`
- Zip release (manifest + prebuilt dist): `npm run package`

This creates a distributable zip under `release/` for self-hosters.

## Content references from finance kit

This project is designed to align with:

- `finance_plugin_kit/function_catalog.json`
- `finance_plugin_kit/content_index.json`
- `finance_plugin_kit/lectures_markdown/`
- `finance_plugin_kit/workbooks_export/`

These can be copied or linked into a local `content/` folder for validation and future prompt/tooltips.
# Managerial Finance Excel Add-in

A VBA Excel add-in that helps you solve managerial finance exam-style problems: pick a topic, enter inputs, and see the formula and result in one pane.

---

## Installing the add-in

1. **Get the add-in file**  
   Use the built add-in file: `ManagerialFinance.xlam` (or the workbook version with macros, if provided).

2. **Install in Excel**
   - Open Excel.
   - Go to **File → Options → Add-ins**.
   - At the bottom, choose **Excel Add-ins** from the dropdown and click **Go**.
   - Click **Browse**, find and select `ManagerialFinance.xlam`, then click **OK**.
   - Check the box next to **Managerial Finance** (or the add-in name) and click **OK**.

3. **Enable macros**  
   If Excel asks to enable macros or content, choose **Enable** so the add-in can run.

---

## Using the add-in

1. **Open the solver**  
   On the ribbon, use the **Managerial Finance** (or add-in) button, or run the macro that opens the solver pane.

2. **Pick a topic**  
   Use the topic dropdown (e.g. TVM, NPV/IRR, CAPM) to select the type of problem.

3. **Enter inputs**  
   Fill in the required values (e.g. rate, cash flows, number of periods).

4. **Solve**  
   Click **Solve**. The pane shows the formula and the result.

5. **Optional**  
   Use **Insert result into active cell** to put the result in the active worksheet cell, if that option is available.

---

## Requirements

- Microsoft Excel (Windows) with macros enabled.
- The add-in file (`.xlam`) installed as described above.
