# CFI DataClean

This is a simple Windows desktop utility for cleaning forest inventory tables in a 32-bit Microsoft Access database (`.mdb` or `.accdb`).

Build `1.0.00` is the current modern-UI copy. It keeps the same core cleaning/export workflow as the prior stable build while adding recent v3 refinements for Total Height Protocol controls, always-split history export columns, selected-period DBH shrinkage, duplicate height-shrinkage cleanup, clearer run-option grouping, the BIA header credit, front-loaded current/previous period number and error-count columns, a TotalHeight tab layout that puts previous/current height comparison columns directly after `RuleName`, a tree IDBH tab layout that separates green IDBH jump checks from red IDBH shrinkage checks, field-tab color zones, detailed performance logging, the `PeriodScope` export column for faster current/previous period filtering, legacy TotalHeight support for old problem codes and live TreeClass/TreeStatus values, clearer regen square-count rows decoded from `RegenMeasKey`, stronger regen required-species/stem-count checks, built-in regen MinorPlot required checks, fixed regen `SpeciesCode` identity columns in report exports, plot/tree/regen rule-coverage diagnostics for likely setup gaps, stronger selected-period filtering for regen rows that use `ProjectPeriodID`, corrected regen StemCount max handling so the configured Max regen stems value is used in SQL and export wording, and startup run-log lines that print the configured range limits for easier troubleshooting.

## Author

Christopher LaCroix, USDI BIA Division of Forestry, Branch of Inventory and Planning

## Acknowledgements

Selected reference checks in CFI DataClean were adapted from `CFI DB Cleaning 2.5.2026c.Rmd`, a CFI data-cleaning R script developed by Jesse Wooten. Jesse's script provided important reference logic for missing measurement rows, shrinkage review, Lazarus-tree review, and problem/severity consistency checks. CFI DataClean implements those checks directly in the app; it does not require R to run.

## How to run

If CFI DataClean was sent as a ZIP file, right-click the ZIP and choose **Extract All** before running anything. Do not run CFI DataClean from inside the compressed-folder preview.

Double-click:

```text
CFIDataClean.cmd
```

For a plain-language user guide that can be opened in Microsoft Word, open `CFIDataClean-How-To-Guide.html`.

`CFIDataClean.cmd` is a portable launcher. It starts the hidden launcher in `_CFIDataClean_AppFiles`, which uses the 32-bit Windows PowerShell host so the app can see 32-bit Access/ACE database drivers. It uses a per-process PowerShell execution-policy bypass for this app launch so shared copies do not fail with "not digitally signed." For managed deployment, IT should code-sign `_CFIDataClean_AppFiles\ForestInventoryCleaner.ps1`.

If the app closes before the window appears, open `_CFIDataClean_AppFiles` and run `CFIDataClean-debug.cmd` instead. The debug launcher shows the PowerShell window, which is useful for troubleshooting startup errors.

The app window and taskbar use the broom icon from `_CFIDataClean_AppFiles\CFIDataClean-broom.ico`. The app also sets a CFI DataClean taskbar identity so Windows does not group the running window under the PowerShell icon.

When CFI DataClean starts, a small loading splash appears while the WinForms controls are being prepared. It closes automatically when the main app window is ready.

After the window opens, choose an Access database with **Browse**. CFI DataClean connects to it automatically after you pick the file.

When a database connects, CFI DataClean auto-fills **Current measurement period #** and **Previous measurement period #** in **Run options** from the highest `PeriodNumber` and next lower `PeriodNumber` it can find. If the database only has one period, CFI DataClean checks **Single measurement project**, fixes the current period at `1`, and disables both period boxes.

Use **Clear all**, next to **Run**, to remove the selected database, uploaded project manual, woodland species-code list, and minor-plot rules from the current app window.

## What it does

- Opens an Access database and runs project-wide checks without requiring table or field-mapping dropdowns.
- Uses the CFI template's `AppTables` list when present, so core inventory tables show first.
- Auto-detects common forest inventory fields such as species, IDBH, height, plot keys, tree keys, and regen keys.
- In **Run** mode, checks a temporary copy and saves findings directly to a separate Excel workbook without changing the connected Access database.
- By default, most cleaning checks focus on the entered current and previous measurement period numbers. For first-measurement projects, check **Single measurement project**; CFI DataClean fixes the current measurement period at `1` and ignores the previous period. TreeHistory transition checks always review all periods because they depend on the full measurement history.
- Flags missing, zero, negative, or unusually large tree IDBH, height, and regen stem-count values. Regen `IDBH` is checked as a class code where valid values are `0`, `20`, and `40`. Optional regen MinorPlot rules can be entered separately for timber seedlings, timber saplings `20`, timber saplings `40`, woodland seedlings, woodland saplings `20`, and woodland saplings `40`. For tree `IDBH`, blanks are allowed when TreeHistory is blank, or when TreeHistory is `1`, `4`, `6`, `8`, or `9`; optional Run options can require DBH for new mortality TreeHistory `2`/`3` or old mortality TreeHistory `7`. For tree `TotalHeight`, blanks are allowed when TreeHistory is blank, or when TreeHistory is `1`, `4`, `6`, `8`, or `9`; optional Run options can require height for new or old mortality. The **Total Height Protocol** option can also use 100% live-tree heights or a subsample protocol so projects that only measure some live-tree heights are not flooded with false missing-height rows. For older databases that have not been crosswalked yet, CFI DataClean treats TreeClass `1`, `2`, `3`, `4`, and `9`, and TreeStatus `1`, `2`, and `3`, as live status for TotalHeight checks.
- Flags duplicate selected keys.
- Validates CFI code fields against the template's `AppColumns` and `AppColumnCodes` metadata.
- Reviews data-entry fields only when they are marked active in `AppColumns`; inactive fields are skipped for cleaning findings.
- Requires tree species for every tree measurement row in the selected cleaning period scope, requires crown ratio and crown class for TreeHistory `0`, `5`, and `10`, requires radial increment for timber TreeHistory `5` missed trees and timber TreeHistory `0`, `5`, or `10` trees with Problem1/Problem2 code `121`, requires Severity1 code `3` when TreeHistory is `2` or `3` and Problem1 is recorded, and flags recorded DBH values for TreeHistory `1`, `4`, and `8`.
- Checks required plot-entry fields when `PlotStatus` is `1` or `2`, including elevation, aspect, slope, UTM, measurement date, crew, stand, and stockability fields. Plot remarks are required when `PlotStatus` is `7`.
- Checks PlotStatus progression across all periods when the PlotStatus field is active in `AppColumns`. A plot should stay blank until its first measured period, the first nonblank PlotStatus should be `2` initial install, and later periods should be `1`, `4`, `5`, `6`, or `7` with no blanks after install.
- Flags active plot measurement and plot custom measurement values recorded for plots with `PlotStatus` `4` missing plot, `5` not measured, `6` dropped - off reservation, or `7` dropped - other. Plot remarks, UTM coordinate fields, management unit, and `FLCCommercial` are allowed because they may be preloaded or needed for dropped-plot documentation.
- Flags active tree measurement and tree custom measurement values recorded for plots with `PlotStatus` `4` missing plot, `5` not measured, `6` dropped - off reservation, or `7` dropped - other.
- Flags active regen measurement and regen custom measurement values recorded for plots with `PlotStatus` `4` missing plot, `5` not measured, `6` dropped - off reservation, or `7` dropped - other. The exported finding includes the matching plot status code.
- Skips per-acre expansion, `CalcSiteIndex`, and `GMP` fields as cleaning items.
- Handles `DMUML` as a special three-position code instead of a normal lookup-list code. Values like `000`, `001`, `020`, `200`, and `300` are allowed when each digit is `0`, `1`, `2`, or `3` and the digit total is no more than `6`.
- Checks custom measurement codes from the database setup metadata; make sure custom fields and valid codes are entered in `AppColumns` and `AppColumnCodes` before running.
- Runs workbook-style count checks for plot, tree, regen, and matching custom-measurement tables.
- Runs selected reference checks from `CFI DB Cleaning 2.5.2026c.Rmd`, including missing plot/tree measurements, live-timber DBH shrinkage reports, height shrinkage reports, Lazarus-tree status changes, and problem/severity gaps.
- Uses `field-manual-text.json`, extracted from the supplied field manual workbook, to add field-specific collection guidance to audit messages.
- Checks remeasured trees for large IDBH/height changes, IDBH decreases on live timber trees, height decreases on live trees, and TreeHistory transitions that conflict with the field manual guidance. TreeHistory `9` include/non-include corrections between periods are allowed. CFI DataClean also checks ingrowth logic against plot status and timing: TreeHistory `10` is only valid for first-recorded new trees on remeasurement plots after the initial project period, while new trees on install plots with PlotStatus `2` should be TreeHistory `0`. When active `TreeClass` or `TreeStatus` source fields are available, CFI DataClean compares them with `TreeHistory` using the supplied crosswalk. Clear mismatches are flagged, while conditional crosswalk cases are exported as review rows instead of guessed corrections. The reference workbook is included at [`References\TreeClass_Status_History_Crosswalk.xlsx`](References/TreeClass_Status_History_Crosswalk.xlsx).
- Recommended workflow: run the main CFI DataClean cleaning review after project codes have been crosswalked to the current standard. A light pre-crosswalk review can still catch obvious structural problems, such as missing keys, duplicates, broken relationships, or impossible numeric values, but it should not be used to make final cleaning decisions on fields whose standard codes changed.
- During a crosswalk, ignore cleaning-style findings that are caused by old standard codes changing meaning. Treat those rows as crosswalk review items, complete the crosswalk, then rerun CFI DataClean and use the post-crosswalk workbook as the main cleaning report. If a finding still appears after crosswalking, confirm it against the project manual, field records, or project-specific code-change notes before editing values.
- Exports audit findings to an `.xlsx` workbook with a Summary tab and separate tabs for each measurement field, such as `IDBH`, `StemCount`, `TotalHeight`, `SpeciesCode`, and `TreeHistory`. Regen findings are kept on separate `Regen ...` tabs.
- Adds identity columns to each export tab only where they apply, including `PlotNumber`, `TreeNumber`, `SpeciesCode`, `MinorPlot`, `PeriodNumber`, `PeriodScope`, `TreeKey`, and `RegenMeasKey`. Most detail tabs place `RuleName`, `CurrentPeriodNumber`, `CurrentPeriodErrorCount`, `PreviousPeriodNumber`, `PreviousPeriodErrorCount`, the recorded-value column, and `Verification SQL` first. The period error-count columns show `1` when the row belongs to that selected period and `0` otherwise, while `PeriodScope` labels rows as current period, previous period, other period, or all/no period. The `TotalHeight` tab is optimized for height review by placing `PreviousPeriodNumber`, `PreviousTotalHeight`, `CurrentPeriodNumber`, `CurrentTotalHeight`, and `HeightChange` directly after `RuleName`. The tree `IDBH` tab is optimized for DBH review by placing jump columns (`PreviousPeriodNumber`, `PreviousIDBH`, `CurrentPeriodNumber`, `CurrentIDBH`, `IDBHJump`) and shrinkage columns (`EarlierPeriodNumber`, `EarlierIDBH`, `LaterPeriodNumber`, `LaterIDBH`, `Shrinkage`) directly after `RuleName` when those values are present. IDBH jump comparison cells are shaded green; IDBH shrinkage comparison cells are shaded red. Regen IDBH remains a separate class-code report. Regen rows use `RegenMeasKey`, not `TreeKey`.
- Adds `TreeHistory`, `TreeRemarks`, and `SpeciesCode` to tree-related export rows when those fields are available, so tree findings are easier to verify against the field record.
- Adds `PlotStatus` and `PlotRemarks` to plot measurement/custom measurement rows, and `RegenRemarks` to regen measurement/custom measurement rows.
- Adds a verification SQL query to each exported audit row when the app can support one.
- Numeric-looking exported values are written as Excel numbers so they sort, filter, and calculate like numbers. Leading-zero codes such as `020` are kept as text so the displayed code is not changed.
- Lets you upload a project field manual or guide (`.pdf`, `.docx`, `.xlsx`, `.xlsm`, `.xls`, `.txt`, `.md`, `.csv`, or `.json`) from the **AI guidance** tab so the AI helper can use project-specific instructions. Manual uploads are AI context only; they do not change CFI DataClean's built-in cleaning checks or decide which records are flagged.
- Includes an optional AI helper on its own **AI guidance** tab. Core cleaning runs do not require AI. When turned on, the helper can connect to the approved DOI Azure AI endpoint with your API key, read the connected database locally, and send compact table counts, sample rows, coded cleaning rules, and template code metadata for insights and cleaning recommendations. It only runs fresh checks for chat context when **Include fresh run findings** is checked in the AI helper.
- Can use AI to put project-specific guidance in `AIMessage` when **Use AI guidance in export** is checked. Repeated findings are grouped so the app sends one compact request per issue type with the coded app rule, relevant template codes, cleaning-workbook tips, and manual excerpts when a manual is loaded.
- Runs checks in a background runspace so the window stays responsive, with progress/status updates, elapsed time, and a **Cancel run** button.

The app does not delete rows.

## Requirements

- Windows.
- Microsoft Access or Microsoft Access Database Engine installed in 32-bit form.
- For `.accdb`, the 32-bit ACE provider is required.
- For older `.mdb`, the app can also try the older 32-bit Jet provider.

This has been checked against the provided CFI `.mdb` template. That template opens with the 32-bit Jet provider on this machine, and its core cleaning targets include `Plots`, `PlotMeasurements`, `PlotCustomMeasurements`, `Trees`, `TreeMeasurements`, `TreeCustomMeasurements`, `RegenMeasurements`, and `RegenCustomMeasurements`.

The **Run** button follows the supplied data-cleaning workbook by checking a temporary copy of the database, then exporting a separate Excel workbook:

- Selected period scope for normal cleaning checks. Leave **Limit normal checks to measurement period(s)** on to focus most checks on the entered current and previous measurement periods. For first-measurement projects, check **Single measurement project**; CFI DataClean fixes the current measurement period at `1` and ignores the previous period. Turn the period limit off to review all periods. TreeHistory checks always review all periods.
- Active `AppColumns` entries against actual table field names. Data-entry cleaning checks only run for fields marked active in `AppColumns`; excluded items such as per-acre expansion, `CalcSiteIndex`, and `GMP` are still skipped.
- Plot measurement counts against plot records by period.
- Plot custom measurements against plot measurements by period.
- Tree measurement counts against tree records by period.
- Tree custom measurements against tree measurements by period.
- Regen measurements counted by period decoded from `RegenMeasKey`, plus regen custom measurements checked against those regen measurement counts; for example, `P005337-01-3-122-040` is plot `5337`, period `1`, minor plot `3`, species `122`, and `IDBH` class `40`.
- Row-level missing PlotMeasurements and TreeMeasurements by period, following the coworker R-script logic.
- CFI code values against the database template metadata.
- `DMUML` custom code values, using the special three-position digit-total rule.
- Custom measurement code values when the custom fields and valid codes are present in `AppColumns` and `AppColumnCodes`.
- Required plot entries for plots with `PlotStatus` code `1` or `2`: elevation, aspect, slope percent, slope position, UTM northing, UTM easting, UTM zone, measurement date, crew, stand class, stand age, stockability percent, and stockability factor. These are checked only when the matching field exists and is active in `AppColumns`. Plot remarks are required for plots with `PlotStatus` code `7`.
- PlotStatus progression across all periods. Before a plot is installed, earlier period values should be blank. The first nonblank PlotStatus should be `2` for initial install. After a plot has received PlotStatus `2`, later periods should be `1`, `4`, `5`, `6`, or `7`, and should not go back to blank. A PlotStatus `5` before any initial install code is flagged.
- Plot data on not-measured or dropped plots. Active `PlotMeasurements` and `PlotCustomMeasurements` fields are flagged when a value is recorded for a plot with `PlotStatus` `4` missing plot, `5` not measured, `6` dropped - off reservation, or `7` dropped - other. Plot remarks, UTM northing/easting/zone, management unit, and `FLCCommercial` are allowed because they are commonly preloaded or used to document dropped plots. Key fields such as `PlotID`, `PlotKey`, `PlotMeasKey`, `MeasurementID`, and `PeriodNumber` are not treated as plot data for this rule.
- Tree species requirements. Every tree measurement row in the selected cleaning period scope must have a joined `Trees.SpeciesCode`, even when `TreeHistory` is blank.
- Tree crown requirements. TreeHistory `0`, `5`, and `10` require crown ratio and crown class values when those fields exist and are active in `AppColumns`. Entered crown values are checked against the valid code setup when `AppColumnCodes` is configured.
- Tree radial increment requirements. Timber trees with TreeHistory `5` (missed) require radial increment. Timber trees with TreeHistory `0`, `5`, or `10` also require radial increment when `Problem1` or `Problem2` is `121` (negative diameter growth). The rule runs when the radial increment field exists and is active in `AppColumns`; woodland species entered in the woodland species option are skipped.
- New mortality Severity1 requirement. When TreeHistory is `2` or `3` and `Problem1` is recorded, `Severity1` must be code `3`. Blank, Null, and `0` count as no Problem1; values greater than `0` count as an entered Problem1.
- Tree data on not-measured or dropped plots. Active `TreeMeasurements` and `TreeCustomMeasurements` fields are flagged when a value is recorded for the matching tree period on a plot with `PlotStatus` `4` missing plot, `5` not measured, `6` dropped - off reservation, or `7` dropped - other. Key fields such as `TreeID`, `TreeKey`, `TreeMeasKey`, `MeasurementID`, and `PeriodNumber` are not treated as tree data for this rule.
- Regen data on not-measured or dropped plots. Active `RegenMeasurements` and `RegenCustomMeasurements` fields are flagged when a value is recorded for the matching regen/plot period on a plot with `PlotStatus` `4` missing plot, `5` not measured, `6` dropped - off reservation, or `7` dropped - other. The check links regen records by `RegenMeasKey`/period and plot identifiers, and the export includes the `PlotStatus` value for verification. Regen remarks are allowed.
- Regen required entries. Every regen row with `RegenMeasKey`, `IDBH`, `StemCount`, or `MinorPlot` data recorded must have `SpeciesCode`. Every regen row with `SpeciesCode` greater than `0` must have a valid `IDBH` class code (`0`, `20`, or `40`) and a valid `StemCount`. Entered regen `StemCount` values are checked for range even when `SpeciesCode` is blank, so a value above the max regen stem count is still flagged. When selected-period filtering is on, CFI DataClean uses `RegenMeasurements.PeriodNumber` if present, then `ProjectPeriodID` joined to `ProjectMeasurementPeriods`, and only falls back to decoding period from `RegenMeasKey`.
- Regen MinorPlot requirements. Regen rows with `SpeciesCode` greater than `0` must have `MinorPlot`. Enter allowed minor plots for six optional project-specific classes: timber seedlings (`IDBH` `0`), timber saplings (`IDBH` `20`), timber saplings (`IDBH` `40`), woodland seedlings (`IDBH` `0`), woodland saplings (`IDBH` `20`), and woodland saplings (`IDBH` `40`). Blank category boxes do not enforce allowed-value checks. Woodland categories use the woodland species-code list in Run options.
- Rule coverage diagnostics. CFI DataClean adds `Rule coverage diagnostic` findings when matching plot, tree, or regen records exist but setup may prevent a full review. Examples include `PlotStatus` missing/inactive, required plot fields missing/inactive, tree species joins unavailable, TreeHistory inactive, crown/radial fields missing/inactive, missing join keys such as `PlotID` or `PeriodNumber`, inactive regen `SpeciesCode`, no regen MinorPlot allowed lists entered, blank allowed-list boxes for regen groups that have matching data, or woodland MinorPlot rules entered without a woodland species-code list. These rows are setup warnings, not direct value edits.
- Tree `IDBH` blank-value requirements by TreeHistory. Blank DBH is always allowed when TreeHistory is blank, or when TreeHistory is `1`, `4`, `6`, `8`, and `9`; TreeHistory `1`, `4`, and `8` should not have DBH recorded. Use the Run option checkboxes when a project requires DBH for new mortality TreeHistory `2`/`3` or old mortality TreeHistory `7`. Required blank DBH findings are labeled `Missing DBH`; recorded DBH where DBH should be blank is labeled `DBH should be blank`; entered zero, negative, or too-large DBH values remain `IDBH range check`.
- Tree `TotalHeight` blank-value requirements by TreeHistory, optional MinorPlot filter, and selected **Total Height Protocol**. Blank height is always allowed when TreeHistory is blank, or when TreeHistory is `1`, `4`, `6`, `8`, and `9`. Use the Run option checkboxes when a project requires height for new mortality TreeHistory `2`/`3` or old mortality TreeHistory `7`; enter MinorPlot values when height is only required in certain minor plots. For live-tree height protocols, choose normal required-height rules, 100% live-tree heights, or subsample by plot/species/2-inch IDBH class. Rare species entered in Run options are checked as 100% live-tree height species for TreeHistory `0`, `5`, and `10`, except selected problem-code no-height cases. Older non-crosswalked TreeClass `1`, `2`, `3`, `4`, and `9`, and TreeStatus `1`, `2`, and `3`, are also treated as live status for these TotalHeight checks.
- Timber problem-code height checks. Optional toggles flag entered `TotalHeight` for timber trees with `Problem1` or `Problem2` timber code `127` or legacy `74` (broken/missing top), timber code `128` or legacy `75` (dead top), or timber code `123` or legacy `72` (lean > 15 degrees); blank height is treated as correct for those timber problem-code cases, including new ingrowth TreeHistory `10`. Woodland species entered in the woodland species option are skipped, including woodland lean code `123`/`72` cases where height can still be recorded.
- Optional tree `StemCount` checks. When enabled, entered tree stem counts are checked for range; timber species may have blank `StemCount`, while woodland/non-timber species entered in the woodland species box require `StemCount`.
- Period-to-period IDBH and height changes on remeasured trees. Shrinking IDBH findings are limited to live timber trees. Species entered in **Woodland species to exclude from shrinking diameters** are skipped for shrinking IDBH checks, and each run logs a verification of that exclusion. Woodland trees can legitimately have a lower measured DBH or EDRC from one period to the next when cut, dead, or broken limbs/stems are no longer counted. If a woodland stem is removed low enough that it no longer qualifies as a measured stem, the measured stem count and overall diameter summary can decrease even though the record is not a data-entry mistake.
- Live-timber IDBH shrinkage for the selected current/previous measurement pair. TotalHeight shrinkage is also reviewed, but previous/current pairs already reported by the remeasurement check are not repeated. CFI DataClean does not flag shorter later TotalHeight when a previously live tree, TreeHistory `0`, `5`, or `10`, becomes a dead/mortality tree whose height is required by the selected Run options, TreeHistory `2`, `3`, or `7`, because dead trees can lose height through breakage.
- TreeHistory transitions between measurements. TreeHistory `9` include/non-include corrections are allowed and are not treated as transition errors.
- Ingrowth TreeHistory checks. CFI DataClean joins `TreeMeasurements` to the matching plot measurement period when possible and checks PlotStatus. TreeHistory `10` on an install plot with PlotStatus `2` is flagged because install-plot new trees should be TreeHistory `0`. TreeHistory `10` in the first project measurement period is flagged because first-period trees should be TreeHistory `0`. TreeHistory `10` with an earlier nonblank TreeHistory value for the same tree is flagged for review; earlier blank TreeHistory rows do not count as prior history. A first-recorded live TreeHistory `0` row on a remeasurement plot with PlotStatus `1` is flagged as possible missed ingrowth.
- TreeClass/TreeStatus conversion review. If `TreeHistory` and a detected `TreeClass` or `TreeStatus` field are active in `AppColumns`, CFI DataClean compares the source code with `TreeHistory`. Measurement-level fields in `TreeMeasurements` are preferred; if they are not present, CFI DataClean falls back to matching fields in `Trees`. Clear crosswalk mismatches are flagged. Conditional cases, such as TreeClass `0`, TreeClass `10`, TreeStatus `4`, and TreeStatus `9`, are listed for hand review because they can depend on problem codes, snag/downed status, previous TreeHistory, or project use. TreeClass `1`, `2`, `3`, `4`, and `9`, and TreeStatus `1`, `2`, and `3`, are treated as live and should generally align with TreeHistory `0`, `5`, or `10`. Complete code crosswalking before using CFI DataClean as the final cleaning workbook; findings seen during the crosswalk should be treated as conversion review until the crosswalk is finished. See the included crosswalk workbook at [`References\TreeClass_Status_History_Crosswalk.xlsx`](References/TreeClass_Status_History_Crosswalk.xlsx).
- "Lazarus tree" cases where a tree previously coded mortality/harvest/thinned/old mortality or harvest later appears live, missed, or new ingrowth.
- Problem1/Severity1 and Problem2/Severity2 pairs where one side is entered and the other side is none. Blank, Null, and `0` count as none; values greater than `0` count as entered.

Keep `field-manual-text.json` in the same folder as the app if you want audit messages to include the field manual guidance.

Keep the `_CFIDataClean_AppFiles` folder beside `CFIDataClean.cmd`. It contains the hidden launcher, debug launcher, app script, icons, README, security notes, and reference files that CFI DataClean needs to start. Reference workbooks, including the TreeClass/TreeStatus/TreeHistory crosswalk, are stored in `_CFIDataClean_AppFiles\References`.

Custom measurement code cleaning uses the database's existing `AppColumns` and `AppColumnCodes` setup tables. Users do not need to enter custom codes into CFI DataClean, but project setup should include all custom measurement fields, their valid codes, and the correct active toggle there so data-entry errors can be found.

## Notes

The default review thresholds are:

- Max IDBH: `500` by default; the app allows values up to `3500` for redwood or other large-tree projects.
- Max height: `150`
- Max regen stem count: `100`
- Check tree StemCount: off by default. Turn it on when tree `StemCount` should be reviewed. The same max value is used, but timber species may be blank.
- Require DBH for new mortality TreeHistory `2`/`3`: off by default.
- Require DBH for old mortality TreeHistory `7`: off by default.
- Require height for new mortality TreeHistory `2`/`3`: off by default.
- Require height for old mortality TreeHistory `7`: off by default.
- Flag entered timber height for problem `127`/legacy `74` (broken/missing top), `128`/legacy `75` (dead top), and `123`/legacy `72` (lean > 15 degrees): on by default. Turn off the specific category when a project requires height for that timber condition.
- Total Height Protocol: normal required-height rules by default. Choose 100% live-tree heights for projects requiring height on every live TreeHistory `0`, `5`, and `10` tree. Choose subsample by species and 2-inch IDBH class for projects that require only a minimum number of live-tree heights per plot/species/diameter class.
- Total Height subsample settings: minimum count `2` and minimum eligible IDBH `50` (`5.0` inches) by default. The **Require all heights at/above IDBH** checkbox is off by default; turn it on only when a project requires every eligible tree above a project-specific IDBH threshold, such as `170` for 17.0 inches. These use recorded IDBH units.
- In subsample mode, CFI DataClean assumes `TreeNumber` starts at north and increases clockwise. It can flag a group when enough heights were recorded overall, but one of the first required eligible TreeNumbers is blank while a later TreeNumber has height.
- Rare species requiring 100% live-tree height: blank by default. Enter numeric `SpeciesCode` values separated by commas, spaces, semicolons, or pipes when uncommon/rare species require height on every live tree.
- Minor plots where tree height is required: blank by default, which applies required-height checks to all minor plots.
- Regen timber seedling, timber sapling `20`, timber sapling `40`, woodland seedling, woodland sapling `20`, and woodland sapling `40` minor plots: blank by default, so no category allowed-value list is enforced until values are entered. Blank `MinorPlot` is still flagged for regen rows with `SpeciesCode` greater than `0`.
- Max period-to-period IDBH jump: `100`
- Max period-to-period height jump: `20`
- Woodland species to exclude from shrinking diameters: blank by default. Enter the exact numeric `SpeciesCode` values used in the database, separated by commas, spaces, semicolons, or pipes. Example format: `65, 66, 68`. Woodland species are excluded because cut, dead, or broken limbs/stems can reduce the number of stems that qualify for DBH measurement, which can reduce the overall EDRC or measured diameter from one period to the next.
- Limit normal checks to measurement period(s): on by default. Current measurement period is filled from the highest `PeriodNumber`; previous measurement period is filled from the next lower `PeriodNumber`. If only one period exists, **Single measurement project** is checked automatically, current period is fixed at `1`, and both period boxes are disabled.

Adjust these in the app before clicking **Run** if your units or forest type need different limits.

## Export and AI helper

Click **Run** to create a fresh Excel workbook without writing audit findings into the Access database. The workbook includes:

- One tab per active reviewed field. Tabs with findings contain the audit rows for that field; active fields with no findings still get a reviewed-clean row saying no data-entry errors were found. Regen records are separated onto `Regen ...` tabs instead of being mixed with tree records.
- Each tab starts with a plain-language `Checked` note that summarizes what CFI DataClean reviewed on that worksheet, so users can see what was checked before doing deeper hand review.
- The `Summary` tab includes total error counts plus period-specific error counts for the selected current measurement period (`CurrentPeriodNumber`) and previous measurement period (`PreviousPeriodNumber`). Rows that are all-period checks, lack a clear period, or fall outside those two periods are counted under `OtherOrAllPeriodErrorCount`. Summary count columns are color-coded, nonzero count values are bold, and field rows are visually separated into plot, tree, regen, and other sections.
- Detail report tabs include front-loaded `CurrentPeriodNumber`, `CurrentPeriodErrorCount`, `PreviousPeriodNumber`, and `PreviousPeriodErrorCount` columns plus `PeriodScope` beside `PeriodNumber`, making it faster to filter rows to current period, previous period, other periods, or all/no-period checks. The `TotalHeight` tab instead starts with previous/current height comparison columns after `RuleName`. The tree `IDBH` tab separates `IDBH jump check` rows from `IDBH shrinkage check` rows by using an `IDBHJump` column for large positive jumps and a `Shrinkage` column for shrinking DBH, with jump cells shaded green and shrinkage cells shaded red. Field tabs use lightweight color zones: blue for review/action columns, yellow for recorded values, green for supporting plot/tree/period context, and gray for extra audit detail columns at the end.
- Workbook tabs are color-coded and ordered for review: `Summary` stays first, `Database Square` and `Missing Trees` always follow it, `Run Data` comes next, red field tabs needing cleaning review follow, and green passed/reviewed-clean field tabs are grouped at the back.
- A `Run Data` tab that records the app build, app run time, source database path, export path, selected period scope, selected manual, run thresholds, woodland species exclusions, tree height minor-plot rules, regen minor-plot rules, AI settings used, audit counts, and the top slow timed steps. The API key is never included.
- A matching text run log beside the workbook, named like `Mescalero_CFIDataCleanCleaning_20260604_1200_RunLog.txt`. If a run fails before the workbook is created, CFI DataClean still writes the failure message, elapsed time when available, major progress step, and troubleshooting detail to this text log. The log also records source database size, temporary-drive free space, major step timings, and slow Access SQL previews so slow runs can be diagnosed and improved.
- A `Database Square` tab directly after `Summary` that reports whether plot, tree, regen, and custom-measurement counts line up for every measurement period in the database. This tab intentionally checks all periods even when **Limit normal checks to measurement period(s)** is on. Regen does not have a separate parent table like `Trees`, so CFI DataClean adds a `RegenMeasurements count by RegenMeasKey` row for each period, then compares `RegenCustomMeasurements` to that count using the period decoded from `RegenMeasKey`. The tab is red when rows need review and green when no square-count errors are found.
- A `Missing Trees` tab directly after `Database Square`, with `SpeciesCode` when available and either the missing tree/period rows or an `OK` row when none are found. The tab is red when missing tree measurements are found and green when none are found.
- Identity columns such as `PlotNumber`, `TreeNumber`, `SpeciesCode`, `MinorPlot`, `PeriodNumber`, `TreeKey`, and `RegenMeasKey` only when they apply to that tab. Tree tabs include `SpeciesCode` and `MinorPlot`; regen tabs use `RegenMeasKey`.
- Tree-related tabs include `TreeRemarks`, plot measurement tabs include `PlotRemarks`, and regen measurement tabs include `RegenRemarks` even when the remarks are blank, so users can confirm that no remarks were recorded.
- `Recorded<FieldName>Value`, such as `RecordedIDBHValue`, which shows the value or comparison that triggered the finding.
- Recorded-value columns are highlighted yellow in field report tabs so the entered value or comparison is easy to find.
- For TreeHistory findings, `RecordedTreeHistoryValue` shows every recorded TreeHistory value for that tree across all periods, such as `P1=0; P2=2; P3=7`, so the transition can be reviewed in one row.
- For conversion findings, `RecordedTreeClassVsTreeHistoryValue` or `RecordedTreeStatusVsTreeHistoryValue` shows the source code, recorded TreeHistory, and source field used for the crosswalk check.
- DBH, height, and TreeHistory history details are always split into separate export columns where those values are available. DBH comparison text such as `Earlier P1 IDBH=74; Later P3 IDBH=69; Shrinkage=5` is exported into separate columns. Height comparisons are split into columns such as `PreviousTotalHeight`, `CurrentTotalHeight`, and `HeightChange`, where positive values mean the current/later height is taller and negative values mean it is shorter. TreeHistory timelines are split into period columns such as `TreeHistoryP1`, `TreeHistoryP2`, and `TreeHistoryP3`.
- Numeric-looking values in field tabs are written as Excel numbers where it is safe to do so. Text columns and leading-zero codes remain text.
- `Verification SQL`, which can be copied into Access as a SELECT query. CFI DataClean smoke-tests each distinct generated verification query against the temporary database during export.
- `AIMessage`, which appears when AI export help is enabled and gives project-specific cleaning guidance.

The AI helper is off by default and is not required for normal cleaning runs. It sits on the **AI guidance** tab so it stays out of the way during normal Run setup. The project manual upload also sits on this tab because uploaded manuals are used only by AI chat and AIMessage export guidance. Turn on **AI helper on**, open the helper, enter your model and API key, then ask questions about the database or audit findings. Use **Test connection** to confirm the endpoint/model/key can reach the model before running an export. Chat answers are requested in Markdown, including fenced SQL blocks when the helper gives Access SQL. The helper sends table counts and selected-table schema/sample rows to the endpoint. The full Access file is not uploaded, and the API key is only kept in memory while the app is open; it is not saved by the app or written to the background-run temp files.

If you click **Run** before AI help is fully set up, CFI DataClean warns you and lists what is missing. You can continue without AI for that run, or stop and enter the AI helper info first.

### API keys and Azure AI access

API keys are granted through the BIA AI point of contact. CFI DataClean is designed to run without an API key, so users who do not have AI access can still run the built-in cleaning checks and export the Excel workbook.

API keys are accessed through the BIA Azure AI endpoint in Microsoft Azure AI Foundry: `https://ai.azure.com/`. To find your model, endpoint, and API key:

1. Go to `https://ai.azure.com/`.
2. Confirm you are in the correct environment.
3. Click **Build** in the top-right corner.
4. Click **Deployments** in the left toolbar.
5. Review the list of available models and click the model you want to use.
6. In CFI DataClean, replace the model name with the selected deployment/model name, such as `gpt-4.1-mini`.
7. In Azure AI Foundry, open the model's **Details** tab.
8. Copy the **Target URL** into the CFI DataClean endpoint box. Use only the base endpoint, without extra text after `.com`. It should look similar to `https://doi-training-foundry.services.ai.azure.com/`.
9. Copy the key from the same Azure AI Foundry deployment/environment and paste it into the CFI DataClean API key box.
10. Click **Test connection** in CFI DataClean before running AI-assisted exports.

The API key is unique to the Azure AI environment. Do not share it with anyone who does not have access to your endpoint. This access path reflects the current Azure AI Foundry interface as of June 8, 2026. Microsoft changes its screens regularly, so the exact button names or layout may change. For current BIA AI access questions, contact Chris LaCroix in Teams at `christopher.lacroix@bia.gov`.

The default AI settings are endpoint `https://doi-training-foundry.services.ai.azure.com/` and model/deployment `gpt-4.1-mini`. This build allows the approved DOI Azure AI Foundry host and the earlier approved `doi-training-foundry.cognitiveservices.azure.com` host. For either approved base endpoint, the app builds the chat-completions URL automatically. If you see an error about `doi-https`, the endpoint box contains a DOI/network prefix instead of the real API URL; remove that prefix and paste only the HTTPS endpoint.

The AI helper does not run fresh checks just because it is opened or turned on. In the AI helper, check **Include fresh run findings** only when you want a chat message to run those checks on a temporary copy before sending context to the model. Leave it unchecked for faster general questions that only need table counts, schema/sample rows, existing audit rows, coded rules, template metadata, and the uploaded manual.

To get AI-assisted export guidance:

- Optional: open the **AI guidance** tab, click **Project manual** > **Browse**, and choose the project field manual or guide.
- On the **AI guidance** tab, turn on **AI helper on**, open the helper, confirm the approved endpoint, and enter the model and API key.
- Check **Use AI guidance in export** before clicking **Run**.
- Leave **Compact AI export** checked unless you intentionally want larger AI prompts with more manual/rule context.
- Review `AIMessage` beside `VerificationSql`.

For speed, the export asks AI for guidance by repeated finding type instead of every individual row. **Compact AI export** is on by default and sends a cached, field-focused manual digest for each group instead of raw PDF/manual text, reducing Azure 400 errors while still keeping project guidance in the prompt. If that is rejected, the app tries an even smaller manual digest before using the no-manual fallback.

After each AI-assisted run, the log says in plain language how many AI finding groups used the uploaded manual, how many AI handled without the manual after Azure rejected the manual-context prompt, and how many had no matching uploaded-manual guidance.

If Azure rejects one AI guidance group during export, the app retries without optional model settings and then with a tiny no-manual prompt. If Azure still rejects it, the workbook still finishes and leaves `AIMessage` blank for that group.

PDF manuals are supported when they contain selectable text. Scanned/image-only PDFs need OCR first.

During a run, the app disables inputs that could change the run, shows the current step and elapsed time in the status area and title bar, and lets you cancel between major checks or AI requests. If you try to close CFI DataClean while it is still processing, it warns you and can request cancellation before closing. When the workbook finishes, the app offers to open it. If a run fails, the failure popup and on-screen log show the saved run-log path.

The non-AI export caches repeated verification SQL and plot/tree identity lookups while building the workbook, so large exports should spend less time doing repeated Access lookups.

## Security

See [SECURITY.md](SECURITY.md) for the current security review notes, data-flow summary, and DOI/BIA deployment recommendations.

## Documentation maintenance

Shared release folders should include the app build number in the folder name, using this format: `CFIDataClean-v1.0.00`. The folder version should match `$script:AppVersion` in `_CFIDataClean_AppFiles\ForestInventoryCleaner.ps1`.

When CFI DataClean is updated, review `CFIDataClean-How-To-Guide.html` and revise it if the update changes buttons, Run options, export columns, report tabs, cleaning rules, AI behavior, launcher files, or normal user workflow. When a new app build is released, update `$script:AppVersion` and rename the shared release folder so the folder name matches the build.

Current hardening highlights:

- `Run` checks a temporary copy and does not modify the selected Access database.
- AI is off by default and locked to approved DOI Azure endpoint hosts in this build.
- The AI API key is kept in memory only and is not written to background-run temp settings files.
- Temporary CFI DataClean files are cleaned after normal runs; stale CFI DataClean temp files are cleaned at startup.
- Run logs are saved as text files beside the selected workbook path and are useful for troubleshooting failed runs.
- Ingrowth checks join tree and plot measurements by both plot and period before filtering, which avoids large temporary Access joins on bigger projects.
- Excel workbook staging is done in the CFI DataClean temp folder rather than beside the saved workbook, which avoids OneDrive/Desktop cleanup locks from failing the run.
- Office automation used for manual extraction opens files read-only with macro automation disabled where supported.
- `CFIDataClean.cmd` is portable and does not contain user-specific absolute paths. The hidden launcher uses a per-process `ExecutionPolicy Bypass` so shared copies can launch without changing the user's computer-wide PowerShell policy. IT should sign the PowerShell script for managed deployment. `_CFIDataClean_AppFiles\CFIDataClean-debug.cmd` is retained as a visible troubleshooting fallback.
