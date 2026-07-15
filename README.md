# Project Pickle

Project Pickle builds CFI databases for new projects, updates tables in existing projects, and upgrades old CFI databases, tables and queries included, to the latest standard database with the click of a button.

Version 2.0.08 keeps the 2.0 workflow updates and uses the logical setup order: Project, Periods, AppColumns, Upgrade Mapping, Plots, Get Field Order, CFIDEER Queries, Plot Assignments, Tree Assignments, Build, and AI Help.

## Author

Christopher LaCroix, USDI BIA Division of Forestry, Branch of Inventory and Planning

## How To Run

Open the main `Project Pickle` folder and double-click:

```text
Project Pickle.vbs
```

The main folder is intentionally simple: it shows a share-safe root launcher, the launcher shortcut, and the `Project Pickle App Files` support folder. `Project Pickle.vbs` and the root `Run-ProjectPickle-32bit.bat` use relative paths, so they continue to work when the whole Project Pickle folder is shared with team members or moved to another location. The shortcut uses the pickle icon and points to the root launcher for convenience on this local copy. Do not open the internal app script directly, because some computers block unsigned PowerShell script files.

For troubleshooting from the main folder, run:

```text
Run-ProjectPickle-32bit.bat
```

## User How-To Guide

Open `Project Pickle How-To Guide.html` in the main Project Pickle folder for a reader-friendly walkthrough of the app. It is formatted as a Word-friendly HTML file, so it can be opened in Microsoft Word for reading, printing, or saving as a Word document.

Whenever Project Pickle is updated, review this guide before sharing the app so the instructions stay current with the tabs, buttons, and setup workflow.

## What It Does

- Copies the correct base database for the selected mode and modifies only the output copy
- Supports New Project, Upgrade Project, and Existing Project workflows
- Sets up project, searchable region/agency/reservation selection, and measurement-period records
- In upgrade mode, can import old plot/tree/regen inventory data into the selected current master database version
- In upgrade mode, can manually map renamed old fields or add old fields into selected upgrade tables
- Optionally imports plots from `.csv`, `.xlsx`, or `.xlsm`
- Sets up inventory crews, assignments, plot assignments, and optional upgrade-only tree assignments
- Uses the `AppColumns` tab to add custom plot/tree/regen measurement fields
- Syncs custom fields into `AppColumns` and optional `AppColumnCodes`
- Creates `AppColumnsArchive` in the saved project database before changing `AppColumns`
- Updates report headers/report columns
- Updates validation rule ProjectID and applicable periods
- Updates the Access Application Title in the saved project database
- Builds or refreshes `_prjCFIDEER_` get/update queries
- Exports the run log to a detailed text file beside the saved project database
- Can open the saved project database automatically after a successful build
- Includes optional AI Help for setup questions

## Main Tab Order

Project Pickle is arranged left to right as:

`Project`, `Periods`, `AppColumns`, `Upgrade Mapping`, `Plots`, `Get Field Order`, `CFIDEER Queries`, `Plot Assignments`, `Tree Assignments`, `Build`, and `AI Help`.

`Plot Assignments` and `Tree Assignments` appear after `CFIDEER Queries` so users can review AppColumns and query setup before rebuilding assignment tables.

Use the `Next >` button in the bottom-right status area to move through the tabs. It checks required entries on the current tab before moving forward.

When `Next >` is used on the `Plots` tab, Project Pickle moves directly to `Get Field Order` so users can order the active and newly created fields before the CFIDEER queries are rebuilt.

`Upgrade Mapping` and `Tree Assignments` are greyed out and skipped unless `Upgrade Project` mode is selected.

## Choosing A Mode

`New Project`

- Use when the tribe is starting a brand-new CFI project with no prior project database.
- Choose the current master database version. The Existing/old project database box is greyed out because there is no prior project database.
- Enter the new project information, create the first measurement period, and optionally add plots, AppColumns, reports, validation, assignments, and CFIDEER queries if those are ready.

`Upgrade Project`

- Use when an older CFI database is not in the current standard database format.
- Choose the current master database version.
- Choose the old database as the Existing/old project database.
- Project Pickle copies the current master database version, imports old plot/tree/regen data, uses Upgrade Mapping for fields that moved or changed, lets you add the new remeasurement period, fills related setup tables, and rebuilds tables/queries.

`Existing Project`

- Use when the project database is already in the current standard format.
- Choose the current project database as the Existing/old project database and create a new output copy. The current master database version box is greyed out because Existing Project mode updates a copy of the selected project database.
- Project name, cycle interval, Region, Agency, and Reservation are read-only/greyed out in this mode. Click `Load lists` if you want to display those existing values, but Project Pickle preserves the `Projects` row from the selected database during build instead of asking you to re-enter it.
- Use this mode only after the period structure already exists. Common reasons include adding new field data items, adding additional plots, arranging Get Field Order, rebuilding CFIDEER queries, rebuilding plot assignment tables, activating/deactivating AppColumns, and updating reports or validation rules.

The `Current master database version` box appears above the `Existing/old project database` box because it is the base file for `New Project` and `Upgrade Project` mode. In `New Project`, Project Pickle copies the current master database version as the starting database. In `Upgrade Project`, Project Pickle copies the current master database version, then imports data from the old project database. The `Existing/old project database` box is used only by `Upgrade Project` and `Existing Project` mode.

The `Keep Region, Agency, and Reservation from Existing Database` box is off and locked for New Project mode because new projects require manual Region, Agency, and Reservation selections. It is on by default for Upgrade Project mode so Project Pickle can use the location already stored in the old database, and it is on and locked for Existing Project mode because the existing database should already contain the correct location. In Existing Project mode, Project Pickle also preserves the existing project name and cycle interval instead of rewriting them from the Project tab.

## Run Log

Project Pickle writes build details to a detailed text log beside the file named in Save Project Location, using a name like `ProjectName.ProjectPickleRunLog.Failed.20260608_123456.txt`. It does not write a run log table into the Access database. If the save folder is not available yet, Project Pickle writes the text log in the `Project Pickle App Files` support folder.

Project Pickle writes the detailed text log during build cleanup, either after a successful run or after a handled failure. The Build tab is the live progress view while the app is running.

Each row or text-log line includes the step number, log time, run status or category, project mode, project name, Existing/old project database when used, current master database version when used, saved project database, and message.

Timing fields are included so slower parts of a build are easier to spot:

- `ElapsedSeconds`: seconds since the build started.
- `StepDurationSeconds`: seconds since the previous log row.
- `ElapsedTime`: readable elapsed time, such as `03:14.225`.
- `StepDuration`: readable duration since the previous log row.

Project Pickle also writes a final total run-time message before exporting the text run log. On failed builds, Project Pickle writes the detailed text log even if the saved project database does not exist.

During upgrade imports, Project Pickle logs each large table copy with the number of rows copied, elapsed time, and rows per second. Use those messages to see whether the app is still making progress and which table is taking the longest.

Long database builds can still take several minutes, especially large assignment or upgrade steps. Project Pickle refreshes the Build tab at log points, but a single large Access write may still briefly make Windows show `Not Responding`. If the Build tab continues to gain new lines, the app is still working.

If you close the app during a build and choose `Yes`, Project Pickle requests cancel and closes after the current safe checkpoint. It cannot safely interrupt Access in the middle of a database write.

After upgrade import, Project Pickle adds diagnostics for the main imported tables. These lines compare old and output row counts, compare measurement rows by period, check whether measurement rows have matching custom measurement rows, and check whether measurement rows have matching parent plot or tree rows. If something is missing, Project Pickle logs sample IDs so the records are easier to find.

During the upgrade repair step, Project Pickle checks whether imported plot/tree/regen measurement rows are missing parent `Plots` or `Trees` rows, then checks whether each measurement row has a matching custom measurement row. This step now logs each table it is checking and reports progress on large tables.

## AppColumns

The `AppColumns` tab is optional, but it can now be used two ways:

- Use the top box, `Load Standard AppColumn`, for the standard/current AppColumns list from the current master or current-format project database.
- Use the bottom box, `New custom AppColumns to add`, for new project custom fields.

Click `Load Standard AppColumn` after choosing the needed database path on the Project tab. This fills the top review grid from the current master/current-format `AppColumns` table. There is no separate Existing AppColumn list. Existing/project rows show in the same top grid only after they already exist in the selected database, usually because custom AppColumns were added during a previous Upgrade Project or Existing Project run. New custom rows stay in the lower box until you build; after that, they become existing AppColumns in future runs. Project Pickle loads those rows into the top review box so you can review the `Active` checkboxes and uncheck standard fields that do not pertain to the project. The loaded `App table` value and `StandardCode` checkbox are read-only in this top box.

Use the `Show` filter in the loaded review box to show `All fields`, `Active only`, or `Inactive only`. Use the `Data` filter for plot/tree/regen rows, and use `Search field` to find loaded rows by field name. These filters only change what you see in the grid; Project Pickle still remembers all loaded rows when building.

Project Pickle hides loaded fields that are always required, standard, calculated, or not useful to turn on and off. This includes standard regen fields such as `RemarksRegen` and `StemCount`.

Loaded standard rows update AppColumns settings. If an active loaded row belongs to `PlotCustomMeasurements`, `TreeCustomMeasurements`, or `RegenCustomMeasurements`, Project Pickle also makes sure the physical custom-table field exists so upgraded values have a place to land. New custom rows add the field to the selected custom measurement table and then sync AppColumns.

Only add or activate AppColumns for fields that will be collected in the current measurement and are not already in the standard list above. Older fields that still need to exist in the upgraded database for project history, but will not be collected now, should be created through the `Upgrade Mapping` step instead. If an older field has a different name and needs to land in a custom table, add that new custom AppColumns field first; then it can be selected in the `New table or table.field` column during Upgrade Mapping.

CFIDEER Get/Update queries are built from fields that are both `Active` and `QueryVisible` in the AppColumns setup. Fields created only through `Upgrade Mapping` can stay in the physical custom tables to preserve older database history, but they will not be added to CFIDEER queries unless they are also active AppColumns for the current measurement.

Before building with CFIDEER query refresh turned on, Project Pickle shows a confirmation listing the custom fields that are currently `Active + QueryVisible`. Those listed fields will be added to the `_prjCFIDEER` Get/Update queries. If an older history field appears in that popup, cancel the build and uncheck `Active` or `QueryVisible` before running. Fields that are inactive or not query-visible remain in the custom tables only for history.

After fields are added, Project Pickle orders the physical custom measurement tables so required key fields appear first, active AppColumns appear next, and inactive Upgrade Mapping history fields appear last. This keeps newly activated fields ahead of older inactive fields in `PlotCustomMeasurements`, `TreeCustomMeasurements`, and `RegenCustomMeasurements`.

Before Project Pickle changes `AppColumns`, it creates a fresh `AppColumnsArchive` table inside the saved project database. This archive is a snapshot of the original `AppColumns` table from the base database copied for that run. If `AppColumnsArchive` already exists in the output copy, Project Pickle replaces it with a new snapshot.

If custom fields are not known yet, load/review standard fields if needed, then leave new custom rows blank and build the database without them. Later, run Project Pickle again in `Existing Project` mode using the database that needs custom fields as the Existing/old project database, choose a new output copy, then return to this tab.

For new custom rows, choose the App table, field name, data type, and category. New `Field name` values should use PascalCase, such as `CountyCode`, with no spaces, underscores, or punctuation. Make sure `Active` is checked for fields that will be collected. `Active`, `ReportVisible`, and `QueryVisible` default on for new rows. If you check `Active` after it was off, Project Pickle also checks `ReportVisible` and `QueryVisible`. If you uncheck `Active`, Project Pickle also unchecks `ReportVisible` and `QueryVisible`.

To remove a new custom row before building, select the row in the bottom `New custom AppColumns to add` box and press `Delete`. Project Pickle shows a warning before removing it from the setup list. This does not remove fields from a database that has already been built.

To copy the setup format from an existing row, select one row in the top review box and click `Copy selected format`. Project Pickle copies only the custom table placement, data type, and category into a new row in the bottom custom box. It does not copy names, codes, validation abbreviations, or report/query settings.

Checkbox meanings:

**Active**

Field is on and collected for the current project. Inactive fields stay out of CFIDEER queries even if the physical field exists in a custom table. Checking `Active` also checks `ReportVisible` and `QueryVisible`; unchecking `Active` also unchecks them.

**ReportVisible**

Lets Project Pickle include the active field when updating `ReportHeaders` and `ReportColumns`.

**QueryVisible**

Lets Project Pickle include the active field when rebuilding CFIDEER get/update queries. `Active` must also be checked.

**StandardCode**

Read-only; appears only in the loaded review box and marks an existing `Code` category row as using a standard code/rule list.

For a typical custom field, leave `Active`, `ReportVisible`, and `QueryVisible` checked. Uncheck `ReportVisible` or `QueryVisible` only when that active field should stay out of reports or CFIDEER data-entry queries.

For a loaded standard field, the main review action is `Active`: leave it checked when the field applies to the project and uncheck it when the field should be left out of the active project setup.

## AppColumns Codes

For one custom field, put all allowed code values in that field's single `Codes` cell. Enter multiple codes as `code=label` pairs separated by semicolons:

```text
1=Yes; 2=No; 9=Unknown
```

Long code lists are entered the same way:

```text
0=None; 1=Trace; 2=Light; 3=Moderate; 4=Heavy; 5=Severe; 6=Dead; 7=Broken; 8=Missing; 9=Unknown; 10=Other; 11=Not applicable
```

Labels are optional, so this is also valid:

```text
0; 1; 2; 3; 4; 5; 6; 7; 8; 9; 10; 11
```

Pasted one-code-per-line lists are accepted too.

When Project Pickle loads standard AppColumns, code lists are displayed in numeric order, such as `0, 1, 2 ... 10 ... 100`, instead of text order.

The tree species code list for AppColumns is still updated manually in the Access database after Project Pickle finishes running. Project Pickle can build the AppColumns field setup, but the final tree species code list maintenance is still a manual database step.

## Measurement Cycle

`Measurement cycle interval (years)` fills the database `Projects.CycleLength` value. It is the planned number of years between full CFI measurements, not the number of rows in the Periods tab.

For a new project, enter the project cycle, usually `10` unless the project uses another interval. For an existing or remeasurement project, leave it blank to keep the value already in the existing database, or enter the correct cycle years if it needs to be changed.

## Region, Agency, And Reservation

On the Project tab, click `Load lists` after choosing the master or existing database. Project Pickle loads region, agency, and reservation choices from that database.

Choose Region first, then Agency, then Reservation name. The selected reservation supplies the `ReservationID`, `ReservationCode`, and linked `AgencyID` used for project setup and reference pruning.

Project Pickle automatically removes non-applicable region, agency, and reservation reference rows from the saved project database. This is the standard behavior for new, existing, and upgrade builds.

## Access Application Title

During the build, Project Pickle updates the Application Title in the saved project database, shown under Access Options > Current Database. In New Project and Upgrade Project mode, the `Project name` field on the Project tab is the name Project Pickle uses for this. In Existing Project mode, Project Pickle keeps the Project tab field read-only and preserves the existing Access Application Title from the selected database.

If the source title contains `[TRIBENAME]` and `[DATE]`, Project Pickle replaces `[TRIBENAME]` with the project name and `[DATE]` with the build date in `yyyyMMdd` format. For example, `[TRIBENAME] CFI [DATE] DB v3.06.00` becomes `Project Name CFI 20260602 DB v3.06.00`.

## Save Project Location Name Hint

When the Save Project Location Browse window opens before a real project name is entered, Project Pickle builds the suggested name from the current master database version for New/Upgrade mode or from the selected project database for Existing mode, but it leaves the database version as a fill-in placeholder:

```text
[TribeName] CFI [Date 20260520] vX.XX.XX.mdb
```

After a real project name is entered, Project Pickle uses that project name and the current build date, then keeps `vX.XX.XX` so the user can type the correct database version manually. This avoids accidentally pulling the version from an old project database.

## Periods

The `Periods` tab creates measurement-period rows for New Project and Upgrade Project work only. In Existing Project mode, the Periods tab is read-only/greyed out and skipped by the `Next >` workflow because the selected database should already contain the correct period structure.

For a New Project, enter the project information in row 1 with `PeriodNumber` 1 and check `IsCurrent`. Dates can stay blank if they are not known yet.

For an Upgrade Project, click `Load periods from database`, verify the information and number of periods, then add the new period for the remeasurement and check `IsCurrent`. A project with five previous measurements and a new sixth measurement should have periods 1 through 6.

Do not use Existing Project mode to add a new measurement period. The standard Existing Project workflow is: add new field data items or additional plots, review Get Field Order, rebuild CFIDEER Queries, rebuild Assignments if needed, then Build.

The `Load periods from database` button on the Periods tab is enabled only in `Upgrade Project` mode. It reloads period rows from the old project database. It is disabled in `New Project` mode so sample period rows from a master database are not imported.

`ApplicablePeriod` is used when updating `ValidationRules.ApplicablePeriods`. It is usually the measurement year or period identifier used by the project.

Only one `IsCurrent` period can be checked at a time. When you check a period as current, Project Pickle unchecks the other period rows. When you add a new period row by hand, Project Pickle checks that new row as current automatically. For a multiperiod upgrade/remeasurement project, check the new measurement period you are setting up.

After Project Pickle writes the `ProjectMeasurementPeriods` rows, it updates every `Analysis*` table that has a `ProjectPeriodID` field so it uses the current period's newly created `ProjectPeriodID`.

## Plots

The `Plots` tab is used to add new plots and import plot data from a previous project.

For a New Project, add plots from a spreadsheet. Project Pickle sets up the `Plots` table and creates initial `PlotMeasurements` and `PlotCustomMeasurements` records.

For an Upgrade Project, Project Pickle imports plots from the previous database into the `Plots`, `PlotMeasurements`, and `PlotCustomMeasurements` tables. It also creates new records for the current measurement in the `PlotMeasurements` and `PlotCustomMeasurements` tables.

For Adding New Plots in Upgrade or Existing Project mode, upload a spreadsheet using plot number and any additional information provided. Project Pickle creates blank records for the existing period rows already in the selected project setup.

In Existing Project mode, this step is optional. If no new plots need to be added, leave `Add New Plots` blank and continue to the next needed tab. The uploaded/current-format project database already contains all existing plots, and Project Pickle can use those existing plot rows for assignments and query/report updates.

The `Add New Plots` file must include a column that contains the plot number, but the spreadsheet header does not have to be named `PlotNumber`. After choosing the file on the `Plots` tab, click `Load columns`, then map each Project Pickle field to the spreadsheet column that contains that value. `PlotNumber` is required; optional mapped fields include `PlotLabel`, `PlotTypeID`, `UTMNorthingCoordinate`, `UTMEastingCoordinate`, `UTMZone`, and `FLCCommercial`.

For an existing project that already has plots, especially a database that was already built through Upgrade Project mode, upload a spreadsheet with only the new plots. You do not need to include plots that are already in the database. A mixed list of existing and new plots is allowed, but not required. The plot import is additive by `PlotNumber`: Project Pickle skips rows whose `PlotNumber` already exists and inserts only new plot numbers. For each new plot, it creates blank `PlotMeasurements` and matching `PlotCustomMeasurements` rows for the period rows already in the selected project setup. If assignment setup is checked, assignments are rebuilt for all plots in the saved project database, including existing plots and newly imported plots.

The `Import Plot, Tree, Regen, Assignment Data from the previous project` checkbox is Upgrade-only. It is checked automatically in Upgrade Project mode and disabled in Existing Project mode on purpose. After a database has already been upgraded, do not use old-data import again; use `Add New Plots` to add only the new plot spreadsheet.

If the build says `Add New Plots row ... is missing PlotNumber`, check two things: click `Load columns` and map `PlotNumber` to the spreadsheet column containing plot numbers, and remove or ignore blank rows in the spreadsheet. Project Pickle skips fully blank spreadsheet rows, but a row with other mapped data still needs a plot number.

Steps to add plots in Existing or Upgrade mode:

1. Choose the spreadsheet in `Add New Plots`.
2. Click `Load columns`.
3. Map `PlotNumber` to the spreadsheet column that contains plot numbers.
4. Optionally map `PlotLabel`, `PlotTypeID`, coordinates, `UTMZone`, or `FLCCommercial`.
5. Leave unused optional mappings blank.
6. Continue to `Get Field Order`, `CFIDEER Queries`, assignments if needed, then `Build`.

When plot data is available, run Project Pickle again in `Existing Project` mode using the database that contains the plot rows as the Existing/old project database, choose a new output copy, then use the assignment tabs to build assignment tables.

When an existing or upgraded project has plots from previous measurements and a new measurement period is added, Project Pickle creates blank `PlotMeasurements` rows for missing plot/period combinations. For example, if the old database has periods 1 through 5 and the setup adds period 6, Project Pickle adds one blank period 6 plot measurement row for each plot. It fills `MeasurementID`, `PlotMeasKey`, `PlotID`, `PlotKey`, `ProjectPeriodID`, `PeriodNumber`, and `UpdUser=Pickle`; other measurement fields are left blank. Project Pickle also creates the matching `PlotCustomMeasurements` row with the same `MeasurementID`.

In `Upgrade Project` mode, the `Import Plot, Tree, Regen, Assignment Data from the previous project` box is turned on automatically. In `New Project` and `Existing Project` mode, it stays unchecked because there is no old project database to upgrade into the current structure. Project Pickle copies old inventory rows into the current master database version output before building assignments. If a field moved between related plot/tree/regen tables, Project Pickle can map it by same-name moved-field logic or by choices in the `Upgrade Mapping` tab. For example, old `PlotMeasurements` fields can be copied into `PlotCustomMeasurements`, old `TreeCustomMeasurements` fields can be copied into `TreeMeasurements`, and old `PlotCustomMeasurements` fields can be copied into `Plots` when that is where the field belongs in the current standard structure. If an old measurement row has a period ID Project Pickle cannot match directly, Project Pickle uses that row's `PeriodNumber` to attach it to the matching output period. Fields with no matching destination are skipped and listed in the run log instead of stopping the whole build. If Access rejects one field value during import, Project Pickle now tries to keep the row and logs the field that could not be copied.

Same-name moved-field mapping normally uses bulk Access queries so large upgrades do not have to copy those values one row at a time. The run log will say `Upgrade moved-field bulk mapping` when that faster route works. If Access rejects the bulk query because of an unusual old table shape or field type, Project Pickle logs that and automatically uses the slower row-by-row fallback.

Some older project databases may have measurement rows whose parent inventory rows are missing. During upgrade, Project Pickle checks for `PlotMeasurements`, `TreeMeasurements`, and `RegenMeasurements` rows that point to missing `Plots` or `Trees` rows. When enough information exists, Project Pickle rebuilds the missing parent row, marks `UpdUser=Pickle`, and writes the repaired `PlotID`, `TreeID`, `MeasurementID`, keys, and numbers to the run log so the rows can be found later. Project Pickle also creates missing matching custom measurement rows when needed.

The `Upgrade Mapping` tab matches older database field names to the new standard field names and adds older fields that are not collected in the current measurement period into the custom tables, preserving them as part of the project history. Click `Load unmapped fields` after choosing the old project database and current master database version. Project Pickle lists old plot/tree/regen fields that do not exist in the same new table and do not have a same-name moved-field destination.

Use Upgrade Mapping for older fields that should stay in the database but are not collected in the current measurement. If the old field should be renamed into a custom table, create that custom field on the `AppColumns` tab first so the exact `Table.Field` target is available in the mapping dropdown.

Upgrade Mapping only runs in `Upgrade Project` mode. If the run log says `Mode = Existing Project`, Project Pickle did not run the upgrade import or field-moving step. Change the Project tab mode to `Upgrade Project`, choose the old project database as the Existing/old project database, choose the current master database version, then build again.

For each field that should carry forward, choose either a target table or an exact `Table.Field` target. The target dropdown puts the most common table-only choices near the top: `PlotCustomMeasurements`, `TreeCustomMeasurements`, `Plots`, `Trees`, `PlotMeasurements`, and `TreeMeasurements`. Project Pickle can map within the same data family, such as plot table to plot custom table, tree custom table to tree measurement table, or plot custom table to plots. Project Pickle skips fields named `NeedsReview` and common standard fields that should not be manually mapped. Those `NeedsReview` fields come from the old database table layout; Project Pickle does not create them.

Click `Check data types` before building an upgrade. Project Pickle compares old source values against the selected current master database version, active custom fields, automatic moved-field targets, and Upgrade Mapping choices. It flags cases such as old `Text` values going into a new `Integer`, `Date/Time`, `Boolean`, or `Guid` field. Build now stops for these datatype mismatches even when the current values look convertible, because fields such as old text `HabitatType` should be crosswalked into the new integer code before they are written to the standard new field. Fix those by cleaning/crosswalking the old data, choosing a `Text`/`Memo` target, or mapping to a new custom field that uses the old field's data type.

For example, old `Plots.HabitatType` defaults to `PlotMeasurements.HabitatType`. If the old field is `Text` and the current master database version field is `Long Integer`, convert/crosswalk the old HabitatType values before building. If the crosswalk is not ready, change the Upgrade Mapping target to `PlotCustomMeasurements` by table name, or to another Text custom field, so the old text is preserved and can be crosswalked later.

When Project Pickle loads the review list, it preselects likely targets:

- Old `TreeRemarks` defaults to `TreeMeasurements.RemarksTree`.
- Old `PlotRemarks` defaults to `PlotMeasurements.RemarksPlot`.
- Old regen `Remarks` defaults to `RegenMeasurements.RemarksRegen`.
- Old `Plots.HabitatType` defaults to `PlotMeasurements.HabitatType`.
- Old `FieldOverCover` defaults to `PlotMeasurements.CoverType`.
- Old `FieldOverDensity` defaults to `PlotMeasurements.DensityClass`.
- Old `FieldOverSize` defaults to `PlotMeasurements.SizeClass`.
- Old `ForestClassification` defaults to `Plots.FLCCommercial`.
- Old unmapped `TreeMeasurements` fields default to `TreeCustomMeasurements`.
- Old unmapped `Plots` and `PlotMeasurements` fields default to `PlotCustomMeasurements`.
- Old `PlotCustomMeasurements` and `TreeCustomMeasurements` fields default back to their matching custom table.

- Choose a table, such as `PlotCustomMeasurements`, `TreeCustomMeasurements`, `Plots`, or `Trees`, when the old field should be added to that table using the old field name. Project Pickle creates the missing field and copies the old data. When source and target tables use different row keys, Project Pickle uses related IDs and keys such as `MeasurementID`, `PlotMeasKey`, `TreeMeasKey`, `RegenMeasKey`, `PlotID`, `PlotKey`, `TreeID`, and `TreeKey` to find the matching rows.
- Choose `Table.Field`, such as `TreeCustomMeasurements.MappedTreeScore`, when the old field should be renamed while it is imported.
- Choose `Skip` for fields that should not carry forward. `Skip` means Project Pickle will not create that old field in the output custom tables and will not copy data for that field. Use it when the old field is empty, has no useful data, or should not be preserved in the upgraded database.

Upgrade Mapping copies old values into the selected output table and can add the physical field when needed. It does not turn the field on in AppColumns. If a mapped field should be active for data entry, reports, or CFIDEER queries, add or activate that field on the `AppColumns` tab separately.

`Save mapping` is optional. It saves the current grid to a JSON file so the same choices can be reused for similar upgrades. If you do not save, Project Pickle still uses the mappings currently shown in the grid when you click Build. If you close Project Pickle before building, unsaved mapping choices are lost.

`Load saved mapping file` opens a file picker because it is only for reusing a mapping JSON file that was created earlier with `Save mapping`. It is not used to choose the old project database, the current master database version, or a plot spreadsheet. For normal upgrades, start with `Load unmapped fields`.

## Get Field Order

The `Get Field Order` tab controls the field display order in the three CFIDEER Get queries used by data collection:

- `_prjCFIDEER_GetPlotMeasurementsForPeriod`
- `_prjCFIDEER_GetTreeMeasurementsForPeriodByKey`
- `_prjCFIDEER_GetRegenMeasurements`

Click `Load fields from AppColumns setup` after reviewing AppColumns. Project Pickle first tries to keep the saved Get query order from the selected database. Then it adds any newly added or newly activated Active + Query Visible AppColumns at the bottom of the list. This lets users keep the current database order and move only the new fields into the correct location.

If the selected database does not have the saved Get query yet, Project Pickle falls back to the field order in the selected table structure. If AppColumns has not been loaded yet, Project Pickle can still use the database selected for the current mode as a fallback field source. Select a row and use `Move up` or `Move down` to set the display order.

Required trailing fields stay locked at the end even if a user tries to move them: `RemarksPlot`, `RemarksTree`, and `RemarksRegen`. Criteria-only fields such as `PlotID`, `ProjectPeriodID`, `TreeKey`, and `MeasurementID` are not shown in this tab.

Fields you do not move are still included in the Get query as long as they are active/query-visible in AppColumns. Project Pickle keeps those unmoved fields in the saved database order where possible, with new or newly activated fields added at the bottom.

## CFIDEER Queries

### Summary

Develops Get/Update Queries used by CFIDEER/data-entry workflows to retrieve and update plot, tree, and regeneration data.

### How It Works

Project Pickle reads the current measurement tables and selected custom fields from the AppColumns tab, generates the `_prjCFIDEER` get/update SQL, and replaces the existing working query definitions.

### Important

This updates only the output copy. It overwrites existing saved queries without creating backups. Always review the generated queries before using them for production data entry.

The `Build/refresh _prjCFIDEER_ get and update queries` option is checked by default in every mode so the saved project database queries are automatically rebuilt to match the current project setup. Leave it checked for normal runs. Uncheck it only when existing saved `_prjCFIDEER_` queries should be left untouched.

To add or remove fields from the queries, run Project Pickle in `Existing Project` mode, choose the current project database as the Existing/old project database, choose a new output copy, update `Active` and `QueryVisible` on the AppColumns tab, then leave the query rebuild option checked on the `CFIDEER Queries` tab before building.

For `_prjCFIDEER_GetProcessPlotList`, `_prjCFIDEER_GetProcessTreeList`, and `_prjCFIDEER_GetProcessRegenList`, Project Pickle now removes any prior explicit custom-table select items from the saved process-list query, then adds back only fields that are active and query-visible in the AppColumns setup. This keeps older fields created through Upgrade Mapping in the custom tables for history, but stops the process-list queries at the last field collected for the current measurement.

For `_prjCFIDEER_GetPlotMeasurementsForPeriod`, `_prjCFIDEER_GetTreeMeasurementsForPeriodByKey`, and `_prjCFIDEER_GetRegenMeasurements`, the same rule applies: only fields that are both `Active` and `QueryVisible` are included. Upgrade Mapping history fields should usually stay inactive and/or not query-visible.

For `_prjCFIDEER_GetPlotMeasurementsForPeriod`, Project Pickle starts the parameter declarations with `PlotID Guid` and `ProjectPeriodID Guid`, followed by `SELECT`. Its `FROM` block uses `PlotMeasurements LEFT JOIN PlotCustomMeasurements ON PlotMeasurements.[MeasurementID] = PlotCustomMeasurements.[MeasurementID]`, and its criteria check `PlotMeasurements.PlotID = [PlotID]` before `PlotMeasurements.ProjectPeriodID = [ProjectPeriodID]`. `PlotMeasurements.PlotID` and `PlotMeasurements.ProjectPeriodID` are criteria-only fields, so their Access query design `Show` boxes are unchecked by default. The select list does not include `CalcSiteIndex`, `GMP`, `Created`, `Updated`, `UpdUser`, or `RetrievalID`, and `RemarksPlot` is always the last selected field before the criteria-only PlotID/ProjectPeriodID fields.

For `_prjCFIDEER_GetTreeMeasurementsForPeriodByKey`, Project Pickle starts the parameter declarations with `TreeKey Text (255)` and `PeriodID Guid`, followed by `SELECT`. Its `FROM` block uses `TreeMeasurements LEFT JOIN TreeCustomMeasurements ON TreeMeasurements.MeasurementID = TreeCustomMeasurements.MeasurementID`, and its criteria check `TreeMeasurements.TreeKey = [Treekey]` before `TreeMeasurements.ProjectPeriodID = [PeriodID]`. `TreeMeasurements.TreeKey` and `TreeMeasurements.ProjectPeriodID` are criteria-only fields, so their Access query design `Show` boxes are unchecked by default. The select list does not include `Created`, `Updated`, `UpdUser`, or `RetrievalID`, and `RemarksTree` is always the last selected field before the criteria-only TreeKey/ProjectPeriodID fields.

For `_prjCFIDEER_GetRegenMeasurements`, Project Pickle starts the parameter declarations with `RegenID Guid`, followed by `SELECT`. It intentionally leaves out `RegenMeasurements.PlotNumber`, `RegenMeasurements.PlotID`, `RegenMeasurements.ProjectPeriodID`, `IsDeleted`, `Created`, `Updated`, `UpdUser`, and `RetrievalID`, and filters by `RegenMeasurements.MeasurementID = [RegenID]`. `RegenMeasurements.MeasurementID` is criteria-only, so its Access query design `Show` box is unchecked by default. `RemarksRegen` is always the last selected field before the criteria-only `MeasurementID` field.

For `_prjCFIDEER_UpdatePlotMeasurement`, Project Pickle starts the update fields with `PlotMeasurements.PlotMeasKey`, `PlotMeasurements.PlotKey`, and `PlotMeasurements.PeriodNumber`. Other in-house apps depend on those fields staying at the start of that update query. Project Pickle does not include `CalcSiteIndex` or `GMP` in this update query so the update field count stays aligned with the Get Plot measurement query.

For `_prjCFIDEER_UpdatePlotMeasurement`, Project Pickle starts the parameter declarations with `PlotID Guid`, `PeriodID Guid`, `PlotMeasKey Text (255)`, and `PlotKey Text (255)`.

For `_prjCFIDEER_UpdatePlotMeasurement`, Project Pickle starts the update SQL body with `UPDATE PlotMeasurements`, then `INNER JOIN PlotCustomMeasurements ON PlotMeasurements.MeasurementID = PlotCustomMeasurements.MeasurementID`, then a multiline `SET` block beginning with `PlotMeasurements.PlotMeasKey`, `PlotMeasurements.PlotKey`, and `PlotMeasurements.PeriodNumber`.

For `_prjCFIDEER_UpdatePlotMeasurement`, Project Pickle also ends the update grid with `PlotMeasurements.Created`, `PlotMeasurements.Updated`, `PlotMeasurements.UpdUser`, `PlotMeasurements.RetrievalID`, `PlotCustomMeasurements.PlotMeasKey`, then the criteria fields `PlotMeasurements.PlotID` and `PlotMeasurements.ProjectPeriodID`.

For `_prjCFIDEER_UpdateTreeMeasurement`, Project Pickle starts the update fields with `TreeMeasurements.TreeMeasKey`, `TreeMeasurements.TreeKey`, and `TreeMeasurements.PeriodNumber`.

For `_prjCFIDEER_UpdateTreeMeasurement`, Project Pickle starts the parameter declarations with `TreeID Guid`, `PeriodID Guid`, `TreeMeasKey Text (255)`, `TreeKey Text (255)`, and `PeriodNumber Short`.

For `_prjCFIDEER_UpdateTreeMeasurement`, Project Pickle starts the update SQL body with `UPDATE TreeMeasurements`, then `INNER JOIN TreeCustomMeasurements ON TreeMeasurements.MeasurementID = TreeCustomMeasurements.MeasurementID`, then a multiline `SET` block beginning with `TreeMeasurements.TreeMeasKey`, `TreeMeasurements.TreeKey`, and `TreeMeasurements.PeriodNumber`.

For `_prjCFIDEER_UpdateTreeMeasurement`, Project Pickle does not include `Created` in the update grid or SQL. It ends the update grid with `TreeMeasurements.Updated`, `TreeMeasurements.UpdUser`, `TreeMeasurements.RetrievalID`, `TreeCustomMeasurements.TreeMeasKey`, then the criteria fields `TreeMeasurements.TreeID` and `TreeMeasurements.ProjectPeriodID`. `Updated` updates from `[UpdatedDate]`, and `ProjectPeriodID` uses the `[PeriodID]` criteria parameter.

For `_prjCFIDEER_UpdateRegenMeasurement`, Project Pickle starts the update fields with `RegenMeasurements.RegenMeasKey`, `RegenMeasurements.PlotID`, `RegenMeasurements.PlotKey`, and `RegenMeasurements.ProjectPeriodID`. The `ProjectPeriodID` field updates from the `[PeriodID]` parameter. Project Pickle does not include `IsDeleted` in this update query because including it can prevent regen data from opening or saving back correctly.

For `_prjCFIDEER_UpdateRegenMeasurement`, Project Pickle starts the parameter declarations with `RegenMeasurementID Guid`, `RegenMeasKey Text (255)`, and `PlotID Guid`.

For `_prjCFIDEER_UpdateRegenMeasurement`, Project Pickle starts the update SQL body with `UPDATE RegenMeasurements`, then `INNER JOIN RegenCustomMeasurements ON RegenMeasurements.MeasurementID = RegenCustomMeasurements.MeasurementID`, then a multiline `SET` block beginning with `RegenMeasurements.RegenMeasKey`, `RegenMeasurements.PlotID`, `RegenMeasurements.PlotKey`, and `RegenMeasurements.ProjectPeriodID`.

For `_prjCFIDEER_UpdateRegenMeasurement`, Project Pickle ends the update grid with `RegenCustomMeasurements.RegenMeasKey`, `RegenMeasurements.Created`, `RegenMeasurements.Updated`, `RegenMeasurements.UpdUser`, `RegenMeasurements.RetrievalID`, then the criteria field `RegenMeasurements.MeasurementID`. `Created` updates from `[CreatedDate]`, `Updated` updates from `[UpdatedDate]`, and `MeasurementID` uses the `[RegenMeasurementID]` criteria parameter.

Use this after custom fields are final enough for data entry. If AppColumns/custom fields are not known yet, leave this off and run Project Pickle later in `Existing Project` mode after those fields are added or removed.

## Plot Assignments

The `Plot Assignments` tab builds `InventoryCrews`, `InventoryAssignments`, and `InventoryAssignmentPlots`. This requires plots in the database or an `Add New Plots` file on the `Plots` tab.

When plot assignment setup is checked, Project Pickle treats assignment setup as a rebuild. If the output copy already has rows in `InventoryAssignmentTrees`, `InventoryAssignmentPlots`, `InventoryAssignments`, or `InventoryCrews`, Project Pickle clears those rows first and then builds a clean assignment set. This prevents repeated Existing Project or Upgrade Project runs from stacking duplicate crews or duplicate assignment rows.

Assignments are built for the current period only. If no period is marked current, Project Pickle uses the highest period number and writes that choice to the run log. For a remeasurement project with periods 1 through 6, check period 6 as current before building assignments.

After assignment setup runs, Project Pickle now verifies the final row counts. If plots exist and plot assignment setup is checked, `InventoryCrews`, `InventoryAssignments`, and `InventoryAssignmentPlots` must contain rows or the build fails with a clear message. If tree assignment setup is checked and tree rows exist, `InventoryAssignmentTrees` must also contain rows. The run log reports the created and final row counts.

Every rebuilt `InventoryAssignmentPlots` row is set to `StatusID = 1`, which is the Awaiting Work code. Every rebuilt `InventoryAssignmentTrees` row is set to `AssignmentStatusID = 1`, also Awaiting Work. Project Pickle verifies those statuses after the rebuild and stops if any assignment row does not have the required value.

## Tree Assignments

The `Tree Assignments` tab builds `InventoryAssignmentTrees`. It is enabled only in `Upgrade Project` mode and requires existing tree rows. If tree assignment setup is turned on, Project Pickle also turns on plot assignment setup because tree assignment rows attach to plot assignment rows.

## AI Help

The `AI Help` tab is optional and does not change the database. It lets users ask setup questions, such as how to enter AppColumns codes, when to build assignments, or how to use Existing Project mode.

When used, Project Pickle sends the question, the Project Pickle README, relevant snippets from `the internal app script`, and optional guide/reference snippets to the approved Azure OpenAI endpoint. Supported guide/reference file types are DOCX, XLSX, XLSM, TXT, MD, CSV, JSON, and PS1. PDF and old XLS uploads are not supported in-app because they require Office automation.

The default endpoint is `https://doi-training-foundry.cognitiveservices.azure.com/` and the default model/deployment is `gpt-4.1-mini`. IT can approve additional exact host names with the `CFI_PROJECT_PICKLE_APPROVED_AI_HOSTS` environment variable.

## Troubleshooting Build Errors

If Project Pickle says `Please fix these entries before building`, follow the listed items. Those are missed or invalid setup entries and Project Pickle has not started changing the output copy yet.

For database write errors, check the Build tab log and the popup. The popup names the table Project Pickle was writing to and gives a checklist.

Project Pickle uses parameterized Access writes for setup rows so Access handles database IDs, dates, blanks, numbers, and text safely. This prevents common `Syntax error in INSERT INTO statement` failures caused by hand-built Access SQL values.

Project Pickle no longer writes `ProjectPickleRunLog` into the saved project database. Use the detailed text run log named in the build popup for troubleshooting.

Common fixes:

- Project tab: choose the database required by the selected mode and Save Project Location. For New Project, choose the current master database version, enter Project name, and select Region, Agency, and Reservation name manually. For Upgrade Project, choose both the old project database and the current master database version. For Existing Project, choose the current-format project database; project name, cycle interval, Region, Agency, and Reservation are preserved from that database.
- Periods tab: use only in New Project or Upgrade Project mode. Enter at least one numeric `PeriodNumber`. Dates can be blank, but typed dates must be valid. Existing Project mode keeps this tab read-only and preserves the period rows already in the selected database.
- AppColumns tab: use one complete row per custom field. If `Category` is `Code`, the `Codes` cell must contain the allowed codes.
- Plots and Assignments tab: leave `Add New Plots` blank if plot data is not known yet. Only turn on assignment setup when plots exist or an `Add New Plots` file is selected.
- CFIDEER Queries tab: query rebuild is on by default in every mode. Leave it checked for normal runs so queries match the current setup; uncheck it only if saved queries should stay untouched.

If the error is still unclear, open the detailed text run log named in the popup. Send that file, or at least the last 20 lines. The detailed log is the best place to check whether a large upgrade lost rows for a specific measurement period, whether custom measurement rows were missing, or whether parent plot/tree rows had to be rebuilt.

## Safety

- The master template and original existing-project database are never edited.
- Existing output files are moved to a timestamped backup before replacement.
- Plot assignment setup is skipped when plots are not known yet.
- Tree assignment setup is optional and intended for remeasurement projects with tree rows.
- CFIDEER query refresh replaces the working queries in the output copy without creating separate backup queries.
- AI Help is optional, advisory, and restricted to approved HTTPS Azure OpenAI hosts.

## Self-Test

To check the builder core without opening the UI:

```powershell
.\Run-ProjectPickle-32bit.bat -SelfTest
```






