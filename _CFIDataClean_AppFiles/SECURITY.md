# CFI DataClean Notes

These notes summarize the current security posture for internal DOI/BIA review. They are not a formal authority-to-operate package, but they identify the intended data flow, controls, and remaining deployment items.

## Scope

CFI DataClean is a Windows PowerShell/WinForms desktop utility for reviewing 32-bit Microsoft Access forest inventory databases and exporting data-quality findings to Excel.

The app runs locally on the user's workstation. It does not host a web server, open a listening port, require administrator privileges, or install a service.

## Local Database Handling

- The user selects an Access `.mdb` or `.accdb` file.
- The app opens the file through local OleDb Access/ACE/Jet drivers.
- The main **Run** workflow copies the selected database to a temporary CFI DataClean working file, runs checks against that temporary copy, and exports findings to a separate Excel workbook.
- The original selected Access database is not modified by **Run**.
- Temporary database copies and Excel staging folders are deleted after normal runs. Stale CFI DataClean temp files are also cleaned at startup.
- The exported workbook includes a `Run Data` tab with run settings, paths, and counts for review traceability. It does not include the AI API key.
- Each **Run** also creates a matching text run log beside the selected workbook path. The log records run status, selected paths, run options, major progress steps, elapsed time, and failure details when a run fails before the workbook can be created. Failed Access write steps include a shortened SQL preview for troubleshooting. The app redacts the AI API key and database password value if either appears in log text.

## AI Data Flow

The AI helper is optional and off by default.

When AI is used, CFI DataClean sends compact text context to the configured Azure AI endpoint. Depending on the selected AI action, that context can include:

- table names and row counts,
- selected table column names,
- limited sample rows,
- grouped audit findings,
- coded cleaning rules,
- template code metadata from `AppColumns` / `AppColumnCodes`,
- field-manual snippets or field-focused manual digests,
- the user's chat question.

CFI DataClean does not upload the full Access database file to the AI endpoint.

This build only allows the approved DOI Azure AI endpoint hosts:

```text
doi-training-foundry.services.ai.azure.com
doi-training-foundry.cognitiveservices.azure.com
```

Requests use HTTPS/TLS 1.2. The endpoint box may contain the base endpoint or an Azure deployment chat-completions URL, but the host must match one of the approved hosts above.

## API Key Handling

- The AI API key is typed by the user in the AI helper dialog.
- The key is held in app memory while the window is open.
- The key is not saved to a config file, README, workbook, database, or registry.
- The key is not written to the text run log.
- Background runs pass settings in memory to the runspace; the app does not create a background-run settings JSON file containing the key.
- API keys are granted through the BIA AI point of contact and are unique to the Azure AI environment. Users should not share keys with anyone who does not have access to the approved endpoint.

Crash dumps, workstation memory capture, or endpoint-side logging are outside the app's direct control and should be assessed by IT under normal endpoint-management policy.

## Manual File Handling

Users may upload project manuals or guides for AI context and built-in guidance extraction.

Supported manual inputs are `.pdf`, `.docx`, `.xlsx`, `.xlsm`, `.xls`, `.txt`, `.md`, `.csv`, and `.json`.

Security-relevant behavior:

- `.docx`, `.xlsx`, and `.xlsm` are read through Open XML package parsing when possible; macros are not executed by that parser.
- Legacy `.xls` files use local Excel automation in read-only mode.
- PDF fallback may use local Word automation to extract text when simple stream extraction is insufficient.
- Word/Excel automation is opened hidden, read-only, with macro automation disabled where supported.

For the strictest deployment posture, IT may choose to allow only text-based/manual formats such as `.pdf` with selectable text, `.docx`, `.xlsx`, `.txt`, and `.md`, and block legacy `.xls`.

## Script Execution

The preferred launcher is the portable `CFIDataClean.cmd` file in the main release folder. It starts `_CFIDataClean_AppFiles\CFIDataClean.vbs`, which uses Windows Script Host to start the 32-bit Windows PowerShell host without showing a console window. The support files are kept in `_CFIDataClean_AppFiles` so most users only need to see the main launcher and documentation.

The old `.lnk` shortcut is not used for shared releases because Windows shortcuts can store user-specific absolute paths. A visible `_CFIDataClean_AppFiles\CFIDataClean-debug.cmd` launcher is retained for troubleshooting startup errors. Both launch paths use the 32-bit Windows PowerShell host when available so the app can access 32-bit Access database drivers.

The launch commands use:

```text
-ExecutionPolicy Bypass
```

This is a per-process launcher setting. It does not change the user's machine-wide PowerShell execution policy. It is used so portable/shared release folders do not fail with "not digitally signed" when the script was copied from another machine, downloaded from OneDrive, or extracted from a ZIP.

Recommended deployment hardening:

- Code-sign `_CFIDataClean_AppFiles\ForestInventoryCleaner.ps1` with an organization-approved certificate.
- Distribute the script and launchers from an approved internal share or software center.
- Preserve file hashes for release tracking.
- If agency policy requires `AllSigned`, use the signed script and remove or replace the per-process execution-policy argument from the launcher. A machine-enforced Group Policy may still block unsigned scripts even when the launcher uses `Bypass`.

## SQL and File Safety Notes

- Table and field names used in generated SQL are bracket-quoted.
- Text values inserted into generated SQL are single-quote escaped.
- The main Run workflow operates on a temporary copy of the database.
- The generated Excel workbook includes verification SQL for analyst review; it is intended to be run manually by the user in Access.

The app still relies on the Access database file being trusted enough to open locally with OleDb drivers. Treat unknown Access files and unknown Office manuals as untrusted until reviewed under agency policy.

## Residual Review Items

Before broad deployment, DOI/BIA IT should review:

- whether the approved Azure AI endpoint is authorized for the sensitivity level of the inventory data,
- whether sample rows and audit findings are allowed to be sent to that endpoint,
- whether legacy `.xls` manual upload should remain enabled,
- whether the PowerShell script should be packaged or signed through an internal deployment system,
- whether endpoint logging, proxy logging, or Azure logging stores prompts/responses and under what retention policy.

## Current Hardening Summary

- AI helper is off by default.
- AI endpoint hosts are locked to the approved DOI Azure hosts.
- API key is memory-only and not written to run temp files.
- Text run logs are saved beside the selected workbook path for troubleshooting and are designed not to include the API key.
- Run uses a temporary copy and leaves the original database unchanged.
- Stale CFI DataClean temp files are cleaned at startup.
- Office automation disables macro automation where supported.
- The launchers use a per-process `ExecutionPolicy Bypass` to avoid unsigned-script launch failures for portable shared folders; managed deployment should replace this with organization code signing or packaging when required.
