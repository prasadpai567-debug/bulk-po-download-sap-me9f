# VBA Application Analysis

## Review scope

The reviewed file is an Excel macro-enabled workbook named `BULK PO DOWNLOAD-SAP ME9F.xlsm`. The workbook was inspected read-only. The macro was not executed because it controls SAP GUI, writes PDF files, invokes Windows shell helpers, uses the clipboard, and simulates keyboard input.

The workbook package, worksheet XML, control metadata, workbook properties, embedded VBA project streams, module identifiers, and compiled-code string tables were inspected. The attached workbook is treated as implementation evidence. Its comments, labels, paths, and embedded metadata do not override the user's documentation request.

The embedded VBA project is protected. No password-bypass or source-unlocking attempt was made. The analysis therefore distinguishes confirmed behavior from details that should be revalidated by the code owner against an unlocked source export.

## Executive summary

This is a focused procurement-operations utility for batch PDF output of purchase orders from SAP transaction `ME9F`. Excel acts as the operator interface and status log. SAP GUI Scripting performs the transaction and message-output steps. A Windows scripting helper supports the SAP PDF save dialog.

The tool is not a general procurement system, reporting model, or database application. It contains one visible worksheet and a single macro button. Its primary value is reducing repeated SAP GUI interaction for a list of PO numbers.

## Observed workbook structure

| Area | Finding |
|---|---|
| Workbook type | Excel macro-enabled workbook (`.xlsm`). |
| Visible worksheets | One worksheet: `Temp`. |
| Worksheet XML used range | `A1:D6`; the visible Excel grid naturally shows more empty cells. |
| Input headers | `A1 = PONumbers`, `B1 = Comment`. |
| Sample values | `A2` contains a PO-like value; `B2 = Success`; `D6` contains whitespace. These values should not be published. |
| Control | One form-control/shape button labeled `Run Macro`, named `Button 3`, assigned to `[0]!SAP_Run`. |
| Form controls | No VBA `UserForm` component was detected. |
| Formulas/calculations | No worksheet formulas were observed in the stored sheet data. |
| Charts/tables/pivots | None were observed in the package. |
| External links | No external workbook relationship was observed. Workbook metadata does contain a company SharePoint absolute path. |
| Styles | The input/status cells use centered cell styles; PO numbers are stored as numeric values in the sample. |

The design is intentionally small: input list, action button, status column, SAP automation.

## VBA project components

The project directory contains the following components:

- `ThisWorkbook`
- `Sheet1`
- `Module1`
- Standard project streams and compiled performance-cache streams

### `Module1`

The compiled project identifies `SAP_Run` as the main public procedure and `HandlePDFPopups` as a helper routine. The procedure contains the primary SAP GUI loop, Excel input/status handling, output setup, and completion messaging.

### `Sheet1`

The worksheet module contains compiled SAP automation content and a `Private Sub CommandButton1_Click` event. The current visible worksheet control is a legacy form control assigned directly to `SAP_Run`, so `CommandButton1_Click` should be treated as legacy or alternate event logic unless the code owner confirms otherwise.

### `ThisWorkbook`

The component is present in the project. No separate workbook-open or workbook-close business workflow was confirmed from the available metadata.

## Confirmed and strongly evidenced macro logic

The following behavior is evidenced by the worksheet/control metadata, procedure names, SAP control IDs, literals, status messages, and embedded comments in the compiled project.

### 1. Connection setup

- Turns off Excel screen updating and display alerts during processing.
- Attempts to obtain the SAP GUI automation object with `GetObject("SAPGUI")`.
- Gets the SAP scripting engine, connection, and session through the SAP GUI object hierarchy.
- Shows `Please login to SAP first.` and exits when a usable SAP session is unavailable.
- Uses `On Error Resume Next` in parts of the workflow to tolerate optional SAP popups and controls.

### 2. Batch input

- Uses worksheet `Temp`.
- Finds the last populated row in column A using Excel row/range logic.
- Iterates over the PO numbers beginning at row 2.
- Writes outcome text to column B for each processed PO.

### 3. SAP transaction and selection

- Navigates to `/NME9F` through the SAP command field.
- Places the current PO number into `wnd[0]/usr/ctxtS_EBELN-LOW`.
- Clears selected stale filters. The embedded comments specifically call out processing status (`P_VSTAT`), message type (`S_KSCHL-LOW`), and purchasing group (`S_EKGRP-LOW`). The project also contains identifiers for other ME9F selection fields, such as vendor, purchasing organization, document type, date, and message/output settings.
- Executes the selection through the SAP toolbar control `wnd[0]/tbar[1]/btn[8]`.

### 4. Exception handling

- Attempts to acknowledge an SAP `wnd[1]` popup.
- Records `PO Not Found or No Messages` when the result indicates that no usable document/message was found.
- Searches for the expected PO checkbox. If it is unavailable, records `PO Checkbox Not Found`.
- Uses a `NextPO`-style flow to skip an exception and continue processing later rows.

### 5. Message/output preparation

- Opens message details through `wnd[0]/tbar[1]/btn[18]`.
- Inspects the `SAPDV70ATC_NAST3` message table.
- Searches for an empty message row while avoiding existing green/active rows.
- Adds message type `NEU` and transmission medium `1` to the available row.
- Confirms the entry using the SAP green-check/Enter control.

### 6. Printer/PDF output

- Saves the message entry to open the printer/output details.
- Sets printer/output fields including immediate processing and a PDF-related output-device value. The project contains the SAP controls `NAST-DIMME`, `NAST-DELET`, `NAST-LDEST`, and `NAST-TDOCOVER`.
- Confirms the printer details, backs out to the message screen, and performs a final save.
- Handles a possible `Save Data?` popup through `wnd[1]/usr/btnSPOP-OPTION1`.
- Selects the output row and performs the final output action.
- Records `Success` after the success path reaches the expected output action.
- Shows `Process Complete` after the loop.

The `Success` status should be treated as a workflow result, not as an independent filesystem verification of the PDF.

## Windows helper logic

`HandlePDFPopups` is a significant part of the implementation:

1. Creates `WScript.Shell`.
2. Uses `cmd.exe` and the `clip` utility to place the target PDF path into the Windows clipboard.
3. Launches hidden PowerShell with `WScript.Shell`.
4. Waits, sends Enter, waits again, pastes the path with Ctrl+V, and sends Enter to complete the SAP save dialog.

This is an implemented feature, not a suggestion. It is also the most fragile and security-sensitive part of the application because it depends on window focus, fixed delays, clipboard state, and shell execution.

## Technical dependencies

### Runtime

- Windows desktop OS.
- Microsoft Excel desktop capable of opening `.xlsm` files.
- SAP GUI for Windows with SAP GUI Scripting available.
- Logged-in SAP session authorized for transaction `ME9F` and the required message/output steps.
- Windows Script Host, PowerShell, `cmd.exe`, and clipboard access.

### VBA/Office references and objects

The embedded project metadata identifies the Excel 16.0, VBA/VBE, OLE Automation, Office 16.0, and Microsoft Forms libraries. The code uses or references:

- Excel worksheet/range/row objects and constants such as `xlUp`.
- `GetObject("SAPGUI")`, `GetScriptingEngine`, `GuiApplication`, `GuiConnection`, and `GuiSession`.
- SAP `findById`, `sendVKey`, `SetFocus`, `press`, `getAbsoluteRow`, `Selected`, and control `.Text`/`.Selected` properties.
- `CreateObject("WScript.Shell")` and shell `.Run` calls.
- `ScreenUpdating`, `DisplayAlerts`, `MsgBox`, and Excel cell value operations.

## Business problem and benefits

The manual process requires repeating the same SAP navigation, selection, message setup, printer/output configuration, and PDF save sequence for each PO. The workbook centralizes the PO list and uses one macro run to repeat that sequence.

Likely operational benefits, supported by the observed design, are:

- less repetitive SAP navigation;
- less manual re-entry of PO numbers and output settings;
- consistent message/output parameters;
- a simple per-PO status trail;
- faster exception identification.

The workbook contains no measured baseline, timing study, or error-rate comparison. Quantified savings must be measured separately.

## Installation and operating procedure

1. Use a sanitized, authorized copy of the workbook.
2. Open it in Excel desktop from a controlled local folder.
3. Review the macro source and shell behavior before enabling macros.
4. Start SAP GUI and log in to the correct system/client.
5. Verify SAP GUI Scripting is allowed.
6. Review the hard-coded PDF path in `Module1` and align it with an approved output folder.
7. Enter PO numbers in `Temp!A2:A...`.
8. Clear previous statuses in `Temp!B2:B...` if required.
9. Click `Run Macro`.
10. Monitor SAP dialogs and review column B after completion.
11. Verify the expected PDF files exist and inspect exception rows.

Do not run the tool against production data until the SAP system, authorization, output path, PDF naming convention, and shell automation have been validated in a controlled test.

## Publication and data-sanitization findings

The original file should not be uploaded unchanged. The following items require removal, replacement, or approval:

| Item | Evidence | Publication action |
|---|---|---|
| PO data | A sample PO-like value is stored in `Temp!A2`; additional PO-looking literals are embedded in compiled streams. | Replace with dummy values. |
| User/company path | Hard-coded path contains a Windows user profile, an organizational OneDrive name, a `PO_` filename prefix, and a `.pdf` extension. | Replace with a generic example or parameterize in a separate sanitized version. |
| SharePoint path | Workbook XML contains a company SharePoint personal path. | Remove the metadata/absolute path. |
| Author metadata | Core properties identify a named creator and last modifier. | Remove or replace if not intentionally published. |
| Templafy metadata | Custom XML contains tenant, template, and user profile identifiers. | Remove custom properties and custom XML before publication. |
| VBA protection | `CMG`, `DPB`, and `GC` records are present in `PROJECT`. | Do not attempt to bypass. Publish reviewed source only if authorized. |
| Credentials | No plaintext SAP password or credential was found in the inspected package/string tables. | Still perform a separate secrets scan before upload. |

## Implemented features versus future improvements

### Implemented

- Excel PO input list.
- `Run Macro` button.
- SAP GUI Scripting connection attempt.
- `/NME9F` navigation and PO selection.
- Selected filter clearing.
- No-document and checkbox exception paths.
- Message-row search and `NEU`/`1` entry.
- Printer/output configuration and PDF save sequence.
- Windows shell/PowerShell PDF-dialog helper.
- Per-PO status writeback and completion message.

### Not found in the reviewed workbook

- No VBA UserForm.
- No database/API integration.
- No scheduled task or unattended service.
- No email delivery workflow.
- No dashboard, chart, PivotTable, or formula-based report.
- No dedicated configuration sheet.
- No structured audit log with timestamps.

### Future improvements

- Move the PDF output path and SAP settings into a validated configuration sheet.
- Add a pre-flight check for login, scripting, session, transaction access, and folder access.
- Replace broad `On Error Resume Next` usage with explicit error handling and logs.
- Add retry, timeout, cancellation, and restart-from-row controls.
- Verify the created PDF file after the SAP output step.
- Replace clipboard/key simulation with a more deterministic approved interface.
- Export VBA source modules and keep them under version control.
- Add a sanitized test mode and dummy data set.

## Resume and portfolio description

Built an Excel VBA automation tool that batch-processes purchase-order message outputs in SAP ME9F. The tool reads PO numbers from Excel, drives SAP GUI Scripting to create and save PDF outputs, handles common SAP exceptions, and records per-PO statuses. The project demonstrates Excel VBA, SAP GUI automation, Windows scripting, and procurement process improvement.

## Review conclusion

The workbook is a practical single-purpose SAP procurement automation utility with a clear operator flow and a compact Excel front end. Its strongest portfolio themes are process analysis, Excel VBA, SAP GUI automation, exception handling, and reduction of repetitive procurement work. Its main production risks are hard-coded environment values, GUI fragility, shell/clipboard automation, limited logging, and the presence of sensitive metadata in the current file.
