# Bulk PO Download – SAP ME9F Automation

An Excel VBA application that batch-processes purchase-order outputs in SAP GUI through transaction `ME9F`. Users enter purchase-order numbers in Excel, run one macro, and receive per-PO processing statuses while the SAP output is saved as a PDF through the configured Windows file-dialog workflow.

> This repository should contain a sanitized demonstration copy only. The supplied workbook contains real-looking PO numbers, user/company metadata, a company SharePoint path, and a hard-coded local output path. Remove or anonymize those items before publishing.

## Project Overview

The application reduces repetitive procurement work when the same SAP message/output process must be completed for many purchase orders. It combines Excel as the input and tracking interface with SAP GUI Scripting for transaction automation.

The workbook is an operator tool rather than a reporting model. It contains one visible worksheet, one input/status area, and a `Run Macro` button.

## Key Features

- Batch processing of PO numbers listed in `Temp!A2:A...`.
- SAP GUI connection through `GetObject("SAPGUI")` and the SAP scripting engine.
- Automated navigation to SAP transaction `/NME9F`.
- Entry of the current PO number and clearing of selected leftover filters, including processing status, message type, and purchasing group.
- Execution of the SAP selection.
- Handling of the SAP “no documents found” path.
- Detection of a PO selection checkbox and navigation to message details.
- Search for an available message row while avoiding existing green/active rows.
- Creation of a message entry using message type `NEU` and transmission medium `1`.
- Automated save and printer-detail steps for PDF output.
- Handling of SAP save-data prompts and output-selection steps.
- Windows clipboard, `WScript.Shell`, `cmd.exe`, and hidden PowerShell helper logic to assist with the PDF save dialog.
- Per-PO status updates in column B, including success and failure messages.
- Final completion message after the batch loop.

## Application Workflow

```text
PO numbers in Temp!A2:A...
            |
            v
       SAP_Run macro
            |
            v
  SAP GUI transaction /NME9F
            |
            v
   Select PO and message details
            |
            v
  Add NEU / medium 1 and save
            |
            v
  Open PDF output and save file
            |
            v
  Write status in Temp!B:B
```

### Step-by-step process

1. Enter one PO number per row in column A, beginning at row 2.
2. Confirm that SAP GUI is open, the user is logged in, and a usable SAP session is available.
3. Click `Run Macro`, which is assigned to `SAP_Run`.
4. The macro connects to SAP GUI and calculates the last populated input row.
5. For each PO, it opens `/NME9F`, fills the PO selection, clears selected stale filters, and executes the search.
6. If SAP reports that no document/message exists, the macro writes `PO Not Found or No Messages` in column B and continues.
7. If the PO checkbox cannot be found, the macro writes `PO Checkbox Not Found` and continues.
8. When a result is available, the macro opens message details, finds an available message row, and enters the required message settings.
9. It opens the printer/output dialog, applies the configured PDF-related settings, confirms the dialog, and saves the SAP data.
10. A Windows scripting helper prepares the expected PDF filename/path and assists with the SAP save dialog.
11. The macro completes the output action, writes `Success` in column B, and continues to the next PO.
12. At the end, it restores Excel display settings and shows `Process Complete`.

## Workbook and VBA Structure

| Component | Observed role |
|---|---|
| `Temp` worksheet | Input and status interface. `A1` is `PONumbers`, `B1` is `Comment`; the current sample has a PO in `A2` and `Success` in `B2`. |
| `Run Macro` button | Legacy Excel form-control/shape button assigned to `[0]!SAP_Run`. |
| `Module1` | Main `SAP_Run` procedure and the `HandlePDFPopups` helper. |
| `Sheet1` | Worksheet-level compiled VBA content, including `CommandButton1_Click` and SAP automation content. The visible button currently points directly to `SAP_Run`. |
| `ThisWorkbook` | Workbook component present in the VBA project; no separate business workflow was confirmed from the available metadata. |

The workbook does not contain a detected VBA `UserForm`. The visible front end is a worksheet plus a form-control button.

## Requirements

- Windows desktop.
- Microsoft Excel desktop with macro-enabled workbook support.
- SAP GUI for Windows.
- SAP GUI Scripting enabled and permitted by the SAP system/security policy.
- An authenticated SAP session with authorization for transaction `ME9F`, PO display/message maintenance, and the required output operation.
- Windows Script Host components, `cmd.exe`, PowerShell, and clipboard access. These are used by the PDF-dialog helper.
- Access to the configured PDF output folder.

The current workbook references the Excel/Office/VBA runtime and Microsoft Forms libraries. SAP GUI scripting objects such as `GuiApplication`, `GuiConnection`, and `GuiSession` are used by the automation.

## Installation and Setup

1. Download or clone the sanitized repository.
2. Store the `.xlsm` file in a trusted local folder approved by your organization.
3. Open the workbook in the Windows desktop version of Excel. Review the VBA project and macro source before enabling content.
4. Start SAP GUI and sign in to the correct system and client.
5. Confirm that SAP GUI Scripting is allowed and that the SAP session is ready for GUI automation.
6. Review the PDF output path in `Module1`. The supplied workbook contains a user/company-specific hard-coded path, so a controlled deployment must align this path with an approved local folder.
7. On the `Temp` sheet, enter PO numbers in column A from row 2 downward. Keep PO numbers as text if leading zeros are possible.
8. Clear old comments/statuses in column B if a new run is required.
9. Click `Run Macro` and monitor SAP and the status column.
10. Verify the generated PDF files and review any non-success statuses before closing Excel.

### Input example

| `PONumbers` | `Comment` |
|---|---|
| `0000000001` | blank before run |
| `0000000002` | blank before run |

### Status examples implemented by the workbook

| Status | Meaning |
|---|---|
| `Success` | The macro reached the output step for the PO. Verify the PDF file exists. |
| `PO Not Found or No Messages` | SAP did not return a usable document/message result. |
| `PO Checkbox Not Found` | The expected SAP result control was not found. |
| `Error: No Empty Rows Found` | No available message row was found for the output step. |
| `Please login to SAP first.` | The SAP GUI scripting connection/session could not be obtained. |
| `Process Complete` | The Excel batch loop finished. |

## Technical Details

- **Language:** VBA.
- **Host:** Microsoft Excel `.xlsm` workbook.
- **SAP integration:** SAP GUI Scripting object model through `GetObject("SAPGUI")`.
- **Excel features:** worksheet cells, legacy form-control button, macro assignment, status messages, `ScreenUpdating`, `DisplayAlerts`, and Excel range/row operations.
- **Windows automation:** `WScript.Shell`, clipboard operations, `cmd.exe`, PowerShell, and simulated key presses for the SAP PDF save workflow.
- **SAP transaction:** `ME9F`.
- **Primary entry point:** `SAP_Run`.
- **PDF helper:** `HandlePDFPopups`.
- **Current workbook scope:** one worksheet named `Temp`; no formulas, tables, charts, pivots, or database connections were observed in the workbook package.

## Business Benefits

- Removes repetitive navigation and data entry from a multi-PO SAP output process.
- Reduces copy/paste and transcription effort between an input list and SAP.
- Provides a simple per-PO audit trail in Excel through the `Comment` column.
- Standardizes message and output settings across processed POs.
- Allows the operator to review exceptions without restarting the whole batch.
- Makes the process easier to demonstrate and hand over than a purely manual SAP procedure.

The workbook does not contain measured cycle-time or error-rate baselines. Claims about percentage savings should be added only after a controlled before/after measurement.

## Known Limitations and Operational Risks

- The automation depends on SAP GUI screen IDs, button positions, table structure, and the exact `ME9F` layout. SAP screen changes can break it.
- The PDF helper depends on the correct window being active and uses simulated key presses, so it is sensitive to timing and focus.
- The PDF path is hard-coded and is not exposed as a user configuration field.
- `On Error Resume Next` is used in parts of the SAP workflow. This can hide unexpected errors and make troubleshooting harder.
- The workbook does not provide a dedicated configuration sheet, structured log, retry control, or run summary.
- The current workbook contains sample/business-identifying data and should not be published unchanged.
- `Success` records that the macro reached its output step; it should not be treated as proof that every PDF was independently verified.

## Security and Publication Checklist

Before uploading to GitHub:

- Replace all real PO numbers with dummy values.
- Remove or anonymize the sample status/output records.
- Remove the hard-coded local path containing the user profile and company name.
- Remove the SharePoint tenant/user path stored in workbook metadata.
- Remove author, last-modified-by, timestamps, Templafy tenant/template/user identifiers, and other custom properties unless intentionally published.
- Do not publish SAP usernames, passwords, session data, exported PDFs, supplier/customer data, or internal screenshots.
- Do not publish the raw `vbaProject.bin` as a separate file.
- Prefer an exported, reviewed VBA source representation plus a sanitized `.xlsm` sample if your organization permits publishing the automation logic.
- Keep macro execution disabled for untrusted copies and review any shell/PowerShell behavior before use.

No plaintext SAP password or credential was found during the review. The VBA project contains protection metadata (`CMG`, `DPB`, and `GC` records); this is not a plaintext password and should not be bypassed for publication.

## Recommended Repository Structure

```text
bulk-po-download-sap-me9f/
├── README.md
├── LICENSE
├── .gitignore
├── workbook/
│   └── bulk-po-download-sap-me9f-sanitized.xlsm
├── src/
│   ├── Module1.bas
│   ├── Sheet1.cls
│   └── ThisWorkbook.cls
├── sample-data/
│   └── po_numbers_sample.csv
├── docs/
│   ├── workflow.md
│   ├── security-review.md
│   └── screenshots/
└── CHANGELOG.md
```

### Files to upload

- `README.md` and supporting documentation.
- Sanitized exported VBA modules, if the source owner permits it.
- A sanitized workbook sample only if macro-enabled files and SAP automation code may be shared.
- Dummy input data and screenshots with all real identifiers removed.

### Files to keep out of GitHub

- The original downloaded workbook.
- Real PO lists, SAP exports, PDF outputs, supplier/customer information, or internal screenshots.
- User/company-specific metadata and local paths.
- Credentials, tokens, cookies, session dumps, or SAP GUI traces.
- Temporary Excel files, lock files (`~$*.xlsm`), and compiled artifacts.

## Suggested Screenshots

1. The sanitized `Temp` sheet showing the `PONumbers` and `Comment` columns.
2. The `Run Macro` button and the input/status area together.
3. A sanitized example showing mixed statuses such as `Success` and `PO Not Found or No Messages`.
4. SAP `ME9F` selection screen with all PO/company/supplier identifiers blurred or replaced.
5. SAP message-details/output screen with internal values removed.
6. The generated PDF folder containing dummy filenames only.
7. A simple workflow diagram showing Excel → SAP GUI → PDF → status update.

Do not include screenshots of credentials, SAP system names, supplier names, real PO numbers, internal paths, or unredacted company branding without approval.

## Future Improvements

These are recommendations, not features confirmed in the supplied workbook:

- Add a configuration sheet for SAP transaction options and PDF output folder.
- Replace hard-coded screen values and paths with validated configuration fields.
- Add explicit error handling, timestamped logging, retry limits, and a final run summary.
- Add a pre-flight check for SAP login, scripting availability, active session, and output-folder access.
- Replace clipboard/key-press PDF handling with a more deterministic SAP GUI or approved document-output interface.
- Export and version VBA source modules so code changes are reviewable in Git.
- Add a sanitized test mode using dummy PO numbers and a mocked status path.
- Add a clear cancellation/stop mechanism for long-running batches.

## Portfolio Description

Built an Excel VBA automation tool that batch-processes purchase-order message outputs in SAP ME9F. The tool reads PO numbers from Excel, drives SAP GUI Scripting to create and save PDF outputs, handles common SAP exceptions, and records per-PO statuses. The project demonstrates Excel VBA, SAP GUI automation, Windows scripting, and procurement process improvement.

## License and Data Ownership

Add a license only after confirming that the VBA code, SAP workflow, screenshots, and sample data may be shared publicly. Company-owned process logic and internal SAP automation may require written approval before publication.

