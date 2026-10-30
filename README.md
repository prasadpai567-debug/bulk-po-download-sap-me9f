            # Bulk PO Download – SAP ME9F Automation

Excel VBA and SAP GUI Scripting concept for batch-processing purchase-order output from SAP transaction ME9F.

> **Public repository status:** This is currently a documentation-only demonstration. The original Excel workbook, VBA source, SAP connection, and PDF-output implementation are intentionally not included. The behavior described below is based on a read-only analysis of the original workbook and is not runnable from this repository.

## Business problem

A manual ME9F process can require repeating the same SAP navigation, PO selection, message setup, output configuration, and PDF save steps for each purchase order. Repeating those steps increases operator effort and creates opportunities for inconsistent settings or transcription mistakes.

The analyzed workbook was designed to centralize PO numbers in Excel, repeat the SAP workflow for each PO, and write a per-PO status back to the worksheet.

No measured cycle-time or error-rate baseline is included, so this project does not claim a percentage time saving.

## How the documented automation works

The original workbook analysis describes this workflow:

1. Enter one PO number per row in an Excel input sheet.
2. Connect to an authenticated SAP GUI session.
3. Open SAP transaction ME9F.
4. Enter the PO number and clear selected stale filters.
5. Execute the SAP search.
6. Handle no-document and missing-control paths.
7. Open message details and select an available message row.
8. Apply message type NEU and transmission medium 1.
9. Configure the output/printer steps for PDF generation.
10. Save the output and write a status for the PO.
11. Continue with the next PO and show a completion message.

The analysis also documents a Windows shell/PowerShell and clipboard helper for the SAP save dialog. That helper is environment-sensitive and should be reviewed carefully before any authorized operational use.

## Repository contents

| Path | Purpose |
|---|---|
| README.md | Project explanation, review status, workflow, limitations, and portfolio summary |
| docs/VBA_APPLICATION_ANALYSIS.md | Read-only analysis of the original workbook and its documented behavior |
| sample-data/po_numbers_sample.csv | Synthetic-looking PO input example for documentation |
| .gitignore | Protection against temporary files, exports, Office files, and private data |

The following are **not** currently included:

- Original or sanitized XLSM workbook
- Exported BAS, CLS, or FRM VBA source
- SAP GUI connection or test environment
- PDF outputs or screenshots
- Automated tests
- CI workflow
- Dependency manifest

## Features documented from the original workbook analysis

The read-only analysis reports the following workbook behavior:

- Excel worksheet input and status tracking
- A Run Macro control assigned to SAP_Run
- SAP GUI Scripting connection through GetObject("SAPGUI")
- Navigation to /NME9F
- PO selection and selected filter clearing
- No-document and missing-checkbox status paths
- Message-row search
- Message type NEU and medium 1
- Printer/output configuration
- PDF save-dialog assistance
- Per-PO status updates and a final completion message

These are documented observations from the analyzed workbook. They are not source-verified or executable in the current public repository.

## Technologies and operational requirements

The analyzed automation requires:

- Windows desktop
- Microsoft Excel desktop with macro-enabled workbook support
- SAP GUI for Windows
- SAP GUI Scripting enabled by the SAP system and security policy
- An authenticated SAP session authorized for the required ME9F operations
- Windows Script Host, PowerShell, cmd.exe, clipboard access, and an approved output folder

The repository itself has no Python dependency requirements because Python is not part of this project.

## Review and installation guidance

There is no runnable installation from the current repository. To review the public demonstration:

1. Clone or browse the repository.
2. Read this README.
3. Review docs/VBA_APPLICATION_ANALYSIS.md.
4. Review sample-data/po_numbers_sample.csv.
5. Do not attempt SAP execution from this repository; the workbook and source are not included.

If an authorized sanitized workbook is added in the future, its deployment instructions must specify the supported Excel version, SAP GUI requirements, scripting permissions, output-folder configuration, and a controlled test procedure.

## Sample input

The sample CSV uses dummy-looking values for documentation:

~~~csv
PONumbers,Comment
0000000001,
0000000002,
0000000003,
~~~

The current repository does not contain an application that consumes this CSV directly. It represents the type of input used by the analyzed Excel workflow.

## Documented status values

| Status | Meaning |
|---|---|
| Success | The documented workflow reached its expected output action; this does not independently verify that a PDF exists |
| PO Not Found or No Messages | SAP did not return a usable result |
| PO Checkbox Not Found | The expected SAP result control was unavailable |
| Error: No Empty Rows Found | No available message row was found |
| Please login to SAP first. | A usable SAP GUI session was not obtained |
| Process Complete | The documented batch loop finished |

## Testing and verification status

No automated tests or CI workflows are currently present.

The following have not been executed from this repository:

- Excel VBA macro
- SAP GUI transaction
- PDF generation
- Windows save-dialog helper
- End-to-end batch processing

The analysis document states that the original workbook was reviewed read-only and was not executed because it controls SAP GUI, writes files, invokes shell helpers, uses the clipboard, and simulates keyboard input.

## Security and publication guidance

Do not publish:

- SAP usernames, passwords, tokens, cookies, session traces, or exports
- Real PO numbers or supplier/customer information
- Internal SAP system names or screenshots
- PDFs generated from business data
- User, company, OneDrive, SharePoint, or network paths
- Workbook metadata, custom XML, or unreviewed vbaProject.bin content
- Unapproved company-owned process logic

The analysis of the original workbook reported a hard-coded output path, SharePoint-related metadata, author metadata, and shell/PowerShell PDF-dialog automation. Those items must be removed, anonymized, or approved before publishing any workbook or source export.

See SECURITY.md for the publication policy.

## Known limitations

- The current repository contains documentation rather than executable implementation.
- SAP GUI screen IDs, controls, timing, and layouts can change.
- The documented PDF helper depends on window focus, fixed delays, clipboard state, and simulated key presses.
- The analyzed workbook reportedly used a hard-coded PDF path.
- Broad On Error Resume Next usage was reported in parts of the analyzed workflow.
- The documented Success status does not independently verify the generated file.
- No structured timestamped log, retry control, cancellation control, or run summary is currently available in this repository.
- No measured business-result baseline is available.

## Future improvements

Subject to publication and business approval:

- Add a sanitized, reviewed workbook or exported VBA source.
- Move output paths and SAP settings into validated configuration.
- Add pre-flight checks for login, scripting, session, transaction access, and folder access.
- Replace broad error suppression with explicit error handling and structured logs.
- Add timeout, retry, cancellation, and restart-from-row controls.
- Verify PDF existence before recording success.
- Replace clipboard/key simulation with a more deterministic approved interface.
- Add static validation, a controlled test mode, and documented test evidence.
- Add anonymized screenshots and a workflow diagram.
- Add a license after confirming code and data ownership.

## Portfolio summary

Built an Excel VBA and SAP GUI Scripting automation concept for procurement operations that batch-processes purchase-order outputs through SAP ME9F. The documented workflow reads PO numbers from Excel, applies standardized SAP message and PDF-output steps, handles common exceptions, and records per-PO statuses. It demonstrates SAP procurement knowledge, Excel automation, Windows scripting, process standardization, and practical workflow improvement.

## License and ownership

No license is currently included. The repository should be treated as all rights reserved until the code owner confirms an appropriate license and confirms that the SAP workflow, workbook, source exports, screenshots, and sample data may be shared publicly.
