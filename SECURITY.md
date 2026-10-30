# Security and Publication Policy

This repository is a public, documentation-only demonstration of an SAP ME9F procurement automation concept. It does not currently contain the original workbook, VBA source, SAP credentials, or generated business documents.

## Do not publish

Do not commit or upload:

- SAP usernames, passwords, tokens, cookies, session traces, or exports
- Real purchase-order numbers or supplier/customer information
- Internal SAP system names, screenshots, or unredacted error messages
- PDFs generated from business data
- User, company, OneDrive, SharePoint, or network paths
- Workbook metadata, custom XML, or unreviewed vbaProject.bin content
- Company-owned process logic without written publication approval

## Before publishing an artifact

1. Confirm that the workbook, VBA source, screenshots, and process logic may be shared publicly.
2. Replace all business identifiers with synthetic values.
3. Remove document properties, custom XML, embedded paths, and company branding unless approved.
4. Review VBA for credentials, hard-coded paths, shell commands, PowerShell calls, clipboard use, and external links.
5. Confirm that PDFs, screenshots, and sample files contain no business data.
6. Review the final Git diff and repository history before making the repository public.

The repository .gitignore excludes common Office workbook types, generated outputs, temporary files, and local development artifacts. Git ignore rules do not remove files that were already committed, so every upload still requires manual review.

## Reporting a suspected exposure

Do not open a public issue containing a secret or sensitive business information. Contact the repository owner privately through GitHub and include the affected path, commit, and a short description of the exposure. Do not include the secret itself in the message.

If an exposed credential or token is found, revoke or rotate it through the owning system immediately and then remove the artifact from the repository history through an approved process.

## Scope and limitations

This policy is guidance for the public demonstration repository. It does not replace company information-security, data-protection, SAP security, or records-management requirements. Organizational approval is required before publishing any production workbook, VBA source, screenshot, export, or internal process documentation.
