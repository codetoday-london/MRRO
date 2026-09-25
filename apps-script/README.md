# MRRO administrator automation

`Code.gs` is the bound Apps Script for the live MRRO Administrator Master workbook. It creates the **MRRO administration** menu and implements validation, import and freeze, refresh, reopen-access, and reopen-and-reset actions.

The script is bound to the live workbook rather than deployed as a web application. Edit and save it from **Extensions → Apps Script** in the Administrator Master. Copy the reviewed source from this folder when restoring or creating the administrator template.

The publisher workbook IDs are read from links in **Import status**; they are not stored in this source file.

The administrator records one publisher email address in column I. Import and freeze changes only that email from editor to viewer. Reopen selected publisher(s) restores editor access without changing the submitted workbook or imported books; after corrections, Refresh replaces that publisher's current-year import and freezes access again. Reopen and reset restores editor access and clears the publisher workbook for the next annual cycle. The Settings end year is the annual submission year. Calculations assign zero points to books outside the eligible Settings range.

Calculations use row-position lookups, and the publisher and author payment tabs use live formulas, so deleting a row anywhere in **All books** updates payments automatically without a trigger. After installing the script, run `upgradePaymentFormulas` once from the bound script editor for the live master and again for the administrator template. It changes only the calculation and payment tabs; it does not touch publisher submissions, sharing, imported books, or Settings.

An empty submission is valid and records a zero-book annual import.
