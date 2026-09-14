# MRRO administrator automation

`Code.gs` is the bound Apps Script for the live MRRO Administrator Master workbook. It creates the **MRRO administration** menu and implements validation, import and freeze, refresh, and reopen-and-reset actions.

The script is bound to the live workbook rather than deployed as a web application. Edit and save it from **Extensions → Apps Script** in the Administrator Master. Copy the reviewed source from this folder when restoring or creating the administrator template.

The publisher workbook IDs are read from links in **Import status**; they are not stored in this source file.

The administrator records one publisher email address in column I. Import and freeze changes only that email from editor to viewer; reopen and reset changes it back to editor. The Settings end year is the annual submission year. Import appends that year's books, while Refresh replaces only that publisher's entries for that year. Calculations assign zero points to books outside the eligible Settings range.

An empty submission is valid and records a zero-book annual import.
