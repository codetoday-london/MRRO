# Changelog

Changes to the Google Sheets administrator workflow are recorded here. The
administrator instructions are in [MRRO Distribution System](docs/MRRO%20Distribution%20System.md).

## 2026-09-25 — Row deletions, payments, and publisher access

- Fixed **All books** row deletion anywhere in the sheet. Calculation references
  now follow the books that remain, rather than breaking or assigning points to
  the wrong book, year, or publisher when rows move. Publisher and author
  payment allocations recalculate from the current books without requiring a
  manual refresh.
- Corrected an existing author overpayment in the administrator master. Before
  the fix, **Author allocation** retained an extra 75-point author entry that
  was no longer represented in **All books**. The stored author allocation had
  75 more points than the live book total, so its payment exceeded half the
  fund. The backup identifies the orphaned allocation but not the removed
  book's title. Author payments now use the same live books and point value as
  publisher payments.
- Added **Reopen selected publisher(s)** to the **MRRO administration** menu.
  It restores the selected publisher's editor access without clearing their
  submission or changing imported books. After corrections, **Refresh selected
  publisher import(s)** replaces that publisher's current-year import and
  freezes access again. **Reopen and reset selected publisher submission(s)**
  remains the separate option for clearing submissions for a new annual cycle.
- Allowed **Refresh** for an already imported current-year submission after
  its last book has been deleted from **All books**, including a zero-book
  submission. First-time submissions still require **Import and freeze**.
- Removed an obsolete template summary label from the publisher payment list;
  it is not a publisher and must not receive a payment row.
- Checked a disposable copy of the full master after deleting the first 2016
  book, duplicate-ISBN rows in 2017 and 2018, a middle 2021 book, and the
  final 2025 book, as well as adding and deleting a multi-author 2026 book.
  Independent checks reconciled books, years, publishers, authors, and totals
  to the full fund and an equal half on each payment sheet. The live master
  was then upgraded without deleting or changing its book rows; its final
  publisher and author totals each equalled half the fund. No real publisher
  access was changed during these tests.

The **Start here** tab in the administrator master and the administrator guide
both describe the new reopen-only option and its distinction from reset.
