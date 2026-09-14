# MRRO Distribution System

The administrator workbook collects publisher book lists, checks them, and calculates payments.

1. Set the fund and eligible years in the yellow Settings cells.
2. Publishers fill in their own workbook until it says **READY TO SUBMIT**.
3. In Import status, record the publisher email address in column I. Validate checks the current submission year. Import and freeze changes only that publisher from editor to viewer and adds a fixed copy of the year's books.
4. All books retains earlier imports. Current-year books are white; earlier eligible books are pale blue; books outside the eligible Settings years are pale red and excluded from payment calculations. The summary row groups matching ISBN and title row numbers in parentheses. Any matching ISBN or non-empty title is flagged, and every matching row is red.

Reopen and reset clears a publisher's input sheet, sets it to the Settings end year, and changes the recorded publisher from viewer to editor. Refresh replaces only that publisher's imported books for the current year.

An empty publisher sheet is valid: it represents no books for that year and imports no rows.

When adding a publisher, insert a row above the procedure block. Enter only the publisher name and workbook link; the status, count, lock and date columns are automatic.

## Publisher-sheet checks

Each publisher workbook checks that the year is a whole number from 2000 to 2200, the title and author are present, multiple author names are separated with commas, pages are a positive whole number, and classification is 1, 2 or 3. Price may be blank, zero, or positive; blank and zero-priced books receive no payment. Duplicate ISBNs and titles are deliberately checked only in **All books**, after import.

## How payments are calculated

Each eligible book receives points for page count and retail-price band, multiplied by its classification. The fund divided by total points gives the value of one point; the book licence amount is its points multiplied by that value.

The workbook totals licence amounts by publisher and splits each between publisher and author. Multiple authors share the author portion equally. Before using results, check that the licence total and both payment totals equal the Settings fund.
