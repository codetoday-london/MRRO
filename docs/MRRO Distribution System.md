# MRRO Distribution System

The administrator workbook collects publisher book lists, checks them, and calculates payments.

1. Set the fund and eligible years in the yellow Settings cells.
2. Publishers fill in their own workbook until it says **READY TO SUBMIT**.
3. In Import status, Validate is optional. Import and freeze changes the publisher from editor to viewer and imports a fixed copy.
4. Check All books for incomplete records, unsuitable years and duplicate ISBNs before using payment figures.

Reset clears one publisher's imported books and returns their register row to open/pending. Reopen changes that publisher back from viewer to editor. Refresh replaces a previous frozen import.

When adding a publisher, insert a row above the procedure block. Enter only the publisher name and workbook link; the status, count, lock and date columns are automatic.

## Publisher-sheet checks

Each publisher workbook checks that the year is a whole number from 2000 to 2200, the title and author are present, multiple author names are separated with commas, pages are a positive whole number, price is positive, and classification is 1, 2 or 3. Duplicate ISBNs are deliberately checked only in **All books**, after import.

## How payments are calculated

Each eligible book receives points for page count and retail-price band, multiplied by its classification. The fund divided by total points gives the value of one point; the book licence amount is its points multiplied by that value.

The workbook totals licence amounts by publisher and splits each between publisher and author. Multiple authors share the author portion equally. Before using results, check that the licence total and both payment totals equal the Settings fund.
