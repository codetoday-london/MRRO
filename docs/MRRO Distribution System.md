# MRRO Distribution System

This is a short guide for the administrator. The workbook collects one publisher submission at a time, keeps the books from earlier years, and calculates the publisher and author payments.

## Normal annual workflow

1. In **Settings**, enter the distributable fund and the earliest and latest eligible publication years in the yellow cells. The latest year is the year publishers are currently submitting.
2. Give each publisher their own copy of the Publisher Template. In **Import status**, record the publisher name, the sheet link, and—if the system should manage their access—their Google-account email address in column I.
3. The publisher completes their sheet. They must correct any messages until it says **READY TO SUBMIT**. An empty sheet is also valid: it simply means there are no books to add that year.
4. In **Import status**, select one or more adjacent publisher rows and use the **MRRO administration** menu. The same action applies to every selected row.
5. Check **All books**, **Publisher payments**, and **Author payments**. Before using the results, confirm that the calculated licence total equals the fund and that publisher and author totals are each half of it.

## MRRO administration menu

### Validate selected publisher(s)

Use this first for a new submission. It checks that the sheet is ready, that its books use the current Settings end year, and that there are no publisher-sheet errors. It does not import books or change sharing access.

### Import and freeze selected publisher(s)

Use this after validation for a new submission. It adds that publisher’s current-year books to **All books** and keeps all previous years. If column I contains the publisher’s Google-account email, it changes only that person from Editor to Viewer. Other administrators keep their access.

### Refresh selected publisher import(s)

Use this when a publisher changes a submission that has already been imported for the current year. It replaces only that publisher’s current-year books. It does not remove older years, then rebuilds the calculations and both payment lists.

### Reopen and reset selected publisher submission(s)

Use this to prepare a blank submission for the current Settings end year. It clears the publisher’s entry area and sets the required year. If column I contains the publisher’s email, it changes that person from Viewer back to Editor. It does not remove historic books from **All books**.

## Preparing the next year

First change **Latest eligible publication year** in Settings to the new year. Then select the publishers in Import status and choose **Reopen and reset selected publisher submission(s)**. Their sheets are cleared and will accept only the new year. When they later submit, use Validate and then Import and freeze.

## Checks and colours

Publisher sheets require the current year, a title, an author name (commas between multiple authors), positive whole-number pages, a classification of 1, 2, or 3, and a blank, zero, or positive price. Blank and zero prices are valid but earn no payment.

Duplicate ISBNs and duplicate titles are checked in **All books**, not on the publisher sheet. Every matching row is red. In All books, dark grey means older than the eligible period and excluded from payment; light blue means an earlier eligible year; white means the current submission year.

## Adding a publisher

Make a copy of the Publisher Template. Insert a row above the procedure notes in **Import status**. Enter the publisher name in column A, the new sheet link in column D, and the publisher’s Google-account email in column I if access should be managed automatically. Columns B, C and E–H are calculated or updated by the menu.

## How payments are calculated

Eligible books receive points based on page count and retail-price band, multiplied by classification. The fund divided by total points gives the value of one point. Each book’s amount is split equally between its publisher and its author share; multiple authors split the author share equally.
