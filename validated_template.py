"""Create copy/paste-friendly, self-validating MRRO submission workbooks."""

from __future__ import annotations

from pathlib import Path
from typing import Any, Iterable

import xlsxwriter


HEADERS = (
    "Year of publication",
    "Name of Book",
    "ISBN",
    "Author(s)",
    "Number of pages",
    "Retail price (inc. VAT)",
    "Melitensia (1), Adult (2), Children's (3)",
    "Issues",
)
FIRST_DATA_ROW = 9  # Excel row number; the zero-based xlsxwriter row is 8.
LAST_DATA_ROW = 508


def validate_records(records: Iterable[dict[str, Any]]) -> list[list[str]]:
    """Return human-readable errors for each record, mirroring the workbook checks."""
    rows = list(records)
    errors: list[list[str]] = [[] for _ in rows]
    fingerprints: dict[tuple[str, ...], list[int]] = {}
    isbns: dict[str, list[int]] = {}

    for index, record in enumerate(rows):
        year = record.get("year")
        title = str(record.get("title", "")).strip()
        authors = str(record.get("authors", "")).strip()
        pages = record.get("pages")
        price = record.get("price")
        category = record.get("category")
        isbn = str(record.get("isbn", "")).strip()

        if not isinstance(year, int) or not 2016 <= year <= 2025:
            errors[index].append("Year must be a whole number from 2016 to 2025.")
        if not title:
            errors[index].append("Book title is required.")
        forbidden_author_separators = ("&", "%", ";")
        if (
            not authors
            or any(token in authors for token in forbidden_author_separators)
            or " AND " in f" {authors.upper()} "
            or authors.startswith(",")
            or authors.endswith(",")
            or ",," in authors
        ):
            errors[index].append("Use comma-separated author names only.")
        if not isinstance(pages, int) or pages <= 0:
            errors[index].append("Pages must be a positive whole number.")
        if not isinstance(price, (int, float)) or isinstance(price, bool) or price <= 0:
            errors[index].append("Price must be a positive number.")
        if not isinstance(category, int) or category not in (1, 2, 3):
            errors[index].append("Classification must be 1, 2, or 3.")

        fingerprint = tuple(str(record.get(key, "")).strip().casefold() for key in ("year", "title", "isbn", "authors", "pages", "price", "category"))
        fingerprints.setdefault(fingerprint, []).append(index)
        if isbn:
            isbns.setdefault(isbn.casefold(), []).append(index)

    for indexes in fingerprints.values():
        if len(indexes) > 1:
            for index in indexes:
                errors[index].append("Duplicate book row.")
    for indexes in isbns.values():
        if len(indexes) > 1:
            for index in indexes:
                errors[index].append("Duplicate ISBN.")
    return errors


def _issue_formula(excel_row: int) -> str:
    row = str(excel_row)
    return (
        f'=IF(COUNTA(A{row}:G{row})=0,"",'
        f'IF(NOT(AND(ISNUMBER(A{row}),A{row}=INT(A{row}),A{row}>=2016,A{row}<=2025)),"Year must be a whole number from 2016 to 2025. ","")&'
        f'IF(B{row}="","Book title is required. ","")&'
        f'IF(OR(D{row}="",ISNUMBER(SEARCH("&",D{row})),ISNUMBER(SEARCH("%",D{row})),ISNUMBER(SEARCH(";",D{row})),ISNUMBER(SEARCH(" AND "," "&UPPER(D{row})&" ")),LEFT(D{row},1)=",",RIGHT(D{row},1)=",",ISNUMBER(SEARCH(",,",D{row}))),"Use comma-separated author names only. ","")&'
        f'IF(NOT(AND(ISNUMBER(E{row}),E{row}=INT(E{row}),E{row}>0)),"Pages must be a positive whole number. ","")&'
        f'IF(NOT(AND(ISNUMBER(F{row}),F{row}>0)),"Price must be a positive number. ","")&'
        f'IF(NOT(AND(ISNUMBER(G{row}),G{row}=INT(G{row}),G{row}>=1,G{row}<=3)),"Classification must be 1, 2, or 3. ","")&'
        f'IF(COUNTIFS($A$9:$A$508,A{row},$B$9:$B$508,B{row},$C$9:$C$508,C{row},$D$9:$D$508,D{row},$E$9:$E$508,E{row},$F$9:$F$508,F{row},$G$9:$G$508,G{row})>1,"Duplicate book row. ","")&'
        f'IF(AND(C{row}<>"",COUNTIF($C$9:$C$508,C{row})>1),"Duplicate ISBN. ",""))'
    )


def create_submission_workbook(
    output_path: str | Path,
    publisher: str = "",
    records: Iterable[dict[str, Any]] = (),
) -> None:
    """Create a protected, self-checking workbook with optional populated records."""
    records = list(records)
    if len(records) > LAST_DATA_ROW - FIRST_DATA_ROW + 1:
        raise ValueError("This template supports at most 500 book rows.")
    record_errors = validate_records(records)
    error_count = sum(bool(errors) for errors in record_errors)

    workbook = xlsxwriter.Workbook(str(output_path))
    workbook.set_properties({"title": "MRRO validated publisher submission"})
    worksheet = workbook.add_worksheet("MRRO submission")
    worksheet.set_column("A:A", 18)
    worksheet.set_column("B:B", 50)
    worksheet.set_column("C:C", 22)
    worksheet.set_column("D:D", 34)
    worksheet.set_column("E:E", 18)
    worksheet.set_column("F:F", 24)
    worksheet.set_column("G:G", 25)
    worksheet.set_column("H:H", 72)

    title = workbook.add_format({"bold": True, "font_size": 15, "align": "center", "valign": "vcenter", "bg_color": "#1F4E78", "font_color": "#FFFFFF"})
    label = workbook.add_format({"bold": True})
    input_cell = workbook.add_format({"border": 1, "locked": False})
    input_int = workbook.add_format({"border": 1, "locked": False, "align": "right", "num_format": "0"})
    input_price = workbook.add_format({"border": 1, "locked": False, "align": "right", "num_format": "€#,##0.00"})
    header = workbook.add_format({"bold": True, "border": 1, "bg_color": "#D9EAF7", "text_wrap": True, "valign": "vcenter"})
    issue = workbook.add_format({"border": 1, "font_color": "#9C0006", "bg_color": "#FFC7CE", "text_wrap": True, "valign": "top"})
    okay = workbook.add_format({"bold": True, "font_color": "#006100", "bg_color": "#C6EFCE", "border": 1})
    not_ready = workbook.add_format({"bold": True, "font_color": "#9C0006", "bg_color": "#FFC7CE", "border": 1})

    worksheet.merge_range("A1:H1", "Malta Reprographic Rights Organisation — Publisher Submission", title)
    worksheet.set_row(0, 28)
    worksheet.write("A3", "Publisher name", label)
    worksheet.write("B3", publisher, input_cell)
    worksheet.write("A4", "Submission status", label)
    status = (
        "NOT READY — add at least one book."
        if not records
        else f"NOT READY TO SUBMIT — {error_count} row(s) need correction."
        if error_count
        else "READY TO SUBMIT"
    )
    worksheet.write_formula(
        "B4",
        '=IF(COUNTA(B9:B508)=0,"NOT READY — add at least one book.",IF(COUNTIF(H9:H508,"<>")>0,"NOT READY TO SUBMIT — "&COUNTIF(H9:H508,"<>")&" row(s) need correction.","READY TO SUBMIT"))',
        not_ready if status.startswith("NOT READY") else okay,
        status,
    )
    worksheet.write("A5", "Books entered", label)
    worksheet.write_formula("B5", "=COUNTA(B9:B508)", None, len(records))
    worksheet.merge_range("D3:H5", "Paste book records directly into the blue table. Red cells show rows that must be corrected. All highlighted rules are checked again before distribution.", workbook.add_format({"text_wrap": True, "valign": "top", "bg_color": "#FFF2CC", "border": 1}))

    for col, heading in enumerate(HEADERS):
        worksheet.write(7, col, heading, header)
    worksheet.set_row(7, 38)

    for zero_row in range(FIRST_DATA_ROW - 1, LAST_DATA_ROW):
        excel_row = zero_row + 1
        worksheet.write_blank(zero_row, 0, None, input_int)
        worksheet.write_blank(zero_row, 1, None, input_cell)
        worksheet.write_blank(zero_row, 2, None, input_cell)
        worksheet.write_blank(zero_row, 3, None, input_cell)
        worksheet.write_blank(zero_row, 4, None, input_int)
        worksheet.write_blank(zero_row, 5, None, input_price)
        worksheet.write_blank(zero_row, 6, None, input_int)
        cached_issue = "; ".join(record_errors[zero_row - (FIRST_DATA_ROW - 1)]) if zero_row < FIRST_DATA_ROW - 1 + len(records) else ""
        worksheet.write_formula(zero_row, 7, _issue_formula(excel_row), issue, cached_issue)

    for offset, record in enumerate(records):
        row = FIRST_DATA_ROW - 1 + offset
        worksheet.write(row, 0, record.get("year"), input_int)
        worksheet.write(row, 1, record.get("title"), input_cell)
        worksheet.write(row, 2, record.get("isbn"), input_cell)
        worksheet.write(row, 3, record.get("authors"), input_cell)
        worksheet.write(row, 4, record.get("pages"), input_int)
        worksheet.write(row, 5, record.get("price"), input_price)
        worksheet.write(row, 6, record.get("category"), input_int)

    worksheet.data_validation(FIRST_DATA_ROW - 1, 0, LAST_DATA_ROW - 1, 0, {"validate": "integer", "criteria": "between", "minimum": 2016, "maximum": 2025, "error_type": "stop", "error_title": "Invalid year", "error_message": "Enter a whole year from 2016 to 2025."})
    worksheet.data_validation(FIRST_DATA_ROW - 1, 4, LAST_DATA_ROW - 1, 4, {"validate": "integer", "criteria": ">", "value": 0, "error_type": "stop", "error_title": "Invalid page count", "error_message": "Enter a positive whole number."})
    worksheet.data_validation(FIRST_DATA_ROW - 1, 5, LAST_DATA_ROW - 1, 5, {"validate": "decimal", "criteria": ">", "value": 0, "error_type": "stop", "error_title": "Invalid price", "error_message": "Enter a positive monetary amount."})
    worksheet.data_validation(FIRST_DATA_ROW - 1, 6, LAST_DATA_ROW - 1, 6, {"validate": "list", "source": [1, 2, 3], "error_type": "stop", "error_title": "Invalid classification", "error_message": "Choose 1, 2, or 3."})
    worksheet.conditional_format(FIRST_DATA_ROW - 1, 0, LAST_DATA_ROW - 1, 7, {"type": "formula", "criteria": "=$H9<>\"\"", "format": issue})
    worksheet.autofilter(7, 0, LAST_DATA_ROW - 1, 7)
    worksheet.freeze_panes(8, 0)
    worksheet.protect("", {"sort": True, "autofilter": True, "select_unlocked_cells": True, "select_locked_cells": False})
    workbook.close()
