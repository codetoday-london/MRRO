"""
MRRO.py - Reads in separate publisher submissions in the form or spreadsheets (Excel)
following the default template. Works out payout per book based on given formula.
Exports CSV files with all books, one with all publisher payments, and one with all
author payments.

All spreadsheets should be in a folder with nothing else in the folder (except files
created by this software itself)

Usage:
First prompt - enter path of folder with spreadsheets (option-command-c will copy
the path on macOS when folder is selected in Finder)

Second prompt - enter amount of money to be distributed (total funds - admin charges)

Three files will be created each time software runs with the required outputs as CSV
documents that can be opened with Excel or any other spreadsheet document

Application runs on macOS only. No Windows version is available at present
"""

import datetime
import os
import pandas as pd

publisher_author_ratio = 0.5

# Welcome message
org_name = "MALTA REPROGRAPHIC RIGHTS ORGANISATION"
print(f"\n\n{org_name}")
print("-" * len(org_name))
print()

# Request user to input folder with all spreadsheets
# (macOS only)
print("Enter full path name of folder with spreadsheets")
print("[Select folder in Finder and use option-command-c to copy path]")
path = input("-> ")

# Create empty DataFrame to contain all data for all spreadsheets/publishers
all_publisher_df = pd.DataFrame()

# Iterate through all files in folder
for path, _, files in os.walk(path):
    for filename in files:
        # print(filename)
        if filename.startswith(".") or filename.startswith("__"):
            continue

        # Read data and create DataFrame
        columns = [
            "Name of Book",
            "ISBN",
            "Author(s)",
            "Number of pages",
            "Retail price (inc. VAT)",
            "Melitensia (1), Adult (2), Children's (3)",
        ]

        data = pd.read_excel(
            path + "/" + filename,
            names=columns,
            dtype="str",
            keep_default_na=False,
        )

        # Remove empty rows
        data = data[~data["Name of Book"].str.fullmatch("")]

        # Extract publisher from row with Name of Company/Self Published Author
        # Note, colon required before publisher name
        publisher = (
            data[
                data.loc[:, "Name of Book"].str.contains("Name of Company")
                | data.loc[:, "Name of Book"].str.contains("Self Published Author")
            ]
            .iloc[0, 0]
            .split(":")[-1]
        )

        # Find rows with sections headings ("List of books published in ...")
        sections = data[data.iloc[:, 0].str.contains("List of books published in")]

        # Create DataFrame containing only rows of books, using sections found above
        output_df = pd.DataFrame()
        # Iterate through yearly sections
        for idx, df_idx in enumerate(sections.index[:-1]):
            temp = data.loc[sections.index[idx] + 1 : sections.index[idx + 1], :].copy()
            # Find part of string starting with 20...
            year_idx = data.loc[df_idx, "Name of Book"].find("20")
            year_value = data.loc[df_idx, "Name of Book"][year_idx : year_idx + 4]
            temp.loc[:, "Year of Publication"] = year_value

            output_df = output_df.append(temp)

        # And add final section
        idx += 1
        df_idx = sections.index[-1]
        temp = data.loc[sections.index[idx] + 1 :, :].copy()
        year_idx = data.loc[df_idx, "Name of Book"].find("20")
        year_value = data.loc[df_idx, "Name of Book"][year_idx : year_idx + 4]
        temp.loc[:, "Year of Publication"] = year_value

        output_df = output_df.append(temp)

        output_df["Year of Publication"] = output_df["Year of Publication"].astype(int)
        # Remove rows with repeated header rows
        output_df = output_df[~output_df["Name of Book"].str.fullmatch("Name of Book")]
        output_df = output_df[
            ~output_df["Name of Book"].str.startswith("List of books published in")
        ]

        output_df.loc[:, "Publisher"] = publisher

        # Replace empty cells with default (fallback) values
        output_df["Number of pages"].replace("", 0, inplace=True)
        output_df["Number of pages"] = output_df["Number of pages"].astype(int)

        output_df["Retail price (inc. VAT)"].replace("", 0, inplace=True)
        output_df["Retail price (inc. VAT)"] = output_df[
            "Retail price (inc. VAT)"
        ].astype(float)

        output_df["Melitensia (1), Adult (2), Children's (3)"].replace(
            "", 0, inplace=True
        )
        output_df["Melitensia (1), Adult (2), Children's (3)"] = output_df[
            "Melitensia (1), Adult (2), Children's (3)"
        ].astype(int)

        # ToDo Remove years prior to cutoff

        # Create Formula terms
        output_df["A"] = (output_df["Number of pages"] - 1) // 100 + 1

        def get_b(item):
            if item <= 0:
                return 0
            elif item <= 10:
                return 1
            elif item <= 20:
                return 2
            elif item <= 50:
                return 3
            elif item <= 80:
                return 4
            elif item <= 100:
                return 5
            elif item <= 150:
                return 6
            else:
                return 7

        output_df["B"] = output_df["Retail price (inc. VAT)"].apply(get_b)

        output_df["Number of authors"] = output_df["Author(s)"].apply(
            lambda x: len(x.split(","))
        )
        output_df["C"] = 1 / output_df["Number of authors"]

        output_df["D"] = output_df["Melitensia (1), Adult (2), Children's (3)"]

        output_df["(A + B) x D"] = (output_df["A"] + output_df["B"]) * output_df["D"]

        all_publisher_df = all_publisher_df.append(output_df)

total_points = all_publisher_df["(A + B) x D"].sum()
# print(all_publisher_df)
# print(total_points)

# Work out E parameter
# Request user input for total funds to be distributed

funds_distributed = float(
    input(
        f"\nEnter amount of funds to be distributed\n"
        f"(Total amount - administrative expenses)\n"
        f"[Enter number without currency or commas]:\n"
    )
)

E = funds_distributed / total_points

all_publisher_df["Funds distributed"] = funds_distributed
all_publisher_df["Total points (A+B)xD for all books"] = total_points
all_publisher_df["E"] = E
all_publisher_df["Licence amount per book (A+B)xDxE"] = (
    all_publisher_df["(A + B) x D"] * all_publisher_df["E"]
)

# Export all books cleaned csv
date = datetime.datetime.now()
unique_ts_id = str(int(date.timestamp()))[-5:]
all_publisher_df.to_csv(
    f"{path}/__all_books_{date.strftime('%d%m%y')}_{unique_ts_id}.csv"
)
print("\nSpreadsheet with all publishers' books created")

# Work out publisher and author payment contributions
publishers = {}
authors = {}

for idx, book in all_publisher_df.iterrows():
    publisher = book["Publisher"].strip()
    book_authors = book["Author(s)"].split(",")

    if publisher not in publishers.keys():
        publishers[publisher] = 0
    publishers[publisher] += (
        publisher_author_ratio * book["Licence amount per book (A+B)xDxE"]
    )

    for author in book_authors:
        author = f"{[publisher]} {author.strip().title()}"
        if author not in authors.keys():
            authors[author] = 0
        authors[author] += (
            (1 - publisher_author_ratio)
            * book["Licence amount per book (A+B)xDxE"]
            * book["C"]
        )

# authors.pop("")

# Export publisher payments as csv
with open(
    f"{path}/__publishers_payment{date.strftime('%d%m%y')}_{unique_ts_id}.csv", "w"
) as file:
    file.write("Publisher,Payment\n")
    for publisher, payment in publishers.items():
        file.write(f"{publisher},{payment}\n")

print("Spreadsheet with all publishers' payments created")

# Export author payments as csv
with open(
    f"{path}/__authors_payment{date.strftime('%d%m%y')}_{unique_ts_id}.csv", "w"
) as file:
    file.write("Author,Payment\n")
    for author, payment in authors.items():
        file.write(f"{author},{payment}\n")

print("Spreadsheet with all authors' payments created")
