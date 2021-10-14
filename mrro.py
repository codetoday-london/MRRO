"""
ToDo Docs

Application runs on macOS. No Windows version is available at present
"""

import pandas as pd
import os

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
        print(filename)

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
            path + "/" + filename, names=columns, dtype="str", keep_default_na=False
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
            temp = data.loc[sections.index[idx] + 1 : sections.index[idx + 1], :]
            # Find part of string starting with 20...
            year_idx = data.loc[df_idx, "Name of Book"].find("20")
            year_value = data.loc[df_idx, "Name of Book"][year_idx : year_idx + 4]
            temp.loc[:, "Year of Publication"] = year_value

            output_df = output_df.append(temp)

        # And add final section
        idx += 1
        df_idx = sections.index[-1]
        temp = data.loc[sections.index[idx] + 1 :, :]
        year_idx = data.loc[df_idx, "Name of Book"].find("20")
        year_value = data.loc[df_idx, "Name of Book"][year_idx : year_idx + 4]
        temp.loc[:, "Year of Publication"] = year_value

        output_df = output_df.append(temp)

        output_df["Year of Publication"] = output_df["Year of Publication"].astype(int)
        # Remove rows with repeated header rows
        output_df = output_df[~output_df["Name of Book"].str.fullmatch("Name of Book")]

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

        output_df["(A + B) x C x D"] = (
            (output_df["A"] + output_df["B"]) * output_df["C"] * output_df["D"]
        )

        all_publisher_df = all_publisher_df.append(output_df)


print(all_publisher_df)
