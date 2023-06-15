# import sys
# import os
import pandas as pd

OPEN_FILE_PATH = "/home/jhspaz988/Downloads/Summary - People Report.csv"
WRITE_FILE_PATH = "/home/jhspaz988/Documents/Sorted_Summary.xlsx"

def sort(ofp: str, wfp: str) -> None:
    headers = ["Username", "First Name", "Last Name", "% Complete", "Is Complete", "Courses Assigned", "Courses Completed", "Crew"]
    data = pd.read_csv(ofp, names=headers)
    data.sort_values(
        by=[headers[-1], headers[2]],
        inplace=True,
        ascending=True
        )
    shift_keys = data[headers[-1]].unique()

    data[headers[3]] = pd.to_numeric(data[headers[3]], errors='coerce')

    with pd.ExcelWriter(wfp, "openpyxl") as writer:
        for shift in shift_keys:
            new_data = data[(data[headers[-1]] == shift) & (data[headers[3]] < 100)]
            if not new_data.empty:
                new_data.to_excel(writer, sheet_name=str(shift), index=False)
        blank_sheet = data[(data[headers[-1]].isna()) & (data[headers[3]] < 100)]
        blank_sheet.to_excel(writer, sheet_name="None", index=False)


# OPEN_FILE_PATH = os.path.abspath(sys.argv[1])
# WRITE_FILE_PATH = os.path.abspath(sys.argv[2])
sort(ofp=OPEN_FILE_PATH, wfp=WRITE_FILE_PATH)
