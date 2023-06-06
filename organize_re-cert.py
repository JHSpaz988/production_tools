import pandas as pd

IN_FILE = "/home/jhspaz988/Downloads/ReCert Dates.xlsx"
OUT_FILE = "/home/jhspaz988/Documents/sorted_recert_file.xlsx"


def sort(in_file, out_file):
    headers = ["Name-Last", "Name-First", "Line", "Cert Date", "Shift", "Notes"]
    data = pd.read_excel(in_file, names=headers)

    data.sort_values(
        by=[headers[-2], headers[0]],
        inplace=True,
        ascending=True
        )
    data[headers[3]] = pd.to_datetime(data[headers[3]]).dt.date


    shift_keys = data[headers[-2]].unique()
    with pd.ExcelWriter(out_file, "openpyxl") as writer:
        for sheet in shift_keys:
            new_data = data[(data[headers[-2]] == sheet)]
            if not new_data.empty:
                new_data.to_excel(writer, sheet_name=sheet, index=False)


sort(IN_FILE, OUT_FILE)
