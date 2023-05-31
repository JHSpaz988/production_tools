import pandas

IN_FILE = "/home/jhspaz988/Downloads/ReCert Dates.xlsx"
OUT_FILE = "/home/jhspaz988/Documents/sorted_recert_file.xlsx"


def sort(in_file, out_file):
    headers = ["Name-Last", "Name-First", "Line", "Cert Date", "Shift", "Notes"]
    data = pandas.read_excel(in_file, names=headers)

    data = data.sort_values(by=headers[-2], ascending=True)

    shift_keys = data[headers[-2]].unique().tolist()
    with pandas.ExcelWriter(out_file, "openpyxl") as writer:
        for sheet in shift_keys:
            new_data = data[(data[headers[-2]] == sheet)]
            new_data = pandas.DataFrame(new_data)
            new_data = new_data.sort_values(by=[headers[0]], ascending=True)
            new_data[headers[3]] = pandas.to_datetime(new_data[headers[3]]).dt.date
            if new_data.empty:
                pass
            else:
                new_data.to_excel(writer, sheet_name=sheet, index=False)


sort(IN_FILE, OUT_FILE)
