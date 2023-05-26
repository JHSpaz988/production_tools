import pandas
import sys
import os


def organize_litmos_report(ofp: str, wfp: str) -> None:
    with open(ofp) as file:
        data = pandas.read_csv(file)

    data_keys = list(data.keys())

    for index in range(1, 3):
        data.drop(data_keys[1], inplace=True, axis=1)
        data_keys.remove(data_keys[1])

    data = data.sort_values(by=data_keys[-1], ascending=True)

    sheets = []
    shift_keys = data[data_keys[-1]].unique().tolist()
    for sheet in shift_keys:
        new_data = data[(data[data_keys[-1]] == sheet)
                        & (data[data_keys[1]] < 100)]
        sheet = pandas.DataFrame(new_data)
        sheets.append(new_data)

    blank_sheet = data[(data[data_keys[-1]].isna())
                       & (data[data_keys[1]] < 100)]
    blank_sheet = pandas.DataFrame(blank_sheet)

    with pandas.ExcelWriter(wfp, "openpyxl") as writer:
        for (index, sheet) in enumerate(sheets):
            if sheet.empty:
                pass
            else:
                sheet.to_excel(
                    writer, sheet_name=shift_keys[index], index=False)
        blank_sheet.to_excel(writer, sheet_name="None", index=False)


open_file_path = os.path.abspath(sys.argv[1])
write_file_path = os.path.abspath(sys.argv[2])
organize_litmos_report(ofp=open_file_path, wfp=write_file_path)
