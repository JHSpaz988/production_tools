import pandas as pd
import sys
import os


def sort(ofp: str, wfp: str) -> None:
    data = pd.read_csv(ofp)
    data_keys = list(data.columns)
    data.sort_values(
        by=[data_keys[-1], data_keys[2]],
        inplace=True,
        ascending=True
        )
    shift_keys = data[data_keys[-1]].unique()

    with pd.ExcelWriter(wfp, "openpyxl") as writer:
        for shift in shift_keys:
            new_data = data[(data[data_keys[-1]] == shift) & (data[data_keys[3]] < 100)]
            if not new_data.empty:
                new_data.to_excel(writer, sheet_name=str(shift), index=False)
        blank_sheet = data[(data[data_keys[-1]].isna()) & (data[data_keys[3]] < 100)]
        blank_sheet.to_excel(writer, sheet_name="None", index=False)


open_file_path = os.path.abspath(sys.argv[1])
write_file_path = os.path.abspath(sys.argv[2])
sort(ofp=open_file_path, wfp=write_file_path)
