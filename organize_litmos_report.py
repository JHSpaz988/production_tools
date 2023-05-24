import pandas

with open("/home/jhspaz988/Downloads/Summary - People Report.csv") as file:
    data = pandas.read_csv(file)

data_keys = list(data.keys())

for index in range(1, 3):
    data.drop(data_keys[1], inplace=True, axis=1)
    data_keys.remove(data_keys[1])

data = data.sort_values(by=data_keys[-1], ascending=True)

sheets = []
shift_keys = data[data_keys[-1]].unique().tolist()
for sheet in shift_keys:
    new_data = data[(data[data_keys[-1]] == sheet) & (data[data_keys[1]] < 100)]
    sheet = pandas.DataFrame(new_data)
    sheets.append(new_data)

blank_sheet = data[(data[data_keys[-1]].isna()) & (data[data_keys[1]] < 100)]
blank_sheet = pandas.DataFrame(blank_sheet)

with pandas.ExcelWriter("/home/jhspaz988/Documents/shift_sorted.xlsx", "openpyxl") as writer:
    for (index, sheet) in enumerate(sheets):
        if sheet.empty:
            pass
        else:
            sheet.to_excel(writer, sheet_name=shift_keys[index], index=False)
    blank_sheet.to_excel(writer, sheet_name="None", index=False)
