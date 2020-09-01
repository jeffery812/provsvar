import pandas as pd


origin_provsvar = 'Provsvar-fixed.xlsx'
output_file = 'provsvar-summary.xlsx'
config = pd.read_excel("config.xlsx")


class LabItem:
    def __init__(self, name, min, max):
        self.name = name
        self.min = min
        self.max = max

# Provsvar-fixed
provsvar = pd.read_excel(origin_provsvar, sheet_name='Sheet1')

# Apply a conditional format to the cell range.
# worksheet.conditional_format(0, 0, 200, 200, {'type': 'cell', 'criteria': '>', 'value': 1.33, 'format': format1})


bad_style = 'background-color: #FFC7CE'

lab_items = []
for i, row in config.iterrows():
    lab_item = LabItem(row['Name'], row['Min'], row['Max'])
    lab_items.append(lab_item)

print(config)

for row in provsvar.index:
    for lab_item in lab_items:
        value = provsvar[lab_item.name].at[row]
        if isinstance(value, str):
            # Fix abnormal cell: "0,33, 0,33", convert "," to "."
            provsvar[lab_item.name].at[row] = float(value.split(", ")[0].replace(",", "."))

headers = list(map(lambda x: x.name, lab_items))
print(f'headers: {headers}')
provsvar_new = pd.DataFrame(provsvar, columns=headers)

provsvar_new.style\
    .applymap(lambda x: bad_style if x > 1.33 else '',
              subset=['P--Calciumjon, fri (mmol/L)'])\
    .to_excel(output_file, engine='openpyxl', sheet_name='summary', index=False)



# provsvar_styled.set_index('Datum', inplace=True)
# provsvar_styled.to_excel(output_file, engine='openpyxl', sheet_name='summary', columns=headers, index=False)