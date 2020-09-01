import pandas as pd


origin_provsvar = 'Provsvar-fixed.xlsx'
output_file = 'provsvar-summary.xlsx'
config = pd.read_excel("config.xlsx")



class LabReference:
    def __init__(self, name, min, max):
        self.name = name
        self.min = min
        self.max = max


# Provsvar-fixed
provsvar = pd.read_excel(origin_provsvar, sheet_name='Sheet1')

# Apply a conditional format to the cell range.
# worksheet.conditional_format(0, 0, 200, 200, {'type': 'cell', 'criteria': '>', 'value': 1.33, 'format': format1})


bad_style = 'background-color: #FFC7CE'


def get_style(val):
    if not isinstance(val, int) and not isinstance(val, float):
        return ''

    if val > 1.33 or val < 1.05:
        return bad_style
    else:
        return ''


lab_references = []
for i, row in config.iterrows():
    lab_item = LabReference(row['Name'], row['Min'], row['Max'])
    lab_references.append(lab_item)


for row in provsvar.index:
    for lab_item in lab_references:
        value = provsvar[lab_item.name].at[row]
        if isinstance(value, str):
            # Fix abnormal cell: "0,33, 0,33", convert "," to "."
            provsvar[lab_item.name].at[row] = float(value.split(", ")[0].replace(",", "."))

headers = list(map(lambda x: x.name, lab_references))

provsvar_new = pd.DataFrame(provsvar, columns=headers)

provsvar_new.style.set_properties(**{'background-color': 'white',
                                     'color': 'black',
                                     'border-color': 'black',
                                     'border-width': '1px',
                                     'border-style': 'solid'})\
    .applymap(lambda x: bad_style if x > lab_references[1].max or x < lab_references[1].min else '', subset=[lab_references[1].name]) \
    .applymap(lambda x: bad_style if x > lab_references[2].max or x < lab_references[2].min else '', subset=[lab_references[2].name]) \
    .applymap(lambda x: bad_style if x > lab_references[3].max or x < lab_references[3].min else '', subset=[lab_references[3].name])\
    .applymap(lambda x: bad_style if x > lab_references[4].max or x < lab_references[4].min else '', subset=[lab_references[4].name])\
    .applymap(lambda x: bad_style if x > lab_references[5].max or x < lab_references[5].min else '', subset=[lab_references[5].name])\
    .applymap(lambda x: bad_style if x > lab_references[6].max or x < lab_references[6].min else '', subset=[lab_references[6].name]) \
    .applymap(lambda x: bad_style if x > lab_references[7].max or x < lab_references[7].min else '', subset=[lab_references[7].name])\
    .applymap(lambda x: bad_style if x > lab_references[8].max or x < lab_references[8].min else '', subset=[lab_references[8].name]) \
    .applymap(lambda x: bad_style if x > lab_references[9].max or x < lab_references[9].min else '', subset=[lab_references[9].name])\
    .applymap(lambda x: bad_style if x > lab_references[10].max or x < lab_references[10].min else '', subset=[lab_references[10].name])\
    .applymap(lambda x: bad_style if x > lab_references[11].max or x < lab_references[11].min else '', subset=[lab_references[11].name])\
    .applymap(lambda x: bad_style if x > lab_references[12].max or x < lab_references[12].min else '', subset=[lab_references[12].name])\
    .applymap(lambda x: bad_style if x > lab_references[13].max or x < lab_references[13].min else '', subset=[lab_references[13].name])\
    .applymap(lambda x: bad_style if x > lab_references[14].max or x < lab_references[14].min else '', subset=[lab_references[14].name])\
    .applymap(lambda x: bad_style if x > lab_references[15].max or x < lab_references[15].min else '', subset=[lab_references[15].name])\
    .applymap(lambda x: bad_style if x > lab_references[16].max or x < lab_references[16].min else '', subset=[lab_references[16].name])\
    .applymap(lambda x: bad_style if x > lab_references[17].max or x < lab_references[17].min else '', subset=[lab_references[17].name])\
    .applymap(lambda x: bad_style if x > lab_references[18].max or x < lab_references[18].min else '', subset=[lab_references[18].name])\
    .applymap(lambda x: bad_style if x > lab_references[19].max or x < lab_references[19].min else '', subset=[lab_references[19].name])\
    .applymap(lambda x: bad_style if x > lab_references[20].max or x < lab_references[20].min else '', subset=[lab_references[20].name])\
    .to_excel(output_file, engine='openpyxl', sheet_name='summary', index=False)

print("complete")

