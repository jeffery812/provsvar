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


bad_style_low = 'background-color: #FFFB00'
bad_style_high = 'background-color: #FFC7CE'


def get_style(val, lab_reference):
    if not isinstance(val, int) and not isinstance(val, float):
        return ''
    if val < lab_reference.min:
        return bad_style_low
    elif val > lab_reference.max:
        return bad_style_high
    else:
        return ''


lab_references = []
for i, row in config.iterrows():
    '''
    Name	                            Min	Max
    Datum		
    S-ACE (E/L)                     	0	70
    P--25-OH Vitamin D2+D3 (nmol/L)	    50	250
    S-1,25-OH-Vitamin D (pmol/L)	    48	190
    ......
    
    '''
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
                                     'border-style': 'solid'}) \
    .applymap(lambda x: get_style(x, lab_references[1]), subset=[lab_references[1].name]) \
    .applymap(lambda x: get_style(x, lab_references[1]), subset=[lab_references[1].name]) \
    .applymap(lambda x: get_style(x, lab_references[2]), subset=[lab_references[2].name]) \
    .applymap(lambda x: get_style(x, lab_references[3]), subset=[lab_references[3].name]) \
    .applymap(lambda x: get_style(x, lab_references[4]), subset=[lab_references[4].name]) \
    .applymap(lambda x: get_style(x, lab_references[5]), subset=[lab_references[5].name]) \
    .applymap(lambda x: get_style(x, lab_references[6]), subset=[lab_references[6].name]) \
    .applymap(lambda x: get_style(x, lab_references[7]), subset=[lab_references[7].name]) \
    .applymap(lambda x: get_style(x, lab_references[8]), subset=[lab_references[8].name]) \
    .applymap(lambda x: get_style(x, lab_references[9]), subset=[lab_references[9].name]) \
    .applymap(lambda x: get_style(x, lab_references[10]), subset=[lab_references[10].name]) \
    .applymap(lambda x: get_style(x, lab_references[11]), subset=[lab_references[11].name]) \
    .applymap(lambda x: get_style(x, lab_references[12]), subset=[lab_references[12].name]) \
    .applymap(lambda x: get_style(x, lab_references[13]), subset=[lab_references[13].name]) \
    .applymap(lambda x: get_style(x, lab_references[14]), subset=[lab_references[14].name]) \
    .applymap(lambda x: get_style(x, lab_references[15]), subset=[lab_references[15].name]) \
    .applymap(lambda x: get_style(x, lab_references[16]), subset=[lab_references[16].name]) \
    .applymap(lambda x: get_style(x, lab_references[17]), subset=[lab_references[17].name]) \
    .applymap(lambda x: get_style(x, lab_references[18]), subset=[lab_references[18].name]) \
    .applymap(lambda x: get_style(x, lab_references[19]), subset=[lab_references[19].name]) \
    .applymap(lambda x: get_style(x, lab_references[20]), subset=[lab_references[20].name]) \
    .applymap(lambda x: get_style(x, lab_references[21]), subset=[lab_references[21].name]) \
    .applymap(lambda x: get_style(x, lab_references[22]), subset=[lab_references[22].name]) \
    .to_excel(output_file, engine='openpyxl', sheet_name='summary', index=False)

print("complete")

