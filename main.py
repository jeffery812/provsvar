import pandas as pd


origin_provsvar = 'Provsvar-fixed.xlsx'
output_file = 'provsvar-summary.xlsx'
config = pd.read_excel("config.xlsx")

# Provsvar-fixed
provsvar = pd.read_excel(origin_provsvar, sheet_name='Sheet1')

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter(output_file, engine='xlsxwriter', options={'strings_to_numbers': True})

provsvar.to_excel(writer, sheet_name='Sheet1')
# Get the xlsxwriter workbook and worksheet objects.
workbook = writer.book
worksheet = writer.sheets['Sheet1']
# Add a format. Light red fill with dark red text.
format1 = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})

# Apply a conditional format to the cell range.
# worksheet.conditional_format(0, 0, 200, 200, {'type': 'cell', 'criteria': '>', 'value': 1.33, 'format': format1})
headers = []
for i in config.index:
    item = config['Name'].at[i]
    headers.append(item)

for row in provsvar.index:
    for header in headers:
        value = provsvar[header].at[row]
        if isinstance(value, str):
            # Fix abnormal cell: "0,33, 0,33", convert "," to "."
            provsvar[header].at[row] = float(value.split(", ")[0].replace(",", "."))

provsvar.set_index('Datum', inplace=True)
provsvar.to_excel(output_file, sheet_name='summary', columns=headers, )