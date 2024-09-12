#filepath for IRR Template:
IRR_filepath = r"C:\Users\Intern\Documents\Intern\QuaterEnd report\2. LP reporting\Claret Fund III IRR Template Q2 2024 v2.xlsx"



#filepath for BBB Template:
BBB_filepath = r"C:\Users\Intern\Documents\Intern\QuaterEnd report\2. LP reporting\Testing\Originals\BBB Q124 Debt Template_CEGCF III - copy.xlsx"


#filepath to save output document to:
Output_filepath = r"C:\Users\Intern\Documents\Intern\QuaterEnd report\2. LP reporting\Testing\Outputs\BBBoutput-final.xlsx"





#proper code beneath here. only change column headers if need be











import pandas as pd
from openpyxl import load_workbook
import xlwings as xw

#Reading in IRR template data
source = pd.read_excel(IRR_filepath, sheet_name='Realized Proceeds Table', skiprows=3, usecols="B:Q")
exclude_columns = ["MoC", "Gross IRR"]
columns_to_multiply = source.columns.difference(exclude_columns)
source[columns_to_multiply] = source[columns_to_multiply].map(lambda x: x * 1000 if pd.api.types.is_numeric_dtype(type(x)) else x)

#Fidning places to split data
table_starts = source[source["Company"] == "Company"].index.tolist()
table_starts.append(len(source))

#Splitting up data
Unrealised_investments = source.iloc[0:(table_starts[0] - 2)].reset_index(drop=True)
Partially_realised_investments = source.iloc[table_starts[0]+2: table_starts[1] -2].reset_index(drop=True)
Fully_realised_proceeds = source.iloc[table_starts[1]+2:table_starts[2] -5].reset_index(drop=True)

#Reading in and sorting KfW data
workbook = load_workbook(BBB_filepath, data_only=False)
sheet = workbook["Debt Template"]
data = sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=1, max_col=78, values_only=False)
columns = [cell.value for cell in sheet[1][0:78]]
BBB_df = pd.DataFrame([[cell.value for cell in row] for row in data], columns=columns)


#Function to change all Values. Fairly simple one
def update_BBB_data(BBB_df, Unrealised_investments):
    for index,row in BBB_df.iterrows():

        company_name = str(row["Enterprise Name"]).strip()

        match = Unrealised_investments[Unrealised_investments['Company'].str.strip()== company_name]
        if not match.empty:
            BBB_df.at[index, 'Gross Investment Amount'] = match['Investment Cost'].values[0]
            BBB_df.at[index, 'Unrealised Cost'] = (match['Investment Cost'].values[0] - match['Principal Repayments / Disposals'].values[0]) 
            BBB_df.at[index, 'Unrealised Value'] = match['Fair Market Value'].values[0]
            BBB_df.at[index, 'Realised Cost'] = match['Principal Repayments / Disposals'].values[0]
            BBB_df.at[index, 'Other Income'] = (match['Fees'].values[0] + match['Realised Gain/(Loss) (Equity / Warrants)'].values[0])
            BBB_df.at[index, 'Interest Payable'] = match['Interest Received'].values[0]
            BBB_df.at[index, 'Interest Received'] = match['Interest Received'].values[0]

#use function to update all data
update_BBB_data(BBB_df, Unrealised_investments)
update_BBB_data(BBB_df, Partially_realised_investments)
update_BBB_data(BBB_df, Fully_realised_proceeds)

print(BBB_df)

#reading data back to worksheet
for row_idx, row in BBB_df.iterrows():
    for col_idx, value in enumerate(row[:78], start=1):
        sheet.cell(row=row_idx+2, column=col_idx, value=value)

workbook.save(Output_filepath)
print("All done yippee")