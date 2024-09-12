#filepath for IRR Template:
IRR_filepath = r"C:\Users\Intern\Documents\Intern\QuaterEnd report\2. LP reporting\Claret Fund III IRR Template Q2 2024 v2.xlsx"



#filepath for EIF Template:
KfW_filepath = r"C:\Users\Intern\Documents\Intern\QuaterEnd report\2. LP reporting\Testing\Originals\CEGCF III Q124 KfW Report.xlsx"


#filepath to save output document to:
Output_filepath = r"C:\Users\Intern\Documents\Intern\QuaterEnd report\2. LP reporting\Testing\Outputs\KwFoutput-Final.xlsx"


















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
print(Unrealised_investments.columns)
Partially_realised_investments = source.iloc[table_starts[0]+2: table_starts[1] -2].reset_index(drop=True)
Fully_realised_proceeds = source.iloc[table_starts[1]+2:table_starts[2] -5].reset_index(drop=True)
Fully_realised_proceeds.fillna(0, inplace=True)
Partially_realised_investments.fillna(0, inplace=True)

# Adding 'Total Income' column
Fully_realised_proceeds['Total Income'] = (
    Fully_realised_proceeds['Interest Received'] + 
    Fully_realised_proceeds['Fees'] + 
    Fully_realised_proceeds['Realised Gain/(Loss) (Equity / Warrants)'])

Partially_realised_investments['Total Income'] = (
    Partially_realised_investments['Interest Received'] + 
    Partially_realised_investments['Fees'] + 
    Partially_realised_investments['Realised Gain/(Loss) (Equity / Warrants)'])


print(Fully_realised_proceeds[['Interest Received', 'Fees', 'Realised Gain/(Loss) (Equity / Warrants)', 'Total Income']])
print(Partially_realised_investments[['Interest Received', 'Fees', 'Realised Gain/(Loss) (Equity / Warrants)', 'Total Income']])

#Reading in and sorting KfW data
workbook = load_workbook(KfW_filepath, data_only=False)
sheet = workbook["Portfolio_Positions"]
data = sheet.iter_rows(min_row=4, max_row=sheet.max_row, min_col=5, max_col=21, values_only=False)
columns = [cell.value for cell in sheet[2][4:21]]
KfW_df = pd.DataFrame([[cell.value for cell in row] for row in data], columns=columns)


#Fucntions to do change and update values
def update_KfW_PR_data(KfW_df, Unrealised_investments):
    for index,row in KfW_df.iterrows():
        company_name = str(row["""Company name 
(Long Name)"""]).strip()
        
        match = Unrealised_investments[Unrealised_investments['Company'].str.strip()== company_name]

        if not match.empty:
            KfW_df.at[index, """Total investment cost
(as of reporting date in reporting currency)"""] = match['Investment Cost'].values[0]
            
            KfW_df.at[index, """Proceeds
(as of reporting date in reporting currency)"""] = match['Principal Repayments / Disposals'].values[0]
            
            KfW_df.at[index, """ Income
 (as of reporting date in reporting currency)"""] = (match['Interest Received'].values[0] + match['Fees'].values[0] + match['Realised Gain/(Loss) (Equity / Warrants)'].values[0])
        
            KfW_df.at[index, """Fair Value 
(as of reporting date in reporting currency)"""] = match['Fair Market Value'].values[0]
            
            KfW_df.at[index, """IRR 
(as of reporting date)"""] = match['Gross IRR'].values[0]
            
            KfW_df.at[index, """MOIC
 (as of reporting date)"""] = match['MoC'].values[0]
            
            KfW_df.at[index, "Company Status"] = "unrealized (Active)"
            
def update_KfW_UR_data(KfW_df, Partially_realised_investments):
    for index,row in KfW_df.iterrows():
        company_name = row["""Company name 
(Long Name)"""]
        match = Partially_realised_investments[Partially_realised_investments['Company']== company_name]
        if not match.empty:
            KfW_df.at[index, """Total investment cost
(as of reporting date in reporting currency)"""] = match['Investment Cost'].values[0]
            
            KfW_df.at[index, """Proceeds
(as of reporting date in reporting currency)"""] = match['Principal Repayments / Disposals'].values[0]
            
            KfW_df.at[index, """ Income
 (as of reporting date in reporting currency)"""] = (match['Interest Received'].values[0] + match['Fees'].values[0] + match['Realised Gain/(Loss) (Equity / Warrants)'].values[0])
            
            KfW_df.at[index, """Fair Value 
(as of reporting date in reporting currency)"""] = match['Fair Market Value'].values[0]
            
            KfW_df.at[index, """IRR 
(as of reporting date)"""] = match['Gross IRR'].values[0]
            
            KfW_df.at[index, """MOIC
 (as of reporting date)"""] = match['MoC'].values[0]
            
            KfW_df.at[index, "Company Status"] = "Partially realized"            

def update_KfW_FR_data(KfW_df, Fully_realised_proceeds):
    for index,row in KfW_df.iterrows():
        company_name = row["""Company name 
(Long Name)"""]
        match = Fully_realised_proceeds[Fully_realised_proceeds['Company']== company_name]
        if not match.empty:
            KfW_df.at[index, """Total investment cost
(as of reporting date in reporting currency)"""] = match['Investment Cost'].values[0]
            
            KfW_df.at[index, """Proceeds
(as of reporting date in reporting currency)"""] = match['Principal Repayments / Disposals'].values[0]
            
            KfW_df.at[index, """ Income
 (as of reporting date in reporting currency)"""] = (match['Interest Received'].values[0] + match['Fees'].values[0] + match['Realised Gain/(Loss) (Equity / Warrants)'].values[0])
            
            KfW_df.at[index, """Fair Value 
(as of reporting date in reporting currency)"""] = match['Fair Market Value'].values[0]
            
            KfW_df.at[index, """IRR 
(as of reporting date)"""] = match['Gross IRR'].values[0]
            
            KfW_df.at[index, """MOIC
 (as of reporting date)"""] = match['MoC'].values[0]
            
            KfW_df.at[index, "Company Status"] = "Realized (exited)"  

#implementing functions
update_KfW_PR_data(KfW_df, Unrealised_investments)
update_KfW_UR_data(KfW_df, Partially_realised_investments)
update_KfW_FR_data(KfW_df, Fully_realised_proceeds)

#fixing data size
if len(KfW_df.columns) > 17:
    eif_df=KfW_df.iloc[:,:17]

#reading data back to worksheet
for row_idx, row in KfW_df.iterrows():
    for col_idx, value in enumerate(row[:17], start=5):
        sheet.cell(row=row_idx+4, column=col_idx, value=value)

workbook.save(Output_filepath)
print("All done yippee")
