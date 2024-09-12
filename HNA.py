#filepath for IRR Template:
IRR_filepath = r"C:\Users\Intern\Documents\Intern\QuaterEnd report\2. LP reporting\Claret Fund III IRR Template Q2 2024 v2.xlsx"


#filepath for Company list to split up Provison loss/FX reevaluation:
Provision_loss_filepath = r"C:\Users\Intern\Documents\Intern\QuaterEnd report\2. LP reporting\ProvisionLoss.xlsx"


#filepath for updated AM PD GP worksheet to change Revenue, EBITDA, and Cash for BCI:
AM_PD_GP_filepath = r"C:\Users\Intern\Documents\Intern\QuaterEnd report\2. LP reporting\CEGCF III Q224 AM PD GP.xlsx"


#filepath for HNA template:
HNA_filepath = r"C:\Users\Intern\Documents\Intern\QuaterEnd report\2. LP reporting\Testing\Originals\HNA Q124 CEGCF III Reporting.xlsx"


#filepath to save output document to:
Output_filepath = r"C:\Users\Intern\Documents\Intern\QuaterEnd report\2. LP reporting\Testing\Outputs\HNAoutput-final.xlsx"



#Make sure columns headers and sheet name match up










import pandas as pd
from openpyxl import load_workbook
import xlwings as xw

#Reading in IRR data
source = pd.read_excel(IRR_filepath, sheet_name='Realized Proceeds Table', skiprows=3, usecols="B:Q")
exclude_columns = ["MoC", "Gross IRR"]
columns_to_multiply = source.columns.difference(exclude_columns)
source[columns_to_multiply] = source[columns_to_multiply].map(lambda x: x * 1000 if pd.api.types.is_numeric_dtype(type(x)) else x)

#Finding places to split up data
table_starts = source[source["Company"] == "Company"].index.tolist()
table_starts.append(len(source))

#Splitting up data
Unrealised_investments = source.iloc[0:(table_starts[0] - 2)].reset_index(drop=True)
Partially_realised_investments = source.iloc[table_starts[0]+2: table_starts[1] -2].reset_index(drop=True)
Fully_realised_proceeds = source.iloc[table_starts[1]+2:table_starts[2] -5].reset_index(drop=True)

#Reading AM PD GP sheet to update monies value
source2 = pd.read_excel(AM_PD_GP_filepath, sheet_name='Companies', skiprows=1, usecols="A:M")
source3 = pd.read_excel(AM_PD_GP_filepath, sheet_name='Instruments', skiprows=2, usecols="A:V")

#Reading in provision loss companies to split column into 2
df = pd.read_excel(Provision_loss_filepath)
provision_loss_companies = df['Company'].tolist()

#Reading in HNA Data
workbook = load_workbook(HNA_filepath, data_only=False)
sheet = workbook.active
data = sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=1, max_col=23, values_only=False)
columns = [cell.value for cell in sheet[1][:23]]
hna_df = pd.DataFrame([[cell.value for cell in row] for row in data], columns=columns)

#Function to match up columns and companies and change data, bit fiddly this one
def update_hna_main_data(hna_df, Unrealised_investments):
    for index,row in hna_df.iterrows():
        company_name = str(row["Issuer"]).strip()

        match = Unrealised_investments[Unrealised_investments['Company'].str.strip()== company_name]
        if not match.empty:
            hna_df.at[index, 'Investment Cost'] = match['Investment Cost'].values[0]
            hna_df.at[index, 'Principal repayments'] = match['Principal Repayments / Disposals'].values[0]
            hna_df.at[index, 'Total Realised Proceeds'] = match['Total Realised Value'].values[0]
            hna_df.at[index, 'Latest Valuation'] = match['Fair Market Value'].values[0]
            fx_gain_loss = match['FX Gain / (Loss) and Provisions (Loans)'].values[0]
            if company_name in provision_loss_companies:
                hna_df.at[index, 'Provision/loan loss'] = fx_gain_loss
            else:
                hna_df.at[index, 'FX Movement'] = fx_gain_loss

        match_source2 = source2[source2['Company Name'].str.strip() == company_name]
        if not match_source2.empty:
            hna_df.at[index, 'Actual EBITDA'] = match_source2['EBITDA'].values[0]   

        match_source3 = source3[(source3['Company Name'].str.strip() == company_name) & (source3['Instrument Name'].str.contains('Loan', case=False, na=False))]
        if not match_source3.empty:
            hna_df.at[index, 'Actual Net Leverage'] = match_source3['Net debt\nper EBITDA.1'].values[0]  

#Using functions on dataset
update_hna_main_data(hna_df, Unrealised_investments)
update_hna_main_data(hna_df, Partially_realised_investments)
update_hna_main_data(hna_df, Fully_realised_proceeds)

#reading data back to worksheet
for row_idx, row in hna_df.iterrows():
    for col_idx, value in enumerate(row, start=1):
        sheet.cell(row=row_idx+2, column=col_idx, value=value)


workbook.save(Output_filepath)
print("All done yippee")