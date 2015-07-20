# Excel Macros 

*Note: Windows only*

Folder structure :

**ASBIS\ (wildcard 2 letters of the file name)/year/FX\ Balance\ Sheet/month'15/File_Name.xlsx**</br>

**AEL/year/File_Name.xslx** 
*Note: File_Name.xslx must contain month in format "mmm"*</br>

**Macros/DailyBS/File_Name.xlsx**</br>

**Macros/script_name.xlsm**</br>

###dailyBS_Subs.xlsm

Generate a file used for taking Debtor&Creditor aging analysis, adding a column for rates, calculating it, adding formatting to AmountTC, AmountDC, creating Pivot table For both EUR & USD. 

Place credit&debit.xlsx file(s) in DailyBS folder.

1. Go to the folder that contains the previous file for previous day, and get the first Worksheet that contains the previous Balance Sheet. The values for Inventory, Receivables, Payables is NOT changed in the new Worksheet. 
2. Save the Workbook in the folder corresponding to the company.
3. In Macros folder,add a folder called DailyBS which will contain the Credit/Debit worksheet(s) for companies (we do not put here the AEL one, cause we have a separate script for it). *Note: Naming of credit debit workbooks - the first 2 letters MUST correspond to folder names (e.g. BY, KZ) - this MUST be the first two letters of the name + wildcard.*
4. Take the file, (e.g. "KZ Debtor-Creditors"), check the name, format, goes back to main folder, finds "ASBIS KZ", goes inside, check folder year, check month year, check for the latest file by, 
5. Append the first worksheet to the credit&debit.xslx file as Sheet 1.
6. Save the file in the corresponding folder "AS" + 2 letter code of credit&debit.xslx folder + Date("dd mm yy")

###balanceSheetMacrosAEL.xlsm

Generate one large FX Balance Sheet, appending credit&debit.xlsx to main one.

Place credit&debit.xlsx file in "AEL*" folder.

*Note: Edit the code to change the folder name if needed*

1. Take credit&debit.xlsx file, adds columns C (Category) & K (Rate), calculates the rate, formats the text according to accounting standards. 
2. Open AEL directory, gets folder by year, finds the latest .xslx file based on month ("mmm" format), appends credit&debit.xlsx worksheets as sheet 1&2, converts the names to "Pay" + Date("dd.mm.yy") & "Rec" + Date("dd.mm.yy").
3. Column 3 is compared to Sheet 3 & 4 accordingly (previous Pay & Rec), to obtain based on Column 1, IndexMatch of company Category Type.
4. Remove any Currency based on Column 5 which is non EUR.
5. Filter Column 3 to show "N/A" only rows.
6. Generate Pivot Table for worksheet 1 & 2.
7. Save file File_Name + Date("dd mm yy")


