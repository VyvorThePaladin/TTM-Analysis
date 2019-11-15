import openpyxl

# Load the two workbooks
wb_orig = openpyxl.load_workbook('1-orig.xlsx')
wb_adj = openpyxl.load_workbook('2-adj.xlsx')

# Creating new workbook and initializing titles
wb_new = openpyxl.Workbook()
sheet_new = wb_new.active
sheet_new.title = 'TTM & Adj Combined'
wb_new.save('3-combinedRes.xlsx')

# Get a handle on the original and adj TTM sheet
sheet_orig = wb_orig[wb_orig.sheetnames[0]]
sheet_adj = wb_adj[wb_adj.sheetnames[0]]

# Inserts Employee, UID, Activity Code, hours_decimal from orig into new sheet's columns
for i in range(1, sheet_orig.max_row+1):
    sheet_new['A' + str(i)] = sheet_orig['H' + str(i)].value
    sheet_new['B' + str(i)] = sheet_orig['G' + str(i)].value
    sheet_new['C' + str(i)] = sheet_orig['I' + str(i)].value
    sheet_new['D' + str(i)] = sheet_orig['M' + str(i)].value

#-----TO-DO-------#
# 1. Modify Employee name in adj_sheet and insert modified val in table

# Inserts Employee, UID, Activity Code, hours_decimal from orig into new sheet's columns
for j in range(2, sheet_adj.max_row+1):
    # sheet_new['A' + str(sheet_orig.max_row+j-1)] = sheet_adj['I' + str(j)].value    
    sheet_new['B' + str(sheet_orig.max_row+j-1)] = sheet_adj['B' + str(j)].value
    sheet_new['C' + str(sheet_orig.max_row+j-1)] = sheet_adj['D' + str(j)].value    
    sheet_new['D' + str(sheet_orig.max_row+j-1)] = sheet_adj['E' + str(j)].value

