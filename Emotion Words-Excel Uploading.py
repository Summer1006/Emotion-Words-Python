import xlrd
Emowords = xlrd.open_workbook('C:\\Users\\niehu\\OneDrive - University of St. Thomas\\MsDS\\A - NPL\\Emowords - Python Practice.xlsx')
sheet = Emowords.sheet_by_index(0)
sheet.cell_value(0, 0)
#print the number of rows and the number of columns
print(sheet.nrows)
print(sheet.ncols)
#extract the column header
for i in range(sheet.ncols):
    print(sheet.cell_value(0, i))
# exrtact the first colmn
for i in range(sheet.nrows):
    print(sheet.cell_value(i,0))
# how to print everything in an excel?







