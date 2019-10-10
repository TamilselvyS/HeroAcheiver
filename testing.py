import openpyxl

path = "HeroAcheiver.xlsx"
  
wb = openpyxl.load_workbook(path, data_only = True) 
  
sheet = wb['details']

bike=sheet.cell(2,1).value
print(bike)

sheet.cell(2,1).value = 'Bike Name'


wb.save('testing.xlsx')    