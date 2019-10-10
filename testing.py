import openpyxl
import datetime
from docxtpl import DocxTemplate
path = "HeroAcheiver.xlsx"
  
wb = openpyxl.load_workbook(path, data_only = True) 
  
sheet = wb['details']
maxrow = sheet.max_row

print(maxrow)

# sheet.cell(2,1).value = 'Bike Name'
# wb.save('testing.xlsx')    

document = DocxTemplate('template.docx')

for x in range(2, maxrow+1):
    bname = sheet.cell(x, 1).value
    bprice = sheet.cell(x, 2).value
    bHelmet = sheet.cell(x, 3).value
    purchaseExpense = sheet.cell(x, 4).value
    serviceExpense = sheet.cell(x, 5).value
    fuelExpense = sheet.cell(x, 6).value
    totalExpense = sheet.cell(x, 7).value
    mileage = sheet.cell(x,8).value
    purchasedDate = sheet.cell(x,9).value
    purchasedDate = datetime.strptime(purchasedDate, '%d/%b/%Y')

    context = { 'BNAME' : bname,
                'BPRICE': bprice,
                'HELMET': bHelmet,
                'PURCHASE_EXPENSE': purchaseExpense,
                'SERVICE_EXPENSE': serviceExpense,
                'FUEL_EXPENSE': fuelExpense,
                'TOTAL_EXPENSE': totalExpense,
                'MILEAGE': mileage,
                'PURCHASED_DATE': purchasedDate

                 }

    document.render(context)
    document.save("generated_doc.docx")