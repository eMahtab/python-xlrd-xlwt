import xlwt

workbook=xlwt.Workbook()
sheet=workbook.add_sheet("awesome")

#Defining Styles
style1 = xlwt.easyxf('font: bold 1, color blue;')
style2= xlwt.easyxf('pattern: pattern solid, fore_color green;')

#Headers
sheet.write(0,0,"Drinks",style1)
sheet.write(0,1,"Price",style1)

sheet.write(1,0,"Coffee")
sheet.write(1,1,150)

sheet.write(2,0,"Latte")
sheet.write(2,1,130)

sheet.write(3,0,"Cappuccino")
sheet.write(3,1,170)

sheet.write(4,0,"Irish Coffee")
sheet.write(4,1,350)

#Applying Formulas
sheet.write(5,1,xlwt.Formula('SUM(B2:B5)'),style2)


#Setting Widths
sheet.col(0).width=6000
sheet.col(1).width=6000

workbook.save("change.xls")