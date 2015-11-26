import os
import xlrd

dataDir="D:/webinarData/git"
files=os.listdir(dataDir)

git_leads=0

for name in files:
    file_name=name
    workbook=xlrd.open_workbook(dataDir+"/"+file_name)
    sheet=workbook.sheet_by_index(0)
    webinar_name=sheet.cell_value(0,1)
    print("Webinar Name : ",webinar_name)

    webinar_date=sheet.cell_value(4,1)
    print("Webinar Date : ",webinar_date)

    registered=(int)(sheet.cell_value(4,3))
    print("Registered ",registered)


    attended=(int)(sheet.cell_value(4,4))
    print("Attended ",attended)

    attendance_ratio=registered/attended
    print("Attendance Ratio : ",attendance_ratio)

    print("-------------------------------------")
    print("-------------------------------------")

