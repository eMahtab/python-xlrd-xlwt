
import os
import xlrd

dataDir="D:/webinarData/python"
files=os.listdir(dataDir)

leads={}
previous_leads=len(leads)

totalRepeats=0
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

    total_rows=sheet.nrows
    #print("Total Rows : ",total_rows)
    previous_leads=len(leads)
    print("Previous Leads : ",previous_leads)
    start=8
    repeats=0;
    for i in range(start,total_rows):
        #print("i : ",i)
        #print("Key : ",sheet.cell_value(i,4))
        #print("Value : ",sheet.cell_value(i,3))
        if(sheet.cell_value(i,4) in leads):
            oldValue=leads[sheet.cell_value(i,4)]
            newValue=oldValue+1
            leads[sheet.cell_value(i,4)]=newValue
        else:
            leads[sheet.cell_value(i,4)]=1


    new_leads=len(leads)-previous_leads
    repeats=registered-new_leads
    totalRepeats=totalRepeats+repeats
    print("Repeats : ",repeats)
    print("New Leads : ",new_leads)
    print("Total Repeats : ",totalRepeats)
    print("Total Leads : ",len(leads))
    #print(leads)
    print("-------------------------------------")
    print("-------------------------------------")


