import xlrd

file_location="D:/webinarData/git/Hacking Git and GitHub - Attendee Report.xlsx"
workbook=xlrd.open_workbook(file_location)
sheet=workbook.sheet_by_index(0)

webinar_name=sheet.cell_value(0,1)
print("Webinar : ",webinar_name)

webinar_date=sheet.cell_value(4,1)
print("Webinar Date : ",webinar_date)

registered=(int)(sheet.cell_value(4,3))
print("Registered ",registered)


attended=(int)(sheet.cell_value(4,4))
print("Attended ",attended)

attendance_ratio=registered/attended
print("Attendance Ratio : ",attendance_ratio)

number_of_rows=sheet.nrows

