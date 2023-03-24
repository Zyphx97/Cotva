
import win32com.client
import xlwt

# Create a new Workbook object
workbook = xlwt.Workbook()

# Add a new Worksheet object to the Workbook object
worksheet = workbook.add_sheet('Hoja1')

# Write some data to the Worksheet object
worksheet.write(0, 0, 'Hello')
worksheet.write(0, 1, 'World!')

# Save the Workbook object to a file
workbook.save('example.xls')

# Open Excel application
excel = win32com.client.Dispatch("Excel.Application")

# Open workbook
workbook = excel.Workbooks.Open(r"C:\Users\PC-AMPM-04\PycharmProjects\Cotva\example.xls")

# Get the custom document properties object
custom_properties = workbook.CustomDocumentProperties

# Set the value of custom property "CustomProperty1" to "CustomValue"
custom_properties.Add("WorkbookGuid", False, 4, "ae5c34e8-16ac-4df2-85b5-0286647fd3c3")

# Save and close workbook
workbook.Save()
workbook.Close()

# Quit Excel application
excel.Quit()