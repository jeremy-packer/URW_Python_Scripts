import win32com.client
import os

xl = win32com.client.Dispatch("Excel.Application")
xl.Visible = False

folder_path = r'C:\Users\jpacker\Desktop\Center Reports'
save_path = r'C:\Users\jpacker\Desktop\Center Reports'

for file in os.listdir(folder_path):
	try:
		wb = xl.Workbooks.Open(folder_path + "\\" + file)
		num_sheets = wb.Worksheets.Count
		
		#reverse order for new occ cost reports
		for i in range(num_sheets,0,-1):
			ws = wb.Worksheets[i - 1]

		wb.WorkSheets.Select()
		wb.ActiveSheet.ExportAsFixedFormat(0, save_path + '\\' + file[:-4] + '.pdf')
		
		wb.close
	except:
		pass
		
	