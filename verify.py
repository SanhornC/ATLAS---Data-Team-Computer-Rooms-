import pandas as pd

file_path = "/Users/sanhorn/Desktop/Internships/ATLAS - Data Team/Computer-Rooms/Dataset/ClassUsedData.xlsx"
full_sheet = pd.read_excel(file_path, sheet_name=None)
combined_df = pd.concat(full_sheet.values(), ignore_index=True)
print("ClassUsedData Length: " + str(len(combined_df)))

first_path = "/Users/sanhorn/Desktop/Internships/ATLAS - Data Team/Computer-Rooms/Dataset/ClassroomsUserAudit.xlsx"
excel = pd.ExcelFile(first_path)
first_df = excel.parse("RAW")
print("ClassUserAudit Length: " + str(len(first_df)))