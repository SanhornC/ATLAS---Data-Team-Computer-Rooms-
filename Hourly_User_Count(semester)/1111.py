"""

This file takes the spring_classlogoonrecord.xlsx file as input. In this input file, we have split the users that logon the computers
during the spring semester by the corresponsding computer rooms. You will be able to see whether a specific student uses the computer at a reserved class 
or by himself/herself. 

This python file takes the input file and counts the logon numbers of each room in every hour on semester bases.
In addition, this file can separate users who logon during classtime and users who logon during non-classtime. 
You can get two types of output from this python file, a file with count of logon records during class time in the spring semester and count of logon records during 
non-classtime in the spring semester.

By Sanhorn Chen (sanhorn2) 2024/11/12.

"""

import pandas as pd
from datetime import datetime

file_path = '/Users/sanhorn/Desktop/Internships/ATLAS - Data Team/Computer-Rooms/Dataset/summer2_classlogonrecord.xlsx'
output_path = '/Users/sanhorn/Desktop/Internships/ATLAS - Data Team/Computer-Rooms/Dataset/1111_summer2timerecord.xlsx'

time_blocks = [f"{hour:02d}:00{'AM' if hour < 12 else 'PM'}" for hour in range(0, 24)]
weekdays = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]

excel_data = pd.ExcelFile(file_path)

with pd.ExcelWriter(output_path) as writer:
    for sheet_name in excel_data.sheet_names:
        room_data = excel_data.parse(sheet_name)
        logon_count_df = pd.DataFrame(0, index=time_blocks, columns=weekdays)
        
        for _, row in room_data.iterrows():
            day_name = row['Logon_Day_Name']
            logon_time = row['Logon_Time']
            
            if day_name in weekdays:
                try:
                    logon_hour = datetime.strptime(logon_time, '%H:%M:%S').hour
                    time_block = f"{logon_hour:02d}:00{'AM' if logon_hour < 12 else 'PM'}"
                    
                    if time_block in logon_count_df.index:
                        logon_count_df.at[time_block, day_name] += 1
                except ValueError:
                    continue

        logon_count_df.to_excel(writer, sheet_name=sheet_name)

print("Processing complete. The output file is saved as:", output_path)
