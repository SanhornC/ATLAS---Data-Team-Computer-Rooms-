"""

This file takes the spring_classlogoonrecord.xlsx file as input. In this file, we have split the users that logon the computers
during the spring semester by their corresponsding rooms. You will be able to see whether a specific student uses the computer in class 
or during non-class period. 

This python file takes the input file and counts the logon numbers of each room in every hour on weekly bases. 
The outputs of this python file will be directed to the 11_13 folder. 
The count of logon numbers are stored weekly in each folder. For example, in the Dav338 file, it contains weekly data in each sheet from January to May.

In addition, there is another column that calculates the utilization percentage. The equation is weekly logon records / room capacity

By Sanhorn Chen (sanhorn2) 2024/11/11.

"""

import os
import pandas as pd
from datetime import datetime, timedelta

file_path = '/Users/sanhorn/Desktop/Internships/ATLAS - Data Team/Computer-Rooms/Dataset/spring_classlogonrecord.xlsx'
output_folder = '/Users/sanhorn/Desktop/Internships/ATLAS - Data Team/Computer-Rooms/Dataset/11_13'

os.makedirs(output_folder, exist_ok=True)

# default number of computers in each room
room_capacities = {
    'DAV 338': 30,
    'LCLB G17': 31,
    'LCLB G27': 40,
    'LCLB G8B': 17,
    'LCLB G52': 10,
    'LCLB G13': 16,
    'LCLB G23': 32,
    'LCLB G3': 23,
    'LCLB G7': 31,
    'LCLB G8A': 18
}

time_blocks = [f"{hour:02d}:00{'AM' if hour < 12 else 'PM'}" for hour in range(7, 20)]
weekdays = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]

start_date = datetime(2024, 1, 15)
end_date = datetime(2024, 5, 10)

all_weeks = []
current_week_start = start_date
while current_week_start <= end_date:
    current_week_end = current_week_start + timedelta(days=4)
    if current_week_end > end_date:
        break
    all_weeks.append((current_week_start, current_week_end))
    current_week_start += timedelta(days=7)

excel_data = pd.ExcelFile(file_path)

for sheet_name in excel_data.sheet_names:
    room_data = excel_data.parse(sheet_name)
    
    if not pd.api.types.is_datetime64_any_dtype(room_data['Logon_Day']):
        room_data['Logon_Day'] = pd.to_datetime(room_data['Logon_Day'], errors='coerce')
    
    room_capacity = room_capacities[sheet_name]
    output_file = os.path.join(output_folder, f"{sheet_name}.xlsx")
    
    with pd.ExcelWriter(output_file) as writer:
        for start, end in all_weeks:
            week_data = room_data[(room_data['Logon_Day'] >= start) & (room_data['Logon_Day'] <= end)]
            
            weekly_logon_count = pd.DataFrame(0, index=time_blocks, columns=weekdays)

            for _, row in week_data.iterrows():
                logon_day_name = row['Logon_Day_Name']
                logon_time = row['Logon_Time']
                
                if logon_day_name in weekdays:
                    try:
                        # breakpoint()
                        logon_hour = datetime.strptime(logon_time, '%H:%M:%S').hour
                        time_block = f"{logon_hour:02d}:00{'AM' if logon_hour < 12 else 'PM'}"
                        
                        if time_block in weekly_logon_count.index:
                            weekly_logon_count.at[time_block, logon_day_name] += 1
                    except ValueError:
                        continue
            
            # calculates the utilization percentage
            weekly_logon_count['Utilization'] = weekly_logon_count[weekdays].max(axis=1) / room_capacity
            
            week_label = f"{start.strftime('%Y%m%d')}_to_{end.strftime('%Y%m%d')}"
            
            weekly_logon_count.to_excel(writer, sheet_name=week_label)

print("Processing complete. Each room's logon data is saved with weekly sheets in individual Excel files in:", output_folder)
