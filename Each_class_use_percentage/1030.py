'''

This python file extracts classnames as the variables. The output of this python file will calculate the weekly logon counts for each class and take the largest count 
as the final number for each class. The final number of each room will be used to count the average usage of computers for the specfic class. 
This file is meant to analyze whether classes are using the classrooms efficiently. 

'''

import pandas as pd

file_path = '/Users/sanhorn/Desktop/Internships/ATLAS - Data Team/Computer-Rooms/Dataset/ClassUsedData.xlsx'
excel_data = pd.ExcelFile(file_path)
weekly_class_counts_path = '/Users/sanhorn/Desktop/Internships/ATLAS - Data Team/Computer-Rooms/Dataset/Refined_Weekly_Class_Counts_for_Valid_Courses.csv'

weekly_class_counts_df = pd.read_csv(weekly_class_counts_path)
weekly_class_counts_dict = dict(zip(weekly_class_counts_df['Course'], weekly_class_counts_df['Weekly_Class_Count']))

class_aggregated_data_final = []

room_capacities = {
    "DAV 338": 30,
    "LCLB G17": 31,
    "LCLB G27": 40,
    "LCLB G8B": 17,
    "LCLB G52": 10,
    "LCLB G13": 16,
    "LCLB G23": 32,
    "LCLB G3": 23,
    "LCLB G7": 31,
    "LCLB G8A": 18
}

for sheet in excel_data.sheet_names:
    class_data = excel_data.parse(sheet)
    spring_data = class_data[class_data['Season'] == 'Spring']
    
    class_groups = spring_data.groupby('Class_Name')
    
    for class_name, group in class_groups:
        total_logons = group.shape[0]
        
        unique_class_days = group['Logon_Day'].nunique()
        
        weekly_class_count = weekly_class_counts_dict.get(class_name, None)  # Default to None if class is not found
        
        if weekly_class_count == 1:
            total_sessions = 15
        elif weekly_class_count == 2:
            total_sessions = 30
        elif weekly_class_count == 3:
            total_sessions = 45
        else:
            total_sessions = None 
        
        class_aggregated_data_final.append({
            'Class_Name': class_name,
            'Total_Logons': total_logons,
            'Unique_Class_Days': unique_class_days,
            'Room_Name': sheet,
            'Computer_Capacity': room_capacities.get(sheet, None),
            'Weekly_Class_Count': weekly_class_count,  
            'Total_Sessions': total_sessions 
        })

class_aggregated_df_final = pd.DataFrame(class_aggregated_data_final)

output_path = '/Users/sanhorn/Desktop/Internships/ATLAS - Data Team/Computer-Rooms/Dataset/1030_Class_Log_On.xlsx'
class_aggregated_df_final.to_excel(output_path, index=False, sheet_name='Aggregated_Data')

print(f"Data has been saved to {output_path}")
