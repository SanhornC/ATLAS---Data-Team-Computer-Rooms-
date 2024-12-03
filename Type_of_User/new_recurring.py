'''

This file gives us the ratio of self users (user that uses computer during non class period), recurring users (user that uses during classtime and after classtime), and
class users (user that uses the computer during class time only).
The purpose of this file is to get the overview of the type of users using the ATLAS computer rooms. 

By: Sanhorn Chen (sanhorn2) 11/03.

'''

import pandas as pd

file_path = "/Users/sanhorn/Desktop/Internships/ATLAS - Data Team/Computer-Rooms/Dataset/ClassUsedData.xlsx"
# sheets = ["DAV 338", "LCLB G17", "LCLB G27", "LCLB G8B", "LCLB G52", "LCLB G13", "LCLB G23", "LCLB G3", "LCLB G7", "LCLB G8A"]
full_sheet = pd.read_excel(file_path, sheet_name=None)
combined_df = pd.concat(full_sheet.values(), ignore_index=True)


self_cnt = 0
self_df = pd.DataFrame()
class_cnt = 0
class_df = pd.DataFrame()
recur_cnt = 0
recur_df = pd.DataFrame()

uniq_users = combined_df["UserName"].unique()
total = len(uniq_users)

for user in uniq_users:
    user_rows = combined_df[combined_df['UserName'] == user].sort_values(by='Logon_Day')

    has_null = user_rows["Class_Name"].isnull().any()
    has_class = user_rows["Class_Name"].notnull().any()

    if has_null and has_class:
        first_row = user_rows.iloc[0]["Class_Name"] is not None
        if first_row: 
            recur_cnt += 1
            recur_df = pd.concat([recur_df, user_rows]) 
        else: 
            self_cnt += 1
            self_df = pd.concat([self_df, user_rows])
    
    elif has_null and not has_class:
        self_cnt += 1
        self_df = pd.concat([self_df, user_rows])

    elif not has_null and has_class:
        class_cnt += 1
        class_df = pd.concat([class_df, user_rows])

# output_file_path = "/Users/sanhorn/Desktop/Internships/ATLAS - Data Team/Computer-Rooms/Dataset/UserType.xlsx"
# with pd.ExcelWriter(output_file_path) as writer:
#     recur_df.to_excel(writer, sheet_name='Recurring Users', index=False)
#     self_df.to_excel(writer, sheet_name='Self Users', index=False)
#     class_df.to_excel(writer, sheet_name='Class Users', index=False)

print("------------ Count ---------------------")
print(f"Recurring user (Users that knew Computer Rooms from Classes): {recur_cnt}")
print(f"Self users (Users that knew Computer Rooms before Classes or Always came to computer rooms alone): {self_cnt}")
print(f"Class users (Users who only used Computer Room during Classes, Never Used after Classes): {class_cnt}")
print(f"Total users (total number of users who have logon records): {total}")

pct_recur = recur_cnt / total
pct_self = self_cnt / total
pct_class = class_cnt / total
print("------------ Percentage ---------------------")
print(f"Recurring user (Users that knew Computer Rooms from Classes): {pct_recur}")
print(f"Self users (Users that knew Computer Rooms before Classes or Always came to computer rooms alone): {pct_self}")
print(f"Class users (Users who only used Computer Room during Classes, Never Used after Classes): {pct_class}")


