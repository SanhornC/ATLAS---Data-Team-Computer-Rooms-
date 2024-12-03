'''

In this file, we have split the users that logon the computers during spring/summer1/summer2 by their corresponsding rooms. 
You will be able to see whether a specific student uses the computer in class or during non-class period. 
If the user is using the computer in class period, the classname will also appear in the row. 
The purpose of this file is to make us easier in analyzing the raw data

By: Sanhorn Chen (sanhorn2) 11/01.

'''

import pandas as pd
from xlsxwriter import Workbook
from datetime import datetime, time


def clean_schedule_df(schedule):
    schedule_df = pd.DataFrame(schedule)
    schedule_df['Start'] = pd.to_datetime(schedule_df['Start'], format='%H:%M:%S').dt.time
    schedule_df['End'] = pd.to_datetime(schedule_df['End'], format='%H:%M:%S').dt.time
    return schedule_df

def assign_class(df, schedule_df, season):
    df["Class_Name"] = "NULL"
    for _, class_info in schedule_df.iterrows():
        mask = (
            (df['Logon_Day_Name'] == class_info['Day']) &
            (df['Logon_Time'] >= class_info['Start']) &
            (df['Logon_Time'] <= class_info['End']) & 
            (df['Season'] == season)
        )
        df.loc[mask, 'Class_Name'] = class_info['Class_Name']

    return df

def final_assign_class(schedule, df, season):
    schedule_df = clean_schedule_df(schedule)
    df = assign_class(df, schedule_df, season)
    return df

def assign_season(row):
    month = row['Logon_Day'].month
    day = row['Logon_Day'].day
    
    # Spring: January 1 - May 31
    if month <= 5:
        return 'Spring'
    # Summer1: June 1 - July 15
    elif (month == 6) or (month == 7 and day <= 15):
        return 'Summer1'
    # Summer2: July 16 - August 21
    elif (month == 7 and day >= 16) or (month == 8 and day <= 21):
        return 'Summer2'
    else:
        return 'Unknown'
def get_roomUsage_bySeasonDayTime(df, season, day, startTime, endTime):
    df = df[((df["Logon_Day_Name"] == day) & ((df["Logon_Time"] >= startTime) & (df["Logon_Time"] < endTime)) & (df["Season"] == season))]
    df = df.drop_duplicates(subset='Computer_Number')
    return df

### Clean Data Format
#---------------------------------------------
# clean data to week name and timeframe
df = pd.read_excel("/Users/sanhorn/Desktop/Internships/ATLAS - Data Team/Computer-Rooms/Dataset/ClassroomsUserAudit.xlsx")
df["LogonTime"] = df["LogonTime"].astype(str)
df[["Logon_Day", "Logon_Time"]] = df["LogonTime"].str.split(" ", expand=True)

df['Logon_Day'] = pd.to_datetime(df['Logon_Day'], errors='coerce')
df["Logon_Day_Name"] = df['Logon_Day'].dt.day_name()
df['Logon_Day'] = df['Logon_Day'].dt.date

# print(df['Logon_Day'].isna().sum(), "rows failed to convert")
df['Logon_Time'] = pd.to_datetime(df['Logon_Time'], format='%H:%M:%S', errors='coerce').dt.time
df['Season'] = df.apply(assign_season, axis=1)
df = df.drop(["LogonTime"], axis=1)

#---------------------------------------------
# clean room name + room computer
df[['Room', 'Room_Section', 'Computer_Number']] = df['MachineName'].str.split("-", expand=True)
df['Room'] = df['Room'] + '-' + df['Room_Section']
df = df.drop(['Room_Section', 'MachineName'], axis=1)

### Separate to each room
#---------------------------------------------
room_name =["DAV-338", "LCLB-G17", "LCLB-G27", "LCLB-G8B", "LCLB-G52", "LCLB-G13", "LCLB-G23", "LCLB-G3", "LCLB-G7", "LCLB-G8A"]

""" Output Logon Data by Computer Rooms """

output_path = "/Users/sanhorn/Desktop/Internships/ATLAS - Data Team/Computer-Rooms/Dataset/ClassroomsLogonData.xlsx"
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
    for name in room_name:
        # Filter data for the current room
        room_data = df[df["Room"] == name]
        # Write to a separate sheet
        room_data.to_excel(writer, sheet_name=name, index=False)

# print(f"Excel file created at: {output_path}")

# Each Room (Total 10)
dav_338 = df[df["Room"] == room_name[0]].dropna()
lclb_g17 = df[df["Room"] == room_name[1]].dropna()
lclb_g27 = df[df["Room"] == room_name[2]].dropna()
lclb_g8b = df[df["Room"] == room_name[3]].dropna()
lclb_g52 = df[df["Room"] == room_name[4]].dropna()
lclb_g13 = df[df["Room"] == room_name[5]].dropna()
lclb_g23 = df[df["Room"] == room_name[6]].dropna()
lclb_g3 = df[df["Room"] == room_name[7]].dropna()
lclb_g7 = df[df["Room"] == room_name[8]].dropna()
lclb_g8a = df[df["Room"] == room_name[9]].dropna()
# print(dav_338)

############### Parse data to Spring Only ####################
dav_338 = dav_338[dav_338["Season"] == "Spring"]
lclb_g17 = lclb_g17[lclb_g17["Season"] == "Spring"]
lclb_g27 = lclb_g27[lclb_g27["Season"] == "Spring"]
lclb_g8b = lclb_g8b[lclb_g8b["Season"] == "Spring"]
lclb_g52 = lclb_g52[lclb_g52["Season"] == "Spring"]
lclb_g13 = lclb_g13[lclb_g13["Season"] == "Spring"]
lclb_g23 = lclb_g23[lclb_g23["Season"] == "Spring"]
lclb_g3 = lclb_g3[lclb_g3["Season"] == "Spring"]
lclb_g7 = lclb_g7[lclb_g7["Season"] == "Spring"]
lclb_g8a = lclb_g8a[lclb_g8a["Season"] == "Spring"]

######################################## SPRING ##################################
season = "Spring"
""" Dav 338 """
# print(dav_338.isna().sum())
dav_schedule = {
    'Class_Name' : ["LCTL 302 - WLF", "ANTH 499 - LD", "ANTH 499 - LD", "LCTL 302 - WLF", "PS 494 - JH", "GGIS 224 - AB1", "GGIS 224 - AB3"],
    "Day": ["Tuesday", "Tuesday", "Wednesday", "Thursday", "Thursday", "Thursday", "Friday"],
    "Start": ["10:00:00", "15:00:00", "15:00:00", "10:00:00", "12:30:00", "14:00:00", "13:00:00"],
    "End": ["11:30:00", "17:00:00", "17:00:00", "11:30:00", "14:00:00", "16:00:00", "14:00:00"]
}
dav_338 = final_assign_class(dav_schedule, dav_338, season)
# for i in range(len(dav_338)):
#     if dav_338[i]["Logon_Day_Name"] == dav_schedule[]
""" LCLB G17 """
lclb_g17_schedule = {
    'Class_Name': [
        "IEI ARW", "IEI ARW", "IEI ARW",
        "LING 514 - H", "LING 514 - H",
        "CONFLICT - LING 448 - G4 & LING 448 - U3", "CONFLICT - LING 448 - G4 & LING 448 - U3",
        "CONFLICT - LING 529 - SM & PSYC 529 - SM", "CONFLICT - LING 529 - SM & PSYC 529 - SM",
        "BTW 250 - E1", "BTW 250 - E1", "BTW 250 - E1",
        "SWAH 404 - CI"
    ],
    "Day": [
        "Monday", "Wednesday", "Friday",
        "Tuesday", "Thursday",
        "Tuesday", "Thursday",
        "Tuesday", "Thursday",
        "Monday", "Wednesday", "Friday",
        "Wednesday"
    ],
    "Start": [
        "09:00:00", "09:00:00", "09:00:00",
        "09:30:00", "09:30:00",
        "11:00:00", "11:00:00",
        "12:30:00", "12:30:00",
        "13:00:00", "13:00:00", "13:00:00",
        "14:00:00"
    ],
    "End": [
        "13:00:00", "13:00:00", "13:00:00",
        "11:30:00", "11:30:00",
        "12:30:00", "12:30:00",
        "14:00:00", "14:00:00",
        "14:00:00", "14:00:00", "14:00:00",
        "16:00:00"
    ]
}
lclb_g17 = final_assign_class(lclb_g17_schedule, lclb_g17, season)
""" LCLB G27 """
lclb_g27_schedule = {
    'Class_Name': [
        "IEI ARW", "IEI ARW", "IEI ARW",
        "EIL 445 - G4", "EIL 445 - G4",
        "IEI pronunciation",
        "EALC 320 - NG", "EALC 320 - NG",
        "INFO 102 - AB1", "INFO 102 - AB2", "INFO 102 - AB3",
        "CONFLICT - Room Closed & INFO 102 - AB4", "CONFLICT - Room Closed & INFO 102 - AB5",
        "UP 519 - 1",
        "CHEM 483 - AL1"
    ],
    "Day": [
        "Monday", "Wednesday", "Friday",
        "Tuesday", "Thursday",
        "Tuesday",
        "Monday", "Wednesday",
        "Tuesday", "Wednesday", "Thursday",
        "Tuesday", "Wednesday",
        "Thursday",
        "Friday"
    ],
    "Start": [
        "10:00:00", "10:00:00", "10:00:00",
        "09:30:00", "09:30:00",
        "13:00:00",
        "14:00:00", "14:00:00",
        "16:00:00", "16:00:00", "16:00:00",
        "17:00:00", "17:00:00",
        "14:00:00",
        "14:00:00"
    ],
    "End": [
        "12:00:00", "12:00:00", "12:00:00",
        "11:30:00", "11:30:00",
        "15:00:00",
        "17:00:00", "17:00:00",
        "17:30:00", "17:30:00", "17:30:00",
        "19:30:00", "19:30:00",
        "17:00:00",
        "17:00:00"
    ]
}
lclb_g27 = final_assign_class(lclb_g27_schedule, lclb_g27, season)
# print(lclb_g27)
""" LCLB G8B """
lclb_g8b_schedule = {
    'Class_Name': [
        "ESL 112 - K", "ESL 112 - K", "ESL 112 - K",
        "ESL 112 - H", "ESL 112 - H", "ESL 112 - H",
        "ESL 112 - O", "ESL 112 - O", "ESL 112 - O",
        "ESL 112 - C", "ESL 112 - C", "ESL 112 - C",
        "ESL 112 - G", "ESL 112 - G", "ESL 112 - G",
        "ESL 112 - I", "ESL 112 - I", "ESL 112 - I",
        "ESL 112 - Q", "ESL 112 - Q", "ESL 112 - Q",
        "ESL 512 - G", "ESL 512 - G",
        "ESL 512 - C", "ESL 512 - C",
        "ESL 512 - H", "ESL 512 - H",
        "ESL 512 - E", "ESL 512 - E",
        "ESL 512 - I", "ESL 512 - I"
    ],
    "Day": [
        "Monday", "Wednesday", "Friday",
        "Monday", "Wednesday", "Friday",
        "Monday", "Wednesday", "Friday",
        "Monday", "Wednesday", "Friday",
        "Monday", "Wednesday", "Friday",
        "Monday", "Wednesday", "Friday",
        "Monday", "Wednesday", "Friday",
        "Tuesday", "Thursday",
        "Tuesday", "Thursday",
        "Tuesday", "Thursday",
        "Tuesday", "Thursday",
        "Tuesday", "Thursday"
    ],
    "Start": [
        "09:00:00", "09:00:00", "09:00:00",
        "10:00:00", "10:00:00", "10:00:00",
        "11:00:00", "11:00:00", "11:00:00",
        "13:00:00", "13:00:00", "13:00:00",
        "14:00:00", "14:00:00", "14:00:00",
        "15:00:00", "15:00:00", "15:00:00",
        "16:00:00", "16:00:00", "16:00:00",
        "10:00:00", "10:00:00",
        "11:00:00", "11:00:00",
        "12:30:00", "12:30:00",
        "14:00:00", "14:00:00",
        "15:30:00", "15:30:00"
    ],
    "End": [
        "10:00:00", "10:00:00", "10:00:00",
        "11:00:00", "11:00:00", "11:00:00",
        "12:00:00", "12:00:00", "12:00:00",
        "14:00:00", "14:00:00", "14:00:00",
        "15:00:00", "15:00:00", "15:00:00",
        "16:00:00", "16:00:00", "16:00:00",
        "17:00:00", "17:00:00", "17:00:00",
        "12:00:00", "12:00:00",
        "12:30:00", "12:30:00",
        "14:00:00", "14:00:00",
        "15:30:00", "15:30:00",
        "17:00:00", "17:00:00"
    ]
}
lclb_g8b = final_assign_class(lclb_g8b_schedule, lclb_g8b, season)
""" LCLB G52 """
lclb_g52_schedule = {
    'Class_Name': [
        "ESL 115 - G", "ESL 115 - G", "ESL 115 - G",
        "ESL 112 - P", "ESL 112 - P", "ESL 112 - P",
        "ESL 112 - F", "ESL 112 - F", "ESL 112 - F",
        "ESL 112 - N", "ESL 112 - N", "ESL 112 - N",
        "ESL 112 - L", "ESL 112 - L", "ESL 112 - L",
        "ESL 115 - I", "ESL 115 - I", "ESL 115 - I",
        "ESL 112 - J", "ESL 112 - J", "ESL 112 - J",
        "CONFLICT - Room Closed & ESL 525 - C", "CONFLICT - Room Closed & ESL 525 - C",
        "IEI Vocabulary"
    ],
    "Day": [
        "Monday", "Wednesday", "Friday",
        "Monday", "Wednesday", "Friday",
        "Monday", "Wednesday", "Friday",
        "Monday", "Wednesday", "Friday",
        "Monday", "Wednesday", "Friday",
        "Monday", "Wednesday", "Friday",
        "Monday", "Wednesday", "Friday",
        "Monday", "Wednesday",
        "Thursday"
    ],
    "Start": [
        "09:00:00", "09:00:00", "09:00:00",
        "10:00:00", "10:00:00", "10:00:00",
        "11:00:00", "11:00:00", "11:00:00",
        "13:00:00", "13:00:00", "13:00:00",
        "14:00:00", "14:00:00", "14:00:00",
        "15:00:00", "15:00:00", "15:00:00",
        "16:00:00", "16:00:00", "16:00:00",
        "17:00:00", "17:00:00",
        "13:30:00"
    ],
    "End": [
        "10:00:00", "10:00:00", "10:00:00",
        "11:00:00", "11:00:00", "11:00:00",
        "12:00:00", "12:00:00", "12:00:00",
        "14:00:00", "14:00:00", "14:00:00",
        "15:00:00", "15:00:00", "15:00:00",
        "16:00:00", "16:00:00", "16:00:00",
        "17:00:00", "17:00:00", "17:00:00",
        "18:30:00", "18:30:00",
        "15:30:00"
    ]
}
lclb_g52 = final_assign_class(lclb_g52_schedule, lclb_g52, season)
""" LCLB G13 """
lclb_g13_schedule = {
    'Class_Name': [
        "ESL 112 - D", "ESL 112 - D", "ESL 112 - D",
        "ESL 112 - E", "ESL 112 - E", "ESL 112 - E",
        "ESL 115 - J", "ESL 115 - J", "ESL 115 - J",
        "ESL 111 - Y", "ESL 111 - Y", "ESL 111 - Y",
        "ESL 112 - A", "ESL 112 - A", "ESL 112 - A",
        "ESL 112 - S", "ESL 112 - S", "ESL 112 - S",
        "ESL 115 - K", "ESL 115 - K", "ESL 115 - K",
        "ESL 515 - E", "ESL 515 - E",
        "PORT 401 - X", "PORT 401 - X",
        "ESL 515 - C", "ESL 515 - C",
        "CONFLICT - Room Closed & ESL 522 - E", "CONFLICT - Room Closed & ESL 522 - E",
        "CONFLICT - Room Closed & ESL 512 - D", "CONFLICT - Room Closed & ESL 512 - D"
    ],
    "Day": [
        "Monday", "Wednesday", "Friday",
        "Monday", "Wednesday", "Friday",
        "Monday", "Wednesday", "Friday",
        "Monday", "Wednesday", "Friday",
        "Monday", "Wednesday", "Friday",
        "Monday", "Wednesday", "Friday",
        "Monday", "Wednesday", "Friday",
        "Tuesday", "Thursday",
        "Tuesday", "Thursday",
        "Tuesday", "Thursday",
        "Monday", "Wednesday",
        "Tuesday", "Thursday"
    ],
    "Start": [
        "09:00:00", "09:00:00", "09:00:00",
        "10:00:00", "10:00:00", "10:00:00",
        "11:00:00", "11:00:00", "11:00:00",
        "13:00:00", "13:00:00", "13:00:00",
        "14:00:00", "14:00:00", "14:00:00",
        "15:00:00", "15:00:00", "15:00:00",
        "16:00:00", "16:00:00", "16:00:00",
        "09:30:00", "09:30:00",
        "11:00:00", "11:00:00",
        "14:00:00", "14:00:00",
        "17:00:00", "17:00:00",
        "17:00:00", "17:00:00"
    ],
    "End": [
        "10:00:00", "10:00:00", "10:00:00",
        "11:00:00", "11:00:00", "11:00:00",
        "12:00:00", "12:00:00", "12:00:00",
        "14:00:00", "14:00:00", "14:00:00",
        "15:00:00", "15:00:00", "15:00:00",
        "16:00:00", "16:00:00", "16:00:00",
        "17:00:00", "17:00:00", "17:00:00",
        "11:00:00", "11:00:00",
        "12:30:00", "12:30:00",
        "15:30:00", "15:30:00",
        "18:30:00", "18:30:00",
        "18:30:00", "18:30:00"
    ]
}
lclb_g13 = final_assign_class(lclb_g13_schedule, lclb_g13, season)
""" LCLB G23 """
lclb_g23_schedule = {
    'Class_Name': [
        "ESL 115 - L1", "ESL 115 - L1", "ESL 115 - L1",
        "ESL 115 - A1", "ESL 115 - A1", "ESL 115 - A1",
        "ESL 115 - C1", "ESL 115 - C1", "ESL 115 - C1",
        "ESL 115 - U", "ESL 115 - U", "ESL 115 - U",
        "ESL 115 - D1", "ESL 115 - D1", "ESL 115 - D1",
        "ESL 112 - M", "ESL 112 - M", "ESL 112 - M",
        "IEI Vocabulary", "IEI Vocabulary",
        "IEI Adv Vocabulary", "IEI Adv Vocabulary",
        "ATLAS TLI Weekly Meeting",
        "ESL 522 - B", "ESL 522 - B",
        "CONFLICT - Room Closed & ESL 515 - D", "CONFLICT - Room Closed & ESL 515 - D"
    ],
    "Day": [
        "Monday", "Wednesday", "Friday",
        "Monday", "Wednesday", "Friday",
        "Monday", "Wednesday", "Friday",
        "Monday", "Wednesday", "Friday",
        "Monday", "Wednesday", "Friday",
        "Monday", "Wednesday", "Friday",
        "Tuesday", "Thursday",
        "Tuesday", "Thursday",
        "Wednesday",
        "Tuesday", "Thursday",
        "Tuesday", "Thursday"
    ],
    "Start": [
        "09:00:00", "09:00:00", "09:00:00",
        "10:00:00", "10:00:00", "10:00:00",
        "13:00:00", "13:00:00", "13:00:00",
        "14:00:00", "14:00:00", "14:00:00",
        "15:00:00", "15:00:00", "15:00:00",
        "16:00:00", "16:00:00", "16:00:00",
        "09:00:00", "09:00:00",
        "11:00:00", "11:00:00",
        "11:00:00",
        "14:00:00", "14:00:00",
        "17:00:00", "17:00:00"
    ],
    "End": [
        "10:00:00", "10:00:00", "10:00:00",
        "11:00:00", "11:00:00", "11:00:00",
        "14:00:00", "14:00:00", "14:00:00",
        "15:00:00", "15:00:00", "15:00:00",
        "16:00:00", "16:00:00", "16:00:00",
        "17:00:00", "17:00:00", "17:00:00",
        "11:00:00", "11:00:00",
        "12:30:00", "12:30:00",
        "12:00:00",
        "15:30:00", "15:30:00",
        "18:30:00", "18:30:00"
    ]
}
lclb_g23 = final_assign_class(lclb_g23_schedule, lclb_g23, season)
""" LCLB G3 """
lclb_g3_schedule = {
    'Class_Name': [
        "ACE 499 - MW", "ACE 499 - MW",
        "ARAB 202 - BE1", "ARAB 202 - BE1", "ARAB 202 - BE1", "ARAB 202 - BE1", "ARAB 202 - BE1",
        "ARAB 404 - B", "ARAB 404 - B", "ARAB 404 - B", "ARAB 404 - B",
        "ARAB 202 - CE1", "ARAB 202 - CE1", "ARAB 202 - CE1", "ARAB 202 - CE1", "ARAB 202 - CE1",
        "EALC 320 - NG", "EALC 320 - NG",
        "HDFS 594 - A",
        "CONFLICT - Room Closed & ESL 522 - Q", "CONFLICT - Room Closed & ESL 522 - Q"
    ],
    "Day": [
        "Monday", "Wednesday",
        "Monday", "Tuesday", "Wednesday", "Thursday", "Friday",
        "Monday", "Tuesday", "Wednesday", "Thursday",
        "Monday", "Tuesday", "Wednesday", "Thursday", "Friday",
        "Monday", "Wednesday",
        "Tuesday",
        "Tuesday", "Thursday"
    ],
    "Start": [
        "09:30:00", "09:30:00",
        "11:00:00", "11:00:00", "11:00:00", "11:00:00", "11:00:00",
        "12:00:00", "12:00:00", "12:00:00", "12:00:00",
        "13:00:00", "13:00:00", "13:00:00", "13:00:00", "13:00:00",
        "14:00:00", "14:00:00",
        "14:00:00",
        "17:00:00", "17:00:00"
    ],
    "End": [
        "11:00:00", "11:00:00",
        "12:00:00", "12:00:00", "12:00:00", "12:00:00", "12:00:00",
        "13:00:00", "13:00:00", "13:00:00", "13:00:00",
        "14:00:00", "14:00:00", "14:00:00", "14:00:00", "14:00:00",
        "15:30:00", "15:30:00",
        "17:00:00",
        "18:30:00", "18:30:00"
    ]
}
lclb_g3 = final_assign_class(lclb_g3_schedule, lclb_g3, season)
""" LCLB G7 """
lclb_g7_schedule = {
    'Class_Name': [
        "TA meeting and Office Hours",
        "ATMS 100 - ADB", "ATMS 100 - ADK",
        "ATMS 100 - ADC", "ATMS 100 - ADL",
        "ATMS 100 - ADM",
        "ATMS 100 - ADE", "ATMS 100 - ADN",
        "ATMS 100 - ADF",
        "IB 451 - AB2", "IB 451 - AB3",
        "SPAN 248 - A11",
        "BTW 250 - T1", "BTW 250 - T1"
    ],
    "Day": [
        "Monday",
        "Tuesday", "Thursday",
        "Tuesday", "Thursday",
        "Thursday",
        "Thursday", "Thursday",
        "Friday",
        "Tuesday", "Wednesday",
        "Wednesday",
        "Wednesday", "Friday"
    ],
    "Start": [
        "10:00:00",
        "10:00:00", "10:00:00",
        "11:00:00", "11:00:00",
        "12:00:00",
        "13:00:00", "13:00:00",
        "14:00:00",
        "12:00:00", "12:00:00",
        "15:00:00",
        "15:30:00", "15:30:00"
    ],
    "End": [
        "12:00:00",
        "11:00:00", "11:00:00",
        "12:00:00", "12:00:00",
        "13:00:00",
        "14:00:00", "14:00:00",
        "15:00:00",
        "14:00:00", "14:00:00",
        "16:00:00",
        "17:00:00", "17:00:00"
    ]
}
lclb_g7 = final_assign_class(lclb_g7_schedule, lclb_g7, season)
""" LCLB G8A """
lclb_g8a_schedule = {
    'Class_Name': [
        "ESL 111 - J", "ESL 111 - J", "ESL 111 - J",
        "ESL 115 - Q", "ESL 115 - Q", "ESL 115 - Q",
        "ESL 111 - A", "ESL 111 - A", "ESL 111 - A",
        "GSD 390 - ANJ",
        "GSD 390 - JVA",
        "GSD 390 - DBU",
        "GSD 504 - AL",
        "CONFLICT - Room Closed & PORT 404 - GR4 & PORT 404 - UG3", "CONFLICT - Room Closed & PORT 404 - GR4 & PORT 404 - UG3"
    ],
    "Day": [
        "Monday", "Wednesday", "Friday",
        "Monday", "Wednesday", "Friday",
        "Monday", "Wednesday", "Friday",
        "Tuesday",
        "Tuesday",
        "Thursday",
        "Tuesday",
        "Tuesday", "Thursday"
    ],
    "Start": [
        "10:00:00", "10:00:00", "10:00:00",
        "13:00:00", "13:00:00", "13:00:00",
        "15:00:00", "15:00:00", "15:00:00",
        "10:00:00",
        "13:00:00",
        "14:00:00",
        "15:30:00",
        "17:30:00", "17:30:00"
    ],
    "End": [
        "11:00:00", "11:00:00", "11:00:00",
        "14:00:00", "14:00:00", "14:00:00",
        "16:00:00", "16:00:00", "16:00:00",
        "12:30:00",
        "15:00:00",
        "16:30:00",
        "17:00:00",
        "19:00:00", "19:00:00"
    ]
}
lclb_g8a = final_assign_class(lclb_g8a_schedule, lclb_g8a, season)





######################################## Summer 2 ##################################
s2_lclb_g13 = df[df["Room"] == room_name[5]].dropna()
s2_lclb_g23 = df[df["Room"] == room_name[6]].dropna()
s2_lclb_g27 = df[df["Room"] == room_name[2]].dropna()
s2_lclb_g17 = df[df["Room"] == room_name[1]].dropna()

############ parse to only have Summer 2 ############
s2_lclb_g13 = s2_lclb_g13[s2_lclb_g13["Season"] == "Summer2"]
s2_lclb_g23 = s2_lclb_g23[s2_lclb_g23["Season"] == "Summer2"]
s2_lclb_g27 = s2_lclb_g27[s2_lclb_g27["Season"] == "Summer2"]
s2_lclb_g17 = s2_lclb_g17[s2_lclb_g17["Season"] == "Summer2"]

season = "Summer2"
""" LCLB G27 """
s2_lclb_g27_schedule = {
    'Class_Name': [
        "Bolashak Program", "Bolashak Program", "Bolashak Program"
    ],
    "Day": [
        "Monday", "Wednesday", "Friday"
    ],
    "Start": [
        "13:00:00", "13:00:00", "13:00:00"
    ],
    "End": [
        "15:00:00", "15:00:00", "15:00:00"
    ]
}
s2_lclb_g27 = final_assign_class(s2_lclb_g27_schedule, s2_lclb_g27, season)
""" LCLB G13 """
s2_lclb_g13_schedule = {
    'Class_Name': [
        "Bolashak Program", "Bolashak Program", "Bolashak Program",
        "Bolashak Program", "Bolashak Program", "Bolashak Program"
    ],
    "Day": [
        "Monday", "Wednesday", "Friday",
        "Monday", "Wednesday", "Friday"
    ],
    "Start": [
        "10:00:00", "10:00:00", "10:00:00",
        "13:00:00", "13:00:00", "13:00:00"
    ],
    "End": [
        "12:00:00", "12:00:00", "12:00:00",
        "15:00:00", "15:00:00", "15:00:00"
    ]
}
s2_lclb_g13 = final_assign_class(s2_lclb_g13_schedule, s2_lclb_g13, season)
""" LCLB G23 """
s2_lclb_g23_schedule = {
    'Class_Name': [
        "Bolashak Program", "Bolashak Program", "Bolashak Program"
    ],
    "Day": [
        "Monday", "Wednesday", "Friday"
    ],
    "Start": [
        "13:00:00", "13:00:00", "13:00:00"
    ],
    "End": [
        "15:00:00", "15:00:00", "15:00:00"
    ]
}
s2_lclb_g23 = final_assign_class(s2_lclb_g23_schedule, s2_lclb_g23, season)
""" LCLB G17 """
s2_lclb_g17_schedule = {
    'Class_Name': [
        "Bolashak Program", "Bolashak Program", "Bolashak Program"
    ],
    "Day": [
        "Monday", "Wednesday", "Friday"
    ],
    "Start": [
        "10:00:00", "10:00:00", "10:00:00"
    ],
    "End": [
        "12:00:00", "12:00:00", "12:00:00"
    ]
}
s2_lclb_g17 = final_assign_class(s2_lclb_g17_schedule, s2_lclb_g17, season)


dataframes = {
    'DAV 338': dav_338,
    'LCLB G17': lclb_g17,
    'LCLB G27': lclb_g27,
    'LCLB G8B': lclb_g8b,
    'LCLB G52': lclb_g52,
    'LCLB G13': lclb_g13,
    'LCLB G23': lclb_g23,
    'LCLB G3': lclb_g3,
    'LCLB G7': lclb_g7,
    'LCLB G8A': lclb_g8a
}


s2_dataframes = {
    'LCLB G27': s2_lclb_g27,
    'LCLB G13': s2_lclb_g13,
    'LCLB G23': s2_lclb_g23,
    'LCLB G17': s2_lclb_g17,
}

# Create an Excel writer object
with pd.ExcelWriter('/Users/sanhorn/Desktop/Internships/ATLAS - Data Team/Computer-Rooms/Dataset/spring_classlogonrecord.xlsx') as writer:
    for sheet_name, df in dataframes.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)

with pd.ExcelWriter('/Users/sanhorn/Desktop/Internships/ATLAS - Data Team/Computer-Rooms/Dataset/summer2_classlogonrecord.xlsx') as writer:
    for sheet_name, df in s2_dataframes.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)