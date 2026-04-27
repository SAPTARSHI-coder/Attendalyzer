import pandas as pd
df = pd.read_excel("B.Tech 4th Semester Attendance Collection (Debarred List) (Responses) (1).xlsx")
col_name = "Student's Name (As per University Records)"
col_event = "Events participation details with date"
skip = ["no","nah",""]
events = df[df[col_event].notna() & (~df[col_event].str.strip().str.lower().isin(skip))]
for _, row in events.iterrows():
    print("NAME :", row[col_name])
    print("EVENT:", repr(row[col_event]))
    print("---")
