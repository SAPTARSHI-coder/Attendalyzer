import pandas as pd
df = pd.read_excel("B.Tech 4th Semester Attendance Collection (Debarred List) (Responses) (1).xlsx")
col_name = "Student's Name (As per University Records)"
col_roll = "Student's University Roll Number"
col_miss = "Any attendance missed but you are present?"
col_event = "Events participation details with date"

print("=== MISSED ATTENDANCE ===")
missed = df[df[col_miss].notna() & (df[col_miss].str.strip().str.lower() != "no")]
for _, row in missed.iterrows():
    print(f"NAME : {row[col_name]}")
    print(f"ROLL : {row[col_roll]}")
    print(f"TEXT : {repr(row[col_miss])}")
    print("---")

print()
print("=== EVENTS ===")
events = df[df[col_event].notna() & (df[col_event].str.strip().str.lower() not in ["no","nah",""])]
for _, row in events.iterrows():
    print(f"NAME : {row[col_name]}")
    print(f"EVENT: {repr(row[col_event])}")
    print("---")
