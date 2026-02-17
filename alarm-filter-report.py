import pandas as pd


input_path = r"E:\alarmlist_2025_08_04_11_04.xlsx"  # OVO JE MJESTO GDJE SE EXCEL file NALAZI, treba ga specificirati
output_path = r"E:\filtered_alarms_with_summary.xlsx"  # OVO JE MJESTO FILTRIRANOG EXCEL FAJLA KOJEG ČE PROGRAM NAPRAVITI

# Alarm type to filter
alarm_type_to_filter = "RedundancyNoRedundancy"  # OVDJE MJENJATE VRSTU ALARMA, u ovom slučaju je RedundancyNoRedundancy

# Load Excel into DataFrame
df = pd.read_excel(input_path)

# Filter DataFrame 
filtered_df = df[df['Alarm type'] == alarm_type_to_filter]

# Statistics
total_alarms = len(df)
filtered_alarms = len(filtered_df)
percentage = (filtered_alarms / total_alarms) * 100 if total_alarms > 0 else 0

summary_df = pd.DataFrame({
    'Metric': ['Total alarms', f"Alarms of type '{alarm_type_to_filter}'", 'Percentage'],
    'Value': [total_alarms, filtered_alarms, f"{percentage:.2f}%"]
})


with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    filtered_df.to_excel(writer, sheet_name='Filtered Data', index=False)
    summary_df.to_excel(writer, sheet_name='Summary', index=False)

print(f"Filtered data and summary saved to {output_path}")
print(f"Total alarms: {total_alarms}")
print(f"Filtered alarms ({alarm_type_to_filter}): {filtered_alarms}")
print(f"Percentage: {percentage:.2f}%")
