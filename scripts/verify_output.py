import pandas as pd
import openpyxl

# Load the output file
output_file = "Washout_Matrix_Output.xlsx"

print("=== Washout Matrix Output Verification ===\n")

# Read the Washout Matrix sheet
df = pd.read_excel(output_file, sheet_name="Washout Matrix")

print(f"Total transitions: {len(df)}")
print(f"\nColumns: {list(df.columns)}\n")

print("Summary Statistics:")
print(f"  - Washout Needed YES: {(df['Washout_Needed'] == 'YES').sum()}")
print(f"  - Washout Needed NO: {(df['Washout_Needed'] == 'NO').sum()}")
print(f"  - HEEL WASH: {(df['Washout_Type'] == 'HEEL WASH').sum()}")
print(f"  - FULL WASH: {(df['Washout_Type'] == 'FULL WASH').sum()}")
print(f"  - Acid Wash YES: {(df['Acid_Wash'] == 'YES').sum()}")
print(f"  - Acid Wash NO: {(df['Acid_Wash'] == 'NO').sum()}")

print("\n" + "=" * 80)
print("Sample Transitions (first 10 rows):")
print("=" * 80)

# Show first 10 rows with key columns
sample_cols = [
    "From_Product",
    "To_Product",
    "Washout_Needed",
    "Washout_Type",
    "Acid_Wash",
    "Duration_Hours",
]
print(df[sample_cols].head(10).to_string(index=False))

print("\n" + "=" * 80)
print("Sample transitions where washout IS needed:")
print("=" * 80)
washout_yes = df[df["Washout_Needed"] == "YES"][sample_cols].head(10)
print(washout_yes.to_string(index=False))

print("\n" + "=" * 80)
print("Sample transitions where ACID WASH is needed:")
print("=" * 80)
acid_wash_yes = df[df["Acid_Wash"] == "YES"][sample_cols].head(10)
print(acid_wash_yes.to_string(index=False))

print("\n" + "=" * 80)
print("Average Duration Statistics:")
print("=" * 80)
print(f"Overall average duration: {df['Duration_Hours'].mean():.2f} hours")
print(
    f"When washout needed: {df[df['Washout_Needed']=='YES']['Duration_Hours'].mean():.2f} hours"
)
print(
    f"HEEL WASH average: {df[df['Washout_Type']=='HEEL WASH']['Duration_Hours'].mean():.2f} hours"
)
print(
    f"FULL WASH average: {df[df['Washout_Type']=='FULL WASH']['Duration_Hours'].mean():.2f} hours"
)
print(
    f"With acid wash: {df[df['Acid_Wash']=='YES']['Duration_Hours'].mean():.2f} hours"
)
