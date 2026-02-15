import pandas as pd

# Step 1: Read Excel file
source_path = "source.xlsx"   # or user can upload later
df = pd.read_excel(source_path)

# Step 2: Select required columns
required_columns = ["src_trans_date", "tot_amount"]  # customize this
filtered_df = df[required_columns]

# Step 3: Sort by date
filtered_df = filtered_df.sort_values(by="src_trans_date")

# Step 4: Save to another Excel file
filtered_df.to_excel("output.xlsx", index=False)

print("✅ Data copied, sorted, and saved to output.xlsx!")

