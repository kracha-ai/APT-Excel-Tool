import pandas as pd
from tkinter import Tk, filedialog, messagebox
import os

def sort_each_date_column():
    root = Tk()
    root.withdraw()

    # Step 1: Choose the Excel file
    input_path = filedialog.askopenfilename(
        title="Select Excel File with Multiple Date Columns",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    if not input_path:
        messagebox.showinfo("Cancelled", "No file selected.")
        return

    try:
        ext = os.path.splitext(input_path)[1].lower()
        if ext == ".xlsx":
            df = pd.read_excel(input_path, engine="openpyxl")
        else:
            df = pd.read_excel(input_path, engine="xlrd")
    except Exception as e:
        messagebox.showerror("Error", f"Cannot read file:\n{e}")
        return

    # Step 2: Find all columns that contain 'date'
    date_cols = [col for col in df.columns if "date" in str(col).lower()]

    if not date_cols:
        messagebox.showwarning("No Date Columns", "No columns with 'date' found.")
        return

    # Step 3: Process each Date–Amount pair
    new_df_parts = []  # will hold sorted pairs

    for date_col in date_cols:
        # try to find the "amount" column right next to this date
        idx = df.columns.get_loc(date_col)
        amount_col = None
        if idx + 1 < len(df.columns):
            next_col = df.columns[idx + 1]
            if "amount" in str(next_col).lower() or "amt" in str(next_col).lower():
                amount_col = next_col

        # extract the date (and amount if exists)
        cols_to_keep = [date_col]
        if amount_col:
            cols_to_keep.append(amount_col)

        temp = df[cols_to_keep].copy()

        # convert and sort only this pair
        temp[date_col] = pd.to_datetime(temp[date_col], errors="coerce", dayfirst=True)
        temp = temp.dropna(subset=[date_col])
        temp = temp.sort_values(by=date_col, ascending=True)

        # convert back to DD-MM-YYYY
        temp[date_col] = temp[date_col].dt.strftime("%d-%m-%Y")

        temp = temp.reset_index(drop=True)
        new_df_parts.append(temp)

    # Step 4: Combine all sorted parts side by side
    final_df = pd.concat(new_df_parts, axis=1)

    # Step 5: Save output
    save_path = filedialog.asksaveasfilename(
        title="Save Sorted File As",
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")]
    )
    if not save_path:
        return

    try:
        final_df.to_excel(save_path, index=False, engine="openpyxl")
        messagebox.showinfo("Success", f"✅ Each date column sorted individually and saved to:\n{save_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Cannot save file:\n{e}")

if __name__ == "__main__":
    sort_each_date_column()
