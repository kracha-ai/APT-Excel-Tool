import pandas as pd
from tkinter import Tk, filedialog, messagebox
import os

def process_multiple_excels():
    root = Tk()
    root.withdraw()

    # Step 1: Ask user to select multiple files
    file_paths = filedialog.askopenfilenames(
        title="Select Excel Files to Combine",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    if not file_paths:
        messagebox.showwarning("Warning", "No files selected!")
        return

    merged_df = pd.DataFrame()

    for file_path in file_paths:
        try:
            ext = os.path.splitext(file_path)[1].lower()
            df = None

            # --- Detect file type and read accordingly ---
            if ext == ".xlsx":
                df = pd.read_excel(file_path, engine="openpyxl")

            elif ext == ".xls":
                try:
                    df = pd.read_excel(file_path, engine="xlrd")
                except Exception:
                    # if it's actually HTML disguised as .xls
                    try:
                        df = pd.read_html(file_path)[0]
                        messagebox.showinfo("Info", f"Read {os.path.basename(file_path)} as HTML table instead of Excel.")
                    except Exception as e2:
                        raise e2

            else:
                messagebox.showwarning("Warning", f"Skipping {file_path}: unsupported file type.")
                continue

        except Exception as e:
            messagebox.showerror("Error", f"Cannot read {file_path}:\n{e}")
            continue

        # --- Try to find Date and Amount columns automatically ---
        date_col = None
        amount_col = None
        for col in df.columns:
            col_lower = str(col).lower()
            if "date" in col_lower:
                date_col = col
            if "amount" in col_lower or "amt" in col_lower:
                amount_col = col

        if not date_col or not amount_col:
            messagebox.showwarning("Warning", f"Skipping {file_path}: missing Date/Amount columns.")
            continue

        # --- Convert date column to datetime for reliable sorting ---
        df[date_col] = pd.to_datetime(df[date_col], errors="coerce")

        # --- Select and sort each file ---
        df = df[[date_col, amount_col]].sort_values(by=date_col)

        # --- Rename columns with file name ---
        file_name = os.path.splitext(os.path.basename(file_path))[0]
        df.columns = [f"{file_name}_Date", f"{file_name}_Amount"]

        # --- Reset index for alignment and combine ---
        df = df.reset_index(drop=True)
        merged_df = pd.concat([merged_df, df], axis=1)

    if merged_df.empty:
        messagebox.showinfo("Info", "No valid data processed.")
        return

    # ✅ Global Sort across all dates (preserves duplicates)
    # Step 1: Create a helper column that holds the first non-null date in each row
    date_columns = [c for c in merged_df.columns if "date" in c.lower()]
    merged_df["__sort_date__"] = merged_df[date_columns].bfill(axis=1).iloc[:, 0]

    # Step 2: Sort the entire dataframe by that helper date
    merged_df = merged_df.sort_values(by="__sort_date__").drop(columns="__sort_date__").reset_index(drop=True)

    # Step 3: Save the output
    save_path = filedialog.asksaveasfilename(
        title="Save Combined File As",
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")]
    )
    if not save_path:
        return

    try:
        merged_df.to_excel(save_path, index=False)
        messagebox.showinfo("Success", f"✅ Combined file saved (sorted by date):\n{save_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Cannot save file:\n{e}")

# Run the process
if __name__ == "__main__":
    process_multiple_excels()
