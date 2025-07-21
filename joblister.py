import sys
import os
import pandas as pd
from PyQt5.QtWidgets import QApplication, QFileDialog, QMessageBox

def get_unique_filename(base_name="joblist.xlsx"):
    if not os.path.exists(base_name):
        return base_name
    name, ext = os.path.splitext(base_name)
    i = 1
    while True:
        new_name = f"{name}_{i:02d}{ext}"
        if not os.path.exists(new_name):
            return new_name
        i += 1

def process_excel(file_path):
    df = pd.read_excel(file_path, engine='xlrd')

    # Remove rows up to and including "Project Number", plus next two rows
    target_index = df[df.iloc[:, 0].astype(str).str.contains("Project Number", na=False)].index
    if not target_index.empty:
        df = df.iloc[target_index[0] + 3:]

    # Remove blank rows and columns
    df.dropna(axis=0, how='all', inplace=True)
    df.dropna(axis=1, how='all', inplace=True)

    # Swap first and second columns
    cols = df.columns.tolist()
    if len(cols) >= 2:
        cols[0], cols[1] = cols[1], cols[0]
        df = df[cols]

    # Keep only first two columns
    df = df.iloc[:, :2]

    # Remove rows with blank first column
    df = df[df.iloc[:, 0].notna()]

    # Remove header row
    df = df.iloc[1:]

    # Save to a unique filename
    output_file = get_unique_filename("joblist.xlsx")
    df.to_excel(output_file, index=False, header=False)

    return len(df), output_file

def main():
    app = QApplication(sys.argv)
    file_dialog = QFileDialog()
    file_path, _ = file_dialog.getOpenFileName(None, "Select Excel File", "", "Excel Files (*.xls *.xlsx)")

    if file_path:
        row_count, saved_file = process_excel(file_path)
        QMessageBox.information(None, "Processing Complete", f"Rows remaining: {row_count}\nSaved as: {saved_file}")
    else:
        QMessageBox.warning(None, "No File Selected", "Please select a valid Excel file.")

if __name__ == "__main__":
    main()
