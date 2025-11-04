from datetime import datetime
import pandas as pd
from pathlib import Path

def export_to_excel(df: pd.DataFrame, folder_path: str):
    folder = Path(folder_path)
    folder.mkdir(parents=True, exist_ok=True)
    file_path = folder / "employees.xlsx"

    with pd.ExcelWriter(file_path, engine="xlsxwriter") as writer:
        # Employees sheet
        df.to_excel(writer, index=False, sheet_name="Employees")

        # Summary sheet
        summary = df.groupby("department")["salary"].mean().reset_index()
        summary.columns = ["Department", "Average Salary"]

        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Write timestamp & summary starting at row 2
        summary.to_excel(writer, index=False, sheet_name="Summary", startrow=2)
        sheet = writer.sheets["Summary"]
        sheet.write(0, 0, f"Exported on: {timestamp}")

    return file_path