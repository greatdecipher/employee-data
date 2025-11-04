import tempfile
import pandas as pd
import sys
from pathlib import Path

# Add the parent directory to the Python path
sys.path.insert(0, str(Path(__file__).parent.parent))

from employee_app.exporter import export_to_excel

def test_export_to_excel():
    df = pd.DataFrame({
        "emp_id": [1, 2],
        "full_name": ["A", "B"],
        "department": ["IT", "HR"],
        "salary": [50000, 60000],
        "hire_date": ["2021-01-01", "2022-01-01"]
    })

    with tempfile.TemporaryDirectory() as tmpdir:
        file_path = export_to_excel(df, tmpdir)
        assert file_path.exists()