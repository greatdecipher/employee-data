import sys
from pathlib import Path

# Add the parent directory to the Python path
sys.path.insert(0, str(Path(__file__).parent.parent))

from employee_app.generator import generate_employee_data

def test_generate_employee_data():
    df = generate_employee_data(5)
    assert len(df) == 5
    assert set(["emp_id", "full_name", "department", "salary", "hire_date"]).issubset(df.columns)