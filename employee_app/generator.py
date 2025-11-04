import random
from datetime import date, datetime
import pandas as pd
from faker import Faker

fake = Faker("en_PH")  # use Filipino/English locale
DEPARTMENTS = ["IT", "HR", "Operations", "Administration", "Finance"]

def generate_employee_data(n: int) -> pd.DataFrame:
    if n <= 0:
        raise ValueError("Number of employees must be positive.")

    data = []
    start_date = datetime(2020, 1, 1).date()
    end_date = date.today()
    
    for i in range(1, n + 1):
        data.append({
            "emp_id": i,
            "full_name": fake.name(),
            "department": random.choice(DEPARTMENTS),
            "salary": random.randint(25000, 120000),
            "hire_date": fake.date_between(start_date=start_date, end_date=end_date)
        })

    return pd.DataFrame(data)