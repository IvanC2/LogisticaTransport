import re

with open('database.py', 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Update init_db to add driver_profiles
new_table = """    # Tabel Driver Profiles
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS driver_profiles (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            driver_name TEXT UNIQUE NOT NULL,
            base_salary REAL,
            day_rate REAL,
            km_rate REAL
        )
    ''')
"""
init_db_end = content.find("conn.commit()")
if "driver_profiles" not in content:
    content = content[:init_db_end] + new_table + "\n    " + content[init_db_end:]

# 2. Add new functions
new_functions = """
def get_driver_profile(driver_name):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT base_salary, day_rate, km_rate FROM driver_profiles WHERE driver_name=?", (driver_name,))
    row = cursor.fetchone()
    conn.close()
    if row:
        return {"base_salary": row[0], "day_rate": row[1], "km_rate": row[2]}
    return {"base_salary": 0.0, "day_rate": 0.0, "km_rate": 0.0}

def save_driver_profile(name, base, day, km):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO driver_profiles (driver_name, base_salary, day_rate, km_rate)
        VALUES (?, ?, ?, ?)
        ON CONFLICT(driver_name) DO UPDATE SET
            base_salary=excluded.base_salary,
            day_rate=excluded.day_rate,
            km_rate=excluded.km_rate
    ''', (name, base, day, km))
    conn.commit()
    conn.close()

def get_trips_for_salary(driver_name, month, year):
    conn = get_connection()
    cursor = conn.cursor()
    # Assuming date is stored as YYYY-MM-DD
    # We can match using LIKE 'YYYY-MM-%'
    date_pattern = f"{year}-{str(month).zfill(2)}-%"
    
    cursor.execute('''
        SELECT date, client, km_total, daily_allowance, bonus, meal_tickets 
        FROM trips 
        WHERE driver_name=? AND date LIKE ?
    ''', (driver_name, date_pattern))
    
    columns = [column[0] for column in cursor.description]
    results = []
    for row in cursor.fetchall():
        results.append(dict(zip(columns, row)))
    conn.close()
    return results
"""

if "def get_driver_profile" not in content:
    content += new_functions

with open('database.py', 'w', encoding='utf-8') as f:
    f.write(content)
