import sqlite3
import hashlib
import os
import sys

# Determinăm calea absolută către folderul în care se află scriptul sau executabilul
if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
else:
    application_path = os.path.dirname(os.path.abspath(__file__))

DB_NAME = os.path.join(application_path, 'transport_app.db')

def get_connection():
    return sqlite3.connect(DB_NAME)

def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

def init_db():
    conn = get_connection()
    cursor = conn.cursor()

    # 1. CREATE ALL TABLES FIRST
    # Tabel Users
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL
        )
    ''')

    # Tabel Trips (Curse)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS trips (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT,
            driver_name TEXT,
            auto_number TEXT,
            auto_type TEXT,
            presence TEXT,
            km_total REAL,
            km_empty REAL,
            km_loaded REAL,
            daily_allowance REAL,
            bonus REAL,
            meal_tickets REAL,
            client TEXT,
            transport_type TEXT,
            cargo_type TEXT,
            trailer TEXT,
            location TEXT,
            trip_count INTEGER,
            notice_number TEXT,
            uit_code TEXT,
            quantity_tons REAL,
            quantity_m3 REAL,
            pump_hours REAL,
            wait_hours REAL,
            price_per_km REAL,
            price_per_trip REAL,
            price_per_ton REAL,
            price_per_m3 REAL,
            price_per_pump_hour REAL,
            price_per_wait_hour REAL,
            total_price REAL,
            route_description TEXT
        )
    ''')

    # Tabel Monthly Expenses (Cheltuieli lunare)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS monthly_expenses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            month_year TEXT,
            driver_name TEXT,
            driver_role TEXT,
            auto_number TEXT,
            fixed_auto_expense REAL,
            fuel_expense REAL,
            maintenance_repair REAL,
            accommodation_parking REAL,
            other_expenses REAL
        )
    ''')

    # Tabel Driver Profiles
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS driver_profiles (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            driver_name TEXT UNIQUE NOT NULL,
            base_salary REAL,
            day_rate REAL,
            km_rate REAL
        )
    ''')

    # 2. RUN ALL ALTER TABLES / MIGRATIONS
    # Migrare automată pentru versiuni vechi ale bazei de date (Trips)
    cursor.execute("PRAGMA table_info(trips)")
    existing_columns = [col[1] for col in cursor.fetchall()]
    if "route_description" not in existing_columns:
        cursor.execute("ALTER TABLE trips ADD COLUMN route_description TEXT")
        
    try:
        cursor.execute("ALTER TABLE trips ADD COLUMN is_external TEXT DEFAULT 'NU'")
    except Exception: pass

    # Migrare automată pentru monthly_expenses
    cursor.execute("PRAGMA table_info(monthly_expenses)")
    existing_exp_columns = [col[1] for col in cursor.fetchall()]
    new_cols = ['start_date', 'end_date', 'total_revenue', 'total_km', 'cars_used', 'gross_salary', 'total_expenses', 'net_profit']
    for c in new_cols:
        if c not in existing_exp_columns:
            cursor.execute(f"ALTER TABLE monthly_expenses ADD COLUMN {c} REAL" if c not in ['start_date', 'end_date', 'cars_used'] else f"ALTER TABLE monthly_expenses ADD COLUMN {c} TEXT")

    # Update driver_profiles for V3
    try: cursor.execute("ALTER TABLE driver_profiles ADD COLUMN diurna_rate REAL DEFAULT 0.0")
    except Exception: pass
    try: cursor.execute("ALTER TABLE driver_profiles ADD COLUMN premiere_rate REAL DEFAULT 0.0")
    except Exception: pass
    try: cursor.execute("ALTER TABLE driver_profiles ADD COLUMN meal_ticket_rate REAL DEFAULT 0.0")
    except Exception: pass
    try: cursor.execute("ALTER TABLE driver_profiles ADD COLUMN weekend_rate REAL DEFAULT 0.0")
    except Exception: pass
    try: cursor.execute("ALTER TABLE driver_profiles ADD COLUMN holiday_rate REAL DEFAULT 0.0")
    except Exception: pass

    # Update driver_profiles for Transport Types
    try: cursor.execute("ALTER TABLE driver_profiles ADD COLUMN agabaritic_rate REAL DEFAULT 0.0")
    except Exception: pass
    try: cursor.execute("ALTER TABLE driver_profiles ADD COLUMN adr_rate REAL DEFAULT 0.0")
    except Exception: pass
    try: cursor.execute("ALTER TABLE driver_profiles ADD COLUMN insotire_rate REAL DEFAULT 0.0")
    except Exception: pass
    try: cursor.execute("ALTER TABLE driver_profiles ADD COLUMN trailer_rate REAL DEFAULT 0.0")
    except Exception: pass

    # Update for External Trips
    try: cursor.execute("ALTER TABLE driver_profiles ADD COLUMN external_rate REAL DEFAULT 0.0")
    except Exception: pass

    # 3. OTHER INITIALIZATIONS
    # Creăm un utilizator implicit (Admin) dacă baza de date e goală
    cursor.execute("SELECT COUNT(*) FROM users")
    if cursor.fetchone()[0] == 0:
        admin_pass = hash_password('admin123')
        cursor.execute("INSERT INTO users (username, password_hash, role) VALUES (?, ?, ?)",
                       ('admin', admin_pass, 'Admin'))

    conn.commit()
    conn.close()


def authenticate_user(username, password):
    conn = get_connection()
    cursor = conn.cursor()
    pwd_hash = hash_password(password)
    cursor.execute("SELECT role FROM users WHERE username = ? AND password_hash = ?", (username, pwd_hash))
    result = cursor.fetchone()
    conn.close()
    if result:
        return result[0]  # Returnează rolul ('Admin' sau 'Operator')
    return None

def add_user(username, password, role):
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO users (username, password_hash, role) VALUES (?, ?, ?)",
                       (username, hash_password(password), role))
        conn.commit()
        return True
    except sqlite3.IntegrityError:
        return False  # Username-ul există deja
    finally:
        conn.close()

# --- TRIPS (CURSE) CRUD ---

def add_trip(data):
    """
    Primește un dicționar cu datele cursei. 
    Total Preț este deja calculat în interfață, deci va fi pasat gata completat în dicționar.
    """
    conn = get_connection()
    cursor = conn.cursor()
    columns = ', '.join(data.keys())
    placeholders = ', '.join(['?'] * len(data))
    values = tuple(data.values())
    
    query = f"INSERT INTO trips ({columns}) VALUES ({placeholders})"
    cursor.execute(query, values)
    conn.commit()
    conn.close()

def get_all_trips():
    conn = get_connection()
    conn.row_factory = sqlite3.Row  # Returnează rezultatele sub formă de dicționar
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM trips ORDER BY date DESC, id DESC")
    rows = cursor.fetchall()
    conn.close()
    return [dict(row) for row in rows]

def get_trips_with_filters(driver=None, auto=None, client=None, start_date=None, end_date=None):
    conn = get_connection()
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    query = "SELECT * FROM trips WHERE 1=1"
    params = []
    
    if driver:
        query += " AND driver_name LIKE ?"
        params.append(f"%{driver}%")
    if auto:
        query += " AND auto_number LIKE ?"
        params.append(f"%{auto}%")
    if client:
        query += " AND client LIKE ?"
        params.append(f"%{client}%")
    if start_date and end_date:
        query += " AND date >= ? AND date <= ?"
        params.extend([start_date, end_date])
        
    query += " ORDER BY date DESC, id DESC"
    cursor.execute(query, params)
    rows = cursor.fetchall()
    conn.close()
    return [dict(row) for row in rows]
    
def get_trip_by_id(trip_id):
    conn = get_connection()
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM trips WHERE id = ?", (trip_id,))
    row = cursor.fetchone()
    conn.close()
    return dict(row) if row else None
    
def delete_trip(trip_id):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM trips WHERE id = ?", (trip_id,))
    conn.commit()
    conn.close()

def update_trip(trip_id, data):
    conn = get_connection()
    cursor = conn.cursor()
    set_clause = ', '.join([f"{k} = ?" for k in data.keys()])
    values = tuple(data.values()) + (trip_id,)
    query = f"UPDATE trips SET {set_clause} WHERE id = ?"
    cursor.execute(query, values)
    conn.commit()
    conn.close()

# --- EXPENSES (CHELTUIELI) CRUD ---

def add_monthly_expense(data):
    conn = get_connection()
    cursor = conn.cursor()
    columns = ', '.join(data.keys())
    placeholders = ', '.join(['?'] * len(data))
    values = tuple(data.values())
    
    query = f"INSERT INTO monthly_expenses ({columns}) VALUES ({placeholders})"
    cursor.execute(query, values)
    conn.commit()
    conn.close()

def get_monthly_expenses(month_year):
    conn = get_connection()
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM monthly_expenses WHERE month_year = ?", (month_year,))
    rows = cursor.fetchall()
    conn.close()
    return [dict(row) for row in rows]

def get_all_monthly_expenses():
    conn = get_connection()
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM monthly_expenses ORDER BY month_year DESC, id DESC")
    rows = cursor.fetchall()
    conn.close()
    return [dict(row) for row in rows]

def get_monthly_revenue(month_year, driver_name, auto_number):
    """
    Returnează suma(total_price) și suma(km_total) pentru o lună (YYYY-MM), un șofer și o mașină.
    """
    conn = get_connection()
    cursor = conn.cursor()
    query = """
        SELECT SUM(total_price), SUM(km_total) FROM trips 
        WHERE date LIKE ? AND driver_name = ? AND auto_number = ?
    """
    cursor.execute(query, (f"{month_year}-%", driver_name, auto_number))
    result = cursor.fetchone()
    conn.close()
    
    total_revenue = result[0] if result[0] is not None else 0.0
    total_km = result[1] if result[1] is not None else 0.0
    return total_revenue, total_km

def delete_financial_report(report_id):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM monthly_expenses WHERE id = ?", (report_id,))
    conn.commit()
    conn.close()

def get_driver_financial_summary(driver_name, start_date, end_date):
    conn = get_connection()
    cursor = conn.cursor()
    query = '''
        SELECT SUM(total_price), SUM(km_total), GROUP_CONCAT(DISTINCT auto_number)
        FROM trips 
        WHERE driver_name = ? AND date >= ? AND date <= ?
    '''
    cursor.execute(query, (driver_name, start_date, end_date))
    result = cursor.fetchone()
    conn.close()
    
    rev = result[0] if result[0] is not None else 0.0
    km = result[1] if result[1] is not None else 0.0
    cars = result[2] if result[2] is not None else ""
    return rev, km, cars

def get_distinct_values(column_name):
    conn = get_connection()
    cursor = conn.cursor()
    # Safely format column_name since it comes from code
    query = f"SELECT DISTINCT {column_name} FROM trips WHERE {column_name} IS NOT NULL AND {column_name} != '' ORDER BY {column_name} ASC"
    try:
        cursor.execute(query)
        rows = cursor.fetchall()
        values = [str(row[0]) for row in rows]
    except Exception as e:
        print(f"Eroare la obținerea valorilor distincte pentru {column_name}: {e}")
        values = []
    conn.close()
    return values

def get_driver_profile(driver_name):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT base_salary, day_rate, km_rate, diurna_rate, premiere_rate, meal_ticket_rate, weekend_rate, holiday_rate, agabaritic_rate, adr_rate, insotire_rate, trailer_rate, external_rate FROM driver_profiles WHERE driver_name=?", (driver_name,))
    row = cursor.fetchone()
    conn.close()
    if row:
        return {
            "base_salary": row[0], "day_rate": row[1], "km_rate": row[2],
            "diurna_rate": row[3] if row[3] is not None else 0.0,
            "premiere_rate": row[4] if row[4] is not None else 0.0,
            "meal_ticket_rate": row[5] if row[5] is not None else 0.0,
            "weekend_rate": row[6] if row[6] is not None else 0.0,
            "holiday_rate": row[7] if row[7] is not None else 0.0,
            "agabaritic_rate": row[8] if row[8] is not None else 0.0,
            "adr_rate": row[9] if row[9] is not None else 0.0,
            "insotire_rate": row[10] if row[10] is not None else 0.0,
            "trailer_rate": row[11] if row[11] is not None else 0.0,
            "external_rate": row[12] if row[12] is not None else 0.0
        }
    return {"base_salary": 0.0, "day_rate": 0.0, "km_rate": 0.0, "diurna_rate": 0.0, "premiere_rate": 0.0, "meal_ticket_rate": 0.0, "weekend_rate": 0.0, "holiday_rate": 0.0, "agabaritic_rate": 0.0, "adr_rate": 0.0, "insotire_rate": 0.0, "trailer_rate": 0.0, "external_rate": 0.0}

def save_driver_profile(name, base, day, km, diurna=0.0, premiere=0.0, meal=0.0, weekend=0.0, holiday=0.0, agabaritic=0.0, adr=0.0, insotire=0.0, trailer=0.0, external=0.0):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO driver_profiles (driver_name, base_salary, day_rate, km_rate, diurna_rate, premiere_rate, meal_ticket_rate, weekend_rate, holiday_rate, agabaritic_rate, adr_rate, insotire_rate, trailer_rate, external_rate)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(driver_name) DO UPDATE SET
            base_salary=excluded.base_salary,
            day_rate=excluded.day_rate,
            km_rate=excluded.km_rate,
            diurna_rate=excluded.diurna_rate,
            premiere_rate=excluded.premiere_rate,
            meal_ticket_rate=excluded.meal_ticket_rate,
            weekend_rate=excluded.weekend_rate,
            holiday_rate=excluded.holiday_rate,
            agabaritic_rate=excluded.agabaritic_rate,
            adr_rate=excluded.adr_rate,
            insotire_rate=excluded.insotire_rate,
            trailer_rate=excluded.trailer_rate,
            external_rate=excluded.external_rate
    ''', (name, base, day, km, diurna, premiere, meal, weekend, holiday, agabaritic, adr, insotire, trailer, external))
    conn.commit()
    conn.close()

def get_trips_for_salary(driver_name, month, year):
    conn = get_connection()
    cursor = conn.cursor()
    # Assuming date is stored as YYYY-MM-DD
    # We can match using LIKE 'YYYY-MM-%'
    date_pattern = f"{year}-{str(month).zfill(2)}-%"
    
    cursor.execute('''
        SELECT date, client, km_total, daily_allowance, bonus, meal_tickets, presence, transport_type, trailer, is_external 
        FROM trips 
        WHERE driver_name=? AND date LIKE ?
    ''', (driver_name, date_pattern))
    
    columns = [column[0] for column in cursor.description]
    results = []
    for row in cursor.fetchall():
        results.append(dict(zip(columns, row)))
    conn.close()
    return results

def delete_driver_profile(driver_name):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM driver_profiles WHERE driver_name=?", (driver_name,))
    conn.commit()
    conn.close()

def get_all_driver_profiles():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT driver_name, base_salary, day_rate, km_rate, diurna_rate, premiere_rate, meal_ticket_rate, weekend_rate, holiday_rate, agabaritic_rate, adr_rate, insotire_rate, trailer_rate, external_rate FROM driver_profiles ORDER BY driver_name ASC")
    columns = [column[0] for column in cursor.description]
    results = []
    for row in cursor.fetchall():
        results.append(dict(zip(columns, row)))
    conn.close()
    return results
