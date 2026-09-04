import customtkinter as ctk
import database
import datetime
import tkinter.ttk as ttk
from tkinter import filedialog, messagebox
import pandas as pd
from tkcalendar import DateEntry
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")


def apply_excel_formatting(filename, date_col_name=None, tip_zi_col_name="Tip Zi"):
    try:
        wb = load_workbook(filename)
        ws = wb.active
        
        # Format Header
        header_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
        header_font = Font(bold=True)
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            
        # Adjust column widths
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)
            if adjusted_width < 15:
                adjusted_width = 15
            elif adjusted_width > 40:
                adjusted_width = 40
            ws.column_dimensions[column].width = adjusted_width
            
        # Highlight weekends and holidays based on Tip Zi
        tip_zi_col_idx = None
        for cell in ws[1]:
            if cell.value == tip_zi_col_name:
                tip_zi_col_idx = cell.column
                break
                
        if tip_zi_col_idx:
            weekend_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid") # Light yellow
            holiday_fill = PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid") # Light orange/red
            
            for row in ws.iter_rows(min_row=2):
                cell_val = row[tip_zi_col_idx - 1].value
                if cell_val == "Weekend":
                    for cell in row:
                        cell.fill = weekend_fill
                elif cell_val == "Sărbătoare legală":
                    for cell in row:
                        cell.fill = holiday_fill
        elif date_col_name:
            # Fallback for Financiar Frame
            date_col_idx = None
            for cell in ws[1]:
                if cell.value == date_col_name:
                    date_col_idx = cell.column
                    break
                    
            if date_col_idx:
                weekend_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid") # Light yellow
                
                for row in ws.iter_rows(min_row=2):
                    date_cell = row[date_col_idx - 1]
                    if date_cell.value:
                        try:
                            # Parse date assuming YYYY-MM-DD
                            dt = datetime.datetime.strptime(str(date_cell.value).strip(), "%Y-%m-%d")
                            if dt.weekday() >= 5: # 5 = Saturday, 6 = Sunday
                                for cell in row:
                                    cell.fill = weekend_fill
                        except Exception:
                            pass
                            
        wb.save(filename)
    except Exception as e:
        print(f"Eroare la formatarea Excel: {e}")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Aplicație Transport - Dashboard")
        self.geometry("900x600")

        # Inițializăm baza de date la pornire (în caz că nu există)
        database.init_db()

        self.current_frame = None
        self.show_login_frame()

    def show_login_frame(self):
        if self.current_frame:
            self.current_frame.destroy()
            
        self.current_frame = LoginFrame(self, self.on_login_success)
        self.current_frame.pack(fill="both", expand=True)

    def on_login_success(self, role):
        if self.current_frame:
            self.current_frame.destroy()
            
        self.current_frame = MainDashboardFrame(self, role, self.show_login_frame)
        self.current_frame.pack(fill="both", expand=True)


class LoginFrame(ctk.CTkFrame):
    def __init__(self, master, login_success_callback):
        super().__init__(master)
        self.login_success_callback = login_success_callback

        # Configuram un grid pentru a centra formularul de login
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(2, weight=1)

        login_box = ctk.CTkFrame(self, corner_radius=15)
        login_box.grid(row=1, column=1, padx=20, pady=20)

        title = ctk.CTkLabel(login_box, text="Autentificare", font=ctk.CTkFont(size=24, weight="bold"))
        title.pack(pady=(20, 10), padx=30)

        self.username_entry = ctk.CTkEntry(login_box, placeholder_text="Username", width=200)
        self.username_entry.pack(pady=10, padx=30)

        self.password_entry = ctk.CTkEntry(login_box, placeholder_text="Parolă", show="*", width=200)
        self.password_entry.pack(pady=10, padx=30)
        # Permitem logarea și la apăsarea tastei Enter
        self.password_entry.bind("<Return>", lambda event: self.login()) 

        self.login_btn = ctk.CTkButton(login_box, text="Login", command=self.login, width=200)
        self.login_btn.pack(pady=(10, 20), padx=30)

        self.error_label = ctk.CTkLabel(login_box, text="", text_color="red")
        self.error_label.pack(pady=(0, 10))

    def login(self):
        username = self.username_entry.get()
        password = self.password_entry.get()

        if not username or not password:
            self.error_label.configure(text="Completează ambele câmpuri!")
            return

        # Verificăm în baza de date
        role = database.authenticate_user(username, password)
        if role:
            self.login_success_callback(role)
        else:
            self.error_label.configure(text="Username sau parolă incorecte!")



class DailyTripsFormFrame(ctk.CTkScrollableFrame):
    def __init__(self, master):
        super().__init__(master, fg_color="transparent")
        
        self.cart_list = []
        self.vars = {}
        self.entries = {}
        
        title = ctk.CTkLabel(self, text="Curse Noi - Coș Zilnic", font=ctk.CTkFont(size=24, weight="bold"))
        title.pack(anchor="w", pady=(0, 20))
        
        # 1. FIXED SECTION
        fixed_frame = ctk.CTkFrame(self)
        fixed_frame.pack(fill="x", pady=10, padx=5)
        
        ctk.CTkLabel(fixed_frame, text="1. Date Generale (Per Șofer / Zi)", font=ctk.CTkFont(weight="bold", size=16), text_color="#1f6aa5").grid(row=0, column=0, columnspan=6, sticky="w", pady=10, padx=10)
        
        self.fixed_fields = [
            ("date", "Data"), ("driver_name", "Nume Șofer"), ("presence", "Prezență"),
            ("daily_allowance", "Diurnă"), ("bonus", "Premiere"), ("meal_tickets", "Bonuri Masă")
        ]
        
        row, col = 1, 0
        for field_id, label in self.fixed_fields:
            ctk.CTkLabel(fixed_frame, text=label).grid(row=row, column=col*2, padx=10, pady=5, sticky="e")
            var = ctk.StringVar()
            self.vars[field_id] = var
            if field_id == "date":
                widget = DateEntry(fixed_frame, textvariable=var, date_pattern='yyyy-mm-dd', background='darkblue', foreground='white', borderwidth=2, font=('Helvetica', 12), width=15)
                var.set(datetime.date.today().strftime("%Y-%m-%d"))
            elif field_id == "presence":
                widget = ctk.CTkOptionMenu(fixed_frame, variable=var, values=["Prezent", "Liber", "Garaj", "Medical", "Concediu"])
                var.set("Liber")
            elif field_id == "driver_name":
                vals = database.get_distinct_values("driver_name")
                widget = ctk.CTkComboBox(fixed_frame, variable=var, values=vals)
            else:
                widget = ctk.CTkEntry(fixed_frame, textvariable=var)
            widget.grid(row=row, column=col*2+1, padx=10, pady=5, sticky="w")
            self.entries[field_id] = widget
            
            col += 1
            if col > 2:
                col = 0
                row += 1

        # 2. DYNAMIC SECTION
        dyn_frame = ctk.CTkFrame(self)
        dyn_frame.pack(fill="x", pady=10, padx=5)
        
        ctk.CTkLabel(dyn_frame, text="2. Adaugă Cursă / Client", font=ctk.CTkFont(weight="bold", size=16), text_color="#1f6aa5").grid(row=0, column=0, columnspan=6, sticky="w", pady=10, padx=10)
        
        self.dyn_fields = [
            ("auto_number", "Număr Auto", "Auto"), ("auto_type", "Tip Auto", "Auto"),
            ("client", "Client", "Marfă"), ("transport_type", "Tip Transport", "Marfă"), ("special_transport_count", "Nr. Curse Speciale", "Marfă"), ("cargo_type", "Tip Marfă", "Marfă"),
            ("trailer", "Remorcă", "Marfă"), ("notice_number", "Nr. Aviz", "Marfă"), ("uit_code", "Cod UIT", "Marfă"),
            ("location", "Locație/Rută", "Traseu"), ("route_description", "Descriere Traseu", "Traseu"), ("is_external", "Cursă Externă", "Traseu"),
            ("km_total", "Km Total", "Traseu"), ("km_empty", "Km Gol", "Traseu"), ("km_loaded", "Km Încărcat", "Traseu"),
            ("quantity_tons", "Cant. Tone", "Cantități"), ("quantity_m3", "Cant. m³", "Cantități"), ("trip_count", "Nr. Curse", "Cantități"),
            ("pump_hours", "Ore Pompă", "Cantități"), ("wait_hours", "Ore Staționare", "Cantități"),
            ("price_per_km", "Preț/Km", "Tarife"), ("price_per_trip", "Preț/Cursă", "Tarife"),
            ("price_per_ton", "Preț/Tonă", "Tarife"), ("price_per_m3", "Preț/m³", "Tarife"),
            ("price_per_pump_hour", "Preț Oră Pompă", "Tarife"), ("price_per_wait_hour", "Preț Oră Stațion.", "Tarife"),
            ("total_price", "Total Preț", "Final")
        ]
        
        self.calc_fields = [
            "km_total", "price_per_km", "trip_count", "price_per_trip",
            "quantity_tons", "price_per_ton", "quantity_m3", "price_per_m3",
            "pump_hours", "price_per_pump_hour", "wait_hours", "price_per_wait_hour"
        ]
        
        row, col = 1, 0
        current_sec = ""
        for field_id, label, sec in self.dyn_fields:
            if sec != current_sec:
                if col != 0:
                    col = 0
                    row += 1
                current_sec = sec
                ctk.CTkLabel(dyn_frame, text=sec, font=ctk.CTkFont(weight="bold", size=12), text_color="#aaaaaa").grid(row=row, column=0, columnspan=6, sticky="w", pady=(10, 0), padx=10)
                row += 1
                
            ctk.CTkLabel(dyn_frame, text=label).grid(row=row, column=col*2, padx=10, pady=5, sticky="e")
            
            if field_id == "route_description":
                widget = ctk.CTkTextbox(dyn_frame, height=60, wrap="word")
                widget.grid(row=row, column=col*2+1, padx=10, pady=5, sticky="ew")
                self.entries[field_id] = widget
            else:
                var = ctk.StringVar()
                self.vars[field_id] = var
                if field_id in self.calc_fields:
                    var.trace_add("write", self.calculate_total)
                if field_id in ["km_total", "km_empty"]:
                    var.trace_add("write", self.calculate_loaded_km)
                    
                if field_id == "total_price":
                    widget = ctk.CTkEntry(dyn_frame, textvariable=var, font=ctk.CTkFont(weight="bold"), text_color="#f39c12")
                elif field_id == "auto_number":
                    vals = database.get_distinct_values("auto_number")
                    widget = ctk.CTkComboBox(dyn_frame, variable=var, values=vals)
                elif field_id == "auto_type":
                    vals = database.get_distinct_values("auto_type")
                    widget = ctk.CTkComboBox(dyn_frame, variable=var, values=vals)
                elif field_id == "client":
                    vals = database.get_distinct_values("client")
                    widget = ctk.CTkComboBox(dyn_frame, variable=var, values=vals)
                elif field_id == "location":
                    vals = database.get_distinct_values("location")
                    widget = ctk.CTkComboBox(dyn_frame, variable=var, values=vals)
                elif field_id == "transport_type":
                    base_vals = ["", "AGABARITIC", "ADR", "INSOTIRE"]
                    db_vals = database.get_distinct_values("transport_type")
                    vals = list(dict.fromkeys(base_vals + db_vals))
                    widget = ctk.CTkComboBox(dyn_frame, variable=var, values=vals)
                elif field_id == "trailer":
                    vals = ["NU", "DA"]
                    widget = ctk.CTkComboBox(dyn_frame, variable=var, values=vals)
                    var.set("NU")
                elif field_id == "is_external":
                    vals = ["NU", "DA"]
                    widget = ctk.CTkComboBox(dyn_frame, variable=var, values=vals)
                    var.set("NU")
                elif field_id == "special_transport_count":
                    widget = ctk.CTkEntry(dyn_frame, textvariable=var)
                    var.set("1")
                else:
                    widget = ctk.CTkEntry(dyn_frame, textvariable=var)
                widget.grid(row=row, column=col*2+1, padx=10, pady=5, sticky="w")
                self.entries[field_id] = widget
                
            col += 1
            if col > 2:
                col = 0
                row += 1
                
        ctk.CTkButton(dyn_frame, text="Adaugă Cursă în Lista de Azi", command=self.add_to_cart, fg_color="#3498db", hover_color="#2980b9").grid(row=row+1, column=0, columnspan=6, pady=20)
        
        # 3. SUMMARY SECTION (Cart)
        cart_frame = ctk.CTkFrame(self)
        cart_frame.pack(fill="x", pady=10, padx=5)
        
        ctk.CTkLabel(cart_frame, text="3. Sumar Curse (Listă Temporară)", font=ctk.CTkFont(weight="bold", size=16), text_color="#1f6aa5").pack(anchor="w", pady=10, padx=10)
        
        self.tree = ttk.Treeview(cart_frame, columns=("client", "auto", "km", "price"), show="headings", height=5)
        self.tree.heading("client", text="Client")
        self.tree.heading("auto", text="Auto")
        self.tree.heading("km", text="Km Total")
        self.tree.heading("price", text="Total Preț")
        self.tree.pack(fill="x", padx=10, pady=5)
        
        btn_frame = ctk.CTkFrame(cart_frame, fg_color="transparent")
        btn_frame.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkButton(btn_frame, text="Șterge Cursă Selectată", command=self.remove_from_cart, fg_color="#c0392b", hover_color="#e74c3c").pack(side="left")
        ctk.CTkButton(btn_frame, text="Salvează Toată Ziua", command=self.save_all_trips, fg_color="#27ae60", hover_color="#2ecc71", font=ctk.CTkFont(weight="bold", size=14)).pack(side="right")
        
        self.msg_label = ctk.CTkLabel(self, text="", font=ctk.CTkFont(weight="bold"))
        self.msg_label.pack(pady=10)

    def get_float(self, var_name):
        try:
            val = self.vars[var_name].get().strip().replace(',', '.')
            return float(val) if val else 0.0
        except ValueError:
            return 0.0

    def calculate_loaded_km(self, *args):
        try:
            total = self.get_float("km_total")
            empty = self.get_float("km_empty")
            self.vars["km_loaded"].set(f"{total - empty:g}")
        except Exception: pass

    def calculate_total(self, *args):
        try:
            total = (
                self.get_float("km_total") * self.get_float("price_per_km") +
                self.get_float("trip_count") * self.get_float("price_per_trip") +
                self.get_float("quantity_tons") * self.get_float("price_per_ton") +
                self.get_float("quantity_m3") * self.get_float("price_per_m3") +
                self.get_float("pump_hours") * self.get_float("price_per_pump_hour") +
                self.get_float("wait_hours") * self.get_float("price_per_wait_hour")
            )
            self.vars["total_price"].set(f"{total:.2f}")
        except Exception: pass

    def add_to_cart(self):
        client = self.vars["client"].get().strip()
        auto = self.vars["auto_number"].get().strip()
        
        trip_data = {}
        for f, l in self.fixed_fields:
            trip_data[f] = self.vars[f].get().strip()
            
        for f, l, s in self.dyn_fields:
            if f == "route_description":
                trip_data[f] = self.entries[f].get("1.0", "end-1c").strip()
            else:
                trip_data[f] = self.vars[f].get().strip()
                
        self.cart_list.append(trip_data)
        
        item_id = self.tree.insert("", "end", values=(client, auto, trip_data["km_total"], trip_data["total_price"]))
        trip_data["_tree_id"] = item_id
        
        # Clear dynamic fields
        for f, l, s in self.dyn_fields:
            if f == "route_description":
                self.entries[f].delete("1.0", "end")
            else:
                self.vars[f].set("")
                
    def remove_from_cart(self):
        selected = self.tree.selection()
        if not selected: return
        item_id = selected[0]
        self.tree.delete(item_id)
        self.cart_list = [t for t in self.cart_list if t.get("_tree_id") != item_id]
        
    def save_all_trips(self):
        if not self.cart_list:
            messagebox.showwarning("Avertizare", "Nu ai adăugat nicio cursă în coșul zilei!")
            return
            
        for index, trip in enumerate(self.cart_list):
            data = {}
            for field in trip:
                if field == "_tree_id": continue
                
                val = trip[field]
                if field in ["presence", "date", "driver_name", "auto_number", "auto_type", "client", "transport_type", "cargo_type", "trailer", "location", "route_description", "notice_number", "uit_code", "is_external"]:
                    data[field] = val
                elif field == "trip_count":
                    try: data[field] = int(val) if val else 0
                    except ValueError: data[field] = 0
                elif field == "special_transport_count":
                    try: data[field] = int(val) if val else 1
                    except ValueError: data[field] = 1
                else:
                    try: data[field] = float(str(val).replace(',', '.')) if val else 0.0
                    except ValueError: data[field] = 0.0
            
            # Logica ANTI-DUPLICARE pentru cheltuielile zilnice
            if index > 0:
                data["daily_allowance"] = 0.0
                data["bonus"] = 0.0
                data["meal_tickets"] = 0.0
                
            database.add_trip(data)
            
        self.cart_list.clear()
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        # Reset fixed fields too
        for f, l in self.fixed_fields:
            if f != "date" and f != "presence":
                self.vars[f].set("")
        self.vars["presence"].set("Prezent")
        self.vars["date"].set(datetime.date.today().strftime("%Y-%m-%d"))
        
        self.msg_label.configure(text="Toate cursele zilei au fost salvate cu succes!", text_color="#2ecc71")
        self.after(3000, lambda: self.msg_label.configure(text=""))


class TripFormFrame(ctk.CTkScrollableFrame):
    def __init__(self, master, edit_trip=None, on_save_callback=None):
        super().__init__(master, fg_color="transparent")
        self.edit_trip = edit_trip
        self.on_save_callback = on_save_callback
        
        title_text = "Editează Cursa" if edit_trip else "Adaugă Cursă Nouă"
        title = ctk.CTkLabel(self, text=title_text, font=ctk.CTkFont(size=24, weight="bold"))
        title.grid(row=0, column=0, columnspan=6, pady=(0, 20), sticky="w")
        
        self.entries = {}
        self.vars = {}
        
        self.fields = [
            ("date", "Data", "Date Generale"),
            ("driver_name", "Nume Șofer", "Date Generale"),
            ("auto_number", "Număr Auto", "Date Generale"),
            ("auto_type", "Tip Auto", "Date Generale"),
            ("presence", "Prezență", "Date Generale"),
            
            ("km_total", "Km Total", "Detalii Traseu"),
            ("km_empty", "Km Gol", "Detalii Traseu"),
            ("km_loaded", "Km Încărcat", "Detalii Traseu"),
            ("location", "Locație/Rută", "Detalii Traseu"),
            ("route_description", "Descriere Traseu", "Detalii Traseu"),
            ("is_external", "Cursă Externă", "Detalii Traseu"),
            
            ("client", "Client", "Detalii Marfă"),
            ("transport_type", "Tip Transport", "Detalii Marfă"),
            ("special_transport_count", "Nr. Curse Speciale", "Detalii Marfă"),
            ("cargo_type", "Tip Marfă", "Detalii Marfă"),
            ("trailer", "Remorcă", "Detalii Marfă"),
            ("notice_number", "Nr. Aviz", "Detalii Marfă"),
            ("uit_code", "Cod UIT", "Detalii Marfă"),
            
            ("quantity_tons", "Cantitate Tone", "Cantități / Timpi"),
            ("quantity_m3", "Cantitate m³", "Cantități / Timpi"),
            ("trip_count", "Nr. Curse", "Cantități / Timpi"),
            ("pump_hours", "Ore Pompă", "Cantități / Timpi"),
            ("wait_hours", "Ore Staționare", "Cantități / Timpi"),
            
            ("price_per_km", "Preț / Km", "Tarifare"),
            ("price_per_trip", "Preț / Cursă", "Tarifare"),
            ("price_per_ton", "Preț / Tonă", "Tarifare"),
            ("price_per_m3", "Preț / m³", "Tarifare"),
            ("price_per_pump_hour", "Preț Oră Pompă", "Tarifare"),
            ("price_per_wait_hour", "Preț Oră Stațion.", "Tarifare"),
            
            ("daily_allowance", "Diurnă", "Altele"),
            ("bonus", "Premiere", "Altele"),
            ("meal_tickets", "Bonuri Masă", "Altele"),
            
            ("total_price", "Total Preț", "Final")
        ]
        
        self.calc_fields = [
            "km_total", "price_per_km", 
            "trip_count", "price_per_trip",
            "quantity_tons", "price_per_ton",
            "quantity_m3", "price_per_m3",
            "pump_hours", "price_per_pump_hour",
            "wait_hours", "price_per_wait_hour"
        ]
        
        current_section = ""
        row = 1
        col = 0
        
        # Folosim 6 coloane logice (label, entry) * 3 seturi = 6 coloane grid
        for field_id, label_text, section in self.fields:
            if section != current_section:
                current_section = section
                col = 0
                row += 1
                sec_lbl = ctk.CTkLabel(self, text=section, font=ctk.CTkFont(size=16, weight="bold"), text_color="#1f6aa5")
                sec_lbl.grid(row=row, column=0, columnspan=6, pady=(15, 5), sticky="w")
                row += 1
                
            lbl = ctk.CTkLabel(self, text=label_text)
            lbl.grid(row=row, column=col*2, padx=(0, 10), pady=5, sticky="e")
            
            if field_id == "route_description":
                widget = ctk.CTkTextbox(self, height=80, wrap="word")
                widget.grid(row=row, column=col*2+1, padx=(0, 20), pady=5, sticky="ew")
                self.entries[field_id] = widget
            else:
                var = ctk.StringVar()
                self.vars[field_id] = var
                
                if field_id in self.calc_fields:
                    var.trace_add("write", self.calculate_total)
                    
                if field_id in ["km_total", "km_empty"]:
                    var.trace_add("write", self.calculate_loaded_km)
                
                if field_id == "presence":
                    widget = ctk.CTkOptionMenu(self, variable=var, values=["Prezent", "Liber", "Garaj", "Medical", "Concediu"])
                    var.set("Prezent")
                elif field_id == "total_price":
                    widget = ctk.CTkEntry(self, textvariable=var, font=ctk.CTkFont(weight="bold"), text_color="#f39c12")
                elif field_id == "date":
                    widget = DateEntry(self, textvariable=var, date_pattern='yyyy-mm-dd', background='darkblue', foreground='white', borderwidth=2, font=('Helvetica', 12), width=15)
                    var.set(datetime.date.today().strftime("%Y-%m-%d"))
                elif field_id == "driver_name":
                    vals = database.get_distinct_values("driver_name")
                    widget = ctk.CTkComboBox(self, variable=var, values=vals)
                elif field_id == "auto_number":
                    vals = database.get_distinct_values("auto_number")
                    widget = ctk.CTkComboBox(self, variable=var, values=vals)
                elif field_id == "auto_type":
                    vals = database.get_distinct_values("auto_type")
                    widget = ctk.CTkComboBox(self, variable=var, values=vals)
                elif field_id == "client":
                    vals = database.get_distinct_values("client")
                    widget = ctk.CTkComboBox(self, variable=var, values=vals)
                elif field_id == "location":
                    vals = database.get_distinct_values("location")
                    widget = ctk.CTkComboBox(self, variable=var, values=vals)
                elif field_id == "transport_type":
                    base_vals = ["", "AGABARITIC", "ADR", "INSOTIRE"]
                    db_vals = database.get_distinct_values("transport_type")
                    vals = list(dict.fromkeys(base_vals + db_vals))
                    widget = ctk.CTkComboBox(self, variable=var, values=vals)
                elif field_id == "trailer":
                    vals = ["NU", "DA"]
                    widget = ctk.CTkComboBox(self, variable=var, values=vals)
                    var.set("NU")
                elif field_id == "is_external":
                    vals = ["NU", "DA"]
                    widget = ctk.CTkComboBox(self, variable=var, values=vals)
                    var.set("NU")
                elif field_id == "is_external":
                    vals = ["NU", "DA"]
                    widget = ctk.CTkComboBox(self, variable=var, values=vals)
                    var.set("NU")
                elif field_id == "special_transport_count":
                    widget = ctk.CTkEntry(self, textvariable=var)
                    if not self.edit_trip:
                        var.set("1")
                else:
                    widget = ctk.CTkEntry(self, textvariable=var)
                    
                widget.grid(row=row, column=col*2+1, padx=(0, 20), pady=5, sticky="w")
                self.entries[field_id] = widget
            
            col += 1
            if col > 2:
                col = 0
                row += 1
                
        if self.edit_trip:
            for field, lbl_text, sec in self.fields:
                val = self.edit_trip.get(field, "")
                if field == "route_description":
                    self.entries[field].insert("1.0", str(val))
                elif field == "date":
                    try:
                        self.entries[field].set_date(datetime.datetime.strptime(str(val), '%Y-%m-%d').date())
                    except:
                        pass
                else:
                    self.vars[field].set(str(val) if val is not None else "")
                    
        row += 1
        btn_text = "Salvează Modificările" if self.edit_trip else "Salvează Cursă"
        self.save_btn = ctk.CTkButton(self, text=btn_text, command=self.save_trip, fg_color="#27ae60", hover_color="#2ecc71")
        self.save_btn.grid(row=row, column=0, columnspan=6, pady=30)
        
        self.msg_label = ctk.CTkLabel(self, text="", font=ctk.CTkFont(weight="bold"))
        self.msg_label.grid(row=row+1, column=0, columnspan=6)

    def get_float(self, var_name):
        try:
            val = self.vars[var_name].get().strip()
            if not val:
                return 0.0
            return float(val.replace(',', '.'))
        except ValueError:
            return 0.0

    def calculate_loaded_km(self, *args):
        try:
            total = self.get_float("km_total")
            empty = self.get_float("km_empty")
            
            loaded = total - empty
            self.vars["km_loaded"].set(f"{loaded:g}")
        except Exception:
            pass

    def calculate_total(self, *args):
        try:
            total = (
                self.get_float("km_total") * self.get_float("price_per_km") +
                self.get_float("trip_count") * self.get_float("price_per_trip") +
                self.get_float("quantity_tons") * self.get_float("price_per_ton") +
                self.get_float("quantity_m3") * self.get_float("price_per_m3") +
                self.get_float("pump_hours") * self.get_float("price_per_pump_hour") +
                self.get_float("wait_hours") * self.get_float("price_per_wait_hour")
            )
            self.vars["total_price"].set(f"{total:.2f}")
        except Exception as e:
            pass

    def save_trip(self):
        data = {}
        for field, lbl, sec in self.fields:
            if field == "route_description":
                data[field] = self.entries[field].get("1.0", "end-1c").strip()
                continue
                
            val = self.vars[field].get().strip()
            if field in ["presence", "date", "driver_name", "auto_number", "auto_type", "client", "transport_type", "cargo_type", "trailer", "location", "notice_number", "uit_code", "is_external"]:
                data[field] = val
            elif field == "trip_count":
                try:
                    data[field] = int(val) if val else 0
                except ValueError:
                    data[field] = 0
            elif field == "special_transport_count":
                try:
                    data[field] = int(val) if val else 1
                except ValueError:
                    data[field] = 1
            else:
                try:
                    data[field] = float(val.replace(',', '.')) if val else 0.0
                except ValueError:
                    data[field] = 0.0
                    
        if self.edit_trip:
            database.update_trip(self.edit_trip["id"], data)
            messagebox.showinfo("Succes", "Modificările au fost salvate cu succes!")
            if self.on_save_callback:
                self.on_save_callback()
        else:
            database.add_trip(data)
            self.msg_label.configure(text="Cursa a fost salvată cu succes!", text_color="#2ecc71")
            
            # Reset form
            for field, lbl, sec in self.fields:
                if field == "route_description":
                    self.entries[field].delete("1.0", "end")
                else:
                    self.vars[field].set("")
            self.vars["presence"].set("Prezent")
            self.vars["date"].set(datetime.date.today().strftime("%Y-%m-%d"))
            
            self.after(3000, lambda: self.msg_label.configure(text=""))


class CentralizatorFrame(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master, fg_color="transparent")
        
        # Filtre
        filter_frame = ctk.CTkFrame(self)
        filter_frame.pack(fill="x", pady=(0, 20))
        
        self.driver_var = ctk.StringVar()
        self.auto_var = ctk.StringVar()
        self.client_var = ctk.StringVar()
        
        self.use_date_var = ctk.BooleanVar(value=False)
        self.start_date_var = ctk.StringVar(value=datetime.date.today().replace(day=1).strftime("%Y-%m-%d"))
        self.end_date_var = ctk.StringVar(value=datetime.date.today().strftime("%Y-%m-%d"))
        
        row1 = ctk.CTkFrame(filter_frame, fg_color="transparent")
        row1.pack(fill="x", pady=5)
        
        ctk.CTkLabel(row1, text="Șofer:").pack(side="left", padx=5)
        ctk.CTkComboBox(row1, variable=self.driver_var, values=[""] + database.get_distinct_values("driver_name"), width=120).pack(side="left", padx=5)
        
        ctk.CTkLabel(row1, text="Auto:").pack(side="left", padx=5)
        ctk.CTkComboBox(row1, variable=self.auto_var, values=[""] + database.get_distinct_values("auto_number"), width=120).pack(side="left", padx=5)
        
        ctk.CTkLabel(row1, text="Client:").pack(side="left", padx=5)
        ctk.CTkComboBox(row1, variable=self.client_var, values=[""] + database.get_distinct_values("client"), width=120).pack(side="left", padx=5)
        
        row2 = ctk.CTkFrame(filter_frame, fg_color="transparent")
        row2.pack(fill="x", pady=5)
        
        ctk.CTkCheckBox(row2, text="Activează Filtru Dată", variable=self.use_date_var).pack(side="left", padx=10)
        ctk.CTkLabel(row2, text="De la:").pack(side="left", padx=5)
        DateEntry(row2, textvariable=self.start_date_var, date_pattern='yyyy-mm-dd', font=('Helvetica', 12), width=15).pack(side="left", padx=5)
        ctk.CTkLabel(row2, text="Până la:").pack(side="left", padx=5)
        DateEntry(row2, textvariable=self.end_date_var, date_pattern='yyyy-mm-dd', font=('Helvetica', 12), width=15).pack(side="left", padx=5)
        
        ctk.CTkButton(row2, text="Filtrează / Caută", command=self.load_data).pack(side="left", padx=20)
        
        ctk.CTkButton(row2, text="Exportă în Excel", fg_color="#27ae60", hover_color="#2ecc71", command=self.export_excel).pack(side="right", padx=10)
        ctk.CTkButton(row2, text="Șterge Cursă Selectată", fg_color="#c0392b", hover_color="#e74c3c", command=self.delete_selected_trip).pack(side="right", padx=10)
        ctk.CTkButton(row2, text="Editează Cursă", fg_color="#f39c12", hover_color="#e67e22", command=self.open_edit_window).pack(side="right", padx=10)
        
        # Middle Frame (Treeview)
        middle_frame = ctk.CTkFrame(self, fg_color="transparent")
        middle_frame.pack(fill="both", expand=True)
        
        # Bottom Frame (Textbox for description)
        bottom_frame = ctk.CTkFrame(self, fg_color="transparent")
        bottom_frame.pack(fill="x", pady=10)
        
        ctk.CTkLabel(bottom_frame, text="Descriere Traseu Completă:").pack(anchor="w")
        self.desc_textbox = ctk.CTkTextbox(bottom_frame, height=80, wrap="word", state="disabled", font=ctk.CTkFont(size=14))
        self.desc_textbox.pack(fill="x", pady=(5, 0))
        
        # Style Treeview
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", background="#2b2b2b", foreground="white", rowheight=25, fieldbackground="#2b2b2b")
        style.map('Treeview', background=[('selected', '#1f538d')])
        style.configure("Treeview.Heading", background="#565b5e", foreground="white")
        
        # Treeview
        self.tree = ttk.Treeview(middle_frame, columns=("id", "date", "driver", "auto", "client", "total_km", "route", "total_price"), show="headings")
        self.tree.heading("id", text="ID")
        self.tree.heading("date", text="Data")
        self.tree.heading("driver", text="Șofer")
        self.tree.heading("auto", text="Auto")
        self.tree.heading("client", text="Client")
        self.tree.heading("total_km", text="Km Total")
        self.tree.heading("route", text="Descriere Traseu")
        self.tree.heading("total_price", text="Total Preț")
        
        self.tree.column("id", width=50)
        
        self.tree.bind("<<TreeviewSelect>>", self.on_trip_select)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(middle_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        
        self.tree.pack(fill="both", expand=True, side="left")
        scrollbar.pack(fill="y", side="right")
        
        self.load_data()
        
    def load_data(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        driver = self.driver_var.get()
        auto = self.auto_var.get()
        client = self.client_var.get()
        sd = self.start_date_var.get() if self.use_date_var.get() else None
        ed = self.end_date_var.get() if self.use_date_var.get() else None
        
        trips = database.get_trips_with_filters(driver, auto, client, sd, ed)
        for t in trips:
            self.tree.insert("", "end", values=(t['id'], t['date'], t['driver_name'], t['auto_number'], t['client'], t['km_total'], t.get('route_description', ''), t['total_price']))
            
    def on_trip_select(self, event):
        selected_item = self.tree.selection()
        if not selected_item:
            return
        item = self.tree.item(selected_item)
        route_text = item['values'][6]
        
        self.desc_textbox.configure(state="normal")
        self.desc_textbox.delete("1.0", "end")
        self.desc_textbox.insert("1.0", str(route_text) if str(route_text) != "None" else "")
        self.desc_textbox.configure(state="disabled")

    def open_edit_window(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Avertizare", "Te rog să selectezi o cursă din tabel pentru editare!")
            return
            
        item = self.tree.item(selected_item)
        trip_id = item['values'][0]
        trip = database.get_trip_by_id(trip_id)
        if not trip:
            return
            
        EditTripWindow(self, trip, self.load_data)

    def delete_selected_trip(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Avertizare", "Te rog să selectezi o cursă din tabel pentru a o șterge!")
            return
            
        confirm = messagebox.askyesno("Confirmare Ștergere", "Ești sigur că vrei să ștergi această cursă? Acțiunea este ireversibilă.")
        if confirm:
            item = self.tree.item(selected_item)
            trip_id = item['values'][0]
            database.delete_trip(trip_id)
            messagebox.showinfo("Succes", "Cursa a fost ștearsă cu succes!")
            self.load_data()

    def export_excel(self):
        popup = ctk.CTkToplevel(self)
        popup.title("Selectează Coloane pentru Export")
        popup.geometry("400x500")
        popup.grab_set()
        
        col_mapping = {
            'date': 'Data', 'tip_zi': 'Tip Zi', 'driver_name': 'Nume Șofer', 'auto_number': 'Număr Auto',
            'auto_type': 'Tip Auto', 'presence': 'Prezență', 'km_total': 'Km Parcurși Total',
            'km_empty': 'Km pe Gol', 'km_loaded': 'Km Încărcat', 'daily_allowance': 'Diurnă',
            'bonus': 'Premiere', 'meal_tickets': 'Bonuri Masă', 'client': 'Client',
            'transport_type': 'Tip Transport', 'cargo_type': 'Tip Marfă', 'trailer': 'Remorcă',
            'location': 'Locație/Rută', 'route_description': 'Descriere Traseu', 'trip_count': 'Nr. Curse', 'notice_number': 'Nr. Aviz',
            'uit_code': 'Cod UIT', 'quantity_tons': 'Cantitate Tone', 'quantity_m3': 'Cantitate m3',
            'pump_hours': 'Ore Pompă', 'wait_hours': 'Ore Staționare', 'price_per_km': 'Preț / Km',
            'price_per_trip': 'Preț / Cursă', 'price_per_ton': 'Preț / Tonă', 'price_per_m3': 'Preț / m3',
            'price_per_pump_hour': 'Preț Oră Pompă', 'price_per_wait_hour': 'Preț Oră Stațion.',
            'total_price': 'Total Preț'
        }
        
        scroll = ctk.CTkScrollableFrame(popup)
        scroll.pack(fill="both", expand=True, padx=10, pady=10)
        
        checkboxes = {}
        for db_col, ro_name in col_mapping.items():
            var = ctk.BooleanVar(value=True)
            cb = ctk.CTkCheckBox(scroll, text=ro_name, variable=var)
            cb.pack(anchor="w", pady=2)
            checkboxes[db_col] = var
            
        def do_export():
            # Get current filter values to build default filename
            driver = self.driver_var.get()
            sd = self.start_date_var.get() if self.use_date_var.get() else ""
            
            month_year_str = ""
            if sd:
                try:
                    import datetime
                    dt = datetime.datetime.strptime(sd, "%Y-%m-%d")
                    month_year_str = f"{dt.month}-{dt.year}"
                except:
                    month_year_str = sd
            
            if not driver or driver == "Toți":
                default_name = f"Activitate Generala {month_year_str}".strip()
            else:
                default_name = f"Activitate {driver} {month_year_str}".strip()

            filename = filedialog.asksaveasfilename(initialfile=default_name, defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
            if not filename:
                return
                
            selected_db_cols = [col for col, var in checkboxes.items() if var.get()]
            
            driver = self.driver_var.get()
            auto = self.auto_var.get()
            client = self.client_var.get()
            sd = self.start_date_var.get() if self.use_date_var.get() else None
            ed = self.end_date_var.get() if self.use_date_var.get() else None
            trips = database.get_trips_with_filters(driver, auto, client, sd, ed)
            df = pd.DataFrame(trips)
            
            import holidays
            import datetime
            ro_holidays = holidays.RO()
            def get_day_type(date_str):
                try:
                    dt = datetime.datetime.strptime(str(date_str).strip(), "%Y-%m-%d").date()
                    if dt in ro_holidays:
                        return "Sărbătoare legală"
                    elif dt.weekday() >= 5:
                        return "Weekend"
                    else:
                        return "Zi de lucru"
                except Exception:
                    return ""
                    
            if 'date' in df.columns:
                df.insert(2, 'tip_zi', df['date'].apply(get_day_type))
            
            cols_to_keep = [c for c in selected_db_cols if c in df.columns]
            df = df[cols_to_keep]
            
            final_mapping = {c: col_mapping[c] for c in cols_to_keep}
            df.rename(columns=final_mapping, inplace=True)
            
            if 'Data' in df.columns:
                df.sort_values(by='Data', ascending=True, inplace=True)
                
            df.drop(columns=['ID'], errors='ignore', inplace=True)
            df.to_excel(filename, index=False)
            
            if 'Data' in df.columns:
                apply_excel_formatting(filename, date_col_name="Data", tip_zi_col_name="Tip Zi")
            else:
                apply_excel_formatting(filename)
                
            messagebox.showinfo("Succes", "Datele au fost exportate cu succes!")
            popup.destroy()
            
        btn = ctk.CTkButton(popup, text="Generează Raport", command=do_export, fg_color="#27ae60", hover_color="#2ecc71")
        btn.pack(pady=10)

class FinanciarFrame(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master, fg_color="transparent")
        
        form_frame = ctk.CTkFrame(self)
        form_frame.pack(fill="x", pady=(0, 20))
        
        self.vars = {}
        self.entries = {}
        
        # 1. Selecție și Import
        import_frame = ctk.CTkFrame(form_frame, fg_color="transparent")
        import_frame.pack(fill="x", pady=10, padx=10)
        
        ctk.CTkLabel(import_frame, text="1. Parametri Import", font=ctk.CTkFont(weight="bold", size=16), text_color="#1f6aa5").grid(row=0, column=0, columnspan=6, sticky="w", pady=(0,10))
        
        ctk.CTkLabel(import_frame, text="Nume Șofer:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.driver_var = ctk.StringVar()
        ctk.CTkComboBox(import_frame, variable=self.driver_var, values=[""] + database.get_distinct_values("driver_name")).grid(row=1, column=1, padx=5, pady=5, sticky="w")
        
        ctk.CTkLabel(import_frame, text="De la:").grid(row=1, column=2, padx=5, pady=5, sticky="e")
        self.start_date_var = ctk.StringVar(value=datetime.date.today().replace(day=1).strftime("%Y-%m-%d"))
        DateEntry(import_frame, textvariable=self.start_date_var, date_pattern='yyyy-mm-dd', font=('Helvetica', 12), width=15).grid(row=1, column=3, padx=5, pady=5, sticky="w")
        
        ctk.CTkLabel(import_frame, text="Până la:").grid(row=1, column=4, padx=5, pady=5, sticky="e")
        self.end_date_var = ctk.StringVar(value=datetime.date.today().strftime("%Y-%m-%d"))
        DateEntry(import_frame, textvariable=self.end_date_var, date_pattern='yyyy-mm-dd', font=('Helvetica', 12), width=15).grid(row=1, column=5, padx=5, pady=5, sticky="w")
        
        ctk.CTkButton(import_frame, text="Importă Date din Centralizator", command=self.import_data, fg_color="#3498db", hover_color="#2980b9").grid(row=2, column=0, columnspan=6, pady=15)
        
        # Date Importate
        self.imported_rev_var = ctk.StringVar(value="0.0")
        self.imported_km_var = ctk.StringVar(value="0.0")
        self.imported_cars_var = ctk.StringVar(value="")
        
        ctk.CTkLabel(import_frame, text="Venituri Realizate:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
        ctk.CTkEntry(import_frame, textvariable=self.imported_rev_var, state="disabled", font=ctk.CTkFont(weight="bold")).grid(row=3, column=1, padx=5, pady=5, sticky="w")
        
        ctk.CTkLabel(import_frame, text="Km Parcurși Total:").grid(row=3, column=2, padx=5, pady=5, sticky="e")
        ctk.CTkEntry(import_frame, textvariable=self.imported_km_var, state="disabled", font=ctk.CTkFont(weight="bold")).grid(row=3, column=3, padx=5, pady=5, sticky="w")
        
        ctk.CTkLabel(import_frame, text="Mașini Folosite:").grid(row=3, column=4, padx=5, pady=5, sticky="e")
        ctk.CTkEntry(import_frame, textvariable=self.imported_cars_var, state="disabled", width=180).grid(row=3, column=5, padx=5, pady=5, sticky="w")
        
        # 2. Introducere Cheltuieli
        exp_frame = ctk.CTkFrame(form_frame, fg_color="transparent")
        exp_frame.pack(fill="x", pady=10, padx=10)
        
        ctk.CTkLabel(exp_frame, text="2. Cheltuieli (Calcul Live)", font=ctk.CTkFont(weight="bold", size=16), text_color="#1f6aa5").grid(row=0, column=0, columnspan=6, sticky="w", pady=(10,10))
        
        self.exp_fields = [
            ("gross_salary", "Salariul Brut Șofer"),
            ("fixed_auto_expense", "Chelt. Fixă/Auto"),
            ("fuel_expense", "Combustibil"),
            ("maintenance_repair", "Reparații/Întreținere"),
            ("accommodation_parking", "Cazare/Parcări"),
            ("other_expenses", "Alte Cheltuieli")
        ]
        
        row, col = 1, 0
        for f, lbl in self.exp_fields:
            ctk.CTkLabel(exp_frame, text=lbl).grid(row=row, column=col*2, padx=5, pady=5, sticky="e")
            var = ctk.StringVar(value="0")
            self.vars[f] = var
            var.trace_add("write", self.calculate_live)
            ctk.CTkEntry(exp_frame, textvariable=var).grid(row=row, column=col*2+1, padx=5, pady=5, sticky="w")
            
            col += 1
            if col > 2:
                col = 0
                row += 1
                
        # Totaluri Finale
        self.total_exp_var = ctk.StringVar(value="0.0")
        self.net_profit_var = ctk.StringVar(value="0.0")
        
        ctk.CTkLabel(exp_frame, text="Total Cheltuieli:", font=ctk.CTkFont(weight="bold")).grid(row=row, column=0, padx=5, pady=15, sticky="e")
        ctk.CTkEntry(exp_frame, textvariable=self.total_exp_var, state="disabled", font=ctk.CTkFont(weight="bold"), text_color="#e74c3c").grid(row=row, column=1, padx=5, pady=15, sticky="w")
        
        ctk.CTkLabel(exp_frame, text="Profit Net:", font=ctk.CTkFont(weight="bold")).grid(row=row, column=2, padx=5, pady=15, sticky="e")
        ctk.CTkEntry(exp_frame, textvariable=self.net_profit_var, state="disabled", font=ctk.CTkFont(weight="bold", size=16), text_color="#2ecc71").grid(row=row, column=3, padx=5, pady=15, sticky="w")
        
        ctk.CTkButton(form_frame, text="Salvează Raport", command=self.save_report, fg_color="#27ae60", hover_color="#2ecc71", font=ctk.CTkFont(weight="bold")).pack(pady=15)
        
        self.msg_label = ctk.CTkLabel(form_frame, text="", font=ctk.CTkFont(weight="bold"))
        self.msg_label.pack()
        
        # 3. Istoric Rapoarte
        hist_frame = ctk.CTkFrame(self)
        hist_frame.pack(fill="both", expand=True)
        
        btn_frame = ctk.CTkFrame(hist_frame, fg_color="transparent")
        btn_frame.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkButton(btn_frame, text="Exportă în Excel", fg_color="#27ae60", hover_color="#2ecc71", command=self.export_excel).pack(side="right", padx=5)
        ctk.CTkButton(btn_frame, text="Șterge Raport Selectat", fg_color="#c0392b", hover_color="#e74c3c", command=self.delete_selected_report).pack(side="right", padx=5)
        
        self.tree = ttk.Treeview(hist_frame, columns=("id", "sofer", "perioada", "venit", "chelt", "profit"), show="headings")
        self.tree.heading("id", text="ID")
        self.tree.heading("sofer", text="Șofer")
        self.tree.heading("perioada", text="Perioadă")
        self.tree.heading("venit", text="Total Venituri")
        self.tree.heading("chelt", text="Total Cheltuieli")
        self.tree.heading("profit", text="Profit Net")
        
        self.tree.column("id", width=50)
        
        scrollbar = ttk.Scrollbar(hist_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        
        self.tree.pack(fill="both", expand=True, side="left")
        scrollbar.pack(fill="y", side="right")
        
        self.load_history()

    def get_float(self, val_str):
        try:
            val = val_str.strip().replace(',', '.')
            return float(val) if val else 0.0
        except ValueError:
            return 0.0

    def import_data(self):
        driver = self.driver_var.get().strip()
        sd = self.start_date_var.get().strip()
        ed = self.end_date_var.get().strip()
        
        if not driver:
            messagebox.showwarning("Avertizare", "Te rog să completezi numele șoferului pentru import!")
            return
            
        rev, km, cars = database.get_driver_financial_summary(driver, sd, ed)
        
        self.imported_rev_var.set(f"{rev:.2f}")
        self.imported_km_var.set(f"{km:g}")
        self.imported_cars_var.set(cars if cars else "Niciuna")
        
        self.calculate_live()
        
    def calculate_live(self, *args):
        try:
            total_exp = 0.0
            for f, l in self.exp_fields:
                total_exp += self.get_float(self.vars[f].get())
                
            rev = self.get_float(self.imported_rev_var.get())
            profit = rev - total_exp
            
            self.total_exp_var.set(f"{total_exp:.2f}")
            self.net_profit_var.set(f"{profit:.2f}")
        except Exception:
            pass

    def save_report(self):
        driver = self.driver_var.get().strip()
        sd = self.start_date_var.get().strip()
        ed = self.end_date_var.get().strip()
        
        if not driver:
            messagebox.showwarning("Avertizare", "Te rog să completezi numele șoferului!")
            return
            
        data = {
            "driver_name": driver,
            "start_date": sd,
            "end_date": ed,
            "total_revenue": self.get_float(self.imported_rev_var.get()),
            "total_km": self.get_float(self.imported_km_var.get()),
            "cars_used": self.imported_cars_var.get(),
            "total_expenses": self.get_float(self.total_exp_var.get()),
            "net_profit": self.get_float(self.net_profit_var.get()),
            "month_year": f"{sd} - {ed}" # backwards compatibility or info
        }
        
        # Add manual expenses
        for f, l in self.exp_fields:
            data[f] = self.get_float(self.vars[f].get())
            
        database.add_monthly_expense(data)
        self.msg_label.configure(text="Raportul a fost salvat cu succes!", text_color="#2ecc71")
        self.load_history()
        
        # Reset form
        for f, l in self.exp_fields:
            self.vars[f].set("0")
        self.imported_rev_var.set("0.0")
        self.imported_km_var.set("0.0")
        self.imported_cars_var.set("")
        self.calculate_live()
        
        self.after(3000, lambda: self.msg_label.configure(text=""))

    def load_history(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        expenses = database.get_all_monthly_expenses()
        for e in expenses:
            perioada = f"{e.get('start_date', '')} : {e.get('end_date', '')}"
            if not e.get('start_date'): perioada = e.get('month_year', '')
            
            self.tree.insert("", "end", values=(
                e['id'], 
                e['driver_name'], 
                perioada,
                f"{(e.get('total_revenue') or 0):.2f}",
                f"{(e.get('total_expenses') or 0):.2f}", 
                f"{(e.get('net_profit') or 0):.2f}"
            ))

    def delete_selected_report(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Avertizare", "Te rog să selectezi un raport din tabel!")
            return
            
        if messagebox.askyesno("Confirmare", "Ești sigur că vrei să ștergi acest raport? Acțiunea este ireversibilă."):
            item = self.tree.item(selected_item)
            report_id = item['values'][0]
            database.delete_financial_report(report_id)
            messagebox.showinfo("Succes", "Raportul a fost șters cu succes!")
            self.load_history()

    def export_excel(self):
        filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if filename:
            expenses = database.get_all_monthly_expenses()
            data = []
            for e in expenses:
                # Use updated fields, fallback to 0 if old DB entries don't have them
                e['Perioadă'] = f"{e.get('start_date', '')} : {e.get('end_date', '')}"
                data.append(e)
            df = pd.DataFrame(data)
            
            # Map columns cleanly
            col_mapping = {
                'id': 'ID', 'driver_name': 'Nume Șofer', 'Perioadă': 'Perioadă',
                'start_date': 'Dată Început', 'end_date': 'Dată Sfârșit',
                'total_revenue': 'Venituri Realizate', 'total_km': 'Total Km', 'cars_used': 'Mașini Folosite',
                'gross_salary': 'Salariu Brut', 'fixed_auto_expense': 'Chelt. Fixă/Auto',
                'fuel_expense': 'Combustibil', 'maintenance_repair': 'Reparații/Întreținere',
                'accommodation_parking': 'Cazare/Parcări', 'other_expenses': 'Alte Chelt.',
                'total_expenses': 'Total Cheltuieli', 'net_profit': 'Profit Net'
            }
            
            # Filter and rename
            cols_to_keep = [c for c in col_mapping.keys() if c in df.columns]
            df = df[cols_to_keep]
            df.rename(columns=col_mapping, inplace=True)
            if 'Dată Început' in df.columns:
                df.sort_values(by='Dată Început', ascending=True, inplace=True)
            df.drop(columns=['ID'], errors='ignore', inplace=True)
            
            df.to_excel(filename, index=False)
            
            # Since financial reports don't map to a specific "Tip Zi" natively for weekends, we just use format
            # without Tip Zi. It will just auto-size cols and paint headers.
            apply_excel_formatting(filename)
            messagebox.showinfo("Succes", "Raportul financiar a fost exportat cu succes!")


class SalariesFrame(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master, fg_color="transparent")
        
        # Zone A: Setare Profil
        zone_a = ctk.CTkFrame(self)
        zone_a.pack(fill="x", pady=(0, 20))
        
        ctk.CTkLabel(zone_a, text="Setare Profil Șofer", font=ctk.CTkFont(weight="bold", size=18)).pack(anchor="w", pady=(10, 5), padx=10)
        
        grid_a = ctk.CTkFrame(zone_a, fg_color="transparent")
        grid_a.pack(fill="x", padx=10, pady=10)
        
        # ROW 0
        ctk.CTkLabel(grid_a, text="Șofer:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.profile_driver_var = ctk.StringVar()
        driver_cb = ctk.CTkComboBox(grid_a, variable=self.profile_driver_var, values=database.get_distinct_values("driver_name"), command=self.load_profile_data)
        driver_cb.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        
        ctk.CTkLabel(grid_a, text="Salariu Bază:").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        self.base_salary_var = ctk.StringVar(value="0.0")
        ctk.CTkEntry(grid_a, textvariable=self.base_salary_var, width=80).grid(row=0, column=3, padx=5, pady=5, sticky="w")
        
        ctk.CTkLabel(grid_a, text="Tarif / Zi:").grid(row=0, column=4, padx=5, pady=5, sticky="e")
        self.day_rate_var = ctk.StringVar(value="0.0")
        ctk.CTkEntry(grid_a, textvariable=self.day_rate_var, width=80).grid(row=0, column=5, padx=5, pady=5, sticky="w")
        
        ctk.CTkLabel(grid_a, text="Tarif / KM:").grid(row=0, column=6, padx=5, pady=5, sticky="e")
        self.km_rate_var = ctk.StringVar(value="0.0")
        ctk.CTkEntry(grid_a, textvariable=self.km_rate_var, width=80).grid(row=0, column=7, padx=5, pady=5, sticky="w")
        
        # ROW 1 (Standard Rates)
        ctk.CTkLabel(grid_a, text="Tarif Diurnă:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.diurna_rate_var = ctk.StringVar(value="0.0")
        ctk.CTkEntry(grid_a, textvariable=self.diurna_rate_var, width=80).grid(row=1, column=1, padx=5, pady=5, sticky="w")
        
        ctk.CTkLabel(grid_a, text="Tarif Premiere:").grid(row=1, column=2, padx=5, pady=5, sticky="e")
        self.premiere_rate_var = ctk.StringVar(value="0.0")
        ctk.CTkEntry(grid_a, textvariable=self.premiere_rate_var, width=80).grid(row=1, column=3, padx=5, pady=5, sticky="w")
        
        ctk.CTkLabel(grid_a, text="Tarif Bon Masă:").grid(row=1, column=4, padx=5, pady=5, sticky="e")
        self.meal_rate_var = ctk.StringVar(value="0.0")
        ctk.CTkEntry(grid_a, textvariable=self.meal_rate_var, width=80).grid(row=1, column=5, padx=5, pady=5, sticky="w")
        
        ctk.CTkLabel(grid_a, text="Spor Weekend/Zi:").grid(row=1, column=6, padx=5, pady=5, sticky="e")
        self.weekend_rate_var = ctk.StringVar(value="0.0")
        ctk.CTkEntry(grid_a, textvariable=self.weekend_rate_var, width=80).grid(row=1, column=7, padx=5, pady=5, sticky="w")
        
        # ROW 2 (New Transport Bonuses)
        ctk.CTkLabel(grid_a, text="Spor Sărbăt/Zi:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.holiday_rate_var = ctk.StringVar(value="0.0")
        ctk.CTkEntry(grid_a, textvariable=self.holiday_rate_var, width=80).grid(row=2, column=1, padx=5, pady=5, sticky="w")
        
        ctk.CTkLabel(grid_a, text="Tarif Agabaritic:").grid(row=2, column=2, padx=5, pady=5, sticky="e")
        self.agabaritic_rate_var = ctk.StringVar(value="0.0")
        ctk.CTkEntry(grid_a, textvariable=self.agabaritic_rate_var, width=80).grid(row=2, column=3, padx=5, pady=5, sticky="w")
        
        ctk.CTkLabel(grid_a, text="Tarif ADR:").grid(row=2, column=4, padx=5, pady=5, sticky="e")
        self.adr_rate_var = ctk.StringVar(value="0.0")
        ctk.CTkEntry(grid_a, textvariable=self.adr_rate_var, width=80).grid(row=2, column=5, padx=5, pady=5, sticky="w")
        
        ctk.CTkLabel(grid_a, text="Tarif Însoțire:").grid(row=2, column=6, padx=5, pady=5, sticky="e")
        self.insotire_rate_var = ctk.StringVar(value="0.0")
        ctk.CTkEntry(grid_a, textvariable=self.insotire_rate_var, width=80).grid(row=2, column=7, padx=5, pady=5, sticky="w")
        
        # ROW 3 (Trailer & External Rates & Buttons)
        ctk.CTkLabel(grid_a, text="Tarif Remorcă:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
        self.trailer_rate_var = ctk.StringVar(value="0.0")
        ctk.CTkEntry(grid_a, textvariable=self.trailer_rate_var, width=80).grid(row=3, column=1, padx=5, pady=5, sticky="w")
        
        ctk.CTkLabel(grid_a, text="Tarif Extern(€/100km):").grid(row=3, column=2, padx=5, pady=5, sticky="e")
        self.external_rate_var = ctk.StringVar(value="0.0")
        ctk.CTkEntry(grid_a, textvariable=self.external_rate_var, width=80).grid(row=3, column=3, padx=5, pady=5, sticky="w")
        
        ctk.CTkButton(grid_a, text="Salvează Profil", command=self.save_profile).grid(row=3, column=4, padx=20, pady=5, sticky="we")
        ctk.CTkButton(grid_a, text="Șterge Profil", command=self.delete_profile, fg_color="red", hover_color="darkred").grid(row=3, column=5, padx=5, pady=5, sticky="we")
        ctk.CTkButton(grid_a, text="Vezi toate", command=self.show_all_profiles, fg_color="#3498db", hover_color="#2980b9").grid(row=3, column=6, columnspan=2, padx=20, pady=5, sticky="we")
        
        # Zone B: Calcul Salariu Lunar
        zone_b = ctk.CTkFrame(self)
        zone_b.pack(fill="both", expand=True)
        
        ctk.CTkLabel(zone_b, text="Calcul Salariu Lunar", font=ctk.CTkFont(weight="bold", size=18)).pack(anchor="w", pady=(10, 5), padx=10)
        
        grid_b = ctk.CTkFrame(zone_b, fg_color="transparent")
        grid_b.pack(fill="x", padx=10, pady=10)
        
        ctk.CTkLabel(grid_b, text="Șofer:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.calc_driver_var = ctk.StringVar()
        ctk.CTkComboBox(grid_b, variable=self.calc_driver_var, values=database.get_distinct_values("driver_name")).grid(row=0, column=1, padx=5, pady=5, sticky="w")
        
        ctk.CTkLabel(grid_b, text="Luna:").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        import datetime
        self.month_var = ctk.StringVar(value=str(datetime.datetime.now().month))
        ctk.CTkOptionMenu(grid_b, variable=self.month_var, values=[str(i) for i in range(1, 13)], width=60).grid(row=0, column=3, padx=5, pady=5, sticky="w")
        
        ctk.CTkLabel(grid_b, text="An:").grid(row=0, column=4, padx=5, pady=5, sticky="e")
        self.year_var = ctk.StringVar(value=str(datetime.datetime.now().year))
        ctk.CTkOptionMenu(grid_b, variable=self.year_var, values=[str(i) for i in range(2023, 2031)], width=80).grid(row=0, column=5, padx=5, pady=5, sticky="w")
        
        ctk.CTkLabel(grid_b, text="Curs Valutar (RON/€):").grid(row=0, column=6, padx=5, pady=5, sticky="e")
        self.exchange_rate_var = ctk.StringVar(value="5.00")
        ctk.CTkEntry(grid_b, textvariable=self.exchange_rate_var, width=60).grid(row=0, column=7, padx=5, pady=5, sticky="w")
        
        ctk.CTkLabel(grid_b, text="Avans (Lei):").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.advance_var = ctk.StringVar(value="0.0")
        ctk.CTkEntry(grid_b, textvariable=self.advance_var).grid(row=1, column=1, padx=5, pady=5, sticky="w")
        
        ctk.CTkButton(grid_b, text="Calculează Lichidare", command=self.calculate_salary, fg_color="#27ae60", hover_color="#2ecc71").grid(row=1, column=2, columnspan=3, padx=20, pady=5, sticky="w")
        self.export_btn = ctk.CTkButton(grid_b, text="Exportă Raport Excel", command=self.export_excel_report, fg_color="#8e44ad", hover_color="#9b59b6", state="disabled")
        self.export_btn.grid(row=1, column=5, columnspan=3, padx=5, pady=5, sticky="e")
        
        self.result_text = ctk.CTkTextbox(zone_b, font=ctk.CTkFont(family="Courier", size=14))
        self.result_text.pack(fill="both", expand=True, padx=10, pady=10)
        
    def load_profile_data(self, driver_name):
        profile = database.get_driver_profile(driver_name)
        if profile and any(v > 0 for v in profile.values()):
            self.base_salary_var.set(str(profile.get("base_salary", 0.0)))
            self.day_rate_var.set(str(profile.get("day_rate", 0.0)))
            self.km_rate_var.set(str(profile.get("km_rate", 0.0)))
            self.diurna_rate_var.set(str(profile.get("diurna_rate", 0.0)))
            self.premiere_rate_var.set(str(profile.get("premiere_rate", 0.0)))
            self.meal_rate_var.set(str(profile.get("meal_ticket_rate", 0.0)))
            self.weekend_rate_var.set(str(profile.get("weekend_rate", 0.0)))
            self.holiday_rate_var.set(str(profile.get("holiday_rate", 0.0)))
            self.agabaritic_rate_var.set(str(profile.get("agabaritic_rate", 0.0)))
            self.adr_rate_var.set(str(profile.get("adr_rate", 0.0)))
            self.insotire_rate_var.set(str(profile.get("insotire_rate", 0.0)))
            self.trailer_rate_var.set(str(profile.get("trailer_rate", 0.0)))
            self.external_rate_var.set(str(profile.get("external_rate", 0.0)))
        else:
            self.base_salary_var.set("0.0")
            self.day_rate_var.set("0.0")
            self.km_rate_var.set("0.0")
            self.diurna_rate_var.set("0.0")
            self.premiere_rate_var.set("0.0")
            self.meal_rate_var.set("0.0")
            self.weekend_rate_var.set("0.0")
            self.holiday_rate_var.set("0.0")
            self.agabaritic_rate_var.set("0.0")
            self.adr_rate_var.set("0.0")
            self.insotire_rate_var.set("0.0")
            self.trailer_rate_var.set("0.0")
            self.external_rate_var.set("0.0")

    def delete_profile(self):
        name = self.profile_driver_var.get()
        if not name:
            messagebox.showwarning("Eroare", "Selectează un șofer!")
            return
            
        if messagebox.askyesno("Confirmare", f"Ești sigur că vrei să ștergi profilul pentru {name}?"):
            database.delete_driver_profile(name)
            messagebox.showinfo("Succes", f"Profilul pentru {name} a fost șters.")
            self.profile_driver_var.set("")
            self.base_salary_var.set("0.0")
            self.day_rate_var.set("0.0")
            self.km_rate_var.set("0.0")
            self.diurna_rate_var.set("0.0")
            self.premiere_rate_var.set("0.0")
            self.meal_rate_var.set("0.0")
            self.weekend_rate_var.set("0.0")
            self.holiday_rate_var.set("0.0")
            self.agabaritic_rate_var.set("0.0")
            self.adr_rate_var.set("0.0")
            self.insotire_rate_var.set("0.0")
            self.trailer_rate_var.set("0.0")
            self.external_rate_var.set("0.0")

    def show_all_profiles(self):
        popup = ctk.CTkToplevel(self)
        popup.title("Toate Profilurile")
        popup.geometry("1100x400")
        
        txt = ctk.CTkTextbox(popup, font=ctk.CTkFont(family="Courier", size=14))
        txt.pack(fill="both", expand=True, padx=10, pady=10)
        
        profiles = database.get_all_driver_profiles()
        if not profiles:
            txt.insert("end", "Nu există profiluri salvate în baza de date.")
        else:
            for p in profiles:
                txt.insert("end", f"Șofer: {p['driver_name']:<15} | Bază: {p['base_salary']} | Zi: {p['day_rate']} | KM: {p['km_rate']} | Diurnă: {p['diurna_rate']} | Bon: {p['meal_ticket_rate']} | Wknd: {p['weekend_rate']} | Sărb: {p['holiday_rate']} | Agab: {p['agabaritic_rate']} | ADR: {p['adr_rate']} | Îns: {p['insotire_rate']} | Remorcă: {p['trailer_rate']} | Extern: {p.get('external_rate', 0.0)}\n")
        
        txt.configure(state="disabled")
        
    def save_profile(self):
        name = self.profile_driver_var.get()
        if not name:
            messagebox.showwarning("Eroare", "Selectează un șofer!")
            return
        try:
            base = float(self.base_salary_var.get())
            day = float(self.day_rate_var.get())
            km = float(self.km_rate_var.get())
            diurna = float(self.diurna_rate_var.get())
            premiere = float(self.premiere_rate_var.get())
            meal = float(self.meal_rate_var.get())
            weekend = float(self.weekend_rate_var.get())
            holiday = float(self.holiday_rate_var.get())
            agab = float(self.agabaritic_rate_var.get())
            adr = float(self.adr_rate_var.get())
            insot = float(self.insotire_rate_var.get())
            trail = float(self.trailer_rate_var.get())
            ext = float(self.external_rate_var.get())
            
            database.save_driver_profile(name, base, day, km, diurna, premiere, meal, weekend, holiday, agab, adr, insot, trail, ext)
            messagebox.showinfo("Succes", f"Profilul pentru {name} a fost salvat.")
        except ValueError:
            messagebox.showerror("Eroare", "Valorile financiare trebuie să fie numere.")
            
    def calculate_salary(self):
        driver = self.calc_driver_var.get()
        month = self.month_var.get()
        year = self.year_var.get()
        if not driver:
            messagebox.showwarning("Eroare", "Selectează un șofer!")
            return
            
        try: advance = float(self.advance_var.get())
        except ValueError: advance = 0.0
        
        try: exchange_rate = float(self.exchange_rate_var.get())
        except ValueError: exchange_rate = 5.00
        
        profile = database.get_driver_profile(driver)
        base_salary = profile.get("base_salary", 0.0)
        day_rate = profile.get("day_rate", 0.0)
        km_rate = profile.get("km_rate", 0.0)
        diurna_rate = profile.get("diurna_rate", 0.0)
        premiere_rate = profile.get("premiere_rate", 0.0)
        meal_rate = profile.get("meal_ticket_rate", 0.0)
        weekend_rate = profile.get("weekend_rate", 0.0)
        holiday_rate = profile.get("holiday_rate", 0.0)
        agab_rate = profile.get("agabaritic_rate", 0.0)
        adr_rate = profile.get("adr_rate", 0.0)
        insotire_rate = profile.get("insotire_rate", 0.0)
        trailer_rate = profile.get("trailer_rate", 0.0)
        external_rate = profile.get("external_rate", 0.0)
        
        trips = database.get_trips_for_salary(driver, month, year)
        
        from collections import defaultdict
        import datetime
        import holidays
        
        ro_holidays = holidays.RO(years=int(year))
        
        daily_data = defaultdict(lambda: {"internal_km": 0.0, "external_km": 0.0, "clients": set(), "presence": set(), "has_external": False})
        
        qty_allowance = 0.0
        qty_bonus = 0.0
        qty_meal = 0.0
        
        agab_count = 0
        adr_count = 0
        insotire_count = 0
        trailer_count = 0
        
        for t in trips:
            d = t["date"]
            is_ext = str(t.get("is_external", "")).strip().upper() == "DA"
            
            km_for_trip = t["km_total"] or 0.0
            if is_ext:
                daily_data[d]["external_km"] += km_for_trip
                daily_data[d]["has_external"] = True
            else:
                daily_data[d]["internal_km"] += km_for_trip
                
            daily_data[d]["clients"].add(str(t["client"]).upper().strip())
            daily_data[d]["presence"].add(str(t.get("presence", "")).strip())
            
            # The values from Centralizator act as Quantities
            qty_allowance += (t["daily_allowance"] or 0.0)
            qty_bonus += (t["bonus"] or 0.0)
            qty_meal += (t["meal_tickets"] or 0.0)
            
            # Count transport types
            sp_raw = t.get("special_transport_count")
            sp_count = int(sp_raw) if sp_raw is not None and str(sp_raw).strip() != "" else 1
            
            ttype = str(t.get("transport_type", "")).upper().strip()
            if ttype == "AGABARITIC": agab_count += 1 * sp_count
            if ttype == "ADR": adr_count += 1 * sp_count
            if ttype == "INSOTIRE": insotire_count += 1 * sp_count
            
            trailer_val = str(t.get("trailer", "")).upper().strip()
            if trailer_val == "DA": trailer_count += 1 * sp_count
            
        paid_days_count = 0
        internal_km_pay_total = 0.0
        days_log = []
        total_internal_km_month = 0.0
        total_external_km_month = 0.0
        
        weekend_days_worked = 0
        holiday_days_worked = 0
        
        export_data = []
        
        for d, data in sorted(daily_data.items()):
            int_km = data["internal_km"]
            ext_km = data["external_km"]
            total_km = int_km + ext_km
            clients = data["clients"]
            presences = data["presence"]
            has_ext = data["has_external"]
            
            total_internal_km_month += int_km
            total_external_km_month += ext_km
            
            try:
                dt = datetime.datetime.strptime(d.strip(), "%Y-%m-%d").date()
                is_weekend = dt.weekday() >= 5
                is_holiday = dt in ro_holidays
            except Exception:
                is_weekend = False
                is_holiday = False
                
            if total_km > 0 and not ("Liber" in presences or "Garaj" in presences):
                if is_weekend:
                    weekend_days_worked += 1
                if is_holiday:
                    holiday_days_worked += 1
                
            valid_clients = [c for c in clients if c and c not in ("NONE", "")]
            is_only_aguaki = (len(valid_clients) == 1 and "AGUAKI" in valid_clients)
            
            day_pay_val = 0.0
            km_pay_val = 0.0
            ext_pay_euro = 0.0
            status_desc = ""
            
            if "Liber" in presences or "Garaj" in presences:
                status_desc = "Nu s-a lucrat (Liber/Garaj)"
                days_log.append(f"{d} | {total_km:.1f} KM | {status_desc} -> Plată = 0 Lei")
            elif total_km == 0:
                has_non_aguaki = any(c != "AGUAKI" for c in valid_clients)
                if len(valid_clients) > 0 and has_non_aguaki:
                    paid_days_count += 1
                    day_pay_val = day_rate
                    status_desc = "Staționare Client (!=AGUAKI)"
                    days_log.append(f"{d} | {total_km:.1f} KM | {status_desc} -> Plată la zi: {day_pay_val:.2f} Lei")
                else:
                    status_desc = "Nu s-a lucrat (KM=0)"
                    days_log.append(f"{d} | {total_km:.1f} KM | {status_desc} -> Plată = 0 Lei")
            elif has_ext:
                # Variant A logic: No day rate. External km paid in Euro, Internal km paid at km_rate
                ext_pay_euro = (ext_km * external_rate) / 100
                if int_km > 0:
                    km_pay_val = int_km * km_rate
                    internal_km_pay_total += km_pay_val
                
                status_desc = f"Mixt/Extern (Anulare plată zi)"
                days_log.append(f"{d} | Ext:{ext_km:.1f}KM Int:{int_km:.1f}KM | -> {ext_pay_euro:.2f} € / {km_pay_val:.2f} Lei")
            else:
                # Normal Internal Logic
                if km_rate > 0 and int_km > 350:
                    km_pay_val = int_km * km_rate
                    internal_km_pay_total += km_pay_val
                    status_desc = "Regula >350KM"
                    days_log.append(f"{d} | {int_km:.1f} KM | {status_desc} -> Bani din KM: {km_pay_val:.2f} Lei")
                else:
                    if is_only_aguaki:
                        status_desc = "Doar AGUAKI"
                        days_log.append(f"{d} | {int_km:.1f} KM | {status_desc} -> Plată la zi = 0 Lei")
                    else:
                        paid_days_count += 1
                        day_pay_val = day_rate
                        status_desc = "Zi internă (Eligibil day_rate)"
                        days_log.append(f"{d} | {int_km:.1f} KM | {status_desc} -> Plată la zi: {day_pay_val:.2f} Lei")
            
            export_data.append({
                "Dată": d,
                "Total KM": total_km,
                "Int KM": int_km,
                "Ext KM": ext_km,
                "Status/Clienți": status_desc + " (" + ", ".join(clients) + ")",
                "Plată Zi (Lei)": day_pay_val,
                "Plată KM (Lei)": km_pay_val,
                "Plată Ext (€)": ext_pay_euro
            })
                    
        days_pay_total = paid_days_count * day_rate
        total_external_euro = (total_external_km_month * external_rate) / 100
        total_external_ron = total_external_euro * exchange_rate
        
        weekend_pay_total = weekend_days_worked * weekend_rate
        holiday_pay_total = holiday_days_worked * holiday_rate
        
        total_allowance = qty_allowance * diurna_rate
        total_bonus = qty_bonus * premiere_rate
        total_meal = qty_meal * meal_rate
        
        agab_pay_total = agab_count * agab_rate
        adr_pay_total = adr_count * adr_rate
        insotire_pay_total = insotire_count * insotire_rate
        trailer_pay_total = trailer_count * trailer_rate
        
        total_salary = (base_salary + days_pay_total + internal_km_pay_total + total_external_ron + weekend_pay_total + holiday_pay_total + 
                       total_allowance + total_bonus + total_meal + 
                       agab_pay_total + adr_pay_total + insotire_pay_total + trailer_pay_total)
        rest_plata = total_salary - advance - total_meal
        
        import pandas as pd
        self.current_salary_report_df = pd.DataFrame(export_data)
        self.summary_data = {
            "driver": driver,
            "month": month,
            "year": year,
            "base_salary": base_salary,
            "total_internal_km_month": total_internal_km_month,
            "total_external_km_month": total_external_km_month,
            "total_external_euro": total_external_euro,
            "exchange_rate": exchange_rate,
            "total_external_ron": total_external_ron,
            "weekend_days_worked": weekend_days_worked,
            "weekend_pay_total": weekend_pay_total,
            "holiday_days_worked": holiday_days_worked,
            "holiday_pay_total": holiday_pay_total,
            "qty_allowance": qty_allowance,
            "total_allowance": total_allowance,
            "qty_bonus": qty_bonus,
            "total_bonus": total_bonus,
            "qty_meal": qty_meal,
            "total_meal": total_meal,
            "agab_count": agab_count,
            "agab_pay_total": agab_pay_total,
            "adr_count": adr_count,
            "adr_pay_total": adr_pay_total,
            "insotire_count": insotire_count,
            "insotire_pay_total": insotire_pay_total,
            "trailer_count": trailer_count,
            "trailer_pay_total": trailer_pay_total,
            "total_salary": total_salary,
            "advance": advance,
            "rest_plata": rest_plata
        }
        
        report = []
        report.append(f"=== RAPORT SALARIZARE: {driver} ({month}/{year}) ===")
        report.append(f"Salariu de bază: {base_salary:.2f} Lei")
        report.append("-" * 50)
        report.append(f"Total KM Interni: {total_internal_km_month:.1f} KM")
        report.append(f"Zile Plătite la Zi (<=350km intern): {paid_days_count} x {day_rate:.2f} = {days_pay_total:.2f} Lei")
        report.append(f"Bani KM Intern (>350km sau Mixt): {internal_km_pay_total:.2f} Lei")
        if total_external_km_month > 0:
            report.append("-" * 50)
            report.append(f"Total KM Externi: {total_external_km_month:.1f} KM")
            report.append(f"Bani Extern: {total_external_euro:.2f} Euro (Curs: {exchange_rate}) -> {total_external_ron:.2f} RON")
        report.append("-" * 50)
        
        if weekend_pay_total > 0: report.append(f"Bani din Weekend ({weekend_days_worked} zile): {weekend_pay_total:.2f} Lei")
        if holiday_pay_total > 0: report.append(f"Bani din Sărbători ({holiday_days_worked} zile): {holiday_pay_total:.2f} Lei")
        if agab_pay_total > 0: report.append(f"Spor Agabaritic ({agab_count} curse): {agab_pay_total:.2f} Lei")
        if adr_pay_total > 0: report.append(f"Spor ADR ({adr_count} curse): {adr_pay_total:.2f} Lei")
        if insotire_pay_total > 0: report.append(f"Spor Însoțire ({insotire_count} curse): {insotire_pay_total:.2f} Lei")
        if trailer_pay_total > 0: report.append(f"Spor Remorcă ({trailer_count} curse): {trailer_pay_total:.2f} Lei")
        
        report.append("-" * 50)
        report.append(f"Diurnă ({qty_allowance} x {diurna_rate:.2f}): {total_allowance:.2f} Lei")
        report.append(f"Premiere ({qty_bonus} x {premiere_rate:.2f}): {total_bonus:.2f} Lei")
        if total_meal > 0: report.append(f"Bani Bonuri Masă ({qty_meal} x {meal_rate:.2f}): {total_meal:.2f} Lei")
        report.append("-" * 50)
        report.append(f"TOTAL SALARIU (Brut): {total_salary:.2f} Lei")
        if advance > 0: report.append(f"Deducere Avans: -{advance:.2f} Lei")
        if total_meal > 0: report.append(f"Deducere Bonuri Masă: -{total_meal:.2f} Lei")
        report.append(f"REST DE PLATĂ (Lichidare): {rest_plata:.2f} Lei")
        report.append("=" * 50)
        report.append("Detalii pe Zile:")
        report.extend(days_log)
        
        self.result_text.delete("1.0", "end")
        self.result_text.insert("end", "\n".join(report))
        self.export_btn.configure(state="normal")

    def export_excel_report(self):
        if not hasattr(self, 'current_salary_report_df') or self.current_salary_report_df is None or self.current_salary_report_df.empty:
            messagebox.showwarning("Eroare", "Calculează întâi lichidarea!")
            return
            
        default_filename = f"Salarizare_{self.summary_data['driver'].replace(' ', '_')}_{self.summary_data['month']}_{self.summary_data['year']}.xlsx"
        filename = filedialog.asksaveasfilename(initialfile=default_filename, defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if not filename:
            return
            
        df = self.current_salary_report_df.copy()
        
        summary_rows = [
            {"Dată": "", "Total KM": "", "Int KM": "", "Ext KM": "", "Status/Clienți": "", "Plată Zi (Lei)": "", "Plată KM (Lei)": "", "Plată Ext (€)": ""},
            {"Dată": "REZUMAT", "Total KM": "", "Int KM": "", "Ext KM": "", "Status/Clienți": "", "Plată Zi (Lei)": "", "Plată KM (Lei)": "", "Plată Ext (€)": ""},
            {"Dată": "Salariu de bază", "Total KM": self.summary_data['base_salary'], "Int KM": "", "Ext KM": "", "Status/Clienți": "", "Plată Zi (Lei)": "", "Plată KM (Lei)": "", "Plată Ext (€)": ""},
            {"Dată": "Total KM Interni", "Total KM": self.summary_data['total_internal_km_month'], "Int KM": "", "Ext KM": "", "Status/Clienți": "", "Plată Zi (Lei)": "", "Plată KM (Lei)": "", "Plată Ext (€)": ""},
            {"Dată": "Total KM Externi", "Total KM": self.summary_data['total_external_km_month'], "Int KM": "", "Ext KM": "", "Status/Clienți": "", "Plată Zi (Lei)": "", "Plată KM (Lei)": "", "Plată Ext (€)": ""}
        ]
        
        if self.summary_data['total_external_km_month'] > 0:
            summary_rows.append({"Dată": "Bani Extern (€ -> RON)", "Total KM": self.summary_data['total_external_ron'], "Int KM": "", "Ext KM": "", "Status/Clienți": f"({self.summary_data['total_external_euro']:.2f} € x {self.summary_data['exchange_rate']})", "Plată Zi (Lei)": "", "Plată KM (Lei)": "", "Plată Ext (€)": ""})
            
        summary_rows.extend([
            {"Dată": "Bani din Weekend", "Total KM": self.summary_data['weekend_pay_total'], "Int KM": "", "Ext KM": "", "Status/Clienți": f"({self.summary_data['weekend_days_worked']} zile)", "Plată Zi (Lei)": "", "Plată KM (Lei)": "", "Plată Ext (€)": ""},
            {"Dată": "Bani din Sărbători", "Total KM": self.summary_data['holiday_pay_total'], "Int KM": "", "Ext KM": "", "Status/Clienți": f"({self.summary_data['holiday_days_worked']} zile)", "Plată Zi (Lei)": "", "Plată KM (Lei)": "", "Plată Ext (€)": ""}
        ])
        
        if self.summary_data['agab_pay_total'] > 0:
            summary_rows.append({"Dată": "Spor Agabaritic", "Total KM": self.summary_data['agab_pay_total'], "Int KM": "", "Ext KM": "", "Status/Clienți": f"({self.summary_data['agab_count']} curse)", "Plată Zi (Lei)": "", "Plată KM (Lei)": "", "Plată Ext (€)": ""})
        if self.summary_data['adr_pay_total'] > 0:
            summary_rows.append({"Dată": "Spor ADR", "Total KM": self.summary_data['adr_pay_total'], "Int KM": "", "Ext KM": "", "Status/Clienți": f"({self.summary_data['adr_count']} curse)", "Plată Zi (Lei)": "", "Plată KM (Lei)": "", "Plată Ext (€)": ""})
        if self.summary_data['insotire_pay_total'] > 0:
            summary_rows.append({"Dată": "Spor Însoțire", "Total KM": self.summary_data['insotire_pay_total'], "Int KM": "", "Ext KM": "", "Status/Clienți": f"({self.summary_data['insotire_count']} curse)", "Plată Zi (Lei)": "", "Plată KM (Lei)": "", "Plată Ext (€)": ""})
        if self.summary_data['trailer_pay_total'] > 0:
            summary_rows.append({"Dată": "Spor Remorcă", "Total KM": self.summary_data['trailer_pay_total'], "Int KM": "", "Ext KM": "", "Status/Clienți": f"({self.summary_data['trailer_count']} curse)", "Plată Zi (Lei)": "", "Plată KM (Lei)": "", "Plată Ext (€)": ""})
            
        summary_rows.extend([
            {"Dată": "Total Diurnă", "Total KM": self.summary_data['total_allowance'], "Int KM": "", "Ext KM": "", "Status/Clienți": f"(cant: {self.summary_data['qty_allowance']})", "Plată Zi (Lei)": "", "Plată KM (Lei)": "", "Plată Ext (€)": ""},
            {"Dată": "Total Premiere", "Total KM": self.summary_data['total_bonus'], "Int KM": "", "Ext KM": "", "Status/Clienți": f"(cant: {self.summary_data['qty_bonus']})", "Plată Zi (Lei)": "", "Plată KM (Lei)": "", "Plată Ext (€)": ""},
            {"Dată": "TOTAL SALARIU (Brut)", "Total KM": self.summary_data['total_salary'], "Int KM": "", "Ext KM": "", "Status/Clienți": "", "Plată Zi (Lei)": "", "Plată KM (Lei)": "", "Plată Ext (€)": ""},
            {"Dată": "Deducere Avans", "Total KM": -self.summary_data['advance'], "Int KM": "", "Ext KM": "", "Status/Clienți": "", "Plată Zi (Lei)": "", "Plată KM (Lei)": "", "Plată Ext (€)": ""},
            {"Dată": "Deducere Bonuri Masă", "Total KM": -self.summary_data['total_meal'], "Int KM": "", "Ext KM": "", "Status/Clienți": f"(cant: {self.summary_data['qty_meal']})", "Plată Zi (Lei)": "", "Plată KM (Lei)": "", "Plată Ext (€)": ""},
            {"Dată": "REST DE PLATĂ", "Total KM": self.summary_data['rest_plata'], "Int KM": "", "Ext KM": "", "Status/Clienți": "", "Plată Zi (Lei)": "", "Plată KM (Lei)": "", "Plată Ext (€)": ""}
        ])
        
        import pandas as pd
        summary_df = pd.DataFrame(summary_rows)
        final_df = pd.concat([df, summary_df], ignore_index=True)
        
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            final_df.to_excel(writer, index=False, startrow=1)
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']
            worksheet.cell(row=1, column=1, value=f"Raport Salarizare - {self.summary_data['driver']} - {self.summary_data['month']}/{self.summary_data['year']}")
            worksheet.merge_cells('A1:H1')
            
        try:
            apply_excel_formatting(filename, date_col_name=None, tip_zi_col_name=None)
        except Exception as e:
            print(f"Formatting warning: {e}")
            
        messagebox.showinfo("Succes", "Raportul a fost exportat în Excel cu succes!")
class MainDashboardFrame(ctk.CTkFrame):
    def __init__(self, master, role, logout_callback):
        super().__init__(master)
        self.role = role
        self.logout_callback = logout_callback

        # Configurăm un grid 1x2 (Meniu în stânga, Conținut în dreapta)
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        # Meniul lateral (Sidebar)
        self.sidebar_frame = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(5, weight=1)

        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text="Transport DB", font=ctk.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))

        self.role_label = ctk.CTkLabel(self.sidebar_frame, text=f"Logat ca: {self.role}", font=ctk.CTkFont(size=12))
        self.role_label.grid(row=1, column=0, padx=20, pady=(0, 20))

        # Butoane meniu
        self.nav_curse_btn = ctk.CTkButton(self.sidebar_frame, text="Curse Noi", command=self.show_curse)
        self.nav_curse_btn.grid(row=2, column=0, padx=20, pady=10)

        self.nav_centralizator_btn = ctk.CTkButton(self.sidebar_frame, text="Centralizator", command=self.show_centralizator)
        self.nav_centralizator_btn.grid(row=3, column=0, padx=20, pady=10)

        self.nav_financiar_btn = ctk.CTkButton(self.sidebar_frame, text="Raport Financiar", command=self.show_financiar)
        self.nav_financiar_btn.grid(row=4, column=0, padx=20, pady=10)
        
        self.nav_salarii_btn = ctk.CTkButton(self.sidebar_frame, text="Salarizare", command=self.show_salarii)
        self.nav_salarii_btn.grid(row=5, column=0, padx=20, pady=10)

        # Logica de protecție a butoanelor în funcție de rol
        if self.role != "Admin":
            self.nav_financiar_btn.configure(state="disabled", fg_color="transparent")
            # Dacă vrei ca butonul să fie total ascuns pentru Operator, poți folosi:
            # self.nav_financiar_btn.grid_remove()

        # Buton Logout (jos)
        self.logout_btn = ctk.CTkButton(self.sidebar_frame, text="Logout", command=self.logout, fg_color="darkred", hover_color="red")
        self.logout_btn.grid(row=7, column=0, padx=20, pady=20, sticky="s")

        # Zona de conținut (Dreapta)
        self.content_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.content_frame.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
        
        self.current_content = None
        
        # Selectăm "Curse Noi" by default la logare
        self.show_curse()

    def set_content(self, title):
        # Distruge widget-urile din conținutul vechi (un reset temporar pentru Placeholder)
        if self.current_content:
            self.current_content.destroy()
            
        self.current_content = ctk.CTkLabel(self.content_frame, text=f"Secțiunea: {title}\n(În curs de dezvoltare...)", font=ctk.CTkFont(size=24))
        self.current_content.place(relx=0.5, rely=0.5, anchor="center")

    def show_curse(self):
        if self.current_content:
            self.current_content.destroy()
            
        self.current_content = DailyTripsFormFrame(self.content_frame)
        self.current_content.pack(fill="both", expand=True)

    def show_centralizator(self):
        if self.current_content:
            self.current_content.destroy()
        self.current_content = CentralizatorFrame(self.content_frame)
        self.current_content.pack(fill="both", expand=True)

    def show_financiar(self):
        if self.current_content:
            self.current_content.destroy()
        self.current_content = FinanciarFrame(self.content_frame)
        self.current_content.pack(fill="both", expand=True)
        
    def show_salarii(self):
        if self.current_content:
            self.current_content.destroy()
        self.current_content = SalariesFrame(self.content_frame)
        self.current_content.pack(fill="both", expand=True)

    def logout(self):
        self.logout_callback()

class EditTripWindow(ctk.CTkToplevel):
    def __init__(self, master, trip_data, on_save_callback):
        super().__init__(master)
        self.title(f"Editează Cursa ID {trip_data['id']}")
        self.geometry("950x700")
        
        self.lift()
        self.attributes("-topmost", True)
        self.after(10, lambda: self.attributes("-topmost", False))
        
        def close_and_refresh():
            on_save_callback()
            self.destroy()
            
        form = TripFormFrame(self, edit_trip=trip_data, on_save_callback=close_and_refresh)
        form.pack(fill="both", expand=True)

if __name__ == "__main__":
    app = App()
    app.mainloop()
