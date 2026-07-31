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
                widget = ctk.CTkOptionMenu(fixed_frame, variable=var, values=["Prezent", "Absent", "Concediu", "Medical", "Liber"])
                var.set("Prezent")
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
            ("client", "Client", "Marfă"), ("transport_type", "Tip Transport", "Marfă"), ("cargo_type", "Tip Marfă", "Marfă"),
            ("trailer", "Remorcă", "Marfă"), ("notice_number", "Nr. Aviz", "Marfă"), ("uit_code", "Cod UIT", "Marfă"),
            ("location", "Locație/Rută", "Traseu"), ("route_description", "Descriere Traseu", "Traseu"),
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
                if field in ["presence", "date", "driver_name", "auto_number", "auto_type", "client", "transport_type", "cargo_type", "trailer", "location", "route_description", "notice_number", "uit_code"]:
                    data[field] = val
                elif field == "trip_count":
                    try: data[field] = int(val) if val else 0
                    except ValueError: data[field] = 0
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
            
            ("client", "Client", "Detalii Marfă"),
            ("transport_type", "Tip Transport", "Detalii Marfă"),
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
                    widget = ctk.CTkOptionMenu(self, variable=var, values=["Prezent", "Absent", "Concediu", "Medical", "Liber"])
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
            if field in ["presence", "date", "driver_name", "auto_number", "auto_type", "client", "transport_type", "cargo_type", "trailer", "location", "notice_number", "uit_code"]:
                data[field] = val
            elif field == "trip_count":
                try:
                    data[field] = int(val) if val else 0
                except ValueError:
                    data[field] = 0
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
            filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
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

        # Logica de protecție a butoanelor în funcție de rol
        if self.role != "Admin":
            self.nav_financiar_btn.configure(state="disabled", fg_color="transparent")
            # Dacă vrei ca butonul să fie total ascuns pentru Operator, poți folosi:
            # self.nav_financiar_btn.grid_remove()

        # Buton Logout (jos)
        self.logout_btn = ctk.CTkButton(self.sidebar_frame, text="Logout", command=self.logout, fg_color="darkred", hover_color="red")
        self.logout_btn.grid(row=6, column=0, padx=20, pady=20, sticky="s")

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
