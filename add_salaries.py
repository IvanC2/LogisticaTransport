import sys

with open('main.py', 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Define SalariesFrame
salaries_class = """
class SalariesFrame(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master, fg_color="transparent")
        
        # Zone A: Setare Profil
        zone_a = ctk.CTkFrame(self)
        zone_a.pack(fill="x", pady=(0, 20))
        
        ctk.CTkLabel(zone_a, text="Setare Profil Șofer", font=ctk.CTkFont(weight="bold", size=18)).pack(anchor="w", pady=(10, 5), padx=10)
        
        grid_a = ctk.CTkFrame(zone_a, fg_color="transparent")
        grid_a.pack(fill="x", padx=10, pady=10)
        
        ctk.CTkLabel(grid_a, text="Șofer:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.profile_driver_var = ctk.StringVar()
        driver_cb = ctk.CTkComboBox(grid_a, variable=self.profile_driver_var, values=database.get_distinct_values("driver_name"), command=self.load_profile)
        driver_cb.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        
        ctk.CTkLabel(grid_a, text="Salariu Bază (Lei):").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        self.base_salary_var = ctk.StringVar(value="0.0")
        ctk.CTkEntry(grid_a, textvariable=self.base_salary_var).grid(row=0, column=3, padx=5, pady=5, sticky="w")
        
        ctk.CTkLabel(grid_a, text="Tarif / Zi (Lei):").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.day_rate_var = ctk.StringVar(value="0.0")
        ctk.CTkEntry(grid_a, textvariable=self.day_rate_var).grid(row=1, column=1, padx=5, pady=5, sticky="w")
        
        ctk.CTkLabel(grid_a, text="Tarif / KM (Lei):").grid(row=1, column=2, padx=5, pady=5, sticky="e")
        self.km_rate_var = ctk.StringVar(value="0.0")
        ctk.CTkEntry(grid_a, textvariable=self.km_rate_var).grid(row=1, column=3, padx=5, pady=5, sticky="w")
        
        ctk.CTkButton(grid_a, text="Salvează Profil", command=self.save_profile).grid(row=1, column=4, padx=20, pady=5)
        
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
        self.month_var = ctk.StringVar(value=str(datetime.datetime.now().month))
        ctk.CTkOptionMenu(grid_b, variable=self.month_var, values=[str(i) for i in range(1, 13)]).grid(row=0, column=3, padx=5, pady=5, sticky="w")
        
        ctk.CTkLabel(grid_b, text="An:").grid(row=0, column=4, padx=5, pady=5, sticky="e")
        self.year_var = ctk.StringVar(value=str(datetime.datetime.now().year))
        ctk.CTkOptionMenu(grid_b, variable=self.year_var, values=[str(i) for i in range(2023, 2031)]).grid(row=0, column=5, padx=5, pady=5, sticky="w")
        
        ctk.CTkLabel(grid_b, text="Avans (Lei):").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.advance_var = ctk.StringVar(value="0.0")
        ctk.CTkEntry(grid_b, textvariable=self.advance_var).grid(row=1, column=1, padx=5, pady=5, sticky="w")
        
        ctk.CTkButton(grid_b, text="Calculează Lichidare", command=self.calculate_salary, fg_color="#27ae60", hover_color="#2ecc71").grid(row=1, column=2, columnspan=2, padx=20, pady=5)
        
        self.result_text = ctk.CTkTextbox(zone_b, font=ctk.CTkFont(family="Courier", size=14))
        self.result_text.pack(fill="both", expand=True, padx=10, pady=10)
        
    def load_profile(self, driver_name):
        profile = database.get_driver_profile(driver_name)
        self.base_salary_var.set(str(profile.get("base_salary", 0.0)))
        self.day_rate_var.set(str(profile.get("day_rate", 0.0)))
        self.km_rate_var.set(str(profile.get("km_rate", 0.0)))
        
    def save_profile(self):
        name = self.profile_driver_var.get()
        if not name:
            messagebox.showwarning("Eroare", "Selectează un șofer!")
            return
        try:
            base = float(self.base_salary_var.get())
            day = float(self.day_rate_var.get())
            km = float(self.km_rate_var.get())
            database.save_driver_profile(name, base, day, km)
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
        
        profile = database.get_driver_profile(driver)
        base_salary = profile.get("base_salary", 0.0)
        day_rate = profile.get("day_rate", 0.0)
        km_rate = profile.get("km_rate", 0.0)
        
        trips = database.get_trips_for_salary(driver, month, year)
        
        # Group by date
        from collections import defaultdict
        daily_data = defaultdict(lambda: {"km": 0.0, "clients": set()})
        
        total_allowance = 0.0
        total_bonus = 0.0
        total_meal = 0.0
        
        for t in trips:
            d = t["date"]
            daily_data[d]["km"] += (t["km_total"] or 0.0)
            daily_data[d]["clients"].add(str(t["client"]).upper().strip())
            total_allowance += (t["daily_allowance"] or 0.0)
            total_bonus += (t["bonus"] or 0.0)
            total_meal += (t["meal_tickets"] or 0.0)
            
        paid_days_count = 0
        km_pay_total = 0.0
        days_log = []
        total_km_month = 0.0
        
        for d, data in sorted(daily_data.items()):
            km = data["km"]
            clients = data["clients"]
            total_km_month += km
            
            is_only_aguaki = (len(clients) == 1 and "AGUAKI" in clients)
            
            if km > 350:
                pay = km * km_rate
                km_pay_total += pay
                days_log.append(f"{d} | {km:.1f} KM | Regula >350KM -> Bani din KM: {pay:.2f} Lei")
            else:
                if is_only_aguaki:
                    days_log.append(f"{d} | {km:.1f} KM | Doar AGUAKI -> Plată la zi = 0 Lei")
                else:
                    paid_days_count += 1
                    days_log.append(f"{d} | {km:.1f} KM | Zi normală -> Plată la zi: {day_rate:.2f} Lei")
                    
        days_pay_total = paid_days_count * day_rate
        total_salary = base_salary + days_pay_total + km_pay_total + total_allowance + total_bonus + total_meal
        rest_plata = total_salary - advance
        
        report = []
        report.append(f"=== RAPORT SALARIZARE: {driver} ({month}/{year}) ===")
        report.append(f"Salariu de bază: {base_salary:.2f} Lei")
        report.append("-" * 50)
        report.append(f"Total KM Parcurși (Lună): {total_km_month:.1f} KM")
        report.append(f"Zile Plătite la Zi (<=350km): {paid_days_count} x {day_rate:.2f} = {days_pay_total:.2f} Lei")
        report.append(f"Suma din KM (>350km): {km_pay_total:.2f} Lei")
        report.append("-" * 50)
        report.append(f"Diurnă: {total_allowance:.2f} Lei")
        report.append(f"Premiere: {total_bonus:.2f} Lei")
        report.append(f"Bonuri Masă: {total_meal:.2f} Lei")
        report.append("-" * 50)
        report.append(f"TOTAL SALARIU: {total_salary:.2f} Lei")
        report.append(f"Avans reținut: {advance:.2f} Lei")
        report.append(f"REST DE PLATĂ (Lichidare): {rest_plata:.2f} Lei")
        report.append("=" * 50)
        report.append("Detalii pe Zile:")
        report.extend(days_log)
        
        self.result_text.delete("1.0", "end")
        self.result_text.insert("end", "\\n".join(report))

"""

idx = content.find("class MainDashboardFrame(ctk.CTkFrame):")
if "class SalariesFrame" not in content:
    content = content[:idx] + salaries_class + "\n" + content[idx:]


# 2. Add button to Sidebar
old_btn = """        self.nav_financiar_btn = ctk.CTkButton(self.sidebar_frame, text="Raport Financiar", command=self.show_financiar)
        self.nav_financiar_btn.grid(row=4, column=0, padx=20, pady=10)

        # Logica de protecție a butoanelor în funcție de rol"""

new_btn = """        self.nav_financiar_btn = ctk.CTkButton(self.sidebar_frame, text="Raport Financiar", command=self.show_financiar)
        self.nav_financiar_btn.grid(row=4, column=0, padx=20, pady=10)
        
        self.nav_salarii_btn = ctk.CTkButton(self.sidebar_frame, text="Salarizare", command=self.show_salarii)
        self.nav_salarii_btn.grid(row=5, column=0, padx=20, pady=10)

        # Logica de protecție a butoanelor în funcție de rol"""
content = content.replace(old_btn, new_btn)


# 3. Add show_salarii method to MainDashboardFrame
old_show = """    def show_financiar(self):
        if self.current_content:
            self.current_content.destroy()
        self.current_content = FinanciarFrame(self.content_frame)
        self.current_content.pack(fill="both", expand=True)"""

new_show = """    def show_financiar(self):
        if self.current_content:
            self.current_content.destroy()
        self.current_content = FinanciarFrame(self.content_frame)
        self.current_content.pack(fill="both", expand=True)
        
    def show_salarii(self):
        if self.current_content:
            self.current_content.destroy()
        self.current_content = SalariesFrame(self.content_frame)
        self.current_content.pack(fill="both", expand=True)"""
content = content.replace(old_show, new_show)

# Adjust row configuration for sidebar if needed.
# row 6 was logout, we can move it to row 7.
content = content.replace("self.logout_btn.grid(row=6", "self.logout_btn.grid(row=7")


with open('main.py', 'w', encoding='utf-8') as f:
    f.write(content)
