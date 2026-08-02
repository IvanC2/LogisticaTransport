import sys

with open('main.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Fix TripFormFrame fields
old_trip_fields = """            ("km_total", "Km Total", "Detalii Traseu"),
            ("km_empty", "Km Gol", "Detalii Traseu"),
            ("km_loaded", "Km Încărcat", "Detalii Traseu"),
            ("location", "Locație/Rută", "Detalii Traseu"),
            ("route_description", "Descriere Traseu", "Detalii Traseu"),"""

new_trip_fields = """            ("km_total", "Km Total", "Detalii Traseu"),
            ("km_empty", "Km Gol", "Detalii Traseu"),
            ("km_loaded", "Km Încărcat", "Detalii Traseu"),
            ("location", "Locație/Rută", "Detalii Traseu"),
            ("route_description", "Descriere Traseu", "Detalii Traseu"),
            ("is_external", "Cursă Externă", "Detalii Traseu"),"""
content = content.replace(old_trip_fields, new_trip_fields)

# Check if is_external renderer was added, if not, add it
if 'elif field_id == "is_external":' not in content[content.find("class TripFormFrame("):content.find("def get_float(self, var_name):")]:
    old_trip_renderer = """                elif field_id == "trailer":
                    vals = ["NU", "DA"]
                    widget = ctk.CTkComboBox(self, variable=var, values=vals)
                    var.set("NU")"""
    new_trip_renderer = """                elif field_id == "trailer":
                    vals = ["NU", "DA"]
                    widget = ctk.CTkComboBox(self, variable=var, values=vals)
                    var.set("NU")
                elif field_id == "is_external":
                    vals = ["NU", "DA"]
                    widget = ctk.CTkComboBox(self, variable=var, values=vals)
                    var.set("NU")"""
    content = content.replace(old_trip_renderer, new_trip_renderer)

with open('main.py', 'w', encoding='utf-8') as f:
    f.write(content)
