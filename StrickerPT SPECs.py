import tkinter as tk
from tkinter import ttk, messagebox
import pyodbc


def reset_dropdowns(starting_from):
    # Resets dropdowns starting from a specific point in the sequence
    dropdowns = [location_dropdown, allowfullcolor_dropdown, tipo_customizacao_dropdown, tablemaxareacm_dropdown]
    starting_index = dropdowns.index(starting_from)
    for dropdown in dropdowns[starting_index:]:
        dropdown.set('')
        dropdown['values'] = []
        dropdown.config(state="disabled")


def fetch_initial_data():
    reset_dropdowns(location_dropdown)  # Reset all dropdowns when new fetch is initiated
    prod_ref = prod_ref_entry.get()
    if not prod_ref.isdigit():
        messagebox.showerror("Error", "Product Reference must be a number")
        return
    try:
        conn = pyodbc.connect(
            'DRIVER={SQL Server};SERVER=xxx;DATABASE=xxxx;UID=xxx;PWD=xxxx')
        cursor = conn.cursor()
        cursor.execute("EXEC dbo.FetchStrickerCustomizationData ?", int(prod_ref))
        records = cursor.fetchall()
        if records:
            app_data['records'] = {}
            for row in records:
                key = (row.Local_Personalizacao, str(bool(row.AllowFullColor)), row.Tipo_Customizacao)
                if key not in app_data['records']:
                    app_data['records'][key] = set()
                app_data['records'][key].add((row.TableMaxAreaCM, row.ServiceCode))
            update_location_dropdown(set([k[0] for k in app_data['records'].keys()]))
        else:
            messagebox.showinfo("Result", "No results for that Product Reference")
        cursor.close()
        conn.close()
    except Exception as e:
        messagebox.showerror("Database Error", str(e))


def update_location_dropdown(options):
    reset_dropdowns(allowfullcolor_dropdown)  # Reset all dependent dropdowns
    location_dropdown['values'] = list(options)
    if options:
        location_dropdown.current(0)
    location_dropdown.config(state="readonly")


def location_selected(event):
    reset_dropdowns(allowfullcolor_dropdown)  # Reset all dependent dropdowns
    selected_location = location_dropdown.get()
    allowfullcolor_options = set(k[1] for k in app_data['records'].keys() if k[0] == selected_location)
    update_allowfullcolor_dropdown(allowfullcolor_options)


def update_allowfullcolor_dropdown(options):
    reset_dropdowns(tipo_customizacao_dropdown)  # Reset all dependent dropdowns
    allowfullcolor_dropdown['values'] = list(options)
    if options:
        allowfullcolor_dropdown.current(0)
    allowfullcolor_dropdown.config(state="readonly")


def allowfullcolor_selected(event):
    reset_dropdowns(tipo_customizacao_dropdown)  # Reset all dependent dropdowns
    selected_location = location_dropdown.get()
    selected_allowfullcolor = allowfullcolor_dropdown.get()
    tipo_customizacao_options = set(
        k[2] for k in app_data['records'].keys() if (k[0], k[1]) == (selected_location, selected_allowfullcolor))
    update_tipo_customizacao_dropdown(tipo_customizacao_options)


def update_tipo_customizacao_dropdown(options):
    reset_dropdowns(tablemaxareacm_dropdown)  # Reset all dependent dropdowns
    tipo_customizacao_dropdown['values'] = list(options)
    if options:
        tipo_customizacao_dropdown.current(0)
    tipo_customizacao_dropdown.config(state="readonly")


def tipo_customizacao_selected(event):
    selected_location = location_dropdown.get()
    selected_allowfullcolor = allowfullcolor_dropdown.get()
    selected_tipo_customizacao = tipo_customizacao_dropdown.get()
    key = (selected_location, selected_allowfullcolor, selected_tipo_customizacao)
    tablemaxareacm_options = app_data['records'].get(key, [])
    update_tablemaxareacm_dropdown(tablemaxareacm_options)


def update_tablemaxareacm_dropdown(options):
    tablemaxareacm_dropdown['values'] = [option[0] for option in options]  # Update to show only dimensions
    if options:
        tablemaxareacm_dropdown.current(0)
    tablemaxareacm_dropdown.config(state="readonly")


def confirm_insert():
    if not all([location_dropdown.get(), allowfullcolor_dropdown.get(), tipo_customizacao_dropdown.get(),
                tablemaxareacm_dropdown.get(), subproduct_id_entry.get().isdigit()]):
        messagebox.showerror("Error", "Please complete all selections and enter a valid Subproduct ID.")
        return
    try:
        conn = pyodbc.connect(
            'DRIVER={SQL Server};SERVER=192.168.10.226;DATABASE=wkb_op;UID=wkb_op_su;PWD=vJYWp8cZTZIuGuvP3sTQ')
        cursor = conn.cursor()
        selected_location = location_dropdown.get()
        selected_allowfullcolor = int(allowfullcolor_dropdown.get() == 'True')
        selected_tipo_customizacao = tipo_customizacao_dropdown.get()
        selected_cm = tablemaxareacm_dropdown.get()
        key = (selected_location, allowfullcolor_dropdown.get(), selected_tipo_customizacao)

        # Fetch all records for the selected key
        records = app_data['records'].get(key, [])

        # Find the matching service code
        matching_service_code = None
        for record in records:
            cm, service_code = record
            if cm == selected_cm:
                matching_service_code = service_code
                break

        if matching_service_code:
            cursor.execute("EXEC dbo.InsertStrickerPrintMappingData ?, ?, ?, ?, ?, ?, ?",
                           int(subproduct_id_entry.get()), int(prod_ref_entry.get()), selected_location,
                           selected_tipo_customizacao, selected_allowfullcolor, selected_cm, matching_service_code)
            conn.commit()
            messagebox.showinfo("Success", "Data successfully inserted.")
        else:
            messagebox.showerror("Error", "No matching service code found for the selected criteria.")

        cursor.close()
        conn.close()
    except Exception as e:
        messagebox.showerror("Database Error", str(e))
        print(str(e))


# Define a custom font
custom_font = ('Helvetica', 10, 'bold')  # Change the font size (12) and weight ('bold') as needed

# Create the main application window
app = tk.Tk()
app.title("StrickerPT SPECs App")
app.geometry("800x600")  # Set initial window size
app.configure(bg="#B6D0E2")  # Change background color

app_data = {'records': {}}

# Label and Entry for Product Reference
tk.Label(app, text="Enter Stricker PT Reference:", bg="#B6D0E2", font=custom_font).pack(pady=(40, 5))
prod_ref_entry = tk.Entry(app, font=custom_font, justify="center", width=30)
prod_ref_entry.pack(pady=5)

# Button to fetch initial data
submit_button = tk.Button(app, text="Search", command=fetch_initial_data, font=custom_font)
submit_button.pack(pady=(5, 25))

# Dropdown for Location
tk.Label(app, text="Select Print Location:", bg="#B6D0E2", font=custom_font).pack(pady=5)
location_dropdown = ttk.Combobox(app, state="disabled", font=custom_font, justify="center", width=30)
location_dropdown.pack(pady=(0, 10))
location_dropdown.bind('<<ComboboxSelected>>', location_selected)

# Dropdown for Allow Full Color
tk.Label(app, text="Allow Full Color?", bg="#B6D0E2", font=custom_font).pack(pady=5)
allowfullcolor_dropdown = ttk.Combobox(app, state="disabled", font=custom_font, justify="center", width=30)
allowfullcolor_dropdown.pack(pady=(0, 10))
allowfullcolor_dropdown.bind('<<ComboboxSelected>>', allowfullcolor_selected)

# Dropdown for Tipo Customizacao
tk.Label(app, text="Select Print Technique:", bg="#B6D0E2", font=custom_font).pack(pady=5)
tipo_customizacao_dropdown = ttk.Combobox(app, state="disabled", font=custom_font, justify="center", width=30)
tipo_customizacao_dropdown.pack(pady=(0, 10))
tipo_customizacao_dropdown.bind('<<ComboboxSelected>>', tipo_customizacao_selected)

# Dropdown for TableMaxAreaCM
tk.Label(app, text="Select Print Size in Centimeters (W x H):", bg="#B6D0E2", font=custom_font).pack(pady=5)
tablemaxareacm_dropdown = ttk.Combobox(app, state="disabled", font=custom_font, justify="center", width=30)
tablemaxareacm_dropdown.pack(pady=(2, 40))

# Entry for Subproduct ID
tk.Label(app, text="Associate to SubProductID:", bg="#B6D0E2", font=custom_font).pack(pady=5)
subproduct_id_entry = tk.Entry(app, font=custom_font, justify="center", width=30)
subproduct_id_entry.pack(pady=(0, 10))

# Button to confirm insert
confirm_button = tk.Button(app, text="Confirm Insert", command=confirm_insert, font=custom_font)
confirm_button.pack(pady=5)

# Run the application
app.mainloop()
