import pikepdf
import pdfplumber
import pandas as pd
import os
import sys
from collections import defaultdict
from datetime import datetime
from openpyxl import load_workbook
import tkinter as tk
from tkinter import simpledialog, messagebox

# Initialize Tkinter
root = tk.Tk()
root.withdraw()  # Hide the main Tkinter window

# Function to show a message box and exit the program
def show_message_and_exit(title, message):
    messagebox.showinfo(title, message)
    root.quit()  # Close the Tkinter window

# Get the current directory (where the script or the .exe is located)
if getattr(sys, 'frozen', False):
    # If running as a frozen .exe file
    current_dir = os.path.dirname(sys.executable)
else:
    # If running as a Python script
    current_dir = os.path.dirname(os.path.abspath(__file__))

# Define where the PDFs are stored and where the output Excel file will be saved
pdf_folder = current_dir
output_excel = os.path.join(current_dir, "InfoTransferencias2-2024.xlsx")

# Check if there are any PDF files in the folder
pdf_files = [f for f in os.listdir(pdf_folder) if f.endswith('.pdf')]

if not pdf_files:
    show_message_and_exit("No se encontraron archivos", "No se encontraron los archivos necesarios, descargue desde su correo los PDF 'Cartola_CuentaRUT' enviados por BancoEstado e inclúyalos en una misma carpeta junto al archivo 'ProcesadorCartola.exe'.")
    sys.exit()

# Prepare an empty dataframe for output
columns = ['Mes-Año', 'Tipo', 'Nombre', 'Fecha', 'Única en mes', 'Única en semestre',
           'TEF en el mes', 'TEF en semestre', 'TEF Únicas mes', 'TEF Únicas semestre']
df = pd.DataFrame(columns=columns)

# Helper function to extract the month and year from the date
def extract_month_year(date_str):
    """Extract 'Mes-Año' from the date string."""
    date_obj = datetime.strptime(date_str, "%d/%m/%Y")
    return date_obj.strftime("%m-%Y")

# Helper function to validate a row (only process if TEF transaction, a deposit, and after 01/07/2024)
def is_valid_entry(row):
    """Check if a row is valid (TEF, with deposit, and after 01/07/2024)."""
    if 'TEF' not in row['Descripción']:
        return False
    if not row['Abonos o Depósitos']:  # Only keep rows with deposits
        return False
    date_obj = datetime.strptime(row['Fecha'], "%d/%m/%Y")
    return date_obj >= datetime(2024, 7, 1)

# Process the description to remove unnecessary text and clean spaces in the name
def process_description(desc):
    """Extract 'TEF' and remove redundant prefixes. Also, remove all spaces from the name."""
    if 'TEF BANCOESTADO DE' in desc:
        return 'TEF', "".join(desc.replace('TEF BANCOESTADO DE ', '').split())
    elif 'TEF DE' in desc:
        return 'TEF', "".join(desc.replace('TEF DE ', '').split())
    return '', ''

# Main processing loop
transactions = []
while True:
    password = simpledialog.askstring("PDF Password",
                                      "Ingrese su contraseña:\n\n(Los 4 últimos números de su RUT antes del dígito verificador)\n\nEj: Si su rut es 12.345.678-9, su contraseña será '5678'",
                                      show='*')  # Hide the password input with '*'

    pdf_processed = False  # To track if any PDF was processed successfully
    incorrect_password = False  # To track incorrect password

    for pdf_file in pdf_files:
        pdf_path = os.path.join(pdf_folder, pdf_file)

        # Unlock the PDF using pikepdf with the provided password
        try:
            with pikepdf.open(pdf_path, password=password) as pdf:
                # Save a temporary version without a password
                unlocked_pdf_path = os.path.join(pdf_folder, f"unlocked_{pdf_file}")
                pdf.save(unlocked_pdf_path)
                print(f"Password removed: {unlocked_pdf_path}")
                pdf_processed = True  # Set to True if any PDF is processed successfully
        except pikepdf.PasswordError:
            print(f"File {pdf_file} is password-protected and could not be unlocked with the provided password.")
            incorrect_password = True
            break  # Exit the loop if password is incorrect

        # Process the unlocked PDF
        with pdfplumber.open(unlocked_pdf_path) as pdf:
            for page in pdf.pages:
                table = page.extract_table()
                if table:
                    for row in table[1:]:
                        row_dict = {
                            'Nº DOCTO.': row[0],
                            'Descripción': row[1],
                            'Fecha': row[5],
                            'Cargos o Giros': row[3],
                            'Abonos o Depósitos': row[4],
                        }
                        if is_valid_entry(row_dict):
                            transactions.append(row_dict)

        # Remove the temporary unlocked PDF after processing
        os.remove(unlocked_pdf_path)

    if incorrect_password:
        # Show a message and continue asking for the correct password
        messagebox.showerror("Contraseña incorrecta", "La contraseña ingresada es incorrecta. Por favor, intente de nuevo.")
        continue
    elif pdf_processed:
        break  # Exit the loop if processing was successful

# Create dictionaries for counting TEFs using formatted names (without spaces)
tef_count_month = defaultdict(lambda: defaultdict(int))
tef_unique_month = defaultdict(lambda: set())
tef_count_semester = defaultdict(lambda: defaultdict(int))
tef_unique_semester = defaultdict(lambda: set())

# Collect rows in a list and then concatenate them
rows_to_append = []

for transaction in transactions:
    # Extract and clean data
    tipo, nombre = process_description(transaction['Descripción'])
    fecha = transaction['Fecha']
    mes_año = extract_month_year(fecha)

    # Ensure we are counting and comparing the formatted name
    tef_count_month[mes_año][nombre] += 1
    tef_unique_month[mes_año].add(nombre)

    # Handle the semester (from July 1, 2024)
    semester_key = "2nd semester 2024"
    tef_count_semester[semester_key][nombre] += 1
    tef_unique_semester[semester_key].add(nombre)

    # Check if it is unique for the month and semester
    unique_in_month = "ÚNICA" if tef_count_month[mes_año][nombre] == 1 else "NO ÚNICA"
    unique_in_semester = "ÚNICA" if tef_count_semester[semester_key][nombre] == 1 else "NO ÚNICA"

    # Prepare the row to append
    row = {
        'Mes-Año': mes_año,
        'Tipo': tipo,
        'Nombre': nombre,
        'Fecha': fecha,
        'Única en mes': unique_in_month,
        'Única en semestre': unique_in_semester,
        'TEF en el mes': tef_count_month[mes_año][nombre],
        'TEF en semestre': tef_count_semester[semester_key][nombre],
        'TEF Únicas mes': len(tef_unique_month[mes_año]),
        'TEF Únicas semestre': len(tef_unique_semester[semester_key])
    }
    rows_to_append.append(row)

# Convert the list of rows to a DataFrame and concatenate it with the main DataFrame
df = pd.concat([df, pd.DataFrame(rows_to_append)], ignore_index=True)

# Save to Excel
df.to_excel(output_excel, index=False)

# Adjust column widths in Excel
wb = load_workbook(output_excel)
ws = wb.active

# Define column widths based on content
column_widths = {
    'A': 10,  # Mes-Año
    'B': 6,   # Tipo
    'C': 30,  # Nombre
    'D': 15,  # Fecha
    'E': 15,  # Única en mes
    'F': 18,  # Única en semestre
    'G': 16,  # TEF en el mes
    'H': 16,  # TEF en semestre
    'I': 19,  # TEF Únicas mes
    'J': 19   # TEF Únicas semestre
}

# Apply the column widths
for col, width in column_widths.items():
    ws.column_dimensions[col].width = width

# Save the workbook with adjusted column widths
wb.save(output_excel)

# Show success message
show_message_and_exit("Procesamiento exitoso", f"Datos procesados exitosamente y guardados en {output_excel}.")
