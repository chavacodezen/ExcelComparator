import tkinter as tk
from tkinter import filedialog, messagebox
import os
import time
import win32com.client  # To control Excel via COM

class ExcelComparatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Comparador de Archivos Excel")
        
        # Set window size and padding
        self.root.geometry("740x500")
        self.root.config(padx=30, pady=30)

        # Initialize file paths and default columns
        self.main_file = ''
        self.compare_file = ''
        self.main_col = 'B'  # Default main column to compare
        self.compare_col = 'C'  # Default compare column
        self.insert_col = 'G'  # Default column for inserting data
        self.data_col = 'N'  # Default data column in compare file

        # Create UI components
        self.create_widgets()

    def create_widgets(self):
        # Create a frame for the file selection section
        file_selection_frame = tk.Frame(self.root, bg="#d0d0d0", padx=15, pady=5)
        file_selection_frame.grid(row=0, column=0, sticky="ew", pady=(5, 5))

        # File selection: Main file
        tk.Label(file_selection_frame, text="Selecciona el archivo donde se insertarán los resultados", bg="#f0f0f0").grid(row=0, column=0, sticky="w", pady=(10, 5))
        self.main_file_button = tk.Button(file_selection_frame, text="Seleccionar archivo", command=self.select_main_file)
        self.main_file_button.grid(row=0, column=1, sticky="w", pady=(10, 5), padx=(10, 0))

        self.main_file_label = tk.Label(file_selection_frame, text="No se ha seleccionado archivo", width=50, anchor="w", bg="#f0f0f0")
        self.main_file_label.grid(row=1, column=1, sticky="w", pady=(0, 10))

        # File selection: Compare file
        tk.Label(file_selection_frame, text="Selecciona el archivo a comparar", bg="#f0f0f0").grid(row=2, column=0, sticky="w", pady=(10, 5))
        self.compare_file_button = tk.Button(file_selection_frame, text="Seleccionar archivo", command=self.select_compare_file)
        self.compare_file_button.grid(row=2, column=1, sticky="w", pady=(10, 5), padx=(10, 0))

        self.compare_file_label = tk.Label(file_selection_frame, text="No se ha seleccionado archivo", width=50, anchor="w", bg="#f0f0f0")
        self.compare_file_label.grid(row=3, column=1, sticky="w", pady=(0, 10))

        # Spacer between sections
        spacer_label_1 = tk.Label(self.root, text="", width=1)
        spacer_label_1.grid(row=5, column=0, pady=(5, 5))

        # Create a frame for the column selection section
        column_selection_frame = tk.Frame(self.root, bg="#d0d0d0", padx=15, pady=5)
        column_selection_frame.grid(row=6, column=0, sticky="ew", pady=(5, 5))

        # Column selection: Main column
        self.main_col_label = tk.Label(column_selection_frame, text="Selecciona la columna que usarás para comparar en el archivo principal", bg="#e0e0e0")
        self.main_col_label.grid(row=0, column=0, sticky="w", pady=(5, 3))

        self.main_col_entry = tk.Entry(column_selection_frame)
        self.main_col_entry.insert(0, self.main_col)
        self.main_col_entry.grid(row=0, column=1, sticky="w", pady=(5, 3), padx=(10, 0))

        # Column selection: Compare column
        self.compare_col_label = tk.Label(column_selection_frame, text="Selecciona la columna que usarás para comparar en el archivo fuente", bg="#e0e0e0")
        self.compare_col_label.grid(row=1, column=0, sticky="w", pady=(5, 3))

        self.compare_col_entry = tk.Entry(column_selection_frame)
        self.compare_col_entry.insert(0, self.compare_col)
        self.compare_col_entry.grid(row=1, column=1, sticky="w", pady=(5, 3), padx=(10, 0))

        # Spacer between sections
        spacer_label_2 = tk.Label(self.root, text="", width=1)
        spacer_label_2.grid(row=8, column=0, pady=(5, 5))

        # Create a frame for the data insertion column selection section
        data_insertion_frame = tk.Frame(self.root, bg="#d0d0d0", padx=15, pady=5)
        data_insertion_frame.grid(row=9, column=0, sticky="ew", pady=(5, 5))

        # Column selection: Insert column
        tk.Label(data_insertion_frame, text="Selecciona la columna donde se insertarán los resultados", bg="#d0d0d0").grid(row=0, column=0, sticky="w", pady=(5, 3))  # Reduced pady here
        self.insert_col_entry = tk.Entry(data_insertion_frame)
        self.insert_col_entry.insert(0, self.insert_col)
        self.insert_col_entry.grid(row=0, column=1, sticky="w", pady=(5, 3), padx=(10, 0))

        # Column selection: Data column
        tk.Label(data_insertion_frame, text="Selecciona la columna con los datos a insertar", bg="#d0d0d0").grid(row=1, column=0, sticky="w", pady=(5, 3))  # Reduced pady here
        self.data_col_entry = tk.Entry(data_insertion_frame)
        self.data_col_entry.insert(0, self.data_col)
        self.data_col_entry.grid(row=1, column=1, sticky="w", pady=(5, 3), padx=(10, 0))

        # Spacer before the button
        spacer_label_3 = tk.Label(self.root, text="", width=1)
        spacer_label_3.grid(row=11, column=0, pady=(5, 5))

        # Start process button
        self.start_button = tk.Button(self.root, text="Iniciar proceso", command=self.start_process)
        self.start_button.grid(row=12, column=0, columnspan=2, pady=(5, 10))

    def select_main_file(self):
        # Open file dialog to select the main file
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if file_path:
            self.main_file = file_path
            file_name = os.path.basename(self.main_file)  # Get the file name without path
            self.main_file_label.config(text=file_name)  # Display file name

            # Update column selection labels dynamically
            self.main_col_label.config(text=f"Selecciona la columna a comparar en el archivo {file_name}")

    def select_compare_file(self):
        # Open file dialog to select the compare file
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if file_path:
            self.compare_file = file_path
            file_name = os.path.basename(self.compare_file)  # Get the file name without path
            self.compare_file_label.config(text=file_name)  # Display file name

            # Update column selection labels dynamically
            self.compare_col_label.config(text=f"Selecciona la columna a comparar en el archivo {file_name}")

    def start_process(self):
        # Ensure both files are selected
        if not self.main_file or not self.compare_file:
            messagebox.showerror("Error", "Por favor, selecciona ambos archivos.")
            return
        
        # Start the timer
        start_time = time.time()
        
        try:
            # Start Excel application
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False  # Make Excel invisible
            
            # Open both files
            main_wb = excel.Workbooks.Open(self.main_file)
            compare_wb = excel.Workbooks.Open(self.compare_file)
            
            # Get the active sheets (or specify by name if necessary)
            main_sheet = main_wb.Sheets(1)  # First sheet in the main file
            compare_sheet = compare_wb.Sheets(1)  # First sheet in the compare file
            
            # Define column indexes (1-based index for Excel columns)
            main_col_idx = ord(self.main_col.upper()) - ord('A') + 1
            compare_col_idx = ord(self.compare_col.upper()) - ord('A') + 1
            insert_col_idx = ord(self.insert_col.upper()) - ord('A') + 1
            data_col_idx = ord(self.data_col.upper()) - ord('A') + 1

            # Start comparing columns row by row
            max_row_main = main_sheet.Cells(main_sheet.Rows.Count, main_col_idx).End(-4162).Row  # xlUp
            max_row_compare = compare_sheet.Cells(compare_sheet.Rows.Count, compare_col_idx).End(-4162).Row  # xlUp
            
            for i in range(1, max_row_main + 1):
                main_value = main_sheet.Cells(i, main_col_idx).Value
                inserted_data = ""  # Default value for insertion
                
                # Compare values between the columns
                for j in range(1, max_row_compare + 1):
                    compare_value = compare_sheet.Cells(j, compare_col_idx).Value
                    if main_value == compare_value:
                        # If a match is found, take the value from the compare file's data column
                        inserted_data = compare_sheet.Cells(j, data_col_idx).Value
                        break
                
                # Insert the result into the specified column
                main_sheet.Cells(i, insert_col_idx).Value = inserted_data
            
            # Time check: If it takes longer than 1 minute, close the program
            if time.time() - start_time > 60:
                messagebox.showwarning("Tiempo límite alcanzado", "El proceso ha tomado más de 1 minuto. Cerrando el programa.")
                main_wb.Close(False)  # Close without saving
                compare_wb.Close(False)  # Close without saving
                excel.Quit()  # Close Excel application
                self.root.quit()  # Exit the application
                return
            
            # Save the changes (if any)
            main_wb.Save()
            compare_wb.Save()

            # Close both workbooks
            main_wb.Close(False)
            compare_wb.Close(False)

            # Quit Excel
            excel.Quit()

            # Show completion message
            messagebox.showinfo("Proceso completado", "El proceso ha finalizado correctamente.")
        
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error: {e}")

# Create the main window
root = tk.Tk()
app = ExcelComparatorApp(root)
root.mainloop()
