import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os

from scripts import createXMLFile, convert_xml_to_excel


class ExcelToXmlConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel2XML")

        self.root.geometry("600x800")

        # Set a pastel color scheme
        self.root.option_add("*TButton*background", "#AED6F1")  # Light Blue
        self.root.option_add("*TButton*foreground", "#2C3E50")  # Dark Blue
        self.root.option_add("*TEntry*background", "#D5F5E3")   # Mint Green
        self.root.option_add("*TEntry*foreground", "#2C3E50")    # Dark Blue
        self.root.option_add("*TLabel*background", "#FDEBD0")    # Cream
        self.root.option_add("*TLabel*foreground", "#2C3E50")    # Dark Blue

        # Main menu buttons
        self.button1 = ttk.Button(root, text="Convert Excel to XML", command=self.convert_excel_to_xml_interface)
        self.button1.pack(pady=(50, 10))

        self.button2 = ttk.Button(root, text="Convert XML to Excel", command=self.convert_xml_to_excel_interface)
        self.button2.pack(pady=10)

    def convert_excel_to_xml_interface(self):
        self.clear_root()

        # Input fields for additional metadata
        self.create_label_and_entry("Name:", "name")
        self.create_label_and_entry("Edition Version: (example: v2.0)", "edition_version")
        self.create_label_and_entry("Year: (example: 2019)", "year")
        self.create_label_and_entry("Month: (example: 3)", "month")
        self.create_label_and_entry("Day: (example: 15)", "day")
        self.create_label_and_entry("Source: (example: www.graphisoft.com)", "source")

        # Input Excel file
        self.input_file_label = ttk.Label(self.root, text="Input Excel File:")
        self.input_file_label.pack(pady=(10, 5))

        self.input_file_entry = ttk.Entry(self.root, state="disabled", width=40)
        self.input_file_entry.pack(pady=5)

        self.choose_file_button = ttk.Button(self.root, text="Choose File", command=lambda: self.choose_file("excel"))
        self.choose_file_button.pack(pady=(5, 10))

        # Output folder and XML file
        self.output_folder_label = ttk.Label(self.root, text="Output Folder:")
        self.output_folder_label.pack(pady=(10, 5))

        self.output_folder_entry = ttk.Entry(self.root, state="disabled", width=40)
        self.output_folder_entry.pack(pady=5)

        self.choose_folder_button = ttk.Button(self.root, text="Choose Folder", command=self.choose_folder)
        self.choose_folder_button.pack(pady=(5, 10))

        self.output_file_label = ttk.Label(self.root, text="XML File Name:")
        self.output_file_label.pack(pady=(10, 5))

        self.output_file_entry = ttk.Entry(self.root, width=40)
        self.output_file_entry.pack(pady=5)

        # Button to run conversion
        self.run_button = ttk.Button(self.root, text="Run Conversion", command=self.run_excel_to_xml_conversion)
        self.run_button.pack(pady=(20, 10))

        # Return button
        self.return_button = ttk.Button(self.root, text="Return to Menu", command=self.return_to_menu)
        self.return_button.pack(pady=(20, 10))

    def create_label_and_entry(self, label_text, attribute_name):
        label = ttk.Label(self.root, text=label_text)
        label.pack(pady=(10, 5))

        entry = ttk.Entry(self.root, width=40)
        entry.pack(pady=5)

        setattr(self, f"{attribute_name}_entry", entry)

    def convert_xml_to_excel_interface(self):
        self.clear_root()

        # Input XML file
        self.input_file_label = ttk.Label(self.root, text="Input XML File:")
        self.input_file_label.pack(pady=(10, 5))

        self.input_file_entry = ttk.Entry(self.root, state="disabled", width=40)
        self.input_file_entry.pack(pady=5)

        self.choose_file_button = ttk.Button(self.root, text="Choose File", command=lambda: self.choose_file("xml"))
        self.choose_file_button.pack(pady=(5, 10))

        # Output folder and Excel file
        self.output_folder_label = ttk.Label(self.root, text="Output Folder:")
        self.output_folder_label.pack(pady=(10, 5))

        self.output_folder_entry = ttk.Entry(self.root, state="disabled", width=40)
        self.output_folder_entry.pack(pady=5)

        self.choose_folder_button = ttk.Button(self.root, text="Choose Folder", command=self.choose_folder)
        self.choose_folder_button.pack(pady=(5, 10))

        self.output_file_label = ttk.Label(self.root, text="Excel File Name:")
        self.output_file_label.pack(pady=(10, 5))

        self.output_file_entry = ttk.Entry(self.root, width=40)
        self.output_file_entry.pack(pady=5)

        # Button to run conversion
        self.run_button = ttk.Button(self.root, text="Run Conversion", command=self.run_xml_to_excel_conversion)
        self.run_button.pack(pady=(20, 10))

        # Return button
        self.return_button = ttk.Button(self.root, text="Return to Menu", command=self.return_to_menu)
        self.return_button.pack(pady=(20, 10))

    def choose_file(self, fileType):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls") if fileType == "excel" else ("XML files", "*.xml")])
        if file_path:
            self.input_file_entry.config(state="normal")
            self.input_file_entry.delete(0, tk.END)
            self.input_file_entry.insert(0, file_path)
            self.input_file_entry.config(state="disabled")

    def choose_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.output_folder_entry.config(state="normal")
            self.output_folder_entry.delete(0, tk.END)
            self.output_folder_entry.insert(0, folder_path)
            self.output_folder_entry.config(state="disabled")

    def run_excel_to_xml_conversion(self):
        input_file = self.input_file_entry.get()
        output_folder = self.output_folder_entry.get()
        output_file_name = self.output_file_entry.get()

        if not input_file or not output_folder or not output_file_name:
            messagebox.showerror("Error", "Please select input, output folder, and provide XML file name.")
            return

        name = self.name_entry.get()
        edition_version = self.edition_version_entry.get()
        year = self.year_entry.get()
        month = self.month_entry.get()
        day = self.day_entry.get()
        source = self.source_entry.get()

        if not name or not edition_version or not year or not month or not day or not source:
            messagebox.showerror("Error", "Please fill in all metadata fields.")
            return
        try:
            year = int(year)
            month = int(month)
            day = int(day)
        except ValueError:
            messagebox.showerror("Error", "Year, month, and day must be valid integers.")
            return
        
        year = str(year)
        month = str(month)
        day = str(day)
        index = output_file_name.find('.')
        if index != -1:
            output_file_name = output_file_name[:index]

        output_file_name += '.xml'

        output_file_path = os.path.join(output_folder, output_file_name)

        try:
            createXMLFile(input_file, output_file_path, name, edition_version, year, month, day, source)  # Your conversion function
            messagebox.showinfo("Success", "Conversion successful!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def run_xml_to_excel_conversion(self):
        input_file = self.input_file_entry.get()
        output_folder = self.output_folder_entry.get()
        output_file_name = self.output_file_entry.get()

        if not input_file or not output_folder or not output_file_name:
            messagebox.showerror("Error", "Please select input, output folder, and provide Excel file name.")
            return

        index = output_file_name.find('.')
        if index != -1:
            output_file_name = output_file_name[:index]

        output_file_name += '.xlsx'

        output_file_path = os.path.join(output_folder, output_file_name)

        try:
            convert_xml_to_excel(input_file, output_file_path)  # Your conversion function
            messagebox.showinfo("Success", "Conversion successful!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def clear_root(self):
        for widget in self.root.winfo_children():
            widget.destroy()

    def return_to_menu(self):
        self.clear_root()
        self.__init__(self.root)


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelToXmlConverterApp(root)
    root.mainloop()
