import tkinter as tk
from tkinter import ttk
from openpyxl import load_workbook, Workbook
from tkinter import Tk, Text, StringVar, BooleanVar, _setit, messagebox, filedialog

class ExtractorApp(tk.Frame):
    def __init__(self, parent, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        self.parent.title('xl')
        self.parent.title('XL Extractor')
        self.parent.geometry('510x582')
        self.parent.resizable(0, 0)

        self.source_workbook = None
        self.extraction_workbook = None

        menu_button_sytyle = ttk.Style()
        menu_button_sytyle.configure("TMenubutton", foreground="#ffffff00", background="gray83")
        secondary_button_style = ttk.Style()
        secondary_button_style.configure('secondary.TButton', font=('Calibri', 8))

        self.source_workbook_label = tk.Label(self, text="Source workbook")
        self.source_workbook_field = tk.Text(self, width=50, height = 1)
        self.open_source_file_button = ttk.Button(self, text="Open", command=self.choose_source_file, style='secondary.TButton')
        self.source_worksheet_label = tk.Label(self, text="Source worksheet")
        self.source_worksheet_field = tk.Text(self, width=50, height = 1)
        self.source_worksheet_choice = tk.StringVar()
        self.source_worksheet_choice.set("")
        self.source_worksheet_choice.trace("w", self.source_worksheet_changed)
        self.source_worksheet_optionmenu = ttk.OptionMenu(self, self.source_worksheet_choice)
        self.extraction_workbook_label = tk.Label(self, text="Extraction data workbook")
        self.extraction_workbook_field = tk.Text(self, width=50, height = 1)
        self.open_extraction_file_button = ttk.Button(self, text="Open", command=self.choose_extraction_file, style='secondary.TButton')
        self.extraction_worksheet_label = tk.Label(self, text="Extraction data worksheet")
        self.extraction_worksheet_field = tk.Text(self, width=50, height = 1)
        self.extraction_worksheet_choice = tk.StringVar()
        self.extraction_worksheet_choice.set("")
        self.extraction_worksheet_choice.trace("w", self.extraction_worksheet_changed)
        self.extraction_worksheet_optionmenu = ttk.OptionMenu(self, self.extraction_worksheet_choice)
        self.extraction_columns_frame = tk.Frame(self)
        self.extraction_columns_label = tk.Label(self, text="Extraction data columns")
        self.extraction_columns_listbox = tk.Listbox(self.extraction_columns_frame, width=24, height=10, selectmode='multiple')
        self.extraction_columns_scroll = tk.Scrollbar(self.extraction_columns_frame)
        self.extraction_columns_listbox.config(yscrollcommand=self.extraction_columns_scroll.set)
        self.extraction_columns_scroll.config(command=self.extraction_columns_listbox.yview)
        self.extract_button = ttk.Button(self, text="EXTRACT", command=self.extract)
        self.progress = ttk.Progressbar(self, orient='horizontal', length=500, mode = 'determinate')
        self.status_text = tk.StringVar()
        self.status_label = tk.Label(self, text='', font=('Arial', 8), fg='grey', textvariable=self.status_text)

        self.source_workbook_label.grid(row=0, column=0, columnspan=5, pady=(20,5))
        self.source_workbook_field.grid(row=1, column=0, columnspan=3, padx=(10,5))
        self.open_source_file_button.grid(row=1, column=4, sticky="wns")
        self.source_worksheet_label.grid(row=2, column=0, columnspan=5, pady=(10,5))
        self.source_worksheet_field.grid(row=3, column=0, columnspan=3, padx=(10,5))
        self.source_worksheet_optionmenu.grid(row=3, column=4, sticky="wns")
        self.extraction_workbook_label.grid(row=4, column=0, columnspan=5, pady=(30,5))
        self.extraction_workbook_field.grid(row=5, column=0, columnspan=3, padx=(10,5))
        self.open_extraction_file_button.grid(row=5, column=4, sticky="wns")
        self.extraction_worksheet_label.grid(row=6, column=0, columnspan=5, pady=(10,5))
        self.extraction_worksheet_field.grid(row=7, column=0, columnspan=3, padx=(10,5))
        self.extraction_worksheet_optionmenu.grid(row=7, column=4, sticky="wns")
        self.extraction_columns_label.grid(row=8, column=0, columnspan=5, pady=(10,5))
        self.extraction_columns_frame.grid(row=9, column=0, columnspan=5)
        self.extraction_columns_listbox.pack(side='left', fill='both', expand=1)
        self.extraction_columns_scroll.pack(side='right', fill='y')
        self.extract_button.grid(row=10, column=0, columnspan=5, pady=(40,10))
        self.status_label.grid(row=11, column=0, columnspan=5)
        self.progress.grid(row=12, column=0, columnspan=5, padx=5)
        self.pack(side="top", fill="both", expand=True)

    def get_workbook(self, file_path, error_message='Workbook loading failed'):
        try:
            return load_workbook(file_path)
        except Exception:
            messagebox.showerror(title="Error", message=error_message)

    def choose_source_file(self):
        source_file = filedialog.askopenfilename(filetypes=(("Excel files", "*.xlsx"),("All files","*.*")))
        if source_file:
            self.source_workbook_field.delete(1.0, 'end')
            self.source_workbook_field.insert(1.0, source_file)
            self.set_status('Source workbook loading...')
            self.source_workbook = self.get_workbook(source_file)
            self.set_status('')
            self.refresh_source_worksheets()

    def refresh_source_worksheets(self):
        self.source_worksheet_choice.set("")
        self.source_worksheet_optionmenu['menu'].delete(0, 'end')
        new_choices = self.source_workbook.sheetnames
        for choice in new_choices:
            self.source_worksheet_optionmenu['menu'].add_command(label=choice, command=_setit(self.source_worksheet_choice, choice))

    def source_worksheet_changed(self, *args):
        self.source_worksheet_field.delete(1.0, 'end')
        self.source_worksheet_field.insert(1.0, self.source_worksheet_choice.get())

    def choose_extraction_file(self):
        extraction_file = filedialog.askopenfilename(filetypes=(("Excel files", "*.xlsx"),("All files","*.*")))
        if extraction_file:
            self.extraction_workbook_field.delete(1.0, 'end')
            self.extraction_workbook_field.insert(1.0, extraction_file)
            self.set_status('Extraction workbook loading...')
            self.extraction_workbook = self.get_workbook(extraction_file)
            self.set_status('')
            self.refresh_extraction_worksheets()
            self.clear_extraction_columns_listbox()

    def refresh_extraction_worksheets(self):
        self.extraction_worksheet_choice.set("")
        self.extraction_worksheet_optionmenu['menu'].delete(0, 'end')
        new_choices = self.extraction_workbook.sheetnames
        for choice in new_choices:
            self.extraction_worksheet_optionmenu['menu'].add_command(label=choice, command=_setit(self.extraction_worksheet_choice, choice))

    def extraction_worksheet_changed(self, *args):
        self.extraction_worksheet_field.delete(1.0, 'end')
        self.extraction_worksheet_field.insert(1.0, self.extraction_worksheet_choice.get())
        self.refresh_extraction_columns()

    def refresh_extraction_columns(self):
        self.clear_extraction_columns_listbox()
        if self.extraction_worksheet_choice.get():
            try:
                for header_cell in self.extraction_workbook[self.extraction_worksheet_choice.get()][1]:
                    self.extraction_columns_listbox.insert('end', header_cell.value)
            except:
                messagebox.showerror(title="Error", message='Columns loading failed')

    def clear_extraction_columns_listbox(self):
        self.extraction_columns_listbox.delete(0, 'end')

    def set_status(self, text):
        self.status_text.set(str(text))
        self.status_label.update()

    def extract(self):
        pass

def main():
    root = tk.Tk()
    app = ExtractorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
