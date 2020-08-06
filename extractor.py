import tkinter as tk
from tkinter import ttk
import time

class ExtractorApp(tk.Frame):
    def __init__(self, parent, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        self.parent.title('xl')
        self.parent.title('XL Extractor')
        self.parent.geometry('510x582')
        self.parent.resizable(0, 0)

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

    def choose_source_file(self):
        pass

    def source_worksheet_changed(self):
        pass

    def choose_extraction_file(self):
        pass

    def extraction_worksheet_changed(self):
        pass

    def extract(self):
        pass

def main():
    root = tk.Tk()
    app = ExtractorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
