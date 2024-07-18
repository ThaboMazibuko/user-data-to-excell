import tkinter as tk
from tkinter import messagebox

class GUIApplication(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.create_widgets()

    def create_widgets(self):
        self.name_label = tk.Label(self, text="Name:", font=("Helvetica", 14), fg="blue")
        self.name_label.grid(column=0, row=0, padx=10, pady=10)
        self.name_entry = tk.Entry(self, width=30, font=("Helvetica", 14))
        self.name_entry.grid(column=1, row=0, padx=10, pady=10)

        self.surname_label = tk.Label(self, text="Surname:", font=("Helvetica", 14), fg="blue")
        self.surname_label.grid(column=0, row=1, padx=10, pady=10)
        self.surname_entry = tk.Entry(self, width=30, font=("Helvetica", 14))
        self.surname_entry.grid(column=1, row=1, padx=10, pady=10)

        self.email_label = tk.Label(self, text="Email:", font=("Helvetica", 14), fg="blue")
        self.email_label.grid(column=0, row=2, padx=10, pady=10)
        self.email_entry = tk.Entry(self, width=30, font=("Helvetica", 14))
        self.email_entry.grid(column=1, row=2, padx=10, pady=10)

        self.cellphone_label = tk.Label(self, text="Cellphone Number:", font=("Helvetica", 14), fg="blue")
        self.cellphone_label.grid(column=0, row=3, padx=10, pady=10)
        self.cellphone_entry = tk.Entry(self, width=30, font=("Helvetica", 14))
        self.cellphone_entry.grid(column=1, row=3, padx=10, pady=10)

        self.save_button = tk.Button(self, text="Save", command=self.save_to_excel, font=("Helvetica", 16), fg="green", bg="light green")
        self.save_button.grid(column=1, row=4, padx=10, pady=10)

        self.data_text = tk.Text(self, width=40, height=10, font=("Helvetica", 12))
        self.data_text.grid(column=0, row=5, columnspan=2, padx=10, pady=10)

    def save_to_excel(self):
        # TO DO: implement save to excel functionality
        pass