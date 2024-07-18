import tkinter as tk
from gui import GUIApplication
from excel import ExcelHandler

def main():
    root = tk.Tk()
    root.state("zoomed")
    app = GUIApplication(master=root)
    excel_handler = ExcelHandler("user_data.xlsx")

    def save_to_excel():
        name = app.name_entry.get()
        surname = app.surname_entry.get() 
        email = app.email_entry.get()
        cellphone = app.cellphone_entry.get()
        excel_handler.save_to_excel([name, surname, email, cellphone])
        app.data_text.delete(1.0, tk.END)
        app.data_text.insert(tk.END, "Data saved successfully!")

    app.save_button.config(command=save_to_excel)
    app.mainloop()

if __name__ == "__main__":
    main()