from tkinter import filedialog, messagebox
import pandas as pd
import tkinter as tk

GUI = tk.Tk()

GUI.title('Exel to CSV Converter')
canvas1 = tk.Canvas(GUI, width=300, height=300, bg='gray93', relief='raised')
canvas1.pack()

label1 = tk.Label(GUI, text='Excel to CSV', bg='gray93')
label1.config(font=('helvetica', 20))
canvas1.create_window(150, 60, window=label1)

def get_excel():
    global read_file

    import_file_path = filedialog.askopenfilename()
    read_file = pd.read_excel(import_file_path)

browseButton_Excel = tk.Button(text="     Import Excel File     ", command=get_excel, bg='green', fg='white',  font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 130, window=browseButton_Excel)


def convert_to_cvs():
    global read_file

    export_file_path = filedialog.asksaveasfilename(defaultextension='.csv')
    read_file.to_csv(export_file_path, index=None, header=True)

saveAsButton_CSV = tk.Button(text='Convert Excel to CSV', command=convert_to_cvs, bg='green', fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 180, window=saveAsButton_CSV)


def exit_application():
    message = tk.messagebox.askquestion('Exit Application', 'Are you sure you want to exit the application', icon='warning')
    if message == 'yes':
        GUI.destroy()

exit_button = tk.Button(GUI, text='                 Exit                 ', command=exit_application, bg='brown', fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 230, window=exit_button)

GUI.mainloop()
