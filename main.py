from tkinter import *
from tkinter import ttk
from tkinter import filedialog, messagebox
import getpass
import pandas as pd

import sys
import os

if getattr(sys, 'frozen', False):
    # If the application is run as a bundle, use the special _MEIPASS
    # attribute to get the path to the bundled resources
    bundle_dir = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
else:
    # If the application is not run as a bundle, use the standard
    # directory structure
    bundle_dir = os.path.abspath(os.path.dirname(__file__))

# Set the path to the icon file
icon_path = os.path.join(bundle_dir, 'CSAT_FCR.ico')

root = Tk()

root.title("CSAT & FCR Achievement")
root.geometry('500x650')
root.iconbitmap(default=icon_path)


def upload_file():
    global excel_file
    username = getpass.getuser()
    downloads_path = f'C:/Users/{username}/Downloads'
    excel_file = filedialog.askopenfilename(
        initialdir=downloads_path, title="Select Excel file", filetypes=(("Excel files", "*.xlsm"),))
    file_name = excel_file.replace(f'{downloads_path}/', '')
    if file_name != '':
        label_filename.config(text=file_name)
    else:
        pass
    entry_employee_number.focus()


def search():
    try:
        xl = pd.read_excel(excel_file, engine='openpyxl',
                           sheet_name='CSAT - CSA')
    except NameError:
        messagebox.showinfo("No excel file selected",
                            "No excel file selected. Please select one.")
        return

    search_number = int(entry_employee_number.get())
    rows_csat = xl[xl['Unnamed: 0'] == search_number]
    rows_fcr = xl[xl['Unnamed: 10'] == search_number]

    values_csat = rows_csat.loc[:, 'Unnamed: 1':'Unnamed: 7']
    values_fcr = rows_fcr.loc[:, 'Unnamed: 11':'Unnamed: 15']

    try:
        output_csat = ' '.join(map(str, values_csat.values.tolist()[0]))
        output_fcr = ' '.join(map(str, values_fcr.values.tolist()[0]))
    except IndexError:
        entry_employee_number.delete(0, 'end')

        entry_satisfied.config(state='normal')
        entry_satisfied.delete(0, 'end')
        entry_satisfied.config(
            state='disabled', disabledbackground="white", disabledforeground="black")

        entry_neutral.config(state='normal')
        entry_neutral.delete(0, 'end')
        entry_neutral.config(
            state='disabled', disabledbackground="white", disabledforeground="black")

        entry_dissatisfied.config(state='normal')
        entry_dissatisfied.delete(0, 'end')
        entry_dissatisfied.config(
            state='disabled', disabledbackground="white", disabledforeground="black")

        entry_csat.config(state='normal')
        entry_csat.delete(0, 'end')
        entry_csat.config(
            state='disabled', disabledbackground="white", disabledforeground="black")

        entry_dsat.config(state='normal')
        entry_dsat.delete(0, 'end')
        entry_dsat.config(
            state='disabled', disabledbackground="white", disabledforeground="black")

        entry_yes.config(state='normal')
        entry_yes.delete(0, 'end')
        entry_yes.config(
            state='disabled', disabledbackground="white", disabledforeground="black")

        entry_no.config(state='normal')
        entry_no.delete(0, 'end')
        entry_no.config(state='disabled',
                        disabledbackground="white", disabledforeground="black")

        entry_fcr.config(state='normal')
        entry_fcr.delete(0, 'end')
        entry_fcr.config(
            state='disabled', disabledbackground="white", disabledforeground="black")

        entry_achievement_csat_75.config(state='normal')
        entry_achievement_csat_75.delete(0, 'end')
        entry_achievement_csat_75.config(
            state='disabled', disabledbackground="white", disabledforeground="black")

        entry_achievement_csat_85.config(state='normal')
        entry_achievement_csat_85.delete(0, 'end')
        entry_achievement_csat_85.config(
            state='disabled', disabledbackground="white", disabledforeground="black")

        entry_achievement_csat_91.config(state='normal')
        entry_achievement_csat_91.delete(0, 'end')
        entry_achievement_csat_91.config(
            state='disabled', disabledbackground="white", disabledforeground="black")

        entry_achievement_dsat_5.config(state='normal')
        entry_achievement_dsat_5.delete(0, 'end')
        entry_achievement_dsat_5.config(
            state='disabled', disabledbackground="white", disabledforeground="black")

        entry_achievement_fcr_90.config(state='normal')
        entry_achievement_fcr_90.delete(0, 'end')
        entry_achievement_fcr_90.config(
            state='disabled', disabledbackground="white", disabledforeground="black")

        frame_employee_details.config(text=f'Employee Details : ')

        messagebox.showerror("Invalid Employee Number",
                             "Either the entered employee number does not exist in the database, or there are no responses for the selected employee. Please verify the employee number and try again.")

        return

    item_list_csat = output_csat.split(' ')
    item_list_fcr = output_fcr.split(' ')

    name = f'({item_list_csat[0]} {item_list_csat[1]})'
    dissatisfied = int(item_list_csat[2].replace('nan', '0'))
    neutral = int(item_list_csat[3].replace('nan', '0'))
    satisfied = int(item_list_csat[4].replace('nan', '0'))
    grand_total_csat = int(item_list_csat[5].replace('nan', '0'))
    csat = float(item_list_csat[6].replace('nan', '0'))
    dsat = float(item_list_csat[7].replace('nan', '0'))

    csat_perc = format(csat * 100, '.2f')
    dsat_perc = format(dsat * 100, '.2f')

    no = int(item_list_fcr[2].replace('nan', '0'))
    yes = int(item_list_fcr[3].replace('nan', '0'))
    grand_total_fcr = int(item_list_fcr[4].replace('nan', '0'))
    fcr = float(item_list_fcr[5].replace('nan', '0'))

    fcr_perc = format(fcr * 100, '.2f')

    frame_employee_details.config(text=f'Employee Details {name}: ')

    entry_satisfied.config(state='normal')
    entry_satisfied.delete(0, 'end')
    entry_satisfied.insert(0, satisfied)
    entry_satisfied.config(
        state='disabled', disabledbackground="white", disabledforeground="black")

    entry_neutral.config(state='normal')
    entry_neutral.delete(0, 'end')
    entry_neutral.insert(0, neutral)
    entry_neutral.config(
        state='disabled', disabledbackground="white", disabledforeground="black")

    entry_dissatisfied.config(state='normal')
    entry_dissatisfied.delete(0, 'end')
    entry_dissatisfied.insert(0, dissatisfied)
    entry_dissatisfied.config(
        state='disabled', disabledbackground="white", disabledforeground="black")

    entry_csat.config(state='normal')
    entry_csat.delete(0, 'end')
    entry_csat.insert(0, f'{csat_perc}%')
    entry_csat.config(state='disabled',
                      disabledbackground="white", disabledforeground="black")

    entry_dsat.config(state='normal')
    entry_dsat.delete(0, 'end')
    entry_dsat.insert(0, f'{dsat_perc}%')
    entry_dsat.config(state='disabled',
                      disabledbackground="white", disabledforeground="black")

    entry_yes.config(state='normal')
    entry_yes.delete(0, 'end')
    entry_yes.insert(0, yes)
    entry_yes.config(state='disabled', disabledbackground="white",
                     disabledforeground="black")

    entry_no.config(state='normal')
    entry_no.delete(0, 'end')
    entry_no.insert(0, no)
    entry_no.config(state='disabled', disabledbackground="white",
                    disabledforeground="black")

    entry_fcr.config(state='normal')
    entry_fcr.delete(0, 'end')
    entry_fcr.insert(0, f'{fcr_perc}%')
    entry_fcr.config(state='disabled', disabledbackground="white",
                     disabledforeground="black")

    if grand_total_csat == 0:
        result_csat_75 = "N/A"
        result_csat_85 = "N/A"
        result_csat_91 = "N/A"
    else:
        if csat >= 0.75:
            result_csat_75 = "MET"
        else:
            result_csat_75 = round(
                (75 * grand_total_csat - 100 * satisfied) / 25)

        if csat >= 0.85:
            result_csat_85 = "MET"
        else:
            result_csat_85 = round(
                (85 * grand_total_csat - 100 * satisfied) / 15)

        if csat >= 0.91:
            result_csat_91 = "MET"
        else:
            result_csat_91 = round(
                (91 * grand_total_csat - 100 * satisfied) / 9)

        if csat >= 0.95:
            result_dsat_5 = "MET"
        else:
            result_dsat_5 = round(
                (95 * grand_total_csat - 100 * satisfied) / 5)

    entry_achievement_csat_75.config(state='normal')
    entry_achievement_csat_75.delete(0, 'end')
    entry_achievement_csat_75.insert(0, result_csat_75)
    entry_achievement_csat_75.config(
        state='disabled', disabledbackground="white", disabledforeground="black")

    entry_achievement_csat_85.config(state='normal')
    entry_achievement_csat_85.delete(0, 'end')
    entry_achievement_csat_85.insert(0, result_csat_85)
    entry_achievement_csat_85.config(
        state='disabled', disabledbackground="white", disabledforeground="black")

    entry_achievement_csat_91.config(state='normal')
    entry_achievement_csat_91.delete(0, 'end')
    entry_achievement_csat_91.insert(0, result_csat_91)
    entry_achievement_csat_91.config(
        state='disabled', disabledbackground="white", disabledforeground="black")

    entry_achievement_dsat_5.config(state='normal')
    entry_achievement_dsat_5.delete(0, 'end')
    entry_achievement_dsat_5.insert(0, result_dsat_5)
    entry_achievement_dsat_5.config(
        state='disabled', disabledbackground="white", disabledforeground="black")

    if grand_total_fcr == 0:
        result_fcr_90 = "N/A"
    else:
        if fcr >= 0.9:
            result_fcr_90 = "MET"
        else:
            result_fcr_90 = round((90 * grand_total_fcr - 100 * yes) / 10)

    entry_achievement_fcr_90.config(state='normal')
    entry_achievement_fcr_90.delete(0, 'end')
    entry_achievement_fcr_90.insert(0, result_fcr_90)
    entry_achievement_fcr_90.config(
        state='disabled', disabledbackground="white", disabledforeground="black")


def search_enter(event):
    search()


frame_upload = LabelFrame(root, text='Upload Excel')
frame_upload.pack(fill=X, padx=10, pady=10)

label_upload = Label(frame_upload, text='File Selected:')
label_upload.pack(side=LEFT, padx=10, pady=10)

label_filename = Label(frame_upload, text='No file selected')
label_filename.pack(side=LEFT, padx=10, pady=10)

button_upload = ttk.Button(
    frame_upload, text='Upload Excel', command=upload_file)
button_upload.pack(side=RIGHT, padx=10, pady=10)

frame_employee_details = LabelFrame(root, text='Employee Details:')
frame_employee_details.pack(fill=X, padx=10, pady=10)

frame_employee_number = Frame(frame_employee_details)
frame_employee_number.pack(fill=X)

label_employee_number = Label(
    frame_employee_number, text='Enter your Employee Number: ')
label_employee_number.pack(side=LEFT, padx=10, pady=10)

entry_employee_number = Entry(frame_employee_number, width=6)
entry_employee_number.pack(side=LEFT, fill=X, padx=10, pady=10)

entry_employee_number.bind('<Return>', search_enter)

button_employee_number = ttk.Button(
    frame_employee_number, text='Search', command=search)
button_employee_number.pack(side=RIGHT, fill=X, padx=10, pady=10)

current_stats_frame = Frame(root)
current_stats_frame.pack(fill=X)

csat_frame = LabelFrame(current_stats_frame, text='CSAT', width=280)
csat_frame.pack(fill=Y, side=LEFT, padx=10, pady=10)

label_satisfied = Label(csat_frame, text='Satisfied: ')
label_satisfied.grid(row=0, column=0, padx=10, pady=10)

entry_satisfied = Entry(csat_frame, width=3, state='disabled',
                        disabledbackground="white", disabledforeground="black")
entry_satisfied.grid(row=0, column=1, padx=10, pady=10)

label_neutral = Label(csat_frame, text='Neutral: ')
label_neutral.grid(row=1, column=0, padx=10, pady=10)

entry_neutral = Entry(csat_frame, width=3, state='disabled',
                      disabledbackground="white", disabledforeground="black")
entry_neutral.grid(row=1, column=1, padx=10, pady=10)

label_dissatisfied = Label(csat_frame, text='Dissatisfied: ')
label_dissatisfied.grid(row=2, column=0, padx=10, pady=10)

entry_dissatisfied = Entry(csat_frame, width=3, state='disabled',
                           disabledbackground="white", disabledforeground="black")
entry_dissatisfied.grid(row=2, column=1, padx=10, pady=10)

label_csat = Label(csat_frame, text='CSAT: ')
label_csat.grid(row=3, column=0, padx=10, pady=10)

entry_csat = Entry(csat_frame, width=8, state='disabled',
                   disabledbackground="white", disabledforeground="black")
entry_csat.grid(row=3, column=1, padx=10, pady=10)

label_dsat = Label(csat_frame, text='DSAT: ')
label_dsat.grid(row=4, column=0, padx=10, pady=10)

entry_dsat = Entry(csat_frame, width=8, state='disabled',
                   disabledbackground="white", disabledforeground="black")
entry_dsat.grid(row=4, column=1, padx=10, pady=10)

fcr_frame = LabelFrame(current_stats_frame, text='FCR')
fcr_frame.pack(fill=Y, side=RIGHT, padx=10, pady=10)

label_yes = Label(fcr_frame, text='Yes: ')
label_yes.grid(row=0, column=0, padx=10, pady=10)

entry_yes = Entry(fcr_frame, width=3, state='disabled',
                  disabledbackground="white", disabledforeground="black")
entry_yes.grid(row=0, column=1, padx=10, pady=10)

label_no = Label(fcr_frame, text='No: ')
label_no.grid(row=1, column=0, padx=10, pady=10)

entry_no = Entry(fcr_frame, width=3, state='disabled',
                 disabledbackground="white", disabledforeground="black")
entry_no.grid(row=1, column=1, padx=10, pady=10)

label_fcr = Label(fcr_frame, text='FCR: ')
label_fcr.grid(row=2, column=0, padx=10, pady=10)

entry_fcr = Entry(fcr_frame, width=8, state='disabled',
                  disabledbackground="white", disabledforeground="black")
entry_fcr.grid(row=2, column=1, padx=10, pady=10)

achievement_frame = LabelFrame(root, text='Achievement Targets')
achievement_frame.pack(fill=X, padx=10, pady=10)

label_achievement_csat_75 = Label(
    achievement_frame, text='The number of CSAT only responses required to achieve >=75% CSAT: ')
label_achievement_csat_75.grid(row=0, column=0, padx=10, pady=10)

entry_achievement_csat_75 = Entry(
    achievement_frame, width=4, state='disabled', disabledbackground="white", disabledforeground="black")
entry_achievement_csat_75.grid(row=0, column=1, padx=10, pady=10)

label_achievement_csat_85 = Label(
    achievement_frame, text='The number of CSAT only responses required to achieve >=85% CSAT: ')
label_achievement_csat_85.grid(row=1, column=0, padx=10, pady=10)

entry_achievement_csat_85 = Entry(
    achievement_frame, width=4, state='disabled', disabledbackground="white", disabledforeground="black")
entry_achievement_csat_85.grid(row=1, column=1, padx=10, pady=10)

label_achievement_csat_91 = Label(
    achievement_frame, text='The number of CSAT only responses required to achieve >=91% CSAT: ')
label_achievement_csat_91.grid(row=2, column=0, padx=10, pady=10)

entry_achievement_csat_91 = Entry(
    achievement_frame, width=4, state='disabled', disabledbackground="white", disabledforeground="black")
entry_achievement_csat_91.grid(row=2, column=1, padx=10, pady=10)

label_achievement_dsat_5 = Label(
    achievement_frame, text='The number of CSAT only responses required to achieve <=5% DSAT: ')
label_achievement_dsat_5.grid(row=3, column=0, padx=10, pady=10)

entry_achievement_dsat_5 = Entry(achievement_frame, width=4, state='disabled',
                                 disabledbackground="white", disabledforeground="black")
entry_achievement_dsat_5.grid(row=3, column=1, padx=10, pady=10)

label_achievement_fcr_90 = Label(
    achievement_frame,
    text='The number of satisfied FCR only responses required to achieve 90% FCR: ')
label_achievement_fcr_90.grid(row=4, column=0, padx=10, pady=10)

entry_achievement_fcr_90 = Entry(achievement_frame, width=4, state='disabled',
                                 disabledbackground="white",
                                 disabledforeground="black")
entry_achievement_fcr_90.grid(row=4, column=1, padx=10, pady=10)

root.mainloop()