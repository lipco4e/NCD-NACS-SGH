import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
import openpyxl 
import os

def load_data():
    path = "C:\\NACS\\nacs.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active

    
    list_values = list(sheet.values)
    print(list_values)
    for col_name in list_values[0]:
        treeview.heading(col_name, text=col_name)

    for value_tuple in list_values[1:]:
        treeview.insert('', tk.END, values=value_tuple)

def insert_row():

    # filename = "nacs.xlsx"

    #try: #error handling line begins

        # if os.path.isfile(filename):
        #     # Open the file
        #     input_file = open(filename, 'rb')

        # else:
        #     # create the file
        #     input_file = open(filename, 'wb')

        #     #close file into a variable
        #     input_file.close()
        #     #Error message if key is missing
        #     #resp_text.insert(END, f"\n\nOops, you need an API Key to speak to ChatGPT. Get one from here:\nhttps://beta.openai.com/account/api-keys")        
            
    date = cal.get()
    art = art_entry.get()
    name = name_entry.get()
    age = int(age_spinbox.get())
    Sex = sex_combobox.get()
    tb_status = TPT_combobox.get()
    cancer_status = cacx_combobox.get()
    #bp_status = "Done" if a.get() else "Not done"
    systolic = int(systolic_spinbox.get())
    diastolic = int(diastolic_spinbox.get())
    height = int(systolic_spinbox.get())
    weight = int(diastolic_spinbox.get())
    bp_gen = "Yes" if a.get() else "Not done"
    bmi_gen = "Yes" if a.get() else "Not done"

    #except Exception as e: #error handling line ends
       # resp_text.insert(END, f"\n\nOops, there was an error in your data check and see if you entered everything \n\n{e}")
    
 

    print(date, art, name, age, Sex, tb_status, cancer_status, bp_status, systolic, diastolic, height, weight, bp_gen, bmi_gen)

    # Insert row into Excel sheet
    path = "C:\\NACS\\nacs.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    row_values = [date, art, name, age, Sex, tb_status, cancer_status, bp_status, systolic, diastolic, height, weight, bp_gen, bmi_gen]
    sheet.append(row_values)
    workbook.save(path)

    # Insert row into treeview
    treeview.insert('', tk.END, values=row_values)
    
    # Clear the values
    cal.delete(0, "end")
    cal.insert(0, "Date")
    art_entry.delete(0, "end")
    art_entry.insert(0, "Art#")
    name_entry.delete(0, "end")
    name_entry.insert(0, "Name")
    age_spinbox.delete(0, "end")
    age_spinbox.insert(0, "Age")
    sex_combobox.set(combo_list0[0])
    TPT_combobox.set(combo_list[0])
    cacx_combobox.set(combo_list[0])
    checkbutton.state(["!selected"])
    systolic_spinbox.delete(0, "end")
    systolic_spinbox.insert(0, "Systolic")
    diastolic_spinbox.delete(0, "end")
    diastolic_spinbox.insert(0, "Diastolic")
    height_spinbox.delete(0, "end")
    height_spinbox.insert(0, "Height")
    weight_spinbox.delete(0, "end")
    weight_spinbox.insert(0, "Weight")
    checkbutton.state(["!selected"])
    checkbutton.state(["!selected"])


def toggle_mode():
    if mode_switch.instate(["selected"]):
        style.theme_use("forest-light")
    else:
        style.theme_use("forest-dark")

root = tk.Tk()

root.title("Nutrition Assessment, Counselling and Support (NACS)")

style = ttk.Style(root)
root.tk.call("source", "forest-light.tcl")
root.tk.call("source", "forest-dark.tcl")
style.theme_use("forest-dark")

combo_list0 = ["Sex", "M", "F"]
combo_list = ["Eligible TPT?", "Yes", "No"]
combo_list1 = ["Due for CaCx?","Yes", "No"]

#
frame = ttk.Frame(root)
frame.pack()

widgets_frame = ttk.LabelFrame(frame, text="Insert RoC Data")
widgets_frame.grid(row=0, column=0, padx=20, pady=10)

cal = DateEntry(widgets_frame, selectmode='day')
#cal.insert(0, "Date")
#cal.bind("<FocusIn>", lambda e: cal.delete('0', 'end'))
cal.grid(row=0, column=0, padx=5, pady=(0, 5), sticky="ew")

art_entry = ttk.Entry(widgets_frame)
art_entry.insert(0, "ARTnumber")
art_entry.bind("<FocusIn>", lambda e: art_entry.delete('0', 'end'))
art_entry.grid(row=1, column=0, padx=5, pady=(0, 5), sticky="ew")

name_entry = ttk.Entry(widgets_frame)
name_entry.insert(0, "Names")
name_entry.bind("<FocusIn>", lambda e: name_entry.delete('0', 'end'))
name_entry.grid(row=2, column=0, padx=5, pady=(0, 5), sticky="ew")

age_spinbox = ttk.Spinbox(widgets_frame, from_=18, to=100)
age_spinbox.insert(0, "Age")
age_spinbox.bind("<FocusIn>", lambda e: age_spinbox.delete('0', 'end'))
age_spinbox.grid(row=3, column=0, padx=5, pady=5, sticky="ew")

sex_combobox = ttk.Combobox(widgets_frame, values=combo_list0)
sex_combobox.current(0)
sex_combobox.grid(row=4, column=0, padx=5, pady=5,  sticky="ew")

TPT_combobox = ttk.Combobox(widgets_frame, values=combo_list)
TPT_combobox.current(0)
TPT_combobox.grid(row=5, column=0, padx=5, pady=5,  sticky="ew")

cacx_combobox = ttk.Combobox(widgets_frame, values=combo_list1)
cacx_combobox.current(0)
cacx_combobox.grid(row=6, column=0, padx=5, pady=5,  sticky="ew")

systolic_spinbox = ttk.Spinbox(widgets_frame, from_=60, to=200)
systolic_spinbox.insert(0, "BP Systolic")
systolic_spinbox.bind("<FocusIn>", lambda e: systolic_spinbox.delete('0', 'end'))
systolic_spinbox.grid(row=7, column=0, padx=5, pady=5, sticky="ew")

diastolic_spinbox = ttk.Spinbox(widgets_frame, from_=60, to=200)
diastolic_spinbox.insert(0, "BP Diastolic")
diastolic_spinbox.bind("<FocusIn>", lambda e: diastolic_spinbox.delete('0', 'end'))
diastolic_spinbox.grid(row=8, column=0, padx=5, pady=5, sticky="ew")

height_spinbox = ttk.Spinbox(widgets_frame, from_=20, to=200)
height_spinbox.insert(0, "Height")
height_spinbox.bind("<FocusIn>", lambda e: height_spinbox.delete('0', 'end'))
height_spinbox.grid(row=9, column=0, padx=5, pady=5, sticky="ew")

weight_spinbox = ttk.Spinbox(widgets_frame, from_=5, to=150)
weight_spinbox.insert(0, "Weight")
weight_spinbox.bind("<FocusIn>", lambda e: weight_spinbox.delete('0', 'end'))
weight_spinbox.grid(row=10, column=0, padx=5, pady=5, sticky="ew")

a = tk.BooleanVar()
checkbutton = ttk.Checkbutton(widgets_frame, text="Generate BP Status?", variable=a)
checkbutton.grid(row=11, column=0, padx=5, pady=5, sticky="nsew")

a = tk.BooleanVar()
checkbutton = ttk.Checkbutton(widgets_frame, text="Generate BMI?", variable=a)
checkbutton.grid(row=12, column=0, padx=5, pady=5, sticky="nsew")

button = ttk.Button(widgets_frame, text="Add RoC data", command=insert_row)
button.grid(row=13, column=0, padx=5, pady=5, sticky="nsew")

separator = ttk.Separator(widgets_frame)
separator.grid(row=14, column=0, padx=(20, 10), pady=10, sticky="ew")

mode_switch = ttk.Checkbutton(
    widgets_frame, text="Mode", style="Switch", command=toggle_mode)
mode_switch.grid(row=15, column=0, padx=5, pady=10, sticky="nsew")

treeFrame = ttk.Frame(frame)
treeFrame.grid(row=0, column=1, pady=10)
treeScroll = ttk.Scrollbar(treeFrame)
treeScroll.pack(side="right", fill="y")

cols = ("Date", "ARTnumber", "Names", "Age", "Sex", "TPT", "cacx", "BP", "Systolic", "Diastolic", "Height", "Weight", "bp_gen", "bmi_gen")
treeview = ttk.Treeview(treeFrame, show="headings",
                        yscrollcommand=treeScroll.set, columns=cols, height=30)

treeview.column("Date", width=60)
treeview.column("ARTnumber", width=150)
treeview.column("Names", width=100)
treeview.column("Age", width=50)
treeview.column("Sex", width=50)
treeview.column("TPT", width=100)
treeview.column("cacx", width=100)
treeview.column("BP", width=70)
treeview.column("Systolic", width=70)
treeview.column("Diastolic", width=70)
treeview.column("Height", width=70)
treeview.column("Weight", width=70)
treeview.column("bp_gen", width=70)
treeview.column("bmi_gen", width=70)
treeview.pack()
treeScroll.config(command=treeview.yview)
load_data()



root.mainloop()
