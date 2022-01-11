from tkinter import *
from tkinter import filedialog, messagebox, ttk
from tkinter.filedialog import asksaveasfile
import openpyxl
import tkinter as tk
import pandas as pd
import numpy
from pathlib import Path
from PIL import Image, ImageTk

# BUTTON TO CHANGE SHEETS ON THE LISTBOX_SHEETS. SAME AS **VISUALIZAR ARQUIVO
# REMOVE DUPLICATED VALUES
# CREATE COLUMN (55), (NUMERO), (,), (NOME)

# photo = Image.open(r"C:\Users\felip\Downloads\LogoJ_Better.png")
# photo_icon = PhotoImage(file=r"C:\Users\felip\Pictures\Logo J.ico")

root = tk.Tk()
root.geometry("600x720")  # window size
root.pack_propagate(False)
root.resizable(1, 1)
root.title("Excel Software - by Felippe Nagy")
# root.iconbitmap(photo_icon)
#root.iconphoto(True, photo_icon)

azul_justa = "#076E95"
special_grey = "#f0f0f0"
root.config(background=special_grey)

#test = ImageTk.PhotoImage(photo)

#logo_justa = tk.Label(root, image=test, width=400, height=350)
#logo_justa.image = test
#logo_justa.pack(side=RIGHT, padx=60)

big_frame = tk.LabelFrame(root, background=azul_justa)
big_frame.place(height=800, width=600)

# image=photo #compound='bottom' -> this makes the pic goes behind our box
excel_view = tk.LabelFrame(big_frame, text='EXCEL',
                           bd=7, bg=azul_justa, fg="black")
excel_view.place(height=550, width=600, rely=0.08, relx=0)

frame_button = tk.LabelFrame(big_frame, background=azul_justa, fg='white')
frame_button.place(height=170, width=600, rely=0.785, relx=0)

# Frame for open file dialog
# relief== style of border(RAISED, SUNKEN).  bd=the size of the bord #You can use: font,fg(font color),
file_frame = tk.LabelFrame(big_frame, text='EXCEL FILE',
                           bg=azul_justa, fg='black')
# bg(background color), padx/pady
file_frame.place(height=80, width=420)

# Buttons
button_BrowseFile = tk.Button(file_frame, text='Browse file', command=lambda: File_dialog(), fg="white", bg="black",
                              activebackground="black", activeforeground="white", height=1, width=15)  # activeforegound=#COLOR -> This define the color that is going to be when click
# activebackground -> same thing but for background, #STATE=DISABLED -> no longer click. #Theres an image button option, just use
button_BrowseFile.place(rely=0.6, relx=0.748)
# image like label

button_LoadFile = tk.Button(file_frame, text="Load file", command=lambda: [
    Load_excel_data()], fg="white", bg="black", activebackground="black", activeforeground="white", height=1, width=15)
button_LoadFile.place(rely=0.6, relx=0.38)

button_ClearFile = tk.Button(file_frame, text='Clear file', command=lambda: clear_all(), fg="white", bg="black", activebackground="black",
                             activeforeground="white", height=1, width=15)  # activeforegound=#COLOR -> This define the color that is going to be when click
# activebackground -> same thing but for background, #STATE=DISABLED -> no longer click. #Theres an image button option, just use
button_ClearFile.place(rely=0.6, relx=-0.005)

label_file = ttk.Label(
    file_frame, text="No file selected", background=azul_justa)
label_file.place(rely=0.1, relx=0)

# TreeView Widget
tv1 = ttk.Treeview(excel_view)
tv1.place(relheight=1, relwidth=1)


treescrolly = tk.Scrollbar(excel_view, orient="vertical", command=tv1.yview)
treescrollx = tk.Scrollbar(excel_view, orient="horizontal", command=tv1.xview)

tv1.configure(xscrollcommand=treescrollx.set, yscrollcommand=treescrolly.set)

treescrollx.pack(side="bottom", fill="x")
treescrolly.pack(side="right", fill="y")


button_Columns = tk.Button(frame_button, text='Get columns', command=lambda: [get_columns(
)], fg="black", bg=special_grey, activebackground=special_grey, activeforeground="black", height=1, width=15)
button_Columns.grid(column=0, row=0)

button_newWindow = tk.Button(
    frame_button, text='Extract data', command=lambda: openNewWindow(), height=1, width=15, bg=special_grey, fg="red")
button_newWindow.grid(column=0, row=1)

button_HelpInfo = tk.Button(
    frame_button, text="Help", command=lambda: help_info(), bg=special_grey, fg="red", height=1, width=10)
button_HelpInfo.grid(column=8, row=0, padx=400)

button_changeSheet = tk.Button(big_frame, text="Change sheet", command=lambda: get_sheet_names(
), height=1, width=15, fg="white", bg="black", activebackground="black", activeforeground="white")
button_changeSheet.place(relx=0.755, rely=0.0685)

button_findDuplicates = tk.Button(frame_button, text="Duplicated values", command=lambda: find_duplicated(
), bg=special_grey, fg="black", height=1, width=15)
button_findDuplicates.grid(column=0, row=2)


def get_sheet_names():
    global xlsx_sheet, listbox_sheets
    button_changeSheet.place_forget()
    xlsx = pd.ExcelFile(filename)
    xlsx_sheet = xlsx.sheet_names
    listbox_sheets = Listbox(
        big_frame, width=30, height=len(xlsx_sheet), selectmode=SINGLE, background=azul_justa, foreground="black")
    listbox_sheets.place(relx=0.705)
    for item in xlsx_sheet:
        listbox_sheets.insert(END, item)

    def CurSelet(evt):  # function to return the values chosen in a list box by the user
        global choose_sheet, sheet_view
        # value=str(listbox_columns.get(listbox_columns.curselection()))
        choose_sheet = [listbox_sheets.get(idx)
                        for idx in listbox_sheets.curselection()]
        sheet_view = choose_sheet[0]
        print(sheet_view)

    listbox_sheets.bind('<<ListboxSelect>>', CurSelet)
    # label_file2["text"] = choose_sheet
    print(xlsx_sheet)
    return None


def File_dialog():  # browse the files on your computer
    global filename
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select file",
                                          filetype=(("xlsx files", "*.xlsx"), ("All Files", "*.*")))
    label_file["text"] = filename
    get_sheet_names()
    return None


def try_pandas():  # transform the archive to a dataframe to use for analysis
    global df
    file_path = label_file["text"]
    try:
        excel_filename = r"{}".format(file_path)
        df = pd.read_excel(excel_filename, sheet_name=sheet_view)
    except ValueError:  # If the above func fail, try ths one below
        tk.messagebox.showerror(
            "ERROR", "The file you chose is not valid!")
        return None
    except FileNotFoundError:  # Event that load and cannot recognize the file
        tk.messagebox.showerror("ERROR", f"No such file as {file_path}")
        return None


def Load_excel_data():  # load the archive chosen by the user
    def clear_data():  # this will clear the actual treeview to show another one chosen by the user
        tv1.delete(*tv1.get_children())
        clearData = []
        tv1["column"] = clearData
        return None

    try_pandas()
    clear_data()
    tv1["column"] = list(df.columns)
    tv1["show"] = "headings"

    for column in tv1["columns"]:
        tv1.heading(column, text=column)

    df_rows = df.to_numpy().tolist()
    for row in df_rows:
        tv1.insert("", "end", values=row)

    listbox_sheets.delete(0, END)
    listbox_sheets.place_forget()
    button_changeSheet.place(relx=0.705, rely=0.068)
    return None


def clear_all():  # this is a function to clear the listbox and the treeview anytime the user wants
    clearData = []
    label_file["text"] = "No file selectec"
    tv1["column"] = clearData
    tv1.delete(*tv1.get_children())
    try:
        listbox_columns.delete(0, END)
    except NameError:
        print("There is nothing to clear anymore")
    return None


def get_columns():  # this function will provide the columns and showed on the listbox, which can be chosen by the user
    global listbox_columns
    listbox_columns = Listbox(root, width=20, height=len(
        tv1["column"]), selectmode=MULTIPLE, background=azul_justa, foreground="white")
    listbox_columns.place(rely=0.5, relx=0.5)
    for item in tv1["column"]:
        listbox_columns.insert(END, item)
    for i in listbox_columns.curselection():
        print(listbox_columns.get(i))
    print(tv1["column"])
    return None


def create_column():
    pass


def find_duplicated():
    listbox_columns = Listbox(root, width=20, height=len(
        tv1["column"]), selectmode=SINGLE, background=azul_justa, foreground="white")
    listbox_columns.place(rely=0.5, relx=0.5)
    for item in tv1["column"]:
        listbox_columns.insert(END, item)

    def CurSelet(evt):  # function to return the values chosen in a list box by the user
        global checkDf, column_dp
        column_dp = [listbox_columns.get(idx)
                     for idx in listbox_columns.curselection()]
        print(column_dp)
        checkDf = df[column_dp]
    listbox_sheets.bind('<<ListboxSelect>>', CurSelet)
    button_goCheck = tk.Button(root, text="Check", height=1, width=15, background="green", foreground="white", activebackground="green", activeforeground="white",
                               command=lambda: duplicated_pandas())
    button_goCheck.place(rely=0.5, relx=0.55)

    def duplicated_pandas():
        global another_df, check_duplicates_sum, check_duplicates
        check_duplicates_sum = checkDf.duplicated(subset=column_dp).sum()
        check_duplicates = checkDf.duplicated(subset=column_dp)
        another_df = df
        another["Duplicatas"] = check_duplicates
        print(check_duplicates_sum)
        print(another_df.head(5))


def view_Treeview():
    pass


def openNewWindow():  # function to open another window and extract, view and analyse new columns
    try_pandas()
    newWindow = Toplevel(root)
    newWindow.title("Select columns")
    newWindow.geometry("800x500")
    newWindow.config(background=azul_justa)

    # Frame to organize labels and buttons
    frame1 = tk.LabelFrame(newWindow, background=azul_justa)
    frame1.place(height=400, width=455)

    # Frame to display de treeview
    frame2 = tk.LabelFrame(newWindow, bg=azul_justa)
    frame2.place(height=400, width=350, relx=0.57)

    button_viewData = tk.Button(
        frame1, text="View data", command=lambda: show_viewer(), height=1, width=15)  # button to see treeview
    button_viewData.grid(column=0, row=0, padx=0)

    frame_listboxColumn = tk.LabelFrame(frame1, background=azul_justa)
    frame_listboxColumn.place(height=450, width=200, relx=0.6)
    listbox_columns = Listbox(frame_listboxColumn, width=20, height=450,
                              selectmode=MULTIPLE, background=azul_justa, foreground="black")
    listbox_columns.place(relwidth=0.95, relheight=1)

    for item in tv1["column"]:
        listbox_columns.insert(END, item)

    button_saveNew = tk.Button(frame1, text="Save new file",
                               command=lambda: [required_entry_newfile()], height=1, width=15)
    button_saveNew.grid(column=0, row=1, padx=10)

    s1 = Label(frame1, text="Sheet name:", height=1,
               width=15, background=azul_justa)
    s1.grid(column=0, row=2)

    s2 = Label(frame1, text="Sheet name:", height=1,
               width=15, background=azul_justa)
    s2.grid(column=0, row=4)

    sheet_name_entry = tk.Entry(frame1)
    sheet_name_entry.grid(column=1, row=2)

    sheet_name_entry2 = tk.Entry(frame1)
    sheet_name_entry2.grid(column=1, row=4)

    button_extractData = tk.Button(
        frame1, text="Save on the file", command=lambda: [required_entry_samefile()], height=1, width=15)
    button_extractData.grid(column=0, row=3)
    # this command will return the values choosed

    tv2 = ttk.Treeview(frame2)
    tv2.place(relheight=1, relwidth=1)

    def show_viewer():  # function to show the new data which will be chosen by the user
        tv2["column"] = list(new_df.columns)
        tv2["show"] = "headings"
        for column in tv2["column"]:
            tv2.heading(column, text=column)
        df_rows = new_df.to_numpy().tolist()
        for row in df_rows:
            tv2.insert("", "end", values=row)

    treescrolly = tk.Scrollbar(frame2, orient="vertical", command=tv2.yview)
    treescrollx = tk.Scrollbar(frame2, orient="horizontal", command=tv2.xview)
    tv2.configure(xscrollcommand=treescrollx.set,
                  yscrollcommand=treescrolly.set)
    treescrollx.pack(side="bottom", fill="x")
    treescrolly.pack(side="right", fill="y")

    def CurSelet(evt):  # function to return the values chosen in a list box by the user
        global values, new_df
        # value=str(listbox_columns.get(listbox_columns.curselection()))
        values = [listbox_columns.get(idx)
                  for idx in listbox_columns.curselection()]
        print(values)
        new_df = df[values]
        print(new_df.head(5))

    listbox_columns.bind('<<ListboxSelect>>', CurSelet)
    print(tv1["column"])

    def required_entry_newfile():  # this function will make the entry (sheet_name) REQUIRED
        global input_sheet

        file_types = (
            ("All files", "*.*"),
            ("xlsx files", "*.xlsx"),
            ("CSV file", "*.csv"))
        save_file = asksaveasfile(title="Save file",
                                  filetypes=file_types, defaultextension=("xlsx files", "*.xlsx"))
        print(save_file)
        input_sheet = sheet_name_entry.get()
        if sheet_name_entry.get():
            print(sheet_name_entry.get())
        else:  # If the above func fail, try ths one below
            tk.messagebox.showerror(
                "WARNING", "PLEASE ENTER A FILE NAME")
            sheet_name_entry.focus_set()
            print("PLEASE WRITE SOMETHING")

    def required_entry_samefile():
        global input_sheet2
        input_sheet2 = sheet_name_entry2.get()
        if sheet_name_entry2.get():
            print(sheet_name_entry2.get())
        else:  # If the above func fail, try ths one below
            tk.messagebox.showerror(
                "WARNING", "PLEASE ENTER A FILE NAME")
            sheet_name_entry.focus_set()
            print("PLEASE WRITE SOMETHING")

        def extract_column():  # function to create a new file with new columns and filtered values to a new/used excel file
            with pd.ExcelWriter(filename, mode="a", engine="openpyxl") as writer:
                new_df.to_excel(writer, sheet_name=input_sheet2)
            print("File saved: {}, in {}".format(
                filename, input_sheet2))
        try:
            extract_column()
        except:
            print("ERROR, SOMETHING WRONG WITH THE SHEET ENTRY")

    def search_nFilter():
        label_filter = tk.Label(newWindow)
        filter_entry = tk.Entry(newWindow)
        pass

    return None


def help_info():  # function to teach the user how the programm works
    newWindow = Toplevel(root)
    newWindow.title("Help/Info")
    newWindow.geometry("400x600")
    newWindow.config(background=special_grey)
    #photo = Image.open(r"C:\Users\felip\Downloads\LogoJ_Better.png")
    #test = ImageTk.PhotoImage(photo)
    info = """
    This is a software created to help with some quick tools in excel as : extract columns, sort, visualize, count and more..
    if you got any problem or sugestion, please let me know contacting me on Github:
    github.com/felippefn
    """
    T = tk.Text(newWindow, height=600, width=400, bg=special_grey)
    T.pack()
    T.insert(tk.END, info)
    T.tag_add("start", "2.4", "2.53")
    T.tag_config("start", background="black", foreground="white")
    return None


root.mainloop()
