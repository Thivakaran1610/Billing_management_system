import sys
import os
import openpyxl

sys.path.append('D:\\Billing')

from tkinter import *
import tkinter.font as tfnt
import tkinter.ttk as ttk
from Work_Here.create_workbook.create_workbook import createbook as cb
from Work_Here.create_workbook.data import months, getlist, years, name_list
from Work_Here.create_workbook.editingsheets import editsheet, list_type

root = Tk()
root.geometry("400x650")
root.resizable(0, 0)


class Home:
    def __init__(self):
        self.MS = MonthlySheet()

    def home(self):
        style = tfnt.Font(family="Gabriola", size=15)
        bgcol = "Teal"
        fgcol = "White"

        frame0 = Frame(root, bg=bgcol)
        frame0.place(x=0, y=0, width=400, height=650)

        title = Label(frame0, text="Welcome to finance Management Software", fg=fgcol, bg=bgcol,
                      font=tfnt.Font(family="Trebuchet MS", size=14))
        title.place(x=15, y=25)

        margin = Label(frame0, text="- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -",
                       fg=fgcol,
                       bg=bgcol, font=tfnt.Font(size=14))
        margin.place(x=0, y=60)

        button1 = Button(frame0, text="Monthly View", font=style, command=self.MS.monthlyview)
        button1.place(x=120, y=118, width=160, height=65)

        button2 = Button(frame0, text="Customer Bill", font=style)
        button2.place(x=120, y=251, width=160, height=65)

        button3 = Button(frame0, text="Daily Expenditure", font=style)
        button3.place(x=120, y=384, width=160, height=65)

        button4 = Button(frame0, text="Market Sheet", font=style)
        button4.place(x=120, y=517, width=160, height=65)


class MonthlySheet:
    def __init__(self):
        self.date_num = IntVar(value=1)
        self.code_number = IntVar(value=0)
        self.active_file = StringVar(value="July.xlsx")

    def nexttab(self, event):
        event.widget.tk_focusNext().focus()
        return "break"

    def popup(self):
        popup = Toplevel(root)
        popup.title("New Workbook")
        popup.geometry("400x200")
        bgcol = "LightPink"
        fgcol = "DarkSlateGray"
        popup.resizable(0, 0)
        frame2_1 = Frame(popup, bg=bgcol)
        frame2_1.place(x=0, y=0, width=400, height=200)
        fil1 = Label(popup, text="File Name :", font=tfnt.Font(family="Georgia", size=15), bg=bgcol, fg=fgcol)
        fil1.place(x=20, y=20)
        e_file1 = Entry(popup, font=tfnt.Font(size=13))
        e_file1.place(x=130, y=23, width=250, height=30)
        mon1 = Label(popup, text="Month :", font=tfnt.Font(family="Georgia", size=15), bg=bgcol, fg=fgcol)
        mon1.place(x=50, y=75)
        e_mon1 = ttk.Combobox(popup, values=months)
        e_mon1.place(x=130, y=75, width=100, height=30)
        e_mon1.set(months[8])
        year = Label(popup, text="Year :", font=tfnt.Font(family="Georgia", size=15), bg=bgcol, fg=fgcol)
        year.place(x=250, y=75)
        e_year = ttk.Combobox(popup, values=years, justify=CENTER, font=tfnt.Font(size=13))
        e_year.place(x=310, y=75, width=80, height=30)
        e_year.set(2023)

        def create():
            alert = Toplevel(root)
            alert.title("Successful")
            alert.geometry("250x100")
            fontstyle = tfnt.Font(family="Franklin Gothic Demi Cond", size=13)

            def close():
                alert.destroy()
                popup.destroy()
                hm = Home()
                hm.home()

            if e_file1.get() != "":
                cb(e_file1.get(), months.index(e_mon1.get()) + 1, e_year.get())
                l1 = Label(alert, text="Your File Have Been Saved!!!", font=fontstyle)
                l1.pack()
                b1 = Button(alert, text="OK", font=fontstyle, command=close)
                b1.pack()
            else:
                l1 = Label(alert, text="Enter a File Name!!!", font=fontstyle)
                l1.pack()
                b1 = Button(alert, text="OK", font=fontstyle, command=alert.destroy)
                b1.pack()

        def cancel():
            popup.destroy()

        b_create = Button(popup, text="Create", bg="LightCoral", fg="FloralWhite",
                          font=tfnt.Font(family="Georgia", size=17),
                          command=create)
        b_create.place(x=30, y=130, width=150, height=40)
        b_cancel = Button(popup, text="Cancel", bg="LightCoral", fg="FloralWhite",
                          font=tfnt.Font(family="Georgia", size=17), command=cancel)
        b_cancel.place(x=220, y=130, width=150, height=40)
        popup.mainloop()

    def monthlyview(self):
        fontstyle = tfnt.Font(family="Franklin Gothic Demi Cond", size=15)
        bgcol = "AliceBlue"
        fgcol = "Black"

        frame2 = Frame(bg=bgcol)
        frame2.place(x=0, y=0, width=400, height=650)

        u1 = Label(frame2, text="~ Monthly Sheet ~", fg=fgcol, bg=bgcol, font=tfnt.Font(family="Lucida Fax", size=18))
        u1.place(x=100, y=20)

        un1 = Label(frame2, text="-----------------------------------------------------------------------------",
                    fg=fgcol,
                    bg=bgcol, font=tfnt.Font(size=14))
        un1.place(x=0, y=55)

        b0 = Button(frame2, text="Home", font=tfnt.Font(family="Latha", size=10), bg=bgcol, fg=fgcol, command=hm.home,
                    activebackground="Teal", activeforeground="White")
        b0.place(x=0, y=0, width=60, height=70)

        file = Label(frame2, text="File :", fg=fgcol, bg=bgcol,
                     font=tfnt.Font(family="Franklin Gothic Demi Cond", size=18))
        file.place(x=20, y=95)

        e_sel_file = ttk.Combobox(frame2, textvariable=self.active_file, font=fontstyle, values=getlist())
        e_sel_file.place(x=70, y=100, width=190)
        self.active_file.set(e_sel_file.get())

        MS = MonthlySheet()
        b_new = Button(frame2, text="New", font=tfnt.Font(family="Latha", size=13), bg="LightBlue", fg=fgcol,
                       command=MS.popup)
        b_new.place(x=280, y=92, width=100, height=40)

        dat = Label(frame2, text="Date :", fg=fgcol, bg=bgcol, font=fontstyle)
        dat.place(x=25, y=175)

        def adddate():
            self.date_num.set(self.date_num.get() + 1)

        def subdate():
            self.date_num.set(self.date_num.get() - 1)

        def reload():
            try:
                name['text'] = name_list[self.code_number.get()]
                name['fg'] = fgcol
                e_variety['state'] = ACTIVE
                e_variety['values'] = list_type(int(self.code_number.get()))
                e_variety.set(list_type(int(self.code_number.get()))[0])
                print(self.code_number.get())
            except Exception as e:
                name['text'] = "Error NO user Found"
                name['fg'] = 'Red'
                e_variety['state'] = DISABLED

        b_adddate = Button(frame2, text="+", font=tfnt.Font(size=13), bg="LightBlue", fg=fgcol, command=adddate)
        b_adddate.place(x=160, y=175, width=30)
        b_subdate = Button(frame2, text="-", font=tfnt.Font(size=13), bg="LightBlue", fg=fgcol, command=subdate)
        b_subdate.place(x=90, y=175, width=30)
        e_date = Label(frame2, font=fontstyle, textvariable=self.date_num, fg=fgcol, bg="White", justify=CENTER)
        e_date.place(x=120, y=175, width=40)

        code_num = Label(frame2, text="Code :", fg=fgcol, bg=bgcol, font=fontstyle)
        code_num.place(x=25, y=250)

        e_code_num = Entry(frame2, font=fontstyle, textvariable=self.code_number, justify=CENTER)
        e_code_num.place(x=90, y=250, width=50)

        name = Label(frame2, font=fontstyle, fg=fgcol, bg=bgcol)
        name.place(x=25, y=280)

        b_try = Button(frame2, text="Try", font=fontstyle, bg="LightBlue", fg=fgcol, command=reload)
        b_try.place(x=160, y=250, width=50, height=30)

        variety = Label(frame2, text="Type :", fg=fgcol, bg=bgcol, font=fontstyle)
        variety.place(x=225, y=250)

        e_variety = ttk.Combobox(frame2, font=fontstyle, values=list_type(int(self.code_number.get())))
        e_variety.place(x=275, y=250, width=110)

        weight = Label(frame2, text="Quantity :", fg=fgcol, bg=bgcol, font=fontstyle)
        weight.place(x=5, y=470)

        e_weight = Entry(frame2, font=fontstyle)
        e_weight.place(x=90, y=470, width=100)

        price = Label(frame2, text="Amount :", fg=fgcol, bg=bgcol, font=fontstyle)
        price.place(x=200, y=470)

        e_price = Entry(frame2, font=fontstyle)
        e_price.place(x=275, y=470, width=110)

        prev = Button(frame2, text="<-- Previous", font=tfnt.Font(family="Latha", size=13), bg="LightBlue", fg=fgcol)
        prev.place(x=30, y=595, width=150, height=40)

        def clear():
            e_code_num.delete(0, END)
            e_price.delete(0, END)
            e_weight.delete(0, END)
            e_variety.delete(0, END)

        def Add():
            if str.isdigit(e_price.get()) and str.isdigit(e_weight.get()):
                editsheet(path=self.active_file.get(), book_code=int(e_code_num.get()), variety=e_variety.get(),
                          date=self.date_num.get(),
                          pri_val2=int(e_price.get()),
                          wei_val1=int(e_weight.get()))
                e_price.delete(0, END)
                e_weight.delete(0, END)
                l_error = Label(frame2, text="Error: Value not Added!!!", font=tfnt.Font(size=11), fg=bgcol, bg=bgcol)
                l_error.place(x=135, y=510)
                print("successful")
            else:
                l_error = Label(frame2, text="Error: Value not Added!!!", font=tfnt.Font(size=11), fg="Red", bg=bgcol)
                l_error.place(x=135, y=510)
                print("failure")

        nex = Button(frame2, text="Next -->", font=tfnt.Font(family="Latha", size=13), bg="LightBlue", fg=fgcol)
        nex.place(x=220, y=595, width=150, height=40)

        edit = Button(frame2, text="Edit", font=tfnt.Font(family="Latha", size=13), bg="LightBlue", fg=fgcol)
        edit.place(x=20, y=535, width=100, height=40)

        clr = Button(frame2, text="Clear", font=tfnt.Font(family="Latha", size=13), bg="LightBlue", fg=fgcol,
                     command=clear)
        clr.place(x=150, y=535, width=100, height=40)

        add = Button(frame2, text="Add", font=tfnt.Font(family="Latha", size=13), bg="LightBlue", fg=fgcol, command=Add)
        add.place(x=280, y=535, width=100, height=40)

        e_code_num.bind("<Return>", MS.nexttab)
        e_weight.bind("<Return>", MS.nexttab)
        e_price.bind("<Return>", MS.nexttab)

        print(self.active_file.get())


hm = Home()
hm.home()

root.mainloop()
