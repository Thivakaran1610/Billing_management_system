from tkinter import *
import tkinter.font as tfnt
from tkinter import ttk


root = Tk()
root.geometry("400x650")
root.resizable(0, 0)


class HomeScreen:
    def __init__(self):
        pass

    def homescreen(self):
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

        button1 = Button(frame0, text="Monthly View", font=style)
        button1.place(x=120, y=118, width=160, height=65)

        button2 = Button(frame0, text="Customer Bill", font=style)
        button2.place(x=120, y=251, width=160, height=65)

        button3 = Button(frame0, text="Daily Expenditure", font=style)
        button3.place(x=120, y=384, width=160, height=65)

        button4 = Button(frame0, text="Market Sheet", font=style)
        button4.place(x=120, y=517, width=160, height=65)


class MonthlySheet:
    def __init__(self):
        pass

    def next_tab(self, event):
        event.widget.tk_focusNext().focus()
        return "Break"

    def popup(self):
        popup = Toplevel(root)
        popup.title("New WorkBook")
        popup.geometry("400x200")
        bgcol = "LightPink"
        fgcol = "DarkSlateGray"
        popup.resizable(0, 0)
        fram = Frame(popup, bg=bgcol)
        fram.place(x=0, y=0, width=400, height=200)
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






root.mainloop()
