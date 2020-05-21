from tkinter import *
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText

import org_count


def selectFile():
    org_count.main()
    pass


def main():
    logs.grid(row=0, column=0, rowspan=6)

    Button(root, text='选择文件', command=selectFile).grid(row=0, column=1)
    Entry(root, textvariable=pathFile).grid(row=1, column=1)
    tabControl.grid(row=2, column=1, rowspan=4)
    root.mainloop()


root = Tk()
logs = ScrolledText(root, width=40, height=30)
pathFile = StringVar()
classes = StringVar()
classes.set('晨曦 晨光 曙光 朝阳 旭日')
chVarDis = BooleanVar()
check1 = Checkbutton(root, text="部分修改", variable=chVarDis)
listbox = Listbox(root, selectmode=MULTIPLE)
startline = -1

# Tab Control introduced here --------------------------------------
tabControl = ttk.Notebook(root)  # Create Tab Control

tab1 = ttk.Frame(tabControl)  # Create a tab
tabControl.add(tab1, text='学员分班')  # Add the tab

tab2 = ttk.Frame(tabControl)  # Add a second tab
tabControl.add(tab2, text='第二页')  # Make second tab visible

tab3 = ttk.Frame(tabControl)  # Add a third tab
tabControl.add(tab3, text='远程监测')  # Make second tab visible
monty = ttk.LabelFrame(tab3, text='控件示范区1', width=40, height=30)
monty.grid(column=0, row=0, padx=8, pady=4)
Button(monty, text='开始', ).grid(row=0, column=0)

if __name__ == '__main__':
    main()