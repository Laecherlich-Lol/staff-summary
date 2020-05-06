import xlrd
import tkinter
from tkinter import ttk
import utils
import create_table


staff = 'O'
month = 1
window = tkinter.Tk()
window.title('选择文件')
window.geometry('500x300')
init = utils.Features(window)
select_file_button = tkinter.Button(
    window,
    text='选择文件',
    font=('Arial', 12),
    command=init.get_files
)
accept_opinions = tkinter.Button(
    window,
    text='确定',
    font=('Arial', 12),
    command=init.finish
)

t1 = ttk.Combobox(window,
                  values=['o', 'd', 'p'])
t2 = ttk.Combobox(window,
                  values=[months for months in range(1, 13)])


def callbackt1(event):
    global staff
    staff = t1.get()


def callbackt2(event):
     # print("New Element Selected")
     global month
     month = t2.get()


t1.bind("<<ComboboxSelected>>", callbackt1)
t2.bind("<<ComboboxSelected>>", callbackt2)

tkinter.Label(window, text='员工:', font=('Arial', 14)).place(x=10, y=30)
tkinter.Label(window, text='月份:', font=('Arial', 14)).place(x=10, y=60)
select_file_button.pack()
t1.pack()
t2.pack()
accept_opinions.pack()
window.mainloop()
print(staff, month)
data = xlrd.open_workbook(init.judge())
table = data.sheets()[int(month)-1]
new_table = create_table.Create(table, staff, month)
new_table.init_table()
new_table.creating()