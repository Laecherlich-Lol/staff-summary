import xlrd
import tkinter
from tkinter import ttk
import utils
import create_table


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
t1 = tkinter.StringVar()
t2 = ttk.Combobox(window, 
                            values=[
                                    "January", 
                                    "February",
                                    "March",
                                    "April"])

def callbackFunc(event):
     print("New Element Selected")
     print(t2.get())


t2.bind("<<ComboboxSelected>>", callbackFunc)

text1 = init.put_in(t1)
tkinter.Label(window, text='员工:', font=('Arial', 14)).place(x=10, y=30)
tkinter.Label(window, text='月份:', font=('Arial', 14)).place(x=10, y=60)
select_file_button.pack()
text1.pack()
t2.pack()
accept_opinions.pack()
window.mainloop()
data = xlrd.open_workbook(init.judge())
table = data.sheets()[int(t2.get())-1]
new_table = create_table.Create(table, t1.get(), t2.get())
new_table.init_table()
new_table.creating()