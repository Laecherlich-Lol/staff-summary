import xlrd
import tkinter
import utils
import create_table
from SelectionBox import SelectionBox


def init_window(title, geometry):
    window = tkinter.Tk()
    window.title(title)
    window.geometry(geometry)

    return window


def main():

    window = init_window("Staff Summary", '500x300')

    init = utils.Features(window)

    select_file_button = tkinter.Button(
        window,
        text='选择文件',
        font=('Arial', 12),
        command=init.get_files
    )
    select_file_button.pack()

    staff_selection = SelectionBox(window, "Staff Name", dict(x=125, y=30))
    month_selection = SelectionBox(window, "Month", dict(x=125, y=10))

    staff_selection.set_values(['Olaf Thomsen', 'Manfred Reyelt', 'Doris Rüdiger'])
    month_selection.set_values([months for months in range(1, 13)])


    accept_opinions = tkinter.Button(
        window,
        text='确定',
        font=('Arial', 12),
        command=init.finish
    )
    accept_opinions.pack()

    window.mainloop()

    data = xlrd.open_workbook(init.get_filename())

    staff = staff_selection.selected
    month = month_selection.selected
    table = data.sheets()[int(month)-1]

    new_table = create_table.Create(table, staff, month)
    new_table.init_table()
    new_table.creating()


if __name__ == "__main__":
    main()