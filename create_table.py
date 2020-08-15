import xlwt
import xlrd


class Create:
    def __init__(self, table, employee, month):
        self.table = table
        self.employee = employee
        self.false_status = {'Frei': '',
                             'Ur': 5,
                             'Kr': 6,
                             'Kr>3': 7
                             }
        self.true_status = {'Arb': '',
                            'Feiertag': ''
                            }
        self.row2 = ['Name:', employee, 'Monat:', month]
        self.row4 = ['Arbeit', 'Kommen', 'Gehen', 'Ist-St.', 'Ur', 'Kr', 'Kr>3']
        self.status_name = {'Arb': 'Arbeit',
                            'Feiertag': 'gesetz. Feiertag',
                            'Frei': 'Freizeit',
                            'Ur': 'Urlaub',
                            'Kr': 'Krank',
                            'Kr>3': 'Krank'}
        self.filename = month + ' ' + employee
        self.workbook = xlwt.Workbook()
        self.worksheet = self.workbook.add_sheet(month)
        self.two_rows = False
        self.row = 0
        self.col = 3
        self.new_table_row = 4

    def set_style(self, name, height, color, bold=False):
        style = xlwt.XFStyle()  # 初始化样式

        font = xlwt.Font()  # 为样式创建字体
        font.name = name
        font.bold = bold
        font.color_index = color
        font.height = height

        style.font = font
        return style

    def init_table(self):
        self.worksheet.write(0, 0, 'Monatübersicht für', self.set_style('Arial', 220, 0, True))
        for block in range(4):
            self.worksheet.write(1, block, self.row2[block], self.set_style('Arial', 220, 0, True))
        for block in range(7):
            self.worksheet.write(3, block+1, self.row4[block])

    def write(self, status):
        self.worksheet.write(self.new_table_row, 1, self.status_name[status])
        if self.table.cell(self.row, self.col + 1).value != '':
            self.worksheet.write(self.new_table_row, 2,
                                 str(xlrd.xldate_as_datetime(self.table.cell(self.row, self.col + 1).value, 1)).split(' ')[1])
            self.worksheet.write(self.new_table_row, 3,
                                 str(xlrd.xldate_as_datetime(self.table.cell(self.row, self.col + 2).value, 1)).split(' ')[1])
            self.worksheet.write(self.new_table_row, 4, self.table.cell(self.row, self.col+3).value)
            if self.two_rows and self.table.cell(self.row + 1, self.col + 1).value != '':
                self.new_table_row += 1
                self.worksheet.write(self.new_table_row, 2,
                                     str(xlrd.xldate_as_datetime(self.table.cell(self.row + 1, self.col + 1).value, 1)).split(' ')[1])
                self.worksheet.write(self.new_table_row, 3,
                                     str(xlrd.xldate_as_datetime(self.table.cell(self.row + 1, self.col + 2).value, 1)).split(' ')[1])
                self.worksheet.write(self.new_table_row, 4, self.table.cell(self.row+1, self.col+3).value)
        else:
            self.worksheet.write(self.new_table_row, 2,
                                 str(xlrd.xldate_as_datetime(self.table.cell(self.row + 1, self.col + 1).value, 1)).split(' ')[1])
            self.worksheet.write(self.new_table_row, 3,
                                 str(xlrd.xldate_as_datetime(self.table.cell(self.row + 1, self.col + 2).value, 1)).split(' ')[1])
            self.worksheet.write(self.new_table_row, 4, self.table.cell(self.row+1, self.col+3).value)

    def creating(self):
        for block in range(2, int(self.table.nrows)-1):
            if self.table.cell(block, 0).value == self.employee:
                self.row = block
                if self.table.cell(block + 1, 0).value == '' and block + 1 < self.table.nrows:
                    self.two_rows = True
                break
        while self.table.cell(0, self.col).value != 'Monate Stunden':
            # print(self.table.cell(0, self.col).value)  #TODO: fix empty column issue
            dates = xlrd.xldate_as_tuple(float(self.table.cell(0, self.col).value), 0)
            date = str(dates[2]) + '.' + str(dates[1]) + '.' + str(dates[0])
            self.worksheet.write(self.new_table_row, 0, date)
            if self.table.cell(self.row, self.col).value in self.true_status:
                self.write(self.table.cell(self.row, self.col).value)
            else:
                status = self.table.cell(self.row, self.col).value
                self.worksheet.write(self.new_table_row, 1, self.status_name[status])
                if status != 'Frei':
                    self.worksheet.write(self.new_table_row, self.false_status[status], self.table.cell(self.row, self.col+3).value)
            self.col += 4
            self.new_table_row += 1
        self.worksheet.write(self.new_table_row, 0, 'Summe:')
        self.worksheet.write(self.new_table_row, 4, self.table.cell(self.row, self.col).value)
        self.worksheet.write(self.new_table_row, 5, self.table.cell(self.row, self.col+5).value)
        self.worksheet.write_merge(self.new_table_row, self.new_table_row, 6, 7, self.table.cell(self.row, self.col+4).value)
        self.new_table_row += 1
        self.worksheet.write(self.new_table_row, 1, 'Gleitzeit:')
        self.worksheet.write(self.new_table_row, 2, 'Urlaub:')
        self.worksheet.write(self.new_table_row, 4, 'Freizeit:')
        self.worksheet.write(self.new_table_row, 5, self.table.cell(self.row, self.col+3).value)
        self.new_table_row += 1
        self.worksheet.write(self.new_table_row, 0, 'Vormonat:')
        self.worksheet.write(self.new_table_row, 1, self.table.cell(self.row, 1).value)
        self.worksheet.write(self.new_table_row, 2, self.table.cell(self.row, 2).value)
        self.worksheet.write(self.new_table_row, 4, 'Arbeitstag:')
        self.worksheet.write(self.new_table_row, 5, self.table.cell(self.row, self.col+2).value)
        self.new_table_row += 1
        self.worksheet.write(self.new_table_row, 0, 'Saldo:')
        self.worksheet.write(self.new_table_row, 1, self.table.cell(self.row, self.col+1).value)
        self.worksheet.write(self.new_table_row, 2, self.table.cell(self.row, self.col+6).value)
        self.worksheet.write(self.new_table_row, 4, 'Krank:')
        self.worksheet.write(self.new_table_row, 5, self.table.cell(self.row, self.col+4).value)
        self.new_table_row += 1
        self.worksheet.write(self.new_table_row, 4, 'Urlaub:')
        self.worksheet.write(self.new_table_row, 5, self.table.cell(self.row, self.col+5).value)
        self.workbook.save(self.filename+'.xls')



