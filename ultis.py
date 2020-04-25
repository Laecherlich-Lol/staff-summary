import tkinter
from tkinter import filedialog


class Features:
    def __init__(self, interface):
        self.interface = interface
        self.interface.filename = None

    def get_files(self):
        self.interface.filename = filedialog.askopenfilename(initialdir='/',
                                                             filetypes=(('excel files', '*.xlsx'),
                                                                        ('all files', '*.*')), )

    def finish(self):
        self.interface.destroy()

    def put_in(self, strVar):
        return tkinter.Entry(self.interface, textvariable=strVar, show=None, font=('Arial', 14))

    def judge(self):
        return self.interface.filename if self.interface.filename is not None else ''



