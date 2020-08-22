import tkinter as tk
from tkinter import ttk
from tkinter import LEFT
from tkinter import RIGHT


class SelectionBox:
    def __init__(self, parent_frame, label_name, position, values=None, size=3, font=('Arial', 14)):
        self.parent_frame = parent_frame
        self.label_name = label_name
        self.values = values
        self.font = font
        self.position = position
        self.size = size
        self.frame = None
        self.selected = None
        self.combobox = None
        self.init_frame()

        if values is not None:
            self.setup_combobox()

    def init_frame(self):

        self.frame = tk.Frame(master=self.parent_frame)
        self.frame.pack(padx=self.position['x'], pady=self.position['y'], anchor='e')

        tk.Label(self.frame, text=self.label_name, font=self.font).pack(side=LEFT)

    def set_values(self, values):
        self.values = values
        self.setup_combobox()

    def setup_combobox(self):

        self.combobox = ttk.Combobox(self.frame, values=self.values)
        self.combobox.bind("<<ComboboxSelected>>", self.select_callback)

        self.combobox.pack()

    def select_callback(self, event):
        self.selected = self.combobox.get()




