import tkinter
from tkinter import *
from tkinter import ttk
from tkinter.ttk import *
from singleton import Singleton
from tkinter import filedialog
from tkinter import messagebox
from generalactions import ParseAndEnter


class Window(Tk, Singleton):
    def init(self):
        super().__init__()

    def __init__(self):
        # Стиль кнопок
        self.style = Style()
        self.style.configure('TButton', font=
        ('calibri', 9, 'bold'), foreground='black',
                             borderwidth='2')
        # Инициализация путей  и переменных двух файлов
        self.pathData = None
        self.pathTemplate = None
        self.buttonWidth = 160
        self.buttonHeight = 50
        self.title("Автоматическое заполнение шаблона Excel")

        # Размеры окна
        self.minsize(600, 400)
        self.resizable(width=False, height=False)

        # icon and background
        self.iconbitmap("logo.ico")
        self.image = PhotoImage(file='background.png')
        bg_logo = Label(self, image=self.image)
        bg_logo.place(x=0, y=0)
        print("calling from __init__")
        self.buttonData()
        self.buttonTemplate()
        self.buttonEven()

    def inputPathData(self):
        self.pathData = filedialog.askopenfilename(filetypes=(("Excel Files", "*.xlsx"), ("all files", "*.*")))

    def inputPathTemplate(self):
        self.pathTemplate = filedialog.askopenfilename(filetypes=(("Excel Files", "*.xlsx"), ("all files", "*.*")))

    def EventWithPaths(self):
        if self.pathData is not None and self.pathTemplate is not None and self.pathData and self.pathTemplate:
            event = ParseAndEnter(self.pathTemplate, self.pathData)
            event.automatic()

        else:
            messagebox.showerror("Ошибка", "Не выбраны файлы")

    def buttonData(self):
        self.button = ttk.Button(self, text="Выбрать входной файл", command=self.inputPathData)
        self.button.place(x=40, y=200, width=self.buttonWidth, height=self.buttonHeight)

    def buttonTemplate(self):
        self.buttontemplate = ttk.Button(self, text="Выбрать \nзаполняемый \nфайл", command=self.inputPathTemplate)
        self.buttontemplate.place(x=40, y=260, width=self.buttonWidth, height=self.buttonHeight)

    def buttonEven(self):
        self.buttonAction = ttk.Button(self, text="Выполнить", command=self.EventWithPaths)
        self.buttonAction.place(x=40, y=320, width=self.buttonWidth, height=self.buttonHeight)
