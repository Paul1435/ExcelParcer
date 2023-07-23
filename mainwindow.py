from tkinter import *
import tkinter as tk
from tkinter import ttk
from tkinter.ttk import *
from singleton import Singleton
from tkinter import filedialog
from tkinter import messagebox
from generalactions import ParseAndEnter
import webbrowser


class Window(Tk, Singleton):
    def init(self):
        super().__init__()

    def __init__(self):
        # Стиль кнопок
        self.style = Style()
        self.style.configure('TButton', font=
        ('calibri', 9, 'bold'), foreground='black',
                             borderwidth=0)
        # Инициализация путей  и переменных двух файлов
        self.pathData = None
        self.pathTemplate = None
        self.buttonWidth = 160
        self.buttonHeight = 50
        self.title("Автоматическое заполнение шаблона Excel")

        # Размеры окна
        self.minsize(600, 400)
        self.resizable(width=False, height=False)

        # меню

        menu = Menu(self)
        self.config(menu=menu)
        file_menu = Menu(menu)
        menu.add_cascade(label="Полезная инфа", menu=file_menu)
        file_menu.add_command(label="Создатели", command=self.show_creators)
        file_menu.add_separator()
        file_menu.add_command(label="Примечание", command=self.instruction)

        # icon and background
        self.iconbitmap("logo.ico")
        self.image = PhotoImage(file='background2.png')
        bg_logo = Label(self, image=self.image)
        bg_logo.place(x=0, y=0)
        print("calling from __init__")
        self.buttonData()
        self.buttonTemplate()
        self.buttonEven()
        self.creators_window1 = None
        self.creators_window2 = None

    def instruction(self):
        if self.creators_window2 is not None and self.creators_window2.winfo_exists():
            self.creators_window2.lift()
            return
        self.creators_window2 = tk.Toplevel(self)
        self.creators_window2.title("Инструкция и рекомендации")

        text = "Внимание, это пробная версия заполнения шаблона.\n\n" \
               "Не исключены баги и подвисания, связанные как с графической оболочкой, так и при заполнении шаблона. \n" \
               "Убедительная просьба перед заполнением закрыть обе таблицы и сделать копии на случай неудачного заполнения.\n" \
               "Заполнение происходит в течение 5-10 минут, если строки не окажется в шаблоне,\n" \
               "то выскочит ошибка о неудачном заполнении строки \n" \
               "Если по каким-то причинам программа перестала работать, свяжитесь с одним из создателей."
        label = tk.Label(self.creators_window2, text=text, padx=20, pady=20)
        label.pack()

    def show_creators(self):
        if self.creators_window1 is not None and self.creators_window1.winfo_exists():
            self.creators_window1.lift()
            return
        self.creators_window1 = tk.Toplevel(self)
        self.creators_window1.title("Создатели")

        text = "Это пробная версия заполнения шаблона.\n\n" \
               "Ссылка на репозиторий: https://github.com/Paul1435 \n" \
               "Авторы:\n" \
               "Шестаков Павел\n" \
               "Шестакова Маргарита \n" \
               "Коркин Андрей"
        label = tk.Label(self.creators_window1, text=text, padx=20, pady=20)
        label.pack()
        label.bind("<Button-1>", self.open_link)

    def open_link(self, event):
        url = "https://github.com/Paul1435"
        webbrowser.open(url)

    def inputPathData(self):
        self.pathData = filedialog.askopenfilename(filetypes=(("Excel Files", "*.xlsx"), ("all files", "*.*")))

    def inputPathTemplate(self):
        self.pathTemplate = filedialog.askopenfilename(filetypes=(("Excel Files", "*.xlsx"), ("all files", "*.*")))

    def EventWithPaths(self):
        if self.pathData is not None and self.pathTemplate is not None and self.pathData and self.pathTemplate:
            event = ParseAndEnter(self.pathTemplate, self.pathData)

            def update_progress(progress_value):
                self.progressbar["value"] = progress_value
                self.update()

            event.automatic(update_progress)
            self.progressbar["value"] = 0
        else:
            messagebox.showerror("Ошибка", "Не выбраны файлы")

    def buttonData(self):
        self.button = ttk.Button(self, text="Выбрать файл F2", command=self.inputPathData)
        self.button.place(x=40, y=200, width=self.buttonWidth, height=self.buttonHeight)

    def buttonTemplate(self):
        self.buttontemplate = ttk.Button(self, text="Выбрать файл НЗО", command=self.inputPathTemplate)
        self.buttontemplate.place(x=40, y=260, width=self.buttonWidth, height=self.buttonHeight)

    def buttonEven(self):
        self.progressbar = ttk.Progressbar(orient="horizontal", mode="determinate")
        self.progressbar.place(x=410, y=380, width=130, height=10)
        self.buttonAction = ttk.Button(self, text="Выполнить", command=self.EventWithPaths)
        self.buttonAction.place(x=40, y=320, width=self.buttonWidth, height=self.buttonHeight)
