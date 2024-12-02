import os
current_directory = os.getcwd()     
import tkinter as tk
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

color_for_sprav = (0, 255, 100)
color_for_report = (168, 100, 255)
font_sprav = 'italic'
font_report = 'Arial'
sizeFont_sprav = 12
sizeFont_report = 12 


'''С 19 СТРОЧКИ ДО 280 автор: Тасуев Магомед'''

'''Эта функция нужна для того, чтобы знать адрес откуда читается файл'''
def open(nameFile): #возвращает адрес откуда читать файл
    # Путь к родительской папке (папка, содержащая обе папки)
    parent_folder = os.path.dirname(current_directory)
    # Путь к папке с файлом
    file_folder = os.path.join(parent_folder, "data")
    # Имя файла, который мы хотим открыть
    filename = nameFile
    # Полный путь к файлу
    file_path = os.path.join(file_folder, filename)
    return file_path

'''Эта функция нужна для того, чтобы знать адрес куда сохраняется файл'''
def save(nameFile): #возвращает адрес куда сохранить файл
    parent_folder = os.path.dirname(current_directory)
    file_folder = os.path.join(parent_folder, "output")
    filename = nameFile
    file_path = os.path.join(file_folder, filename)
    return file_path

'''Функция для того, чтобы знать адрес куда сохранять график'''
def save_graf(nameFile):
    parent_folder = os.path.dirname(current_directory)
    file_folder = os.path.join(parent_folder, "graphics")
    filename = nameFile
    file_path = os.path.join(file_folder, filename)
    return file_path

def rgb_hack(rgb):
    return "#%02x%02x%02x" % rgb


def spravochnik_Com():
    def store_excel():
       for i in range(height):
            for j in range(width):
                GDS.iloc[i, j] = pnt[i, j].get()
       filename = "Компании.xlsx"
       file_path = save(filename)
       GDS.to_excel(file_path, index=False)
    def store_csv():
       for i in range(height1):
            for j in range(width1):
                GDS.iloc[i, j] = pnt[i, j].get()
       filename = "Компании.csv"
       file_path = save(filename)
       GDS.to_csv(file_path, index=True)
    def clear_rep():
         """
         Очистка фрейма top_win
         """
         for widgets in top.winfo_children():
           widgets.destroy()

    filename = "Компании.csv"
    file_path = open(filename)
    GDS = pd.read_csv(file_path, delimiter = ';') # Читаем фрейм
    height = GDS.shape[0]
    width = GDS.shape[1]
#------------------------
    height1 = GDS.shape[0]
    width1 = GDS.shape[1]
# Формируем массив указателей на виджеты Entry
    pnt = np.empty(shape=(height, width), dtype="O")
# Формируем массив указателей на текстовые буферы для передачи данных Tcl/Tk
    vrs = np.empty(shape=(height, width), dtype="O")
# Построение изображения
    #root = tk.Tk() # До любых обращений к tkinter
# Инициализация указателей на буферы
    top = tk.LabelFrame(root, text = 'Компании',
                     bg = rgb_hack(color_for_sprav))
    top.grid(column=0, row=1)

    bottom = tk.LabelFrame(root, text = 'Управление',
                     bg = rgb_hack(color_for_sprav))
    bottom.grid(column=0, row=0, sticky = 'w')

    for i in range(height):
        for j in range(width):
            vrs[i, j] = tk.StringVar()
# Построение таблицы

    for i in range(height):
        for j in range(width):
            pnt[i, j] = tk.Entry(top, textvariable = vrs[i, j])
            pnt[i, j].grid (row=i, column=j)
# Заполнение таблицы значениями

    for i in range(height):
        for j in range(width):
            cnt = GDS.iloc[i, j]
            vrs[i, j].set(str(cnt))

    btn_2 = tk.Button(bottom, text='Сохранить в Excel',
    font=(font_sprav, sizeFont_sprav, 'italic'), bg='blue', fg='white', command=store_excel)
    btn_2.grid(column=0, row=0, sticky="w")
#-----------------------------------------
    btn_2 = tk.Button(bottom, text='Сохранить в Csv',
    font=(font_sprav, sizeFont_sprav, 'italic'), bg='blue', fg='white', command=store_csv)
    btn_2.grid(column=0, row=1, sticky="w")

    btn_2 = tk.Button(bottom, text='Очистить все', font=(font_sprav, sizeFont_sprav, 'italic'),
                  bg='blue',fg='white', command=clear_rep)
    btn_2.grid(row=0, column=1, sticky = "w")

    btn_2 = tk.Button(bottom, text='Завершить', font=(font_sprav, sizeFont_sprav, 'italic'),
                  bg='blue',fg='white', command=root.destroy)
    btn_2.grid(row=1, column=1, sticky = "w")

# Измененные значения можно затем получить из буферов методом get()

def spravochnik_Inv():
    def store_excel():
       for i in range(height):
            for j in range(width):
                GDS.iloc[i, j] = pnt[i, j].get()
       filename = "Инвесторы.xlsx"
       file_path = save(filename)
       GDS.to_excel(file_path, index=False)
    def store_csv():
       for i in range(height1):
            for j in range(width1):
                GDS.iloc[i, j] = pnt[i, j].get()
       filename = "ИнвесторыNew.csv"
       file_path = save(filename)
       GDS.to_csv(file_path, index=False)
    def clear_rep():
         """
         Очистка фрейма top_win
         """
         for widgets in top.winfo_children():
           widgets.destroy()
    
    filename = "Инвесторы.csv"
    file_path = open(filename)
    GDS = pd.read_csv(file_path, delimiter = ';') # Читаем фрейм
    height = GDS.shape[0]
    width = GDS.shape[1]
#------------------------
    height1 = GDS.shape[0]
    width1 = GDS.shape[1]
# Формируем массив указателей на виджеты Entry
    pnt = np.empty(shape=(height, width), dtype="O")
# Формируем массив указателей на текстовые буферы для передачи данных Tcl/Tk
    vrs = np.empty(shape=(height, width), dtype="O")
# Построение изображения
    #root = tk.Tk() # До любых обращений к tkinter
# Инициализация указателей на буферы
    top = tk.LabelFrame(root, text = 'Инвесторы',
                     bg = rgb_hack(color_for_sprav))
    top.grid(column=0, row=1)

    bottom = tk.LabelFrame(root, text = 'Управление',
                     bg = rgb_hack(color_for_sprav))
    bottom.grid(column=0, row=0, sticky = 'w')

    for i in range(height):
        for j in range(width):
            vrs[i, j] = tk.StringVar()
# Построение таблицы

    for i in range(height):
        for j in range(width):
            pnt[i, j] = tk.Entry(top, textvariable = vrs[i, j])
            pnt[i, j].grid (row=i, column=j)
# Заполнение таблицы значениями

    for i in range(height):
        for j in range(width):
            cnt = GDS.iloc[i, j]
            vrs[i, j].set(str(cnt))

    btn_1 = tk.Button(bottom, text='Сохранить в Excel',
    font=(font_sprav, sizeFont_sprav, 'italic'), bg='blue', fg='white', command=store_excel)
    btn_1.grid(column=0, row=0, sticky="w")
#-----------------------------------------
    btn_1 = tk.Button(bottom, text='Сохранить в Csv',
    font=(font_sprav, sizeFont_sprav, 'italic'), bg='blue', fg='white', command=store_csv)
    btn_1.grid(column=0, row=1, sticky="w")

    btn_1 = tk.Button(bottom, text='Очистить все', font=(font_sprav, sizeFont_sprav, 'italic'),
                  bg='blue',fg='white', command=clear_rep)
    btn_1.grid(row=0, column=1, sticky = "w")

    btn_1 = tk.Button(bottom, text='Завершить', font=(font_sprav, sizeFont_sprav, 'italic'),
                  bg='blue',fg='white', command=root.destroy)
    btn_1.grid(row=1, column=1, sticky = "w")

def spravochnik_D(): #Deals
    def store_excel():
       for i in range(height):
            for j in range(width):
                GDS.iloc[i, j] = pnt[i, j].get()
       filename = "Сделки.xlsx"
       file_path = save(filename)
       GDS.to_excel(file_path, index=False)
    def store_csv():
       for i in range(height1):
            for j in range(width1):
                GDS.iloc[i, j] = pnt[i, j].get()
       filename = "Сделки.csv"
       file_path = save(filename)
       GDS.to_csv(file_path, index=False)
    def clear_rep():
         """
         Очистка фрейма top_win
         """
         for widgets in top.winfo_children():
           widgets.destroy()
    filename = "Сделки.csv"
    file_path = open(filename)
    GDS = pd.read_csv(file_path, delimiter = ';') # Читаем фрейм
    height = GDS.shape[0]
    width = GDS.shape[1]
#------------------------
    height1 = GDS.shape[0]
    width1 = GDS.shape[1]
# Формируем массив указателей на виджеты Entry
    pnt = np.empty(shape=(height, width), dtype="O")
# Формируем массив указателей на текстовые буферы для передачи данных Tcl/Tk
    vrs = np.empty(shape=(height, width), dtype="O")
# Построение изображения
    #root = tk.Tk() # До любых обращений к tkinter
# Инициализация указателей на буферы
    top = tk.LabelFrame(root, text = 'Сделки',
                     bg = rgb_hack(color_for_sprav))
    top.grid(column=0, row=1)

    bottom = tk.LabelFrame(root, text = 'Управление',
                     bg = rgb_hack(color_for_sprav))
    bottom.grid(column=0, row=0, sticky = 'w')

    for i in range(height):
        for j in range(width):
            vrs[i, j] = tk.StringVar()
# Построение таблицы

    for i in range(height):
        for j in range(width):
            pnt[i, j] = tk.Entry(top, textvariable = vrs[i, j])
            pnt[i, j].grid (row=i, column=j)
# Заполнение таблицы значениями

    for i in range(height):
        for j in range(width):
            cnt = GDS.iloc[i, j]
            vrs[i, j].set(str(cnt))

    btn_3 = tk.Button(bottom, text='Сохранить в Excel',
    font=(font_sprav, sizeFont_sprav, 'italic'), bg='blue', fg='white', command=store_excel)
    btn_3.grid(column=0, row=0, sticky="w")
#-----------------------------------------
    btn_3 = tk.Button(bottom, text='Сохранить в Csv',
    font=(font_sprav, sizeFont_sprav, 'italic'), bg='blue', fg='white', command=store_csv)
    btn_3.grid(column=0, row=1, sticky="w")

    btn_3 = tk.Button(bottom, text='Очистить все', font=(font_sprav, sizeFont_sprav, 'italic'),
                  bg='blue',fg='white', command=clear_rep)
    btn_3.grid(row=0, column=1, sticky = "w")

    btn_3 = tk.Button(bottom, text='Завершить', font=(font_sprav, sizeFont_sprav, 'italic'),
                  bg='blue',fg='white', command=root.destroy)
    btn_3.grid(row=1, column=1, sticky = "w")


'''Автор: Киемидинова Мухсинахон c 286 до 594'''

'''Функции для показа каждого отчета'''
def report_for_call():
    '''Функции для сохранения отчета в csv и xlsx файлы'''
    def store_excel():
       for i in range(height):
            for j in range(width):
                GDS.iloc[i, j] = pnt[i, j].get()
       filename = "Отчет для Колл-Центра.xlsx"
       file_path = save(filename)
       GDS.to_excel(file_path, index=False)
    def store_csv():
       for i in range(height1):
            for j in range(width1):
                GDS.iloc[i, j] = pnt[i, j].get()
       filename = "Отчет для Колл-Центра.csv"
       file_path = save(filename)
       GDS.to_csv(file_path, index=False)
    '''функция: очистить все'''
    def clear_rep():
         """
         Очистка фрейма top_win
         """
         for widgets in top.winfo_children():
           widgets.destroy()
           
    filename = "Отчет для Колл-Центра.csv"
    file_path = open(filename)
    GDS = pd.read_csv(file_path, delimiter=';') # Читаем фрейм
    height = GDS.shape[0]
    width = GDS.shape[1]
#------------------------
    height1 = GDS.shape[0]
    width1 = GDS.shape[1]
# Формируем массив указателей на виджеты Entry
    pnt = np.empty(shape=(height, width), dtype="O")
# Формируем массив указателей на текстовые буферы для передачи данных Tcl/Tk
    vrs = np.empty(shape=(height, width), dtype="O")
# Построение изображения
    #root = tk.Tk() # До любых обращений к tkinter
# Инициализация указателей на буферы
    top = tk.LabelFrame(root, text = 'Отчет для Колл-Центра',
                     bg = rgb_hack(color_for_report))
    top.grid(column=0, row=1)

    bottom = tk.LabelFrame(root, text = 'Управление',
                     bg = rgb_hack(color_for_report))
    bottom.grid(column=0, row=0, sticky = 'w')

    for i in range(height):
        for j in range(width):
            vrs[i, j] = tk.StringVar()
# Построение таблицы

    for i in range(height):
        for j in range(width):
            pnt[i, j] = tk.Entry(top, textvariable = vrs[i, j])
            pnt[i, j].grid (row=i, column=j)
# Заполнение таблицы значениями

    for i in range(height):
        for j in range(width):
            cnt = GDS.iloc[i, j]
            vrs[i, j].set(str(cnt))

    btn_3 = tk.Button(bottom, text='Сохранить в Excel',
    font=(font_report, sizeFont_report, 'italic'), bg='green', fg='white', command=store_excel)
    btn_3.grid(column=0, row=0, sticky="w")
#-----------------------------------------
    btn_3 = tk.Button(bottom, text='Сохранить в Csv',
    font=(font_report, sizeFont_report, 'italic'), bg='green', fg='white', command=store_csv)
    btn_3.grid(column=0, row=1, sticky="w")

    btn_3 = tk.Button(bottom, text='Очистить все', font=(font_report, sizeFont_report, 'italic'),
                  bg='green',fg='white', command=clear_rep)
    btn_3.grid(row=0, column=1, sticky = "w")

    btn_3 = tk.Button(bottom, text='Завершить', font=(font_report, sizeFont_report, 'italic'),
                  bg='green',fg='white', command=root.destroy)
    btn_3.grid(row=1, column=1, sticky = "w")
    
    
def report_for_inv():
    def store_excel():
       for i in range(height):
            for j in range(width):
                GDS.iloc[i, j] = pnt[i, j].get()
       filename = "Отчет для Инвесторов.xlsx"
       file_path = save(filename)
       GDS.to_excel(file_path, index=False)
    def store_csv():
       for i in range(height1):
            for j in range(width1):
                GDS.iloc[i, j] = pnt[i, j].get()
       filename = "Отчет для Инвесторов.csv"
       file_path = save(filename)
       GDS.to_csv(file_path, index=False)
    def clear_rep():
         """
         Очистка фрейма top_win
         """
         for widgets in top.winfo_children():
           widgets.destroy()
           
    filename = "Отчет для Инвесторов.csv"
    file_path = open(filename)
    GDS = pd.read_csv(file_path, delimiter = ';') # Читаем фрейм
    height = GDS.shape[0]
    width = GDS.shape[1]
#------------------------
    height1 = GDS.shape[0]
    width1 = GDS.shape[1]
# Формируем массив указателей на виджеты Entry
    pnt = np.empty(shape=(height, width), dtype="O")
# Формируем массив указателей на текстовые буферы для передачи данных Tcl/Tk
    vrs = np.empty(shape=(height, width), dtype="O")
# Построение изображения
    #root = tk.Tk() # До любых обращений к tkinter
# Инициализация указателей на буферы
    top = tk.LabelFrame(root, text = 'Отчет для Инвесторов',
                     bg = rgb_hack(color_for_report))
    top.grid(column=0, row=1)

    bottom = tk.LabelFrame(root, text = 'Управление',
                     bg = rgb_hack(color_for_report))
    bottom.grid(column=0, row=0, sticky = 'w')

    for i in range(height):
        for j in range(width):
            vrs[i, j] = tk.StringVar()
# Построение таблицы

    for i in range(height):
        for j in range(width):
            pnt[i, j] = tk.Entry(top, textvariable = vrs[i, j])
            pnt[i, j].grid (row=i, column=j)
# Заполнение таблицы значениями

    for i in range(height):
        for j in range(width):
            cnt = GDS.iloc[i, j]
            vrs[i, j].set(str(cnt))

    btn_3 = tk.Button(bottom, text='Сохранить в Excel',
    font=(font_report, sizeFont_report, 'italic'), bg='green', fg='white', command=store_excel)
    btn_3.grid(column=0, row=0, sticky="w")
#-----------------------------------------
    btn_3 = tk.Button(bottom, text='Сохранить в Csv',
    font=(font_report, sizeFont_report, 'italic'), bg='green', fg='white', command=store_csv)
    btn_3.grid(column=0, row=1, sticky="w")

    btn_3 = tk.Button(bottom, text='Очистить все', font=(font_report, sizeFont_report, 'italic'),
                  bg='green',fg='white', command=clear_rep)
    btn_3.grid(row=0, column=1, sticky = "w")

    btn_3 = tk.Button(bottom, text='Завершить', font=(font_report, sizeFont_report, 'italic'),
                  bg='green',fg='white', command=root.destroy)
    btn_3.grid(row=1, column=1, sticky = "w")
    
def report_for_analyst():
    def store_excel():
       for i in range(height):
            for j in range(width):
                GDS.iloc[i, j] = pnt[i, j].get()
       filename = "Отчет для Аналитика.xlsx"
       file_path = save(filename)
       GDS.to_excel(file_path, index=False)
    def store_csv():
       for i in range(height1):
            for j in range(width1):
                GDS.iloc[i, j] = pnt[i, j].get()
       filename = "Отчет для Аналитика.csv"
       file_path = save(filename)
       GDS.to_csv(file_path, index=False)
    def clear_rep():
         """
         Очистка фрейма top_win
         """
         for widgets in top.winfo_children():
           widgets.destroy()
    filename = "Отчет для Аналитика.csv"
    file_path = open(filename)
    GDS = pd.read_csv(file_path, delimiter = ';') # Читаем фрейм
    height = GDS.shape[0]
    width = GDS.shape[1]
#------------------------
    height1 = GDS.shape[0]
    width1 = GDS.shape[1]
# Формируем массив указателей на виджеты Entry
    pnt = np.empty(shape=(height, width), dtype="O")
# Формируем массив указателей на текстовые буферы для передачи данных Tcl/Tk
    vrs = np.empty(shape=(height, width), dtype="O")
# Построение изображения
    #root = tk.Tk() # До любых обращений к tkinter
# Инициализация указателей на буферы
    top = tk.LabelFrame(root, text = 'Отчет для Аналитика',
                     bg = rgb_hack(color_for_report))
    top.grid(column=0, row=1)

    bottom = tk.LabelFrame(root, text = 'Управление',
                     bg = rgb_hack(color_for_report))
    bottom.grid(column=0, row=0, sticky = 'w')

    for i in range(height):
        for j in range(width):
            vrs[i, j] = tk.StringVar()
# Построение таблицы

    for i in range(height):
        for j in range(width):
            pnt[i, j] = tk.Entry(top, textvariable = vrs[i, j])
            pnt[i, j].grid (row=i, column=j)
# Заполнение таблицы значениями

    for i in range(height):
        for j in range(width):
            cnt = GDS.iloc[i, j]
            vrs[i, j].set(str(cnt))

    btn_3 = tk.Button(bottom, text='Сохранить в Excel',
    font=(font_report, sizeFont_report, 'italic'), bg='green', fg='white', command=store_excel)
    btn_3.grid(column=0, row=0, sticky="w")
#-----------------------------------------
    btn_3 = tk.Button(bottom, text='Сохранить в Csv',
    font=(font_report, sizeFont_report, 'italic'), bg='green', fg='white', command=store_csv)
    btn_3.grid(column=0, row=1, sticky="w")

    btn_3 = tk.Button(bottom, text='Очистить все', font=(font_report, sizeFont_report, 'italic'),
                  bg='green',fg='white', command=clear_rep)
    btn_3.grid(row=0, column=1, sticky = "w")

    btn_3 = tk.Button(bottom, text='Завершить', font=(font_report, sizeFont_report, 'italic'),
                  bg='green',fg='white', command=root.destroy)
    btn_3.grid(row=1, column=1, sticky = "w")
    
def report_for_soisk():
    def store_excel():
       for i in range(height):
            for j in range(width):
                GDS.iloc[i, j] = pnt[i, j].get()
       filename = "Отчет для Соискателя.xlsx"
       file_path = save(filename)  
       GDS.to_excel(file_path, index=False)
    def store_csv():
       for i in range(height1):
            for j in range(width1):
                GDS.iloc[i, j] = pnt[i, j].get()
       filename = "Отчет для Соискателя.csv"
       file_path = save(filename)  
       GDS.to_csv(file_path, index=False)
    def clear_rep():
         """
         Очистка фрейма top_win
         """
         for widgets in top.winfo_children():
           widgets.destroy()
    filename = "Отчет для Соискателя.csv"
    file_path = open(filename)
    GDS = pd.read_csv(file_path, delimiter=';') # Читаем фрейм
    height = GDS.shape[0]
    width = GDS.shape[1]
#------------------------
    height1 = GDS.shape[0]
    width1 = GDS.shape[1]
# Формируем массив указателей на виджеты Entry
    pnt = np.empty(shape=(height, width), dtype="O")
# Формируем массив указателей на текстовые буферы для передачи данных Tcl/Tk
    vrs = np.empty(shape=(height, width), dtype="O")
# Построение изображения
    #root = tk.Tk() # До любых обращений к tkinter
# Инициализация указателей на буферы
    top = tk.LabelFrame(root, text = 'Отчет для Соискателя',
                     bg = rgb_hack(color_for_report))
    top.grid(column=0, row=1)

    bottom = tk.LabelFrame(root, text = 'Управление',
                     bg = rgb_hack(color_for_report))
    bottom.grid(column=0, row=0, sticky = 'w')

    for i in range(height):
        for j in range(width):
            vrs[i, j] = tk.StringVar()
# Построение таблицы

    for i in range(height):
        for j in range(width):
            pnt[i, j] = tk.Entry(top, textvariable = vrs[i, j])
            pnt[i, j].grid (row=i, column=j)
# Заполнение таблицы значениями

    for i in range(height):
        for j in range(width):
            cnt = GDS.iloc[i, j]
            vrs[i, j].set(str(cnt))

    btn_3 = tk.Button(bottom, text='Сохранить в Excel',
    font=(font_report, sizeFont_report, 'italic'), bg='green', fg='white', command=store_excel)
    btn_3.grid(column=0, row=0, sticky="w")
#-----------------------------------------
    btn_3 = tk.Button(bottom, text='Сохранить в Csv',
    font=(font_report, sizeFont_report, 'italic'), bg='green', fg='white', command=store_csv)
    btn_3.grid(column=0, row=1, sticky="w")

    btn_3 = tk.Button(bottom, text='Очистить все', font=(font_report, sizeFont_report, 'italic'),
                  bg='green',fg='white', command=clear_rep)
    btn_3.grid(row=0, column=1, sticky = "w")

    btn_3 = tk.Button(bottom, text='Завершить', font=(font_report, sizeFont_report, 'italic'),
                  bg='green',fg='white', command=root.destroy)
    btn_3.grid(row=1, column=1, sticky = "w") 
db = pd.DataFrame()


'''Автор: Тасуев Магомед с 599 до 855'''
'''Функция, чтобы можно было создать текстовый отчет'''
def show_report():
    def show_data():
        """
        Демонстрация данных
        """
        global pnt
        global W
        global num_cols
        def show_rep():
            global pnt
            global W
            global num_cols
            clear_data()
            num_cols = lb.curselection() # Номера выбранных столбцов
            height = W.shape[0]
            width = len(num_cols)
            # Вывод названий столбцов
            for j in range(width):
                e = tk.Entry(top, relief=tk.RIDGE, background='green')
                e.grid(row=0, column=j, sticky=tk.E)
                e.insert(tk.END, W.columns[num_cols[j]])  
            # Самим - вывод названий строк
            rows = []
            for i in range(height):
                cols = []
                for j in range(width):
                    e = tk.Entry(top, relief=tk.RIDGE)
                    e.grid(row=i+1, column=j, sticky=tk.E)
                    e.insert(tk.END, str(W.iloc[i, num_cols[j]]))
                    cols.append(e)
                rows.append(cols)
            pnt = np.array(rows)

        lb = tk.Listbox(distr, selectmode=tk.EXTENDED, height=W.shape[0], 
                         font="Times 14")
        for name in W.columns:
            lb.insert(tk.END, name)
        lb.grid(column=0, row=1)
        btn = tk.Button(distr, text='Показать отчет', font=('Arial', 12, 'italic'),
                      bg='green',fg='white', command=show_rep)
        btn.grid(column=0, row=0)
    def store_data():
        """
        Получение даных таблицы из Tcl/Tk
        с помощью метода get() и сохранение в Excel
        """
        from tkinter import filedialog as fld
        global pnt, W
        global num_cols
        ftypes = [('Excel файлы', '*.xlsx'), ('Все файлы', '*')]
        dlg = fld.SaveAs(filetypes = ftypes)
        fl = dlg.show()
        # Определяем тип файла
        n = fl.rfind('.')
        ext = fl[n+1:]
        # Читаем данные из Entry
        U = np.empty_like(pnt)
        for i in range(pnt.shape[0]): 
            for j in range(pnt.shape[1]): 
                U[i, j] = pnt[i, j].get()
        U = pd.DataFrame(U)
        U.columns = W.columns[list(num_cols)]
        # Записываем на диск
        if ext == 'xlsx':
            U.to_excel(fl, index=False)
        else:
            print("Unknown extention")
            exit()

    def read_data():
        """
        Чтение и демонстрация данных
        """
        import pandas as pd
        from tkinter import filedialog as fld
        global W
        ftypes = [('Excel файлы', '*.xlsx'), ('Все файлы', '*')]
        dlg = fld.Open(filetypes = ftypes)
        fl = dlg.show()
        # Определяем тип файла
        n = fl.rfind('.')
        ext = fl[n+1:]
        if ext == 'xlsx':
            W = pd.read_excel(fl)
        else:
            print("Unknown extention")
            exit()

    def clear_rep():
         """
         Очистка фрейма top_win
         """
         for widgets in top.winfo_children():
           widgets.destroy()
         for widgets in distr.winfo_children():
           widgets.destroy()

    def clear_data():
         """
         Очистка фрейма top_win
         """
         for widgets in top.winfo_children():
           widgets.destroy()
           
           
    rgb = color_for_report
    top = tk.LabelFrame(root, text="Отчет", bg=rgb_hack(rgb))
    top.grid(column=0, row=0)

    # Создаем фрейм (контейнер) в котором будет размещено управление
    ctrl = tk.LabelFrame(root, text="Управление")
    ctrl.grid(column=0, row=1, sticky="w")
    btn_1 = tk.Button(ctrl, text='Чтение данных', font=('Arial', 12, 'italic'),
                      bg='green',fg='white', command=read_data)
    btn_1.grid(row=0, column=0)
    btn_2 = tk.Button(ctrl, text='Просмотр отчета', font=('Arial', 12, 'italic'),
                      bg='green',fg='white', command=show_data)
    btn_2.grid(row=0, column=1)
    btn_3 = tk.Button(ctrl, text='Сохранить даные', font=('Arial', 12, 'italic'),
                      bg='green',fg='white', command=store_data)
    btn_3.grid(row=1, column=0)
    btn_4 = tk.Button(ctrl, text='Очистить все', font=('Arial', 12, 'italic'),
                      bg='green',fg='white', command=clear_rep)
    btn_4.grid(row=1, column=1)
    btn_5 = tk.Button(ctrl, text='Завершить', font=('Arial', 12, 'italic'),
                      bg='green',fg='white', command=root.destroy)
    btn_5.grid(row=2, column=0, columnspan=2)

    # # Создаем фрейм (контейнер) в котором будет определена структура просмотра
    distr = tk.LabelFrame(root, text="Структура отчета")
    distr.grid(column=1, row=1, sticky="w")
    lbl = tk.Label(distr, text="Выберите столбцы", bg='green', fg='white' )
    lbl.grid(column=0, row=0)
    lb = tk.Listbox(distr, selectmode=tk.EXTENDED, height=W.shape[0], 
                     font="Times 14")
    for name in W.columns:
        lb.insert(tk.END, name)
    lb.grid(column=0, row=1)


def draw_histogram(file_path, column, chart_title, save_path, name_column, value):
    # Чтение данных из файла Excel с помощью pandas
    df = pd.read_excel(file_path)

    # Извлечение данных из выбранного столбца
    data = df[column].dropna().values.tolist()

    # Создание гистограммы с помощью Matplotlib
    plt.figure(figsize=(8, 6))
    plt.hist(data, color='blue')
    plt.xlabel(name_column)
    plt.ylabel(value)
    plt.title(chart_title)

    # Сохранение гистограммы в файл
    plt.savefig(save_path)

    # Отображение гистограммы
    plt.show()
  
def histogram1():
    filename = "Компании.xlsx"
    file_path = open(filename)
    column = "price"  # Выбранный столбец
    name_column = "Цена акции (в рублях)"
    value = "Количество компаний"
    chart_title = "Распределение цен"  # Название диаграммы
    filename1 = "Распределение цен.png"
    save_path = save_graf(filename1)
    draw_histogram(file_path, column, chart_title, save_path, name_column, value)

def histogram2():
    filename = "Компании.xlsx"
    file_path = open(filename)
    column = "number of employees"  # Выбранный столбец
    name_column = "Количество сотрудников компаний (в млн)"
    value = "Количество компаний"
    chart_title = "Количество сотрудников компаний"  # Название диаграммы
    filename1 = "Количество сотрудников компаний.png"
    save_path = save_graf(filename1)
    draw_histogram(file_path, column, chart_title, save_path, name_column, value)
    
def draw_scatter_plot_from_excel(file_path, x_column, y_column, chart_title, save_path):
    # Чтение данных из файла Excel с помощью pandas
    df = pd.read_excel(file_path)

    # Извлечение данных из выбранных столбцов
    x_data = df[x_column].dropna().values.tolist()
    y_data = df[y_column].dropna().values.tolist()

    # Создание диаграммы рассеивания с помощью Matplotlib
    plt.figure(figsize=(8, 6))
    plt.scatter(x_data, y_data, color='blue')
    plt.xlabel(x_column)
    plt.ylabel(y_column)
    plt.title(chart_title)

    # Сохранение диаграммы рассеивания в файл
    plt.savefig(save_path)

    # Отображение диаграммы рассеивания
    plt.show()

def func2():
    filename = "Сделки.xlsx"
    file_path = open(filename)
    x_column = "id компании"  # Выбранный столбец для оси X
    y_column = "id инвестора"  # Выбранный столбец для оси Y
    chart_title = "Количество купленных акций компании"  # Название диаграммы
    filename1 = "Распределение цен.png"
    save_path = save_graf(filename1)
    draw_scatter_plot_from_excel(file_path, x_column, y_column, chart_title, save_path)   

# Формируем окно с меню
root = tk.Tk()
root.geometry('600x250+100+80')
root.title("Акционерное сообщество")
root.resizable (True, True)



mainmenu = tk.Menu(root, tearoff=0) # Создаем объект класса Menu
printmenu1 = tk.Menu(mainmenu, tearoff=0)
printmenu1.add_command(label="Инвесторы", command = spravochnik_Inv)
printmenu1.add_command(label="Компании", command = spravochnik_Com)
printmenu1.add_command(label="Сделки", command = spravochnik_D)

printmenu2 = tk.Menu(mainmenu, tearoff=0) # Создаем объект класса Menu
printmenu2.add_command(label="Отчет для Колл-Центра", command = report_for_call)
printmenu2.add_command(label="Отчет для Инвесторов", command = report_for_inv)
printmenu2.add_command(label="Отчет для Аналитика", command = report_for_analyst)
printmenu2.add_command(label="Отчет для Соискателя", command = report_for_soisk)
printmenu2.add_command(label="Создать текстовый отчет", command = show_report)

printmenu3 = tk.Menu(mainmenu, tearoff=0)
printmenu3.add_command(label="Распределение цен", command = histogram1)
printmenu3.add_command(label="Количество купленных акций компании", command = func2)
printmenu3.add_command(label="Количество сотрудников компаний", command = histogram2)

settings_menu = tk.Menu(mainmenu, tearoff = 0)
settings_menu.add_command(label="Распределение цен", command = histogram1)

# Размещение пунктов в меню
mainmenu.add_cascade(label="Справочники", menu=printmenu1)
mainmenu.add_cascade(label="Текстовые отчеты", menu=printmenu2)
mainmenu.add_cascade(label="Графические отчеты", menu=printmenu3)

mainmenu.add_command(label="Выход", command = root.destroy)
root.config(menu=mainmenu)
# Поле с приветствием
lbl = tk.Label(root, text="Зравствуйте! Выберите нужную вкладку в меню")
# Относительные координаты - относительно левого верхнего угла
# (1,1) - правый нижний угол
lbl.place(relx=0.25, rely=0.25)
# lbl.grid(column=0, row=0)

root.mainloop()