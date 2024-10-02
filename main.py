# -*- coding: utf-8 -*-
import re
import os
import pandas as pd
import glob
import warnings
import time
# import asyncio
from tkinter import *
from tkinter import ttk
from tkinter.ttk import Progressbar
from tkinter.messagebox import *
import tkinter.filedialog as fd
from PIL import Image, ImageTk
from threading import *

# Игнорируем уведомление об ошибки метода Append
warnings.filterwarnings("ignore")


def threading():
    # Вызывает функцию main в асинхронном режиме
    t1=Thread(target=main)
    t1.start()


# Функция обновления прогресс бара
def update_bar(x):
    if bar['value'] < 100:
        bar['value'] += 100 / len(x)
        time.sleep(0.12)
        value_label['text'] = f"Результат: {round(bar['value'])} %"
        ws.update_idletasks()

# Функция открытия каталога (папки)
def open_folder():
    os.system(f'start explorer {Entry.get()}')

# Функция выбора каталога (папки)
def choose_directory():
    directory = fd.askdirectory(title="Выбрать папку", initialdir="/").replace('/', '\\')
    if directory:
        Entry.insert(0, directory)


def main():
    dataf = read_files()
    dataf = pd.DataFrame(dataf)
    if dataf is None:
        pass
    else:
        dataf.to_excel('Ответы.xlsx', sheet_name='Ответы', index=False)
        showinfo(message='Программа выполнена успешно! Благодарим за пользование.', title="Answer's")

# Читаем файл и записываем результаты в excel таблицу
def read_files():
    data_f = []
    data_f = pd.DataFrame(data_f)
    direc = Entry.get()
    os.chdir(direc)
    files = glob.glob("**/*.ans", recursive=True)
    if files is not None:
        for path in range(0, len(files)):
            with open(files[path], 'r', encoding='CP1251') as file:
                dates = file.read()
                answer = re.findall(r'\d+', dates)
                if (len(answer)) > 0:
                    del answer[0:13]
                    numbers = answer[4] + answer[0] + answer[1] + answer[2] + answer[3]
                    answer_list = list(numbers)
                    fil = str(files[path])
                    fil_out = re.findall(r'^\d{5}', fil)
                    print(fil_out)
                    answer_list.append(fil_out)
                    file.close()

                    data_f = data_f.append([answer_list])
                    update_bar(files)
    return data_f


ws = Tk()
img = Image.open('ico.ico')
font = ImageTk.PhotoImage(img)
ws.title("Добро пожаловать в программу Answer's!")
ws.geometry('430x400')
ws.resizable(width=False, height=False)
Label(ws, image=font).pack(padx=0, pady=0)

Label(ws, text='Разработчик: Паршин А.Е. Куратор: Дядин Е.А.').place(x=0, y=381)

bar = Progressbar(ws, orient=HORIZONTAL, length=251, mode='determinate', value=0, maximum=100)
bar.place(y=260, x=85)

Button(ws, text='Запустить', command=threading, width=12, height=1).place(y=340, x=165)
Button(ws, text='Выбрать папку', command=choose_directory, width=12, height=1).place(y=340, x=50)
Button(ws, text='Открыть папку', command=open_folder, width=12, height=1).place(y=340, x=280)

Entry = ttk.Entry(ws, width=41)
Entry.place(y=306, x=85)

value_label = Label(ws, text='Результат: ')
value_label.place(x=82, y=283)

ws.mainloop()