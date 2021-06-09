from tkinter import *
from tkinter import filedialog
from tkinter.filedialog import askopenfilename
from tkinter import ttk
from nnov import proga
from tkinter import messagebox as mb
import os
from perm import proga_perm

root = Tk()
root.title('Сравнение РЭС by Nikolaev v2.1')
root.geometry('570x320')
root.resizable(False, False)
provar = IntVar()
provar.set(0)

def openfile1():
    '''
    Вызывает окно с выбором файла с зарагестрированными РЭС. Проверяет расширение.
    '''
    global filename1
    root.withdraw()
    filename1 = askopenfilename(filetypes=(("Excel xlsx", "*.xlsx"), ("All files", "*.*")))
    root.deiconify()
    e1.delete(0, END)
    e1.insert(0, filename1)


def openfile2():
    '''
    Вызывает окно для выбора истекших РЭС
    '''
    global filename2
    root.withdraw()
    filename2 = askopenfilename(filetypes=(("Excel", "*.xlsx"), ("All files", "*.*")))
    root.deiconify()
    e2.delete(0, END)
    e2.insert(0, filename2)


def selectFolderPath():
    '''
    Вызывает окно для выбора папки сохранения
    '''
    global select_folder
    root.withdraw()
    select_folder = filedialog.askdirectory()
    root.deiconify()
    e4.delete(0, END)
    e4.insert(0, select_folder)


def check():
    '''
    Проверяет правильность заполнения всех окон
    :return: True or False
    '''
    okno1 = e1.get()
    okno2 = e2.get()
    okno3 = e4.get()
    if '.xlsx' not in okno1:
        mb.showerror('Ошибочка', 'Неверно выбран файл в первом окне')
        return True
    elif '.xlsx' not in okno2:
        mb.showerror('Ошибочка', 'Неверно выбран файл во втором окне')
        return True
    elif okno3 == '':
        mb.showerror('Сохранять-то куда?!', 'Не выбран каталог для выгрузки!!!')
        return True
    else:
        return False


def zareg():
    c = e1.get()
    return c


def istekli():
    c = e2.get()
    return c


def katalog():
    c = e4.get()
    return c


def start():
    '''
    Запускает алгоритм расчета по Нижнему или Перми, в зависимости от выбора пользователем
    '''
    if cmb.get()=='Форматированная таблица(НН)':
        proga(zareg(), istekli(), katalog())
    elif cmb.get()=='Неформатированная таблица(Пермь)':
        proga_perm(zareg(), istekli(), katalog())



def go():
    '''
    Функция при нажатии кнопки "Начать". Последовательно: Блокирует все поля ввода и кноипки. Запускает функцию start.
    Выводит сообщения в строке прогресса. Разблокирует все поля ввода и кнопки. Выводит ссылку на сохраненный файл.
    '''
    btn3['state'] = 'disable'
    btn2['state'] = 'disable'
    btn1['state'] = 'disable'
    btn4['state'] = 'disable'
    e1['state'] = 'disable'
    e2['state'] = 'disable'
    e4['state'] = 'disable'
    if check() == False:
        lbl = Label(root, text='Подготовка файлов...')
        lbl.place(x=20, y=250)
        root.after(2000, provar.set(25))
        pb.update()
        lbl.destroy()
        lbl = Label(root, text='Сравнение...')
        lbl.place(x=20, y=250)
        root.after(2000, provar.set(50))
        pb.update()
        start()
        lbl.destroy()
        lbl = Label(root, text='Сохранение файла')
        lbl.place(x=20, y=250)
        root.after(1000, provar.set(75))
        pb.update()
        lbl.destroy()
        lbl = Label(root, text='Готово!!!')
        lbl.place(x=20, y=250)
        root.after(2000, provar.set(100))
        pb.update()

        btn3['state'] = 'normal'
        btn2['state'] = 'normal'
        btn1['state'] = 'normal'
        btn4['state'] = 'normal'
        e1['state'] = 'normal'
        e2['state'] = 'normal'
        e4['state'] = 'normal'
        lbl6 = Label(root, fg='blue', text=f'{katalog()}/Итоговый список незарегистрированных РЭС.xlsx', cursor="hand2")
        lbl6.place(x=10, y=270)
        lbl6.bind('<Button-1>', lambda e: os.startfile(f'{katalog()}/Итоговый список незарегистрированных РЭС.xlsx'))
    else:
        btn3['state'] = 'normal'
        btn2['state'] = 'normal'
        btn1['state'] = 'normal'
        btn4['state'] = 'normal'
        e1['state'] = 'normal'
        e2['state'] = 'normal'
        e4['state'] = 'normal'



f_top = LabelFrame(root, text='Выберите файл с вновь зарегистрированными РЭС')
e1 = Entry(f_top, width=70)
e1.pack(side=LEFT, padx=3, pady=3)
btn1 = Button(f_top, text='Обзор...', width=20, command=openfile1)
btn1.pack(side=LEFT, padx=3, pady=3)
f_top.place(x=10, y=10, width=550)

f_down = LabelFrame(root, text='Выберите файл снятых с учета РЭС')
e2 = Entry(f_down, width=70)
e2.pack(side=LEFT, padx=3, pady=3)
btn2 = Button(f_down, text='Обзор...', width=20, command=openfile2)
btn2.pack(side=LEFT, padx=3, pady=3)
f_down.place(x=10, y=80, width=550)

f_katalog = LabelFrame(root, text='Выберите каталог для выгрузки итогового файла')
e4 = Entry(f_katalog, width=70)
e4.pack(side=LEFT, padx=3, pady=3)
btn4 = Button(f_katalog, text='Обзор...', width=20, command=selectFolderPath)
btn4.pack(side=LEFT, padx=3, pady=3)
f_katalog.place(x=10, y=150, width=550)

btn3 = Button(root, text='Начать!!', width=20, height=2, command=go)
btn3.place(x=190, y=220)

#Полоса прогресса
pb = ttk.Progressbar(root, variable=provar, length=570)
pb.pack(side=BOTTOM)

#Поле для выбора региона (алгоритма)
cmb = ttk.Combobox(root, width = 32, value=('Форматированная таблица(НН)','Неформатированная таблица(Пермь)'))
cmb.current(0)
cmb.place(x=350, y=220)

if __name__ == "__main__":
    mainloop()