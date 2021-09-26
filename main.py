from tkinter import *
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from tkinter import filedialog

filename = ''


def get_fil_sum(df, name):
    """
    Функция из датафрейма достает группы, по которым строить график и вычисляет
    стоимость каждой из этих групп

    :param df:
    :param name:
    :return:
    """

    sort_by = set()

    for index, row in df.iterrows():
        sort_by.add((row[name], 0))

    sort_by = list(sort_by)
    for k, elem in enumerate(sort_by):
        sort_by[k] = (elem[0], df.loc[df[name] == elem[0], 'ЛКВ (базовый вариант)'].sum())
    return [i[0] for i in sort_by], [i[1] for i in sort_by]


def get_pie(name, data1, data2, data3, step):
    """
    Функция рисует график по трем датафреймам

    :param name:
    :param data1:
    :param data2:
    :param data3:
    :param step:
    :return:
    """
    diff = False
    if step:
        diff = True
    fil1, sum1 = get_fil_sum(data1, name)
    fil2, sum2 = get_fil_sum(data2, name)
    fil3, sum3 = get_fil_sum(data3, name)
    rasnost1 = set(fil1) - set(fil2)
    rasnost2 = set(fil1) - set(fil3)
    empty = {}
    sum = 0
    for i in sum1:
        sum += i
    for fil in zip(fil1, sum1):
        empty[fil[0]] = [fil[1] / sum]
    sum = 0
    for i in sum2:
        sum += i
    for fil in zip(fil2, sum2):
        empty[fil[0]].append(fil[1] / sum)
    for i in rasnost1:
        empty[i].append(0)
    if diff:
        sum = 0
        for i in sum3:
            sum += i
        for fil in zip(fil3, sum3):
            empty[fil[0]].append(fil[1] / sum)
        for i in rasnost2:
            empty[i].append(0)
    df1 = pd.DataFrame(dict(sorted(empty.items(), key=lambda x: x[1], reverse=True)))
    fig, axes = plt.subplots(figsize=(4, 6), dpi=65)
    canvas = FigureCanvasTkAgg(fig, window)
    canvas.get_tk_widget().place(relx=0.5, y=465, anchor=CENTER)
    df1.plot(kind='bar', colormap='Paired', stacked=True, grid=True, ax=axes)
    plt.title(name, fontsize=10)
    if diff:
        plt.xticks([0, 1, 2], [f'{data1["ЛКВ (базовый вариант)"].sum():1.0f}',
                               f'{data1["ЛКВ (базовый вариант)"].sum():1.0f}\n- {(1 - data2["ЛКВ (базовый вариант)"].sum() / data1["ЛКВ (базовый вариант)"].sum()) * 100:1.0f}%',
                               f'{data1["ЛКВ (базовый вариант)"].sum():1.0f}\n- {(1 - data3["ЛКВ (базовый вариант)"].sum() / data1["ЛКВ (базовый вариант)"].sum()) * 100:1.0f}%'],
                   rotation='horizontal')
    else:
        plt.xticks([0, 1], [f'{data1["ЛКВ (базовый вариант)"].sum():1.0f}',
                            f'{data1["ЛКВ (базовый вариант)"].sum():1.0f}\n- {(1 - data2["ЛКВ (базовый вариант)"].sum() / data1["ЛКВ (базовый вариант)"].sum()) * 100:1.0f}%'],
                   rotation='horizontal')
    canvas.draw()


def result():
    """
    Функция получает параметры и обрабатывает их для построения графика

    :return:
    """
    summ = text1.get('1.0', END)
    percent = text2.get('1.0', END)
    step = text3.get('1.0', END)
    if variable1.get() == 'Сумма' and filename and summ:
        summ = float(summ)
        try:
            step = float(step)
        except Exception:
            step = 0
        df = pd.read_excel(filename, "Рейтинг", usecols="B,D,E,AA", header=8)
        df.drop([0], axis=0, inplace=True)
        df.sort_values(by=['Интегральный рейтинг'], inplace=True, ascending=False)
        df1 = df.copy()
        while df1['ЛКВ (базовый вариант)'].sum() > summ:
            df1 = df1[:-1]
        df2 = df.copy()
        while df2['ЛКВ (базовый вариант)'].sum() > summ + step:
            df2 = df2[:-1]
        get_pie(variable2.get(), df, df1, df2, step)
    elif variable1.get() == 'Сумма и проценты' and filename and summ and percent:
        summ = float(summ)
        percent = float(percent)
        try:
            step = float(step)
        except Exception:
            step = 0
        df = pd.read_excel(filename, "Рейтинг", usecols="B,D,E,AA", header=8)
        df.drop([0], axis=0, inplace=True)
        df.sort_values(by=['Интегральный рейтинг'], inplace=True, ascending=True)
        df1 = df.copy()
        for index, row in df.iterrows():
            tmp = df1.at[index, 'ЛКВ (базовый вариант)']
            df1.at[index, 'ЛКВ (базовый вариант)'] = tmp * percent / 100
            if df1['ЛКВ (базовый вариант)'].sum() < summ:
                break
        df2 = df.copy()
        for index, row in df.iterrows():
            tmp = df2.at[index, 'ЛКВ (базовый вариант)']
            df2.at[index, 'ЛКВ (базовый вариант)'] = tmp * percent / 100
            if df2['ЛКВ (базовый вариант)'].sum() < summ + step:
                break
        get_pie(variable2.get(), df, df1, df2, step)
    else:
        pass


def check(event):
    """
    Переключатель двух опций подсчета

    :param event:
    :return:
    """
    if event == 'Сумма':
        label3.pack_forget()
        text3.pack_forget()
        w2.pack_forget()
        button1.pack_forget()
        label2.pack_forget()
        text2.pack_forget()
        w2.pack()
        label3.pack()
        text3.pack()
        button1.pack()
    if event == 'Сумма и проценты':
        label3.pack_forget()
        text3.pack_forget()
        w2.pack_forget()
        button1.pack_forget()
        label2.pack()
        text2.pack()
        w2.pack()
        label3.pack()
        text3.pack()
        button1.pack()


def browse_files():
    global filename
    filename = filedialog.askopenfilename(initialdir="/", title="Select a File")
    text.insert(1.0, filename)


window = Tk()

window.title('Белоус Лев')

window.geometry("700x700")

window.config(background="white")

text = Text(width=50, height=1, fg='black', wrap=WORD)
text.pack()

button_explore = Button(window, text="Выбрать файл", command=browse_files)
button_explore.pack()
variable1 = StringVar(window)
variable1.set("Сумма")
w1 = OptionMenu(window, variable1, "Сумма", "Сумма и проценты", command=check)
variable2 = StringVar(window)
variable2.set("Филиал")
w2 = OptionMenu(window, variable2, "Филиал", "Программа проекта", command=check)
w1.pack()
text1 = Text(width=50, height=1, fg='black', wrap=WORD)
text2 = Text(width=50, height=1, fg='black', wrap=WORD)
text3 = Text(width=50, height=1, fg='black', wrap=WORD)
label1 = Label(text='Введите максимальную сумму', justify=LEFT)
label2 = Label(text='Введите максимальный процент', justify=LEFT)
label3 = Label(text='Введите шаг', justify=LEFT)
button1 = Button(text="Получить результат", width=15, height=1, command=result)
label1.pack()
text1.pack()
w2.pack()
label3.pack()
text3.pack()
button1.pack()
window.mainloop()
