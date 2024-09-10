import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import warnings
import matplotlib.pyplot as plt
import numpy as np

warnings.simplefilter("ignore")

# Main window
root = tk.Tk()
root.title( "Адаптивное прогнозирование")
root.geometry("500x450")
root.pack_propagate(False)
root.resizable(0, 0)

# Frames
file_frame = tk.LabelFrame(root, text="Дынные для работы")
file_frame.place(height=120, width=490, relx=0.01)

file_frame1 = tk.LabelFrame(root, text="Построение трендовой модели")
file_frame1.place(height=120, width=300, rely=0.3, relx=0.01)

file_frame2 = tk.LabelFrame(root, text="Расчет прогноза")
file_frame2.place(height=120, width=300, rely=0.6, relx=0.01)


# Buttons
button1 = tk.Button(file_frame, text="Загрузить данные",
                    command=lambda: File_dialog())
button1.place(rely=0.60, relx=0.15)

button2 = tk.Button(file_frame, text="Показать начальные данные",
                    command=lambda: Load_excel_data())
button2.place(rely=0.30, relx=0.55)

button3 = tk.Button(file_frame, text="График динамики",
                    command=lambda: dinamic())
button3.place(rely=0.60, relx=0.61)

button4 = tk.Button(file_frame1, text="Построить трендовую модель",
                    command=lambda: func())
button4.place(rely=0.20, relx=0.20)

button5 = tk.Button(file_frame1, text="График функции",
                    command=lambda: graf())
button5.place(rely=0.60, relx=0.31)

button6 = tk.Button(file_frame2, text="Рассчитать прогноз",
                    command=lambda: coff_window())
button6.place(rely=0.30, relx=0.29)

button7 = tk.Button(root, text="Экспорт в Excel",
                    command=lambda: export_to_excel())
button7.place(y=270, x=353)

button8 = tk.Button(root, text="Вывести данные",
                    command=lambda: show())
button8.place(y=230, x=350)

# The file/file path text
label_file = ttk.Label(file_frame, text="Файл не выбран")
label_file.place(rely=0, relx=0.03)

# Global vars
new = []
a = None
b = None
i1 = None
i2 = None
i3 = None
i4 = None

#Выбор файла и первоначальная обработка
def File_dialog():
    file_path = filedialog.askopenfilename(initialdir="/",
                                          title="Выбрать файл",
                                          filetype=(("xlsx files", "*.xlsx"),
                                                    ("All Files", "*.*")))
    label_file["text"] = file_path
    global df
    try:
        excel_filename = r"{}".format(file_path)
        if excel_filename[-4:] == ".csv":
            df = pd.read_csv(excel_filename, engine="openpyxl")
        else:
            df = pd.read_excel(excel_filename, engine="openpyxl")

    except ValueError:
        tk.messagebox.showerror("Ошибика",
                                "Неверный формат файла. Выберете .xlsx или .csv")
        return False
    except FileNotFoundError:
        tk.messagebox.showinfo("Информация",
                                f"Вы не выбрали файл или данный файл отсутствует")
        return False
    
    basa = df.values
    global new
    del new[:]
    a = []
    for i in range(1,len(basa[0])):
        a +=[row[i] for row in basa]
    a = [item for item in a if not(pd.isnull(item)) == True]
    
    for i in range(len(a)):
        new.append([i+1, a[i]])
    
    return None

#Вывод таблицы в новом окне
def Load_excel_data():
   
    newwindow = tk.Tk()
    newwindow.title( "Таблица данных")
    newwindow.geometry("900x500")
    newwindow.pack_propagate(False)
    newwindow.resizable(0, 0)

    tv1 = ttk.Treeview(newwindow)
    tv1.place(relheight=1, relwidth=1)

    treescrolly = tk.Scrollbar(newwindow, orient="vertical",
                               command=tv1.yview)
    treescrollx = tk.Scrollbar(newwindow, orient="horizontal",
                               command=tv1.xview)
    tv1.configure(xscrollcommand=treescrollx.set,
                  yscrollcommand=treescrolly.set)
    treescrollx.pack(side="bottom", fill="x")
    treescrolly.pack(side="right", fill="y")
    
    tv1["column"] = list(df.columns)
    tv1["show"] = "headings"
    for column in tv1["columns"]:
        tv1.heading(column, text=column)

    df_rows = df.to_numpy().tolist()
    for row in df_rows:
        tv1.insert("", "end", values=row)
    return None

def dinamic():
    global new
    if len(new) == 0:
        tk.messagebox.showerror("Ошибка",
                                "Данные не были загружены в программу. Загрузите файл")
    else:
        r = pd.DataFrame(new)
        plt.gcf().canvas.manager.set_window_title("График годовой динамики")
        plt.title("График годовой динамики")
        r[1].plot()
        plt.show()

    return None

def func():
    global new
    global a
    global b
    global i1
    global i2
    global i3
    global i4
    if len(new) == 0:
        tk.messagebox.showerror("Ошибка",
                                "Данные не были загружены в программу. Загрузите файл")
    elif len(new[0]) > 2:
        tk.messagebox.showinfo("Информация", "Расчет трендовой модели уже был осуществлен")
    else:
        xs = 0
        ys = 0
        for i in range(len(new)):
            xs += new[i][0]
            ys += new[i][1]
        xs = xs/len(new)
        ys = ys/len(new)
        
        summ = 0
        kvad = 0
        for i in range(len(new)):
            summ += (new[i][0]-xs)*(new[i][1]-ys)
            kvad += (new[i][0]-xs)**2

        a = summ/kvad
        b = ys - a*xs
        
        for i in range(len(new)):
            new[i].append(a*new[i][0]+b)
            new[i].append(new[i][1]/new[i][2])

        i1 = int(len(new)*0.2)
        i2 = int(len(new)*0.4)
        i3 = int(len(new)*0.8)

        for i in range(i1):
            new[i].append(0)
        for i in range(0,len(new),2):
            new[i1].append((new[i][3]+new[i+1][3])/2)
            i1 += 1
        for i in range(i1,len(new)):
            new[i].append(0)

        i4 = i1 - int(len(new)*0.2)
        i1 = int(len(new)*0.2)
    
    return None

def graf():
    global new
    global a
    global b
    if len(new) == 0:
        tk.messagebox.showerror("Ошибка", "Данные не были загружены в программу. Загрузите файл и расчитайте трендовую модель")
    elif len(new[0]) <= 2:
        tk.messagebox.showerror("Ошибка", "Трендовая модель не была построена")
    else:
        x = [row[0] for row in new]
        y = [row[2] for row in new]
        g = [row[1] for row in new]
        plt.gcf().canvas.manager.set_window_title("График функции и динамики")
        plt.title("График функции и динамики")
        plt.plot(x, y)
        plt.plot(x, g)
        plt.legend(["$y = {}*x + {}$".format("%.2f" % a, "%.2f" % b),"Годовая динамика"], loc=2)
        plt.show()

    return None

def coff():
    global new
    global i1
    global i2
    global i4
    global a
    global b
    j1 = i1
    j2 = i2
    j4 = i4
    for i in range(j2):
        new[i].append(0)
        new[i].append(0)
        new[i].append(0)
        
    new[j2].append(a1*new[j2][1]/new[j1][4] + (1-a1)*(a+b))
    new[j2].append(a2*new[j2][1]/new[j2][5] + (1-a2)*new[j1][4])
    new[j2].append(a3*(new[j2][5]-b) + (1-a3)*a)

    j1 +=1
    j4 = j2+j4

    j2 +=1

    for i in range(j2,j4):    
        new[i].append(a1*new[i][1]/new[j1][4]
                      + (1-a1)*(new[i-1][5]+new[i-1][7]))
        new[i].append(a2*new[i][1]/new[i][2] + (1-a2)*new[j1][4])
        new[i].append(a3*(new[i][5]-new[i-1][5]) + (1-a3)*new[i-1][7])
        j1 +=1
    for i in range(j4,len(new)):
        new[i].append(0)
        new[i].append(0)
        new[i].append(0)
        
    j2 = len(new) - 2*(len(new)-j4)
    j1 = 0
    for i in range (j2):
        new[i].append(0)
        new[i].append(0)        
    for i in range (j2,len(new)):
        new[i].append((new[j4-1][5]
                       + new[j4-1][7]*new[j1][0])*new[j2-(len(new)-j4)][6])
        new[i].append((new[i][1]-new[i][8])/new[i][1] * 100)
        j1 += 1
        j2 += 1
    
    return None
    

def coff_window():
    if len(new) == 0:
        tk.messagebox.showerror("Ошибка",
                                "Данные не были загружены в программу. Загрузите файл")
    elif len(new[0]) == 2:
        tk.messagebox.showerror("Ошибка", "Трендовая модель не была построена")
    else:    
        newwindow = tk.Tk()
        newwindow.title( "Параметры адаптации")
        newwindow.geometry("300x200")
        newwindow.pack_propagate(False)
        newwindow.resizable(0, 0)

        heading_label = tk.Label(newwindow, text = "Коэффициент a1")
        heading_label.pack()
        heading_label.place(x=30, y=20)
        name_field = tk.Entry(newwindow, width = 15)
        name_field.pack()
        name_field.place(x = 170,y = 20)

        heading_label1 = tk.Label(newwindow, text = "Коэффициент a2")
        heading_label1.pack()
        heading_label1.place(x=30, y=50)
        name_field1 = tk.Entry(newwindow, width = 15)
        name_field1.pack()
        name_field1.place(x = 170, y = 50)

        heading_label2 = tk.Label(newwindow, text = "Коэффициент a3")
        heading_label2.pack()
        heading_label2.place(x = 30, y = 80)
        name_field2 = tk.Entry(newwindow, width = 15)
        name_field2.pack()
        name_field2.place(x = 170, y =80)

        button1 = tk.Button(newwindow, text = "Рассчитать",
                            command = lambda:entry(), width=25).place(x=57, y=130)


        def entry():
            global a1
            global a2
            global a3
            
            if len(new[0]) > 5:
                for i in range(len(new)):
                    new[i].pop()
                    new[i].pop()
                    new[i].pop()
                    new[i].pop()
                    new[i].pop()
            
            try:
                float(name_field.get())
                float(name_field1.get())
                float(name_field2.get())
            except ValueError:
                messagebox.showerror("Ошибка",
                                    "Неверный формат данных или не все поля заполнены")
                return False
        
            a1 = float(name_field.get())
            a2 = float(name_field1.get())
            a3 = float(name_field2.get())
            newwindow.destroy()
            coff()
            button6["text"] = "Пересчитать прогноз\nс новыми коэффициентами"
            button6.place(rely=0.30, relx=0.21)
                
def show():
    global new
    newwindow = tk.Tk()
    newwindow.title( "Таблица расчета")
    newwindow.geometry("1000x600")
    newwindow.pack_propagate(False)
    newwindow.resizable(0, 0)

    tv1 = ttk.Treeview(newwindow)
    tv1.place(relheight=1, relwidth=1)
    treescrolly = tk.Scrollbar(newwindow, orient="vertical",
                               command=tv1.yview)
    treescrollx = tk.Scrollbar(newwindow, orient="horizontal",
                               command=tv1.xview)
    tv1.configure(xscrollcommand=treescrollx.set,
                  yscrollcommand=treescrolly.set)
    treescrollx.pack(side="bottom", fill="x")
    treescrolly.pack(side="right", fill="y")
    df = pd.DataFrame(new)
    df.rename(columns={0: "№", 1: "x", 2: "Трендовая модель", 3: "Сезонная составляющая", 4: "Усредненная оценка сезонной составляющей",
                      5: "a1", 6: "f", 7: "a2", 8:"Прогнозные расчеты", 9: "Относительная ошибка"}, inplace=True)
    tv1["column"] = list(df.columns)
    tv1["show"] = "headings"
    for column in tv1["columns"]:
        tv1.heading(column, text=column)

    df_rows = df.to_numpy().tolist()
    for row in df_rows:
        tv1.insert("", "end", values=row)

    return None

def export_to_excel():
    global new
    r = pd.DataFrame(new)
    r = r.replace(0, np.nan)
    r.rename(columns={0: "№", 1: "x", 2: "Трендовая модель", 3: "Сезонная составляющая", 4: "Усредненная оценка сезонной составляющей",
                      5: "a1", 6: "f", 7: "a2", 8:"Прогнозные расчеты", 9: "Относительная ошибка"}, inplace=True)

    file_path = filedialog.asksaveasfilename(initialdir="/",
                                             title="Сохранить файл",
                                             filetype=(("xlsx files", "*.xlsx"),
                                                       ("All Files", "*.*")))

    r.to_excel(file_path + ".xlsx", index=False)

root.mainloop()
