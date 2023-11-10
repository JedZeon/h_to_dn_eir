import os
import time
import zipfile
import tempfile
from sys import exit
from tkinter import *
from tkinter import Tk, ttk, X, IntVar, filedialog, messagebox


class Zapis:
    def __init__(self):
        self.enp = ""
        self.plan_month = 0
        self.place = 1  # 1 - поликлиника
        self.ds = ""
        self.idspec = 0
        self.doc_snils = ""
        self.c_zab = 0
        self.is_dn = False  # это случай диспансерного наблюдения


class ParsFile:
    def __init__(self):
        self.file_path = self.name_file()
        self.file_name = ""
        self.list_zapis = []  # список данных для выгрузки массив формата Zapis

    def name_file(self):
        file_path = filedialog.askopenfile(filetypes=(("Пакеты для фонда H* zip", "H*.zip"), ("все файлы", "*.*")))
        return file_path

    def get_value(self, atribute_, text_):
        result = {"value": "", "founding": False}
        position = text_.find("<" + atribute_ + ">", 0)
        if position >= 0:
            start_position = len("<" + atribute_ + ">") + position
            end_position = text_.find("</" + atribute_ + ">", 0)
            result["value"] = text_[start_position:end_position]
            result["founding"] = True

        return result

    def pars_file(self):
        # Является ли файл zip архивом
        if zipfile.is_zipfile(self.file_path.name):
            z = zipfile.ZipFile(self.file_path.name, 'r')

            # month_ = selected_month.get()
            month_ = months.index(combobox.get())+1

            # Создаём временный каталог
            with tempfile.TemporaryDirectory() as tmp_dir:
                # распаковываем во временный каталог
                z.extractall(path=tmp_dir)

                # перебираем файлы и ищем файл H
                for filename in os.listdir(path=tmp_dir):
                    # print(filename)
                    if filename[0].upper() == "H":

                        path_filename = os.sep.join([tmp_dir, filename])

                        # Открываем и читаем файл построчно
                        xxx = ""  # трехзначный  номер медицинской организации из регионального справочника
                        gggg = ""  # отчетный год

                        self.list_zapis = []
                        line_count = sum(1 for line in open(path_filename))

                        with open(path_filename, 'r') as file_:

                            progress_ = 1
                            if line_count < 100:
                                grad_progr = round(100 / line_count, 1)
                            else:
                                grad_progr = round(line_count / 100, 1)

                            zagolovok = True  # заполнение первичных параметров

                            kol = 0
                            for line in file_:
                                kol += 1
                                if kol > grad_progr:
                                    progress_ += 1
                                    value_var.set(progress_)
                                    kol = 0

                                if zagolovok:
                                    data = self.get_value("DATA", line)
                                    if data["founding"]:
                                        gggg = data["value"].split('-')[0][-4::]

                                    data = self.get_value("CODE_MO", line)
                                    if data["founding"]:
                                        xxx = data["value"][-3::]

                                    data = self.get_value("ZAP", line)
                                    if data["founding"]:
                                        zagolovok = False
                                        self.file_name = f"dn{xxx}_{gggg}.csv"

                                if not zagolovok:
                                    data = self.get_value("ZAP", line)
                                    if data["founding"]:
                                        zap_ = Zapis()
                                        zap_.plan_month = month_

                                    data = self.get_value("NPOLIS", line)
                                    if data["founding"]:
                                        zap_.enp = data["value"]

                                    data = self.get_value("P_CEL", line)
                                    if data["founding"]:
                                        if data["value"] == "1.3":
                                            zap_.is_dn = True

                                    data = self.get_value("DS1", line)
                                    if data["founding"]:
                                        zap_.ds = data["value"]

                                    data = self.get_value("PRVS", line)
                                    if data["founding"]:
                                        zap_.idspec = data["value"]

                                    data = self.get_value("IDDOKT", line)
                                    if data["founding"]:
                                        zap_.doc_snils = data["value"]

                                    data = self.get_value("C_ZAB", line)
                                    if data["founding"]:
                                        if data["value"] == 2:
                                            zap_.doc_snils = 2
                                        else:
                                            zap_.doc_snils = 0

                                    position = line.find("</ZAP>", 0)
                                    if position >= 0 and zap_.is_dn:
                                        self.list_zapis.append(zap_)

    def save_file(self):
        path_file = filedialog.asksaveasfilename(initialdir=self.file_path, initialfile=self.file_name, filetypes=(("Файлы для ЕИР csv", "*.csv"), ("все файлы", "*.*")))
        if path_file != "":
            with open(path_file, "w") as file:
                file.write("num_str;enp;plan_month;place;ds;idspec;doc_snils;c_zab\n")

                num = 0
                for el in self.list_zapis:
                    num += 1
                    str_ = f"{num};{el.enp};{el.plan_month};{el.place};{el.ds};{el.idspec};{el.doc_snils};{el.c_zab}"
                    file.write(str_ + "\n")


def open_and_pars():
    pars = ParsFile()
    pars.pars_file()
    pars.save_file()


def stop():
    progressbar.stop()  # останавливаем progressbar


main_win = Tk()
main_win.title("Создание файла ДН, для ЕИР из H*")

value_var = IntVar(master=main_win, value=0)

# Оформление и размещение элементов окна

# Добавим выбор месяца
def selected(event):
    # получаем выделенный элемент
    selection = combobox.get()
    header["text"] = f"вы выбрали: {selection}"
    start_btn['state'] = 'enabled'

position = {"padx": 6, "pady": 6, "anchor": NW}
months = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь",
          "Декабрь"]
selected_month = StringVar()

header = ttk.Label(text="Запланируем на месяц")
header.grid(column=0, row=0)

combobox = ttk.Combobox(values=months, state="readonly")
combobox.grid(column=1, row=0)
combobox.bind("<<ComboboxSelected>>", selected)


start_btn = ttk.Button(text="Открыть файл", command=open_and_pars)
start_btn.grid(column=2, row=0)
start_btn['state'] = 'disabled'
stop_btn = ttk.Button(text="Прервать", command=stop)
stop_btn.grid(column=3, row=0)

def show_help_window():
    messagebox.showinfo('Справка', 'Назначение: \n'
                                   '- анализ архивных файлов амбулаторных пакетов H*,\n'
                                   '- поиск случаев с целью 1.3 - Диспансерное наблюдение\n'
                                   '- формирование файла, готового для загрузки в ЕИР dnXXX_ГГГГ.csv\n'
                                   '* планирование месяца посещения, выбирается из списка\n\n'
                                   'enp         - Единый номер полиса гражданина\n'
                                   'plan_month  - Планируемый месяц проведения диспансерного наблюдения\n'
                                   'place       - Место проведения: 1 - поликлиника, 2 - на дому\n'
                                   'ds          - Диагноз\n'
                                   'idspec      - Код врачебной специальности\n'
                                   'doc_snils   - СНИЛС врача\n'
                                   'c_zab       - Характер основного заболевания: \n'
                                   '            2 – Впервые в жизни установленное хроническое, 0 -в ином случае\n\n'
                                   'обратная связь: jedcrb@mail.ru')

stop_btn = ttk.Button(text="?", command=show_help_window, width=3)
stop_btn.grid(column=4, row=0)

progressbar = ttk.Progressbar(length=350, variable=value_var)
progressbar.grid(column=0, row=1, columnspan=3)

label = ttk.Label(textvariable=value_var)
label.grid(column=3, row=1)

label = ttk.Label(text="© Jed КГБУЗ Пограничная ЦРБ")
label.grid(column=0, row=2, columnspan=4 )

main_win.mainloop()
