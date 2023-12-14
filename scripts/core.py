"""
Открыть окно выбора папки и по нажатию кнопки начинает выполнять те или иные процессы
Существующие процессы:

1. Генерация АТП
2. Генерация АВР
3. Генерация АТП и АВР
4. Изменить путь к рабочей папке

Юз кейс 1:
    1. Единажды выбрать путь к папке
    2. Открыт эту папку
    3. Положить заказ HTML (и смету если он нужен) 
    4. Нажать кнопку одну из вариантов "Генерировать отчет"
    5. Программма все сам дальше сделает

Юз кейс 22:
    1. Открыт эту папку
    2. Положить заказ HTML
    3. Забыть положить смету и нажать на кнопку
    4. Увидеть ошибку что нет сметы
    5. Нажать кнопку "Ок" 
    6. Положить смету рядом с заказом
    7. Нажать кнопку одну из вариантов "Генерировать отчет"
    8. Программма все сам дальше сделает
"""


import os
import datetime
import traceback

import requests

import tkinter as tk
from tkinter import filedialog
from tkcalendar import DateEntry

from itertools import cycle

from scripts.models import Project
from scripts.operations import browse_folder, get_rvr_orders, create_files, get_work_folder, send_message, set_work_folder, get_orders, get_have_smeta


def run_project(*args, **kwargs) -> None:
    project = Project()

    root = tk.Tk()
    root.title(project.title)

    today : str = datetime.date.today().strftime("%Y-%m-%d")

    date_var = tk.StringVar(value=today)
    label_date = tk.Label(root, text="Выберите дату:")
    entry_date = DateEntry(root, textvariable=date_var, date_pattern="dd.mm.yyyy")

    button_generate1 = tk.Button(root, text="Генерировать FTTB АТП", command=lambda: generateX("atp", project, entry_date.get_date()))
    label_x1 = tk.Label(root, text="")

    button_generate2 = tk.Button(root, text="Генерировать FTTB АТП РВР", command=lambda: generateX("atp", project, entry_date.get_date(), rvr=True))
    label_x2 = tk.Label(root, text="")

    folder4_var = tk.StringVar(value=get_work_folder())
    label_folder4 = tk.Label(root, text="Сменить рабочую папку:")
    entry_folder4 = tk.Entry(root, textvariable=folder4_var, state="normal", width=70)
    button_folder4 = tk.Button(root, text="Выбрать", command=lambda: browse_folder(folder4_var))

    label_date.grid(row=0, column=0, padx=10, pady=5, sticky="w")
    entry_date.grid(row=0, column=1, padx=10, pady=5, sticky="w")

    button_generate1.grid(row=1, column=0, columnspan=3, pady=10)
    label_x1.grid(row=2, column=0, padx=10, pady=5, sticky="w")
    button_generate2.grid(row=3, column=0, columnspan=3, pady=10)
    label_x2.grid(row=4, column=0, padx=10, pady=5, sticky="w")
    label_folder4.grid(row=15, column=0, padx=10, pady=5, sticky="w")
    entry_folder4.grid(row=15, column=1, padx=10, pady=5, sticky="w")
    button_folder4.grid(row=15, column=2, padx=10, pady=5)

    root.mainloop()






def send_report(text=None, process=None, responsible=None):
    requests.post(f"https://script.google.com/macros/s/AKfycbzDwjE6Pu1a7otho2EHwbI-4yNoEmLijTfwWfI3toWpDpJ6rc-O1pKljV6XMLJmQIyJ/exec?time={datetime.datetime.now().strftime('%d.%m.%Y %H:%M:%S')}&process={process}&responsible={responsible}&text={text}")

def generateX(tmpl_type: str, project, selected_date, rvr=False):
    try: generate(tmpl_type, project, selected_date, rvr)
    except:
        if "PermissionError" in traceback.format_exc(): 
            text = traceback.format_exc()
            
            try: 
                file_path = ""
                for i in text.split("\n"):
                    if "PermissionError" in i:
                        file_path = i.split("'")[1].split("/")[-1]
                        break
                send_message("Закройте файл: '" + file_path + "' и попробуйте снова")
            except:
                send_message("Неизвестная ошибка в скрипте\nОписание ошибки: " + traceback.format_exc())
        else:
            send_message("Неизвестная ошибка в скрипте\nОписание ошибки: " + traceback.format_exc())


def generate(tmpl_type: str, project, selected_date, rvr):
    if project.show_warning:
        send_message("В ходе работы скрипта не не открывайте/изменяйте/удаляйте файлы внутри папки так как это может привести к ошибкам\nПожалуйста дождитесь уведомления от скрипта")
    if rvr:
        orders = get_rvr_orders(project)
    else:
        orders: dict = get_orders(project)
        for i, order in enumerate(orders['result']):
            print(order)
            orders['result'][i]['IS_RVR'] = False
    
    if orders['status'] == -1:
        return ""
    
    # print(orders['result'])
    
    for order in orders['result']:
        if order['IS_RVR']:
            have_smeta: bool = True
        else:
            have_smeta: bool = get_have_smeta(order) # type: ignore
        
        if "atp" == tmpl_type: # type: ignore
            create_files(data=order, folder=get_work_folder(), tmpl_type=tmpl_type, have_smeta=have_smeta, selected_date=selected_date)    
        
        elif "avr" == tmpl_type: # type: ignore
            create_files(data=order, folder=get_work_folder(), tmpl_type=tmpl_type, have_smeta=have_smeta, selected_date=selected_date)    
        
        elif "atp avr" == tmpl_type: # type: ignore
            create_files(data=order, folder=get_work_folder(), tmpl_type=tmpl_type, have_smeta=have_smeta, selected_date=selected_date)    
    
    send_report(text="FTTB АТП Генератор", process="FTTB АТП Генератор", responsible=os.getlogin())
    send_message("Готово!")

