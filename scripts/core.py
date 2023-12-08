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
import tkinter as tk
from tkinter import filedialog
import traceback
from scripts.models import Project
from itertools import cycle


from scripts.operations import browse_folder, create_files, get_work_folder, send_message, set_work_folder, get_orders, get_have_smeta


def run_project(*args, **kwargs) -> None:
    project = Project()

    root = tk.Tk()
    root.title(project.title)

 
    # 1. Генерировать АТП
    button_generate1 = tk.Button(root, text="Генерировать FTTB АТП", command=lambda: generateX("atp", project))

    # разделитель
    label_x1 = tk.Label(root, text="")

    # 2. Генерировать АВР
    # button_generate2 = tk.Button(root, text="Генерировать АВР", command=lambda: generate("avr", project))

    # разделитель
    # label_x2 = tk.Label(root, text="")

    # 3. Генерировать АТП и АВР
    # button_generate3 = tk.Button(root, text="Генерировать АТП и АВР", command=lambda: generate("atp avr", project))

    # разделитель
    # label_x3 = tk.Label(root, text="")
    
    # 4. Генерация АТП и АВР с сметой
    folder4_var = tk.StringVar(value=get_work_folder())
    label_folder4 = tk.Label(root, text="Сменить рабочую папку:")
    entry_folder4 = tk.Entry(root, textvariable=folder4_var, state="normal" , width=70)
    button_folder4 = tk.Button(root, text="Выбрать", command=lambda: browse_folder(folder4_var))
    # button_generate4 = tk.Button(root, text="Сохранить новый путь", command=lambda: set_work_folder(folder4_var.get()))

    # позиции элементов 1
    button_generate1.grid(row=0, column=0, columnspan=3, pady=10)

    # позиция разделителя 1
    label_x1.grid(row=1, column=0, padx=10, pady=5, sticky="w")

    # позиции элементов 2
    # button_generate2.grid(row=2, column=0, columnspan=3, pady=10)

    # позиция разделителя 2
    # label_x2.grid(row=3, column=0, padx=10, pady=5, sticky="w")

    # позиции элементов 3
    # button_generate3.grid(row=4, column=0, columnspan=3, pady=10)

    # # позиция разделителя 3
    # label_x3.grid(row=5, column=0, padx=10, pady=5, sticky="w")

    # позиции элементов 4
    label_folder4.grid(row=14, column=0, padx=10, pady=5, sticky="w")
    entry_folder4.grid(row=14, column=1, padx=10, pady=5, sticky="w")
    button_folder4.grid(row=14, column=2, padx=10, pady=5)
    # button_generate4.grid(row=15, column=0, columnspan=3, pady=10)

    root.mainloop()


import requests
import datetime

def send_report(text=None, process=None, responsible=None):
    requests.post(f"https://script.google.com/macros/s/AKfycbzDwjE6Pu1a7otho2EHwbI-4yNoEmLijTfwWfI3toWpDpJ6rc-O1pKljV6XMLJmQIyJ/exec?time={datetime.datetime.now().strftime('%d.%m.%Y %H:%M:%S')}&process={process}&responsible={responsible}&text={text}")

def generateX(tmpl_type: str, project):
    try: generate(tmpl_type, project)
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
def generate(tmpl_type: str, project):
    if project.show_warning:
        send_message("В ходе работы скрипта не не открывайте/изменяйте/удаляйте файлы внутри папки так как это может привести к ошибкам\nПожалуйста дождитесь уведомления от скрипта")
    orders: dict = get_orders(project)
   
    if orders['status'] == -1:
        return
    
    # print(orders['result'])
    
    for order in orders['result']:
        

        have_smeta: bool = get_have_smeta(order) # type: ignore
        
        if "atp" == tmpl_type: # type: ignore
            create_files(data=order, folder=get_work_folder(), tmpl_type=tmpl_type, have_smeta=have_smeta)    
        
        elif "avr" == tmpl_type: # type: ignore
            create_files(data=order, folder=get_work_folder(), tmpl_type=tmpl_type, have_smeta=have_smeta)    
        
        elif "atp avr" == tmpl_type: # type: ignore
            create_files(data=order, folder=get_work_folder(), tmpl_type=tmpl_type, have_smeta=have_smeta)    
        
    send_report(text="FTTB АТП Генератор", process="FTTB АТП Генератор", responsible=os.getlogin())
    send_message("Готово!")

