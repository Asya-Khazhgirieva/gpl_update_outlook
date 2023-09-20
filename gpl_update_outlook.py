import win32com.client as win32
import time as t
import pywinauto
from pywinauto.application import Application #для Outlook
from datetime import *
import shutil #работа с папками и файлами Windows
from win32com.client import Dispatch
import os
import zipfile

print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <Outlook>: Создание COM-объекта')
outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")

print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <Входящие>: Поиск папки GPL')
inbox = outlook.Folders["mail@domain.ru"].Folders["Входящие"].Folders["GPL"]
print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <GPL>: Получение писем')
messages = inbox.Items
print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <GPL>: Просмотр последнего письма')
messages = messages.GetLast()
print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <GPL>: Проверка даты письма')
date = messages.SentOn

# Получение текущего года и месяца
year_month = datetime.now().strftime("%Y-%m")
new_GPL_gpl_name = "GPL_" + year_month + ".xlsx"

tz = date.tzinfo
today = datetime.now(tz)
today_gpl = datetime.now()

print('Открываем файл для записи')
f = open('C:/Users/user/Desktop/GPL_log.txt', 'a', encoding='utf-8')

#функции
def copying_column(column_number):
    print("Копируем столбец")
    column = worksheet.Columns(column_number)
    column.Copy()
    column.Insert()

def paste_formula_vpr(cell, formula):
    print("Вставляем формулу ВПР")
    formula_range = worksheet.Range(cell)
    formula_range.FormulaLocal = formula

def stretching_formula(rnge):
    print("Растягиваем формулу до конца столбца")
    range_fill = worksheet.Range(rnge.format(last_row))
    range_fill.FillDown()

def the_sleep(sec):
    print("Задержка на 5 секунд")
    t.sleep(sec)

def filtering_column(column, value):
    print(f"Фильтруем {column} столбец по значению {value}")
    filter_range = worksheet.Range("A1").AutoFilter(column, value)

def transfer_old_value(rng):
    print("Переношу старые значения в новый столбец")
    for cell in worksheet.Range(rng.format(last_row)).SpecialCells(12):
        cell.Offset(1, 2).Value = cell.Value
        cell.Offset(1, 2).Interior.Color = 255

def deleting_formulas(rng):
    print("Выделяю столбец от первой ячейки с ВПР и до последней заполненной ячейки и копирую")
    range_copy = worksheet.Range(rng.format(last_row))
    range_copy.Copy()
    print("Вставляю только значения, без формулы ВПР")
    range_copy.PasteSpecial(-4163)

def delete_old_column(cln):
    print("Удаляю старый столбец")
    column = worksheet.Columns(cln)
    column.Delete(-4159)

def gpl_update_date(cell):
    print("Меняю дату обновления GPL")
    date_str = today.strftime('%d.%m.%Y')
    date_range = worksheet.Range(cell)
    date_s = str(date_str)
    date_range.Value = date_s

def change_fill_eos(rng):
    print("Меняю заливку в соответствии со значением")
    for cell in worksheet.Range(rng.format(last_row)):
        if cell.Value == "есть":
            cell.Interior.ColorIndex = 4
        elif cell.Value == "x":
            cell.Interior.ColorIndex = 37

def saving_closing_workbook():
    print("Сохранение и закрытие книги")
    workbook.Save()
    workbook.Close()
    print("Закрытие Excel")
    excel.Quit()

if today - date < timedelta(days=8):
    print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <GPL>: Скачивание вложений последнего письма')
    attachments = messages.Attachments
    for attachment in attachments:
        if attachment.FileName.startswith("GPL RU pricelist Reseller"):
            attachment_path = os.path.join("C:/Users/user/Desktop/!_GPLи для компани/" + new_GPL_gpl_name)
            attachment.SaveAsFile(attachment_path)
            print(f"Скачано вложение {attachment.FileName}")
            lists.append("Скачано вложение GPL")
            f.write(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + 'Скачано вложение GPL\n')

            #обновляем GPL
            print("Создание объекта Excel")
            excel = win32.Dispatch('Excel.Application')
            excel.Visible = 1
            print("Открытие книги")
            workbook = excel.Workbooks.Open('C:\\Users\\user\\Desktop\\PriceRU\\PriceRU_.xlsx')
            print("Установка активного листа")
            worksheet = workbook.ActiveSheet

            copying_column(25)
            
            now = datetime.now()
            month_abbr = now.strftime('%b')  # вернет сокращенное название текущего месяца
            folder_name = month_abbr + " MoRU"

            paste_formula_vpr('Z4', '=ВПР(F4;\'\\\\IP\\Users\\user\\Desktop\\!_GPLи для компани\\[' + new_GPL_gpl_name + ']' + folder_name + '\'!$A:$J;10;ЛОЖЬ)/$AE$1')

            print("Определяем последнюю строку диапазона")
            last_row = worksheet.Cells(worksheet.Rows.Count, 26).End(-4162).Row

            stretching_formula("Z4:Z{}")
            the_sleep(5)
            filtering_column(26, "#Н/Д")
            filtering_column(25, "<>#Н/Д")
            transfer_old_value("Y4:Y{}")

            print("Убираю фильтры с листа")
            worksheet.AutoFilterMode = False
            worksheet.Names("_FilterDatabase").Delete()

            deleting_formulas("Z4:Z{}")
            delete_old_column(25)
            gpl_update_date("Y3")


            #Обновление EOS
            copying_column(4)
            paste_formula_vpr("E4", '=ЕСЛИ(C4="x";"x";ЕСЛИОШИБКА(ЕСЛИ(ВПР(F4;\'\\\\IP\\Users\\user\\Desktop\\!_GPLи для компани\\[' + new_GPL_gpl_name + ']' + folder_name + '\'!$A:$B;2;ЛОЖЬ)="EOL";"EOS";"есть");"есть"))')
            stretching_formula("E4:E{}")
            the_sleep(5)
            filtering_column(5, "есть")
            filtering_column(4, "EOS")
            transfer_old_value("D4:D{}")

            print("Убираю фильтры с листа")
            worksheet.AutoFilterMode = False

            deleting_formulas("E4:E{}")
            delete_old_column(4)
            change_fill_eos("D4:D{}")
            saving_closing_workbook()
            
else:
    print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <GPL>: Дата последнего письма более недели назад')
    f.write(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' Дата последнего письма GPL более недели назад\n')


#GPL Dist
print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <Входящие>: Поиск папки GPL Dist')
inbox = outlook.Folders["mail@domain.ru"].Folders["Входящие"].Folders["GPL Dist"]
print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <GPL>: Получение писем')
messages = inbox.Items
print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <GPL>: Просмотр последнего письма')
messages = messages.GetLast()
print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <GPL>: Проверка даты письма')
date = messages.SentOn

new_GPL_gpl__Dist_name = "GPL_Dist_" + year_month + ".xlsx"

if today - date < timedelta(days=8):
    print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <GPL>: Скачивание вложений последнего письма')
    attachments = messages.Attachments
    for attachment in attachments:
        if attachment.FileName.startswith("GPL_price_Dist.zip"):
            src = 'C:/Users/user/Desktop/'
            filename = 'GPL_price_Dist.zip'
            attachment.SaveASFile(src + filename)
            print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <GPL>: Извлечение содержимого архива GPL_price_Dist.zip в Desktop')
            z = zipfile.ZipFile('C:/Users/user/Desktop/GPL_price_Dist.zip', 'r')
            z.extractall('C:/Users/user/Desktop/')
            z.close()

            #сохраняем новый GPL
            print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <GPL>: Создание COM-объекта Excel')
            Dist = win32.Dispatch('Excel.Application')
            Dist.Visible = 1
            print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <GPL>: Назначение имени файла GPL_price_Dist.zip')
            list_files=list()
            for name in z.namelist(): 
                list_files.append(name) 
            wb = Dist.Workbooks.Open('C:/Users/user/Desktop/' + name, None, True)
            print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <GPL>: Сохранение GPL_price_Dist.zip в !_GPLи для компани')
            wb.SaveCopyAs("C:/Users/user/Desktop/!_GPLи для компани/" + new_GPL_gpl__Dist_name)
            print(f"Скачано вложение {attachment.FileName}")
            f.write(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' Скачано вложение GPL_price_Dist\n')

            wb.Close()
            print("Закрытие Excel")
            Dist.Quit()

            os.remove('C:/Users/user/Desktop/GPL_price_Dist.zip')
            os.remove('C:/Users/user/Desktop/' + name)

            #обновляем GPL Дилер Dist
            print("Создание объекта Excel")
            excel = win32.Dispatch('Excel.Application')
            excel.Visible = 1
            print("Открытие книги")
            workbook = excel.Workbooks.Open('C:\\Users\\user\\Desktop\\PriceRU\\PriceRU_GPL.xlsx')
            print("Установка активного листа")
            worksheet = workbook.ActiveSheet

            copying_column(23)
            paste_formula_vpr('X4', '=ВПР(F4;\'\\\\IP\\Users\\user\\Desktop\\!_GPLи для компани\\[' + new_GPL_gpl__Dist_name + ']Price_EU_GPL_Dist\'!$B:$H;7;ЛОЖЬ)')

            print("Определяем последнюю строку диапазона")
            last_row = worksheet.Cells(worksheet.Rows.Count, 26).End(-4162).Row

            stretching_formula("X4:X{}")
            the_sleep(5)
            filtering_column(24, "#Н/Д")
            filtering_column(23, "<>#Н/Д")
            transfer_old_value("W4:W{}")

            print("Убираю фильтры с листа")
            worksheet.AutoFilterMode = False
            worksheet.Names("_FilterDatabase").Delete()

            deleting_formulas("X4:X{}")
            delete_old_column(23)
            gpl_update_date("W3")

            #обновляем GPL РРЦ Dist
            copying_column(24)
            paste_formula_vpr('Y4', '=ВПР(F4;\'\\\\IP\\Users\\user\\Desktop\\!_GPLи для компани\\[' + new_GPL_gpl__Dist_name + ']Price_EU_GPL_Dist\'!$B:$G;6;ЛОЖЬ)')

            stretching_formula("Y4:Y{}")
            the_sleep(5)
            filtering_column(25, "#Н/Д")
            filtering_column(24, "<>#Н/Д")
            transfer_old_value("X4:X{}")

            print("Убираю фильтры с листа")
            worksheet.AutoFilterMode = False

            deleting_formulas("Y4:Y{}")
            delete_old_column(24)
            gpl_update_date("X3")
            saving_closing_workbook()
            
else:
    print(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <GPL>: Дата последнего письма более недели назад')
    f.write(datetime.strftime(datetime.now(), "%d.%m.%Y %H:%M:%S") + ' Дата последнего письма от Dist более недели назад\n')
f.write('\n')
f.close()
