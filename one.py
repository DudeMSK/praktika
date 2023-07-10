import sys,os
import re
from collections import Counter
import win32com.client
import docx
from docx import Document
import textract
import shutil
from pathlib import Path
import glob
import docx2txt
from subprocess import Popen, PIPE
import pytesseract

def find_uin_number(file_path):
    doc = Document(file_path)

    for para in doc.paragraphs:
        text = para.text
        match = re.search(r"УИН\s*№\s*(\d+)", text, re.UNICODE)
        if match:
            uin_number = match.group(1)
            return uin_number

    return None

master_folder_path = r"C:\Users\anaconda\Desktop\praktika\ТП_2022\ТП"
source_folder_path = r"C:\Users\anaconda\Desktop\praktika\ТП_2022\ТП"
dest_notTXT = r"C:\Users\anaconda\Desktop\praktika\ml\2022\notTXT"
dest_TXT = r"C:\Users\anaconda\Desktop\praktika\ml\2022\TXT"
pol_folder_path = os.path.join(dest_TXT, "pol")

for filename in os.listdir(source_folder_path):
    file_path = os.path.join(source_folder_path, filename)
    if file_path.endswith('.docx') and not filename.startswith('~$'):
        uin_number = find_uin_number(file_path)
        if uin_number:
            digits_only = re.sub(r"\D", "", uin_number)
            print(f"УИН № {digits_only} - {filename}")
            match = re.search(r'(\S+).*?(пол|отр)', filename)
            if match:
                numbers = match.group(1)
                word = match.group(2)
                print("Файл:", file_path)
                print("Категория заключения:", word)
                print("Номер реестра:", numbers)
                print()
                found_folders = []
                for root, dirs, files in os.walk(master_folder_path):
                    for dir_name in dirs:
                        if uin_number in dir_name:
                            folder_path = os.path.join(root, dir_name)
                            found_folders.append((uin_number, folder_path))
                if found_folders:
                    for uin, folder_path in found_folders:
                        print(f"Найдена папка с УИН № {uin} по пути: {folder_path}")
                        for root, dirs, files in os.walk(folder_path):
                            for file_name in files:
                                file_path = os.path.join(root, file_name)
                                if "Материалы по обоснованию в текстовой форме" in file_name:
                                    print(f"Найден файл 'Материалы по обоснованию в текстовой форме' по пути: {file_path}")
                                    dest_folder = dest_notTXT if file_path.endswith('.pdf') else dest_TXT
                                    new_file_name = f'{word} {numbers}{os.path.splitext(file_name)[1]}'
                                    dest_path = os.path.join(dest_folder, new_file_name)
                                    try:
                                        shutil.copy(file_path, dest_path)
                                        print(f"Файл {file_name} успешно скопирован в папку {dest_folder}!")
                                        new_file_path = os.path.join(dest_folder, new_file_name)
                                        os.rename(dest_path, new_file_path)
                                        print(f"Файл {dest_path} успешно переименован в {new_file_path}")
                                        if 'пол' in new_file_name:
                                            pol_folder = os.path.join(dest_folder, 'pol')
                                            if not os.path.exists(pol_folder):
                                                os.makedirs(pol_folder)
                                            shutil.move(new_file_path, pol_folder)
                                            print(f"Файл {new_file_name} перемещен в папку 'pol'")
                                        elif 'отр' in new_file_name:
                                            otr_folder = os.path.join(dest_folder, 'otr')
                                            if not os.path.exists(otr_folder):
                                                os.makedirs(otr_folder)
                                            shutil.move(new_file_path, otr_folder)
                                            print(f"Файл {new_file_name} перемещен в папку 'otr'")
                                    except Exception as e:
                                        print(f"Ошибка при копировании файла {file_name} в папку {dest_folder}: {str(e)}")
                else:
                    print(f"Папки с УИН {uin_number} не найдены в мастер-папке.")
                    print()
        else:
            print(f"Значение УИН не найдено в файле {filename}")
            print()
    else:
        
        print(f"Неверный формат файла: {filename}")

def convert_doc_to_docx(doc_file_path, docx_file_path):
    # Создаем объект приложения Word
    word_app = win32com.client.Dispatch("Word.Application")
    # Открываем документ .doc
    doc = word_app.Documents.Open(doc_file_path)
    # Сохраняем его в формате .docx
    doc.SaveAs(docx_file_path, 16)  # 16 означает формат .docx
    # Закрываем документ и выходим из приложения Word
    doc.Close()
    word_app.Quit()

for new_file_name in os.listdir(pol_folder_path):
    file_path = os.path.join(pol_folder_path, new_file_name)
    print(f"{new_file_name} находится в {pol_folder_path}!")
    try:
        if file_path.endswith('.doc'):
            # Конвертировать .doc в .docx
            docx_filename = os.path.splitext(new_file_name)[0] + '.docx'
            docx_file_path = os.path.join(pol_folder_path, docx_filename)
            convert_doc_to_docx(file_path, docx_file_path)
            print(f"Файл {new_file_name} успешно сконвертирован в {docx_filename}!")
    except Exception as e:
        print(f"Ошибка при конвертации файла {new_file_name} в папку {docx_filename}: {str(e)}")

for new_file_name in os.listdir(pol_folder_path):
    file_path = os.path.join(pol_folder_path, new_file_name)
    print(f"{new_file_name} находится в {pol_folder_path}!")
    try:
        if file_path.endswith('.docx'):
            txt_filename = os.path.splitext(new_file_name)[0] + '.txt'
            txt_file_path = os.path.join(pol_folder_path, txt_filename)
            doc = Document(file_path)
            text = "\n".join([paragraph.text for paragraph in doc.paragraphs])
            with open(txt_file_path, "w", encoding="utf-8") as txt_file:
                txt_file.write(text)
            print(f"Файл {new_file_name} успешно сконвертирован в {txt_filename}!")
    except Exception as e:
        print(f"Ошибка при конвертации файла {new_file_name} в папку {pol_folder_path}: {str(e)}")
        
#Определение функции find_uin_number, которая ищет номер УИН (уникального идентификационного номера) в переданном документе. Для этого функция открывает документ с помощью библиотеки docx, проходит по параграфам документа и ищет совпадения с помощью регулярного выражения.
Задание путей к папкам и файлам, которые будут использоваться в процессе. Например, master_folder_path - это путь к мастер-папке, source_folder_path - путь к папке, из которой будут браться файлы для обработки, и т.д.

Цикл for для перебора файлов в папке source_folder_path. Для каждого файла выполняются следующие шаги:

Проверка, что файл имеет расширение .docx и не является временным файлом (~$).
Поиск номера УИН в файле с помощью функции find_uin_number.
Если номер УИН найден, происходит извлечение информации о категории, номере реестра и т.д. из имени файла.
Поиск папок с соответствующими номерами УИН в мастер-папке.
Перебор файлов в найденных папках для выполнения дополнительных действий (например, копирования, переименования, перемещения файлов).
Определение функции convert_doc_to_docx, которая использует библиотеку win32com.client для конвертации документов .doc в .docx. Для этого функция открывает документ .doc с помощью приложения Word, сохраняет его в формате .docx и закрывает документ и приложение Word.

Цикл for для конвертации файлов .doc в .docx. Для каждого файла в папке pol_folder_path выполняются следующие шаги:

Проверка, что файл имеет расширение .doc.
Конвертация файла .doc в .docx с использованием функции convert_doc_to_docx.
Цикл for для конвертации файлов .docx в .txt. Для каждого файла в папке pol_folder_path выполняются следующие шаги:

Проверка, что файл имеет расширение .docx.
Открытие файла .docx с помощью библиотеки docx.
Извлечение текста из файла и сохранение его в формате .txt.
#Внутри каждой папки проходит обработка каждого файла.
#Если найден файл "Материалы по обоснованию в текстовой форме", он копируется в соответствующую папку назначения (в зависимости от расширения файла) и переименовывается.
#Если в новом имени файла содержится подстрока "пол", файл перемещается в папку pol. Если в новом имени файла содержится подстрока "отр", файл перемещается в папку otr.
#В случае возникновения ошибки при копировании файла, выводится соответствующее сообщение об ошибке.
#Если не найдены папки с номером УИН, выводится сообщение о том, что папки не найдены.
#Если номер УИН не найден в файле, выводится сообщение об этом.
#Если файл имеет неправильный формат, выводится сообщение о неправильном формате файла.
#Этот скрипт предназначен для обработки файлов с определенным форматом и структурой и выполняет операции по поиску и обработке этих файлов. Он использует различные библиотеки и модули, такие как docx, docx2txt, pytesseract, PIL и другие, для работы с документами, извлечения текста и переименования файлов.
