import sys,os
import re
from collections import Counter
#from PIL import Image
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

# Конвертировать файлы .docx и .doc в папке "pol" в .txt
for new_file_name in os.listdir(pol_folder_path):
    file_path = os.path.join(pol_folder_path, new_file_name)
    print(f"{new_file_name} находится в {pol_folder_path}!")
    try:
        if file_path.endswith('.docx'):
            txt_filename = os.path.splitext(new_file_name)[0] + '.txt'
            txt_file_path = os.path.join(pol_folder_path, txt_filename)
            docx_to_txt(file_path, txt_file_path)
            print(f"Файл {new_file_name} успешно сконвертирован в {txt_filename}!")
        elif file_path.endswith('.doc'):
            txt_filename = os.path.splitext(new_file_name)[0] + '.txt'
            txt_file_path = os.path.join(pol_folder_path, txt_filename)
            text = textract.process(file_path, encoding='utf-8')
            with open(txt_file_path, "w", encoding="utf-8") as txt_file:
                txt_file.write(text.decode("utf-8"))
            print(f"Файл {new_file_name} успешно сконвертирован в {txt_filename}!")
    except Exception as e:
        print(f"Ошибка при конвертации файла {new_file_name} в папку {pol_folder_path}: {str(e)}")
        
#Этот скрипт выполняет следующую последовательность действий:
#Определяет функцию find_uin_number, которая ищет номер УИН (уникальный идентификационный номер) в документе формата .docx.
#Задает пути к различным папкам и файлам, включая папку-источник файлов, папку назначения для файлов, не являющихся текстовыми документами (.docx или .doc), папку назначения для текстовых файлов (.txt) и папку pol внутри папки назначения для файлов.
#Циклически обрабатывает каждый файл в папке-источнике.
#Если файл имеет расширение .docx и не начинается с ~$, выполняются следующие действия:
#Ищется номер УИН в файле с помощью функции find_uin_number.
#Если номер УИН найден, извлекаются категория заключения, номер реестра и другая информация из имени файла.
#Поиск папок с соответствующим номером УИН в мастер-папке.
#Если найдены папки с номером УИН, проходит обработка каждой папки.
#Внутри каждой папки проходит обработка каждого файла.
#Если найден файл "Материалы по обоснованию в текстовой форме", он копируется в соответствующую папку назначения (в зависимости от расширения файла) и переименовывается.
#Если в новом имени файла содержится подстрока "пол", файл перемещается в папку pol. Если в новом имени файла содержится подстрока "отр", файл перемещается в папку otr.
#В случае возникновения ошибки при копировании файла, выводится соответствующее сообщение об ошибке.
#Если не найдены папки с номером УИН, выводится сообщение о том, что папки не найдены.
#Если номер УИН не найден в файле, выводится сообщение об этом.
#Если файл имеет неправильный формат, выводится сообщение о неправильном формате файла.
#Этот скрипт предназначен для обработки файлов с определенным форматом и структурой и выполняет операции по поиску и обработке этих файлов. Он использует различные библиотеки и модули, такие как docx, docx2txt, pytesseract, PIL и другие, для работы с документами, извлечения текста и переименования файлов.
