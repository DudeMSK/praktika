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
import numpy as np
import sklearn
from sklearn.model_selection import train_test_split
from sklearn.naive_bayes import MultinomialNB
from sklearn.metrics import accuracy_score
from sklearn.feature_extraction.text import CountVectorizer
from sklearn.svm import SVC

def find_uin_number(file_path):
    doc = Document(file_path)

    for para in doc.paragraphs:
        text = para.text
        match = re.search(r"УИН\s*№\s*(\d+)", text, re.UNICODE)
        if match:
            uin_number = match.group(1)
            return uin_number

    return None

#master_folder_path = r".\ТП_2021"
#source_folder_path = r".\ТП_2021"
#master_folder_path = r".\ТП_2022\ТП"
#source_folder_path = r".\ТП_2022\ТП"
master_folder_path = r".\ТП_2023\ТП"
source_folder_path = r".\ТП_2023\ТП"

#dest_notTXT = r".\ml\2021\notTXT"
#dest_TXT = r".\ml\2021\TXT"
#dest_notTXT = r".\ml\2022\notTXT"
#dest_TXT = r".\ml\2022\TXT"
dest_notTXT = r".\ml\2023\notTXT"
dest_TXT = r".\ml\2023\TXT"

pol_folder_path = os.path.join(dest_TXT, "pol")
otr_folder_path = os.path.join(dest_TXT, "otr")

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

# Конвертировать файлы .docx и .doc в папке "pol" в .txt
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

# Конвертировать файлы .docx и .doc в папке "otr" в .txt
for new_file_name in os.listdir(otr_folder_path):
    file_path = os.path.join(otr_folder_path, new_file_name)
    print(f"{new_file_name} находится в {otr_folder_path}!")
    try:
        if file_path.endswith('.doc'):
            # Конвертировать .doc в .docx
            docx_filename = os.path.splitext(new_file_name)[0] + '.docx'
            docx_file_path = os.path.join(otr_folder_path, docx_filename)
            convert_doc_to_docx(file_path, docx_file_path)
            print(f"Файл {new_file_name} успешно сконвертирован в {docx_filename}!")
    except Exception as e:
        print(f"Ошибка при конвертации файла {new_file_name} в папку {docx_filename}: {str(e)}")

for new_file_name in os.listdir(otr_folder_path):
    file_path = os.path.join(otr_folder_path, new_file_name)
    print(f"{new_file_name} находится в {otr_folder_path}!")
    try:
        if file_path.endswith('.docx'):
            txt_filename = os.path.splitext(new_file_name)[0] + '.txt'
            txt_file_path = os.path.join(otr_folder_path, txt_filename)
            doc = Document(file_path)
            text = "\n".join([paragraph.text for paragraph in doc.paragraphs])
            with open(txt_file_path, "w", encoding="utf-8") as txt_file:
                txt_file.write(text)
            print(f"Файл {new_file_name} успешно сконвертирован в {txt_filename}!")
    except Exception as e:
        print(f"Ошибка при конвертации файла {new_file_name} в папку {otr_folder_path}: {str(e)}")

############################################################################################################
############################################################################################################
############################################################################################################



# Функция для предобработки текста
def preprocess_text(text):
    text = text.lower()  # Приводим к нижнему регистру
    text = re.sub(r"[^\w\s]", "", text)  # Удаляем пунктуацию
    return text

# Функция для чтения текста из файла формата .txt
def read_text_from_txt(txt_file):
    with open(txt_file, 'r', encoding='utf-8') as file:
        text = file.read()
    return text

# Функция для преобразования файла формата .docx в .txt
def convert_docx_to_txt(docx_file):
    doc = Document(docx_file)
    paragraphs = [paragraph.text for paragraph in doc.paragraphs]
    text = "\n".join(paragraphs)
    return text

# Создание набора данных для обучения модели
def create_dataset(greetings_files, microbiology_files):
    dataset = []
    labels = []

    for file in greetings_files:
        text = read_text_from_txt(file)
        preprocessed_text = preprocess_text(text)
        dataset.append(preprocessed_text)
        labels.append("пол")

    for file in microbiology_files:
        text = read_text_from_txt(file)
        preprocessed_text = preprocess_text(text)
        dataset.append(preprocessed_text)
        labels.append("отр")

    return dataset, labels

# Каталог с положительными и отрицательными файлами
greetings_folder = r"C:\Users\anaconda\Desktop\praktika\ml\pol"
microbiology_folder = r"C:\Users\anaconda\Desktop\praktika\ml\otr"

# Получить список файлов .txt из каталогов
greetings_files = [os.path.join(greetings_folder, f) for f in os.listdir(greetings_folder) if f.endswith(".txt")]
microbiology_files = [os.path.join(microbiology_folder, f) for f in os.listdir(microbiology_folder) if f.endswith(".txt")]

# Создание набора данных для обучения
dataset, labels = create_dataset(greetings_files, microbiology_files)

# Преобразование текстовых данных в матрицу счетчиков признаков
vectorizer = CountVectorizer()
X = vectorizer.fit_transform(dataset)

# Разделение набора данных на обучающий и тестовый наборы
X_train, X_test, y_train, y_test = train_test_split(X, labels, test_size=0.2, random_state=42)

# Обучение модели
model = MultinomialNB()
model.fit(X_train, y_train)

# Оценка производительности модели
y_pred = model.predict(X_test)
accuracy = accuracy_score(y_test, y_pred)
print()
print("Accuracy:", accuracy)

# Классификация новых файлов
classifications_folder = ".\\files"
new_files = [os.path.join(classifications_folder, f) for f in os.listdir(classifications_folder) if f.endswith(".txt")]
for file_path in new_files:
    text = read_text_from_txt(file_path)
    preprocessed_text = preprocess_text(text)
    features = vectorizer.transform([preprocessed_text])
    category = model.predict(features)[0]
    print("Файл:", file_path)
    print("Категория:", category)
    print()

##############################################################################################
        
# 1. Определение функции find_uin_number, которая ищет номер УИН (уникального идентификационного номера) в переданном документе. Для этого функция открывает документ с помощью библиотеки docx, проходит по параграфам документа и ищет совпадения с помощью регулярного выражения.
# 2. Задание путей к папкам и файлам, которые будут использоваться в процессе. Например, master_folder_path - это путь к мастер-папке, source_folder_path - путь к папке, из которой будут браться файлы для обработки, и т.д.
# 3. Цикл for для перебора файлов в папке source_folder_path. Для каждого файла выполняются следующие шаги:
# 3.1 Проверка, что файл имеет расширение .docx и не является временным файлом (~$).
# 3.2 Поиск номера УИН в файле с помощью функции find_uin_number.
# 3.3 Если номер УИН найден, происходит извлечение информации о категории, номере реестра и т.д. из имени файла.
# 3.4 Поиск папок с соответствующими номерами УИН в мастер-папке.
# 3.5 Перебор файлов в найденных папках для выполнения дополнительных действий (например, копирования, переименования, перемещения файлов).
# 4. Определение функции convert_doc_to_docx, которая использует библиотеку win32com.client для конвертации документов .doc в .docx. Для этого функция открывает документ .doc с помощью приложения Word, сохраняет его в формате .docx и закрывает документ и приложение Word.
# 5. Цикл for для конвертации файлов .doc в .docx. Для каждого файла в папке pol_folder_path выполняются следующие шаги:
# 5.1 Проверка, что файл имеет расширение .doc.
# 5.2 Конвертация файла .doc в .docx с использованием функции convert_doc_to_docx.
# 6. Цикл for для конвертации файлов .docx в .txt. Для каждого файла в папке pol_folder_path выполняются следующие шаги:
# 6.1 Проверка, что файл имеет расширение .docx.
# 6.2 Открытие файла .docx с помощью библиотеки docx.
# 6.3 Извлечение текста из файла и сохранение его в формате .txt.

## В результате выполнения кода происходит обработка файлов в указанных папках.
## Он ищет номер УИН в документах, находит соответствующие папки в мастер-папке, копирует, переименовывает и перемещает файлы, а также конвертирует документы .doc в .docx и .docx в .txt.
