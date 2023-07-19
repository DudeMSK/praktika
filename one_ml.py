import sys,os
import re
from collections import Counter
from collections import defaultdict
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
from sklearn.ensemble import RandomForestClassifier
from sklearn.ensemble import GradientBoostingClassifier
from sklearn.svm import SVC
from sklearn.svm import LinearSVC
import torch
from torch import nn, optim
import torch.nn as nn
import torch.optim as optim
from torch.nn.utils import weight_norm
from sklearn.inspection import permutation_importance
#import eli5
#from eli5.sklearn import PermutationImportance



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
        labels.append(0)  # Класс "пол"

    for file in microbiology_files:
        text = read_text_from_txt(file)
        preprocessed_text = preprocess_text(text)
        dataset.append(preprocessed_text)
        labels.append(1)  # Класс "отр"

    return dataset, labels

# Каталог с положительными и отрицательными файлами
greetings_folder = r".\ml\pol"
microbiology_folder = r".\ml\otr"

# Получить список файлов .txt из каталогов
greetings_files = [os.path.join(greetings_folder, f) for f in os.listdir(greetings_folder) if f.endswith(".txt")]
microbiology_files = [os.path.join(microbiology_folder, f) for f in os.listdir(microbiology_folder) if f.endswith(".txt")]

# Создание набора данных для обучения
dataset, labels = create_dataset(greetings_files, microbiology_files)

# Препроцессинг данных
preprocessed_dataset = [preprocess_text(text) for text in dataset]

# Преобразование текстовых данных в матрицу счетчиков признаков
vectorizer = CountVectorizer()
X = vectorizer.fit_transform(preprocessed_dataset)

# Разделение набора данных на обучающий и тестовый наборы
X_train, X_test, y_train, y_test = train_test_split(X, labels, test_size=0.2, random_state=42)

X_train = torch.tensor(X_train.toarray(), dtype=torch.float32)
X_test = torch.tensor(X_test.toarray(), dtype=torch.float32)

device = torch.device("cuda" if torch.cuda.is_available() else "cpu")
y_train = torch.tensor(y_train, dtype=torch.long).to(device)
y_test = torch.tensor(y_test, dtype=torch.long).to(device)

# Определение архитектуры нейронной сети
class NeuralNet(nn.Module):
    def __init__(self, input_size, hidden_size, num_classes):
        super(NeuralNet, self).__init__()
        self.fc1 = nn.Linear(input_size, hidden_size)
        self.relu = nn.ReLU()
        self.fc2 = nn.Linear(hidden_size, num_classes)

    def forward(self, x):
        out = self.fc1(x)
        out = self.relu(out)
        out = self.fc2(out)
        return out

# Параметры модели
input_size = X_train.shape[1]
hidden_size = 100
num_classes = 2
learning_rate = 0.001
num_epochs = 10

# Создание и обучение модели
model = NeuralNet(input_size, hidden_size, num_classes)
criterion = nn.CrossEntropyLoss()
optimizer = torch.optim.Adam(model.parameters(), lr=learning_rate)

for epoch in range(num_epochs):
    optimizer.zero_grad()
    outputs = model(X_train)
    loss = criterion(outputs, y_train)
    loss.backward()
    optimizer.step()

# Оценка производительности модели
with torch.no_grad():
    outputs = model(X_test)
    _, predicted = torch.max(outputs.data, 1)
    accuracy = accuracy_score(y_test.numpy(), predicted.numpy())
    print("Accuracy:", accuracy)

# Классификация новых файлов
classifications_folder = ".\\files"
new_files = [os.path.join(classifications_folder, f) for f in os.listdir(classifications_folder) if f.endswith(".txt")]

# Создаем словарь для хранения причин классификации
reasons = {}

# Функция для получения вклада каждого признака в классификацию
def get_feature_contributions(features):
    feature_contributions = {
        "пол": defaultdict(float),
        "отр": defaultdict(float)
    }

    # Получение значений весов из модели
    weights = model.fc1.weight.detach().cpu().numpy()

    for feature_idx, feature_name in enumerate(feature_names):
        for class_idx, class_name in enumerate(["пол", "отр"]):
            feature_contributions[class_name][feature_name] += weights[class_idx][feature_idx]

    return feature_contributions

# Проходим через новые файлы и классифицируем их
for file_path in new_files:
    text = read_text_from_txt(file_path)
    preprocessed_text = preprocess_text(text)
    features = vectorizer.transform([preprocessed_text])
    features = torch.tensor(features.toarray(), dtype=torch.float32)
    
    with torch.no_grad():
        output = model(features)
        _, predicted = torch.max(output.data, 1)
    
    category = "пол" if predicted.item() == 0 else "отр"
    
    feature_names = vectorizer.get_feature_names()
    feature_contributions = get_feature_contributions(features)
    
    reasons[file_path] = {
        "Категория": category,
        "Причины": feature_contributions
    }

# Вывод причин классификации для каждого файла
for file_path, classification in reasons.items():
    print()
    print("Файл:", file_path)
    print("Категория:", classification["Категория"])
    print("Причины:")
    for class_name, contributions in classification["Причины"].items():
        print(f"Класс: {class_name}")
        sorted_contributions = sorted(contributions.items(), key=lambda x: abs(x[1]), reverse=True)
        top_reason = sorted_contributions[0][0]  # Получаем первую причину с наибольшим вкладом
        print(f"Причина: {top_reason}")
    print()
    print()

# Создаем словарь для хранения изменений значений признаков
feature_changes = defaultdict(list)

# Регистрируем хук, чтобы получить значения признаков перед прохождением через слои модели
def get_feature_values(module, input, output):
    for i, inp in enumerate(input[0]):
        feature_changes[i].extend(inp.detach().cpu().numpy())

# Регистрируем хук для всех слоев модели
for module in model.modules():
    if isinstance(module, nn.Linear):
        module.register_forward_hook(get_feature_values)

# Проходим через тестовый набор данных и сохраняем изменения значений признаков
with torch.no_grad():
    model.eval()
    model(X_test)

# Вычисляем важность каждого признака на основе изменений значений
feature_importance = np.std([np.array(values) for values in feature_changes.values()], axis=1)

# Вывод значений важности признаков, отсортированных по убыванию
sorted_indices = np.argsort(feature_importance)[::-1]
for idx in sorted_indices:
    print(f"Признак: {feature_names[idx]}, Важность: {feature_importance[idx]}")


#Код выполняет следующие действия:
#Определяет функцию preprocess_text, которая выполняет предобработку текста: приводит его к нижнему регистру и удаляет пунктуацию.
#Определяет функцию read_text_from_txt, которая читает текст из файла формата .txt.
#Определяет функцию convert_docx_to_txt, которая преобразует файл формата .docx в текстовый формат.
#Определяет функцию create_dataset, которая создает набор данных для обучения модели. Она считывает текст из файлов и применяет предобработку, затем формирует датасет и метки классов для положительных и отрицательных файлов.
#Задает пути каталогов с положительными и отрицательными файлами.
#Получает список файлов .txt из каталогов положительных и отрицательных файлов.
#Создает набор данных для обучения вызовом функции create_dataset на основе списков файлов из пункта 6.
#роизводит предобработку данных над текстом из набора данных.
#Преобразует текстовые данные в матрицу счетчиков признаков с помощью CountVectorizer.
#Разделяет набор данных на обучающий и тестовый наборы.
#Преобразует данные в тензоры и отправляет на устройство (GPU или CPU).
#Определяет архитектуру нейронной сети с одним скрытым слоем и двумя линейными слоями.
#Определяет параметры модели, такие как размер входа, размер скрытого слоя, количество классов, скорость обучения и количество эпох.
#оздает экземпляр модели и определяет функцию потерь и оптимизатор.
#Обучает модель на обучающем наборе данных с использованием цикла обучения и обновления весов.
#Оценивает производительность модели на тестовом наборе данных с помощью точности.
#Задает путь каталога с новыми файлами для классификации.
#Получает список новых файлов .txt из каталога.
#Создает словарь reasons для хранения причин классификации новых файлов.
#Определяет функцию get_feature_contributions, которая вычисляет вклад каждого признака в классификацию и сохраняет его в словаре feature_contributions.
#Проходит через новые файлы, выполняет их классификацию, сохраняет категорию и вычисляет вклад каждого признака с помощью функции get_feature_contributions.
#Выводит причины классификации для каждого файла, отображая категорию и причину с наибольшим вкладом.
#Создает словарь feature_changes для хранения изменений значений признаков.
#Регистрирует хук для получения значений признаков перед прохождением через слои модели.
#Проходит через тестовый набор данных с помощью модели, сохраняет изменения значений признаков.
#Вычисляет важность каждого признака на основе изменений значений.
#Выводит значения важности признаков, отсортированные по убыванию.
