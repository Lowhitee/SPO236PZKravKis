import pandas as pd
import math //Импортируем библиотеки

lkPzPuples = pd.read_excel('C:\Vassev Kiselev\pz12\Предмет 1 - Оценки.xlsx')
pos = pd.read_excel('C:\Vassev Kiselev\pz12\Предмет1_Посещаемость.xlsx')
predm = pd.read_excel('C:\Vassev Kiselev\pz12\Предмет1-шкала.xlsx')
//Импортируем эксель файлы в переменныт

arr = []
grades = []
//Создаем массивы для того чтобы в них записывать учеников и их баллы

def inter(a):
    a = a.replace(' ', '')
    a = a.split('/')
    return (int(a[0]) / int(a[1]))
    //Считаем баллы посещаемости

def interPzLk(a):
    if a != '-':
        return int(a)
    else:
        return 0
    //Считаем баллы пз и лк

for i in range(0, len(pos['Фамилия'])):
    sum = 0
    arr.append(pos['Фамилия'][i])
    b = (inter(pos['Баллы'][i]) * 15)
    b = math.ceil(b)
    //Вычисляем посещаемость

    pz1 = math.ceil((interPzLk(lkPzPuples['Задание:Пз 1(234Б) (Значение)'][i]) / 4))
    pz2 = math.ceil((interPzLk(lkPzPuples['Задание:Пз1(236) (Значение)'][i]) / 4))
    pz3 = math.ceil((interPzLk(lkPzPuples['Задание:Пз10(234) (Значение)'][i]) / 4))
    lk1 = math.ceil((interPzLk(lkPzPuples['Задание:Лк 1 (Значение)'][i]) / 10))
    lk2 = math.ceil((interPzLk(lkPzPuples['Задание:Лк 2 (Значение)'][i]) / 10))
    lk3 = math.ceil((interPzLk(lkPzPuples['Задание:Лк 3 (Значение)'][i]) / 10))
    lk4 = math.ceil((interPzLk(lkPzPuples['Задание:Лк 4 (Значение)'][i]) / 10))
    //Вычисляем все работы учеников

    sum += (b + pz1 + pz2 + pz3 + lk1 + lk2 + lk3 + lk4)
    //Суммируем всю их работу

    if sum >= 100:
        sum = 100

    grades.append(sum)
    //Добавляем под одним индексом фамилию и его баллы

for i in range(0, len(arr)):
    print(arr[i], 'балл -', grades[i])
    //Выводим построчно баллы и фамилию

predm.insert(5, "фамилия", arr)
predm.insert(5, "Баллы", grades)
//Создаем в таблице столбики с нашими новыми данными

predm.to_excel('C:\Vassev Kiselev\pz12\Финал.xlsx', index=False)
//Сохраняем старую таблицу как новую