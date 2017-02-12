import os
from excel import OpenExcel
import sys
import re
import copy

# Каталог из которого будем брать файлы
# Home PC
directory = 'C:\\Users\\Andrey\\Desktop\\Timesheets 01-31\\'
# Work PC
# directory = 'D:\\Reports\\For Bruce\\2017\\Timesheets 01-31\\'
# directory = '\\\\KH-FSRV\\Public\\Timesheets\\2017\\01.31\\'

# Получаем список файлов в переменную timesheets
timesheets = os.listdir(directory)
employees = ['Aleksandr Sydorenko', 'Aleksey Chernyshov', 'Aleksey Tsyhanenko', 'Alesya Samylova', 'Alexandr Terentiev',
             'Alexandr Yaresko', 'Alexey Chuiko', 'Alexey Tronchuk', 'Andrey Belyaev', 'Andrey Sheplyakov',
             'Andrey Tovstik', 'Andrey Yakovenko', 'Andrey Kovalyov', 'Andrii Lashchev', 'Anton Doronin',
             'Artem Bondarenko', 'Bohdan Talalaiev', 'Boris Serdyuk', 'Danil Lutsenko', 'Dmitriy Khomenko',
             'Dmytro Matiushkin', 'Egor Raketskyy', 'Elya Sikerin', 'Eugene Sheplyakov', 'Evgenii Ilchenko',
             'Igor Makarenko', 'Ilya Volchanetskiy', 'Iurii Lukianov', 'Kirill Rukkas', 'Kirill Yarosh',
             'Konstantin Leyba', 'Konstantyn Glushko', 'Maksym Trykoz', 'Maksym Yepaneshnikov', 'Olha Omelianenko',
             'Pavel Gaydidey', 'Pavel Shulik', 'Polina Selianinova', 'Roman Bubyr', 'Roman Solianik',
             'Ruslan Mordvinov', 'Sava Shevchenko', 'Sergey Cherepnin', 'Sergey Kovalenko', 'Sergey Malovitsa',
             'Sergey Nikolayenko', 'Sergey Pasichnyi', 'Sergey Savchenko', 'Sergii Boldovskyi', 'Sergii Paveliev',
             'Svyatoslav Kolimbrovsky', 'Taras Dolya', 'TEMPLATE', 'Valentin Lukyanets', 'Valerii Zhelezniakov',
             'Viktor Bondarchuk', 'Viktor Iliukha', 'Viktor Khabalevskyi', 'Viktor Taemnickyy', 'Vladimir Agafonov',
             'Vladyslav Zahrebeniuk']

# Создаем копию списка с сотрудниками
tmp_employee = copy.deepcopy(employees)
# Удаляем из копии списка сотрудников людей, чьи таймшиты находятся в папке.
for i in timesheets:
    name_surname = re.search('(.*) \d', i)
    # employees.append(name_surname.group(1))
    for name in employees:
        if name == name_surname.group(1):
            tmp_employee.remove(name)

# Выводим список людей, кто не создал таймшит
print('     Готово ' + str(len(timesheets)) + ' таймшит(а). \nДолжно быть ' + str(len(employees)) + ' таймшит(а).')
print('===============================================================================')
print('Отсутствует ' + str(len(tmp_employee)) + ' таймшит(а).')
for i in range(0, int(len(tmp_employee))):
    print(str(i + 1) + ') ' + tmp_employee[i])

# Проверка поля Имя и поля Проект
print('===============================================================================')
wrong_timesheets = []
for i in timesheets:
    # print(i)
    try:
        f = OpenExcel(directory + i)
        username = f.read('C3')
        project_1 = f.read('B8')
        # Извлекаем имена сотрудников из названия файла. Копируем строку до пробела и цифры.
        name_surname = re.search('(.*) \d', i)
        # Печатаем имя и фамилию сотрудников
        # print(name_surname.group(1))
        # Проверяем корректно ли заполнены поля Имя сотрудника и название проекта в excel файле
        if username == 'Username' or project_1 == 'Project1':
            wrong_timesheets.append(name_surname.group(1))

    except Exception:
        e = sys.exc_info()[1]
        print(e.args[0])
        print(i)

# Печатаем списсок людей с некорректным excel файлом
if len(wrong_timesheets) > 0:
    print('\nОшибка! Имя сотрудника или название проекта указано неверно.\nПнуть следующих людей:')
    for i in range(0, int(len(wrong_timesheets))):
        print(str(i + 1) + ') ' + wrong_timesheets[i])
print('===============================================================================')