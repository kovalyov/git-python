import os
from excel import OpenExcel
import sys
import re
import copy

# Считываем список сотрудников из файла. Существует 2 типа написания имени и фамилии, мы раздилили их запятой.
filename = 'C:\\Users\\akovalyo\\Desktop\\employeesList.txt'
file = open(filename, 'r')
s = file.read()
employees = re.split(r'; ', s)
file.close()

# Каталог из которого будем брать таймшиты
# Home PC
# directory = 'C:\\Users\\Andrey\\Desktop\\Timesheets 01-31\\'
# Work PC
directory = 'D:\\Reports\\For Bruce\\2017\\Timesheets 01-31\\'
# directory = '\\\\KH-FSRV\\Public\\Timesheets\\2017\\02.15\\'

# Получаем список таймшитов в переменную timesheets
timesheets = os.listdir(directory)

# Создаем копию списка с сотрудниками
tmp_employees = copy.deepcopy(employees)

# Удаляем из копии списка сотрудников людей, чьи таймшиты находятся в папке.
for i in timesheets:
    name_surname = re.search('(.*) \d', i)
    # employees.append(name_surname.group(1))
    for name in employees:
        # Так как имя сотрудника имеет 2 варианта записи разделенных запятой, проверяем первый и второй варианты отдельно. Если какой-либо из вариантов совпал - удаляем элементы из списка tmp_employees
        one_name = re.split(r', ', name)
        if one_name[0] == name_surname.group(1) or one_name[1] == name_surname.group(1):
            tmp_employees.remove(name)

# Выводим список людей, кто не создал таймшит
print('     Готово ' + str(len(timesheets)) + ' таймшит(а). \nДолжно быть ' + str(len(employees)) + ' таймшит(а).')
print('===============================================================================')
print('Отсутствует ' + str(len(tmp_employees)) + ' таймшит(а).')
for i in range(0, int(len(tmp_employees))):
    name = re.split(r', ', tmp_employees[i])
    print(str(i + 1) + ') ' + str(name[1]))

# Проверка поля Имя и поля Проект
print('===============================================================================')
wrong_timesheets = []
for i in timesheets:
    # print(i)
    try:
        # Считывваем имя сотрудника и название его первого проекта из таймшита
        f = OpenExcel(directory + i)
        username = f.read('C3').strip()
        project_1 = f.read('B8')
        # Считываем имя сотрудника из названия файла. Копируем строку до пробела и цифры.
        name_surname = re.search('(.*) \d', i)
        # Печатаем имя и фамилию сотрудников
        # print(name_surname.group(1))
        # Проверяем корректно ли заполнены поля Имя сотрудника и название проекта в таймшите
        if username != name_surname.group(1) or project_1 == 'Project1':
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