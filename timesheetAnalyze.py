import os
from excel import OpenExcel
import sys
import re
import copy

# Считываем список сотрудников из файла. Существует 2 типа написания имени и фамилии, мы раздилили их запятой.



# Каталог из которого будем брать таймшиты
# Home PC
# directory = 'C:\\Users\\Andrey\\Desktop\\Timesheets 01-31\\'
# filename = 'C:\\Users\\Andrey\\Desktop\\employeesList.txt'
# Work PC
filename = 'C:\\Users\\akovalyo\\Desktop\\employeesList.txt'
directory = 'D:\\Reports\\For Bruce\\2017\\Timesheets 01-31\\'
# directory = '\\\\KH-FSRV\\Public\\Timesheets\\2017\\02.28\\'


file = open(filename, 'r')
s = file.read()
employees = re.split(r'; ', s)
file.close()


# Получаем список таймшитов в переменную timesheets
timesheets = os.listdir(directory)

# ****************************** Удаляем временные файлы из списка файлов считанных из директории ***********************************
# создать временную копию списка таймшита
tmp_timesheets = copy.deepcopy(timesheets)

for i in tmp_timesheets:
    if i[0] == '~':
        # print (i)
        timesheets.remove(i)
del tmp_timesheets

#
# ****************************** Проверяем все ли создали таймшиты ***********************************
# Создаем копию списка с сотрудниками
tmp_employees = copy.deepcopy(employees)

# Удаляем из копии списка сотрудников людей, чьи таймшиты находятся в папке.
try:
    for i in timesheets:
        name_surname = re.search('(.*) \d', i)
        # employees.append(name_surname.group(1))
        for name in employees:
            # Так как имя сотрудника имеет 2 варианта записи разделенных запятой, проверяем первый и второй варианты отдельно. Если какой-либо из вариантов совпал - удаляем элементы из списка tmp_employees
            one_name = re.split(r', ', name)
            a = str(name_surname.group(1))
            if one_name[0] == name_surname.group(1) or one_name[1] == name_surname.group(1):
                tmp_employees.remove(name)
                break
except Exception:
    e = sys.exc_info()[1]
    # print(e.args[0])
    print('==============================Exception========================================')
    print(e)


# Выводим список людей, кто не создал таймшит
print('     Готово ' + str(len(timesheets)) + ' таймшит(а). \nДолжно быть ' + str(len(employees)) + ' таймшит(а).')
print('===============================================================================')
print('Отсутствует ' + str(len(tmp_employees)) + ' таймшит(а).')
for i in range(0, int(len(tmp_employees))):
    name = re.split(r', ', tmp_employees[i])
    print(str(i + 1) + ') ' + str(name[1]))

del tmp_employees


# ***************** Проверка праздничных дней, поля "Имя" и поля "Проект" ***********************************
# Внесение праздничных выходных дней пользователем
is_holiday = input("Есть ли выходные праздничные дни в этом месяце? (yes/no): ")
holiday = [] # массив содержащий номера ячеек праздничных выходных
if is_holiday == 'yes':
    sum_holiday_day = int(input("Сколько выходных дней в этот период?: "))
    # print(str(sum_holiday_day))
    while sum_holiday_day > 0:
        holiday_cell = input("Введите номер ячейки с праздником на латинице: \n")
        holiday.append(holiday_cell)
        sum_holiday_day -= 1
else:
    print("Нечего отдыхать! Труд сделал из обезьяны человека!")

print(holiday)

# ************************ Проверка  ************************************************
print('===============================================================================')
# wrong_timesheets = []
# for i in timesheets:
#     # print(i)
#     try:
#         # Считывваем имя сотрудника и название его первого проекта из таймшита
#         f = OpenExcel(directory + i)
#         username = f.read('C3').strip()
#         project_1 = f.read('B8')
#         # Считываем имя сотрудника из названия файла. Копируем строку до пробела и цифры.
#         name_surname = re.search('(.*) \d', i)
#         # Печатаем имя и фамилию сотрудников
#         # print(name_surname.group(1))
#         # Проверяем корректно ли заполнены поля Имя сотрудника и название проекта в таймшите
#         if username != name_surname.group(1) or project_1 == 'Project1':
#             wrong_timesheets.append(name_surname.group(1))
#
#         # Проверка, есть ли ошибка в файле в поле F4.
#         error_message=f.read('F4')
#         if len(error_message) > 0:
#             print('В таймшите %s ошибка! Она выглядит вот так - %s' % (username, error_message))
#             wrong_timesheets.append(name_surname.group(1))
#
#     except Exception:
#         e = sys.exc_info()[1]
#         # print(e.args[0])
#         print('==============================Exception========================================')
#         print(e)
#
# print('===============================================================================')
# # Печатаем списсок людей с некорректным excel файлом
# if len(wrong_timesheets) > 0:
#     print('\nОшибка! Имя сотрудника или название проекта указано неверно.\nПнуть следующих людей:')
#     for i in range(0, int(len(wrong_timesheets))):
#         print(str(i + 1) + ') ' + wrong_timesheets[i])
# print('===============================================================================')
#

for i in timesheets:
        # Считывваем имя сотрудника и название его первого проекта из таймшита
        f = OpenExcel(directory + i)    #         f = OpenExcel(directory + i)
        username = f.read('C3').strip()
        project_1 = f.read('B8')
        error_message = f.read('F4')  # Ячейка F4 содержит ошибку, в случае если сумма часов за каждый рабочий день не равна 8
        # Считываем имя сотрудника из названия файла. Копируем строку до пробела и цифры.
        name_surname = re.search('(.*) \d', i)

        # Проверка, есть ли ошибка в файле в поле F4.
        if len(error_message) > 0:
            print(f'В таймшите {name_surname.group(1)} ошибка! Она выглядит вот так - {error_message}')
            # wrong_timesheets.append(name_surname.group(1))

        # Проверяем корректно ли заполнены праздничные дни
        if len(holiday) > 0:
            for j in holiday:
                holiday_day = f.read(str(j).upper()) # Проверка праздничных дней
                if holiday_day != 8:
                    print(f'В таймшите {name_surname.group(1)} ошибка! Часы для строки "Holiday" незаполнены.')

        # Проверяем корректно ли заполнены поля Имя сотрудника и название проекта в таймшите
        if username != name_surname.group(1):
            print('В таймшите %s ошибка в имени пользователя !' % name_surname.group(1))

        if project_1 == 'Project1':
            print('В таймшите %s ошибка в имени проекта !' % name_surname.group(1))

        # elif holiday_day != 8:
        #     print("Выходной не отмечен.")
        #     # wrong_timesheets.append(name_surname.group(1))

print ('===============================================================================')
# # Печатаем списсок людей с некорректным excel файлом
# if len(wrong_timesheets) > 0:
#     print('\nОшибка! Имя сотрудника или название проекта указано неверно.\nПнуть следующих людей:')
#     for i in range(0, int(len(wrong_timesheets))):
#         print(str(i + 1) + ') ' + wrong_timesheets[i])
# print('===============================================================================')