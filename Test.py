
# def print_spam():
#     print('spam \n' * 3)
# print_spam()
# def multiply(number):
#     print(list(range(number)))
# multiply(5)
###########################################
# def max_2(x, y):
#     '''Инструкция к применению''' # вызывается командой print(max_2.__doc__)
#     if x > y:
#         return x
#     else:
#         return y
#     print('THE END')
#
# print(max_2.__doc__)
# x = float(input('Item #1: '))
# y = float(input('Item #2: '))
# print(max_2(x, y))
###########################################
# r - Read
# w - Write
# a - Add to file
# b - Binary mode
###########################################
## Чтение файла
# filename = input('Укажите файл: ')
# file = open(filename, 'r')
# print(file.read())
# file.close()
###########################################
# # Запись в файл
# filename = input('Введите желаемое имя файла: ')
# text = input('insert text: ')
# file = open(filename, 'w')
# file.write(text)
# # print(file.read())
# file.close()
###########################################
# # Добавление в файл
# file = open('test.txt', 'a')
# file.write(' TEST ')
# file.close()

###########################################
# #Программа для копирования файлов
# filename1 = input('Backup of file: ')
# filename2= 'backup_'+filename1
#
# file1 = open(filename1,'r')
# file2 = open(filename2,'w')
#
# file2.write(file1.read())
#
# file1.close()
# file2.close()
#
# print('DONE')

###########################################
# # Чтение файла построчно
# with open('employees.txt','r') as f:
#     print(f.read())