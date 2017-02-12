# import pyowm
#
# owm = pyowm.owm('0616cee1399d6bcbbed8aaa6fa3b5fa7')  # You MUST provide a valid API key
#
# observation = owm.weather_at_place('London, uk')
# w = observation.get_weather()
# print(w)                      # <Weather - reference time=2013-12-18 09:20,
#                               # status=Clouds>


# from excel import OpenExcel
# f = OpenExcel("C:\\Users\\akovalyo\Desktop\\test.xlsx")
# # f.read() # read all
# #data = f.read('A') # read 'A' row
# #data = f.read(1) # f.read('1'), read '1' column
# data = f.read('B5') # read 'A5' position
# print (data)
#



from excel import OpenExcel

path = 'C:\\Users\\akovalyo\Desktop\\'
# file1 = 'test.xlsx'
# file2 = 'test2.xls'
f = OpenExcel(path+'test.xlsx')
data = f.read('A1')  # read 'A5' position

print(data)
# list = [file1, file2]
#
# for i in list:
#     f = OpenExcel('test.xlsx')
#     data = f.read('A1')  # read 'A5' position
#     if data != 'Username':
#         print(file1 + ' is ok.')
#     else:
#         print('ERROR ' + file1 + ' is WRONG!.')
#
print("============================THE END!============================")
