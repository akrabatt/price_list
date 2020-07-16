import xlrd  # импортируем библиотеку
import xlwt

book1 = xlrd.open_workbook('C:\goPY\./price_mpls.xlsx')  # указываем путь к файлу
book2 = xlrd.open_workbook('C:\goPY\./price_elc.xlsx')
sheet1 = book1.sheet_by_index(0)  # указываем страницу
sheet2 = book2.sheet_by_index(0)

book_for_write = xlwt.Workbook('utf8')  # создаём книгу
sheet_for_write = book_for_write.add_sheet('ОБОРУДОВАНИЕ')  # создаём лист в этой книге

row_number1 = sheet1.nrows  # количество строк на нашей странице
row_number2 = sheet2.nrows

product1 = []  # список оборудования
price1 = []  # список цен
articuls1 = []  # список артикулов
# индекс списка оборудования(product) соответствует индексу списка цене(price) и артикул(articuls)
product2 = []
price2 = []
articuls2 = []

''' 1 - Й ПРАЙС ЛИСТ mpls '''

if row_number1 > 0:  # проверка не пустой ли документ
    for row in range(0, row_number1):  # пробегаемся по всем строчкам
        product1.append(sheet1.cell_value(row, 4))  # оборудование

if row_number1 > 0:
    for row in range(0, row_number1):
        price1.append(sheet1.cell_value(row, 14))  # цена

if row_number1 > 0:
    for row in range(0, row_number1):
        articuls1.append(sheet1.cell_value(row, 0))  # артикул

'''2-Й ПРАЙС ЛИСТ ec'''

if row_number2 > 0:  # проверка не пустой ли документ
    for row in range(0, row_number2):  # пробегаемся по всем строчкам
        product2.append(sheet2.cell_value(row, 1))

if row_number2 > 0:
    for row in range(0, row_number2):
        price2.append(sheet2.cell_value(row, 2))

if row_number2 > 0:
    for row in range(0, row_number2):
        articuls2.append(sheet2.cell_value(row, 0))

del product1[:7], price1[:7], articuls1[:7]  # удвляем мусор из списков
del product2[:8], price2[:8], articuls2[:8]

#print('введите количестово оборудования: ')
#v = input()
#v = int(v)
aa = 0  # переменная для строк в листе
v = ''

#for i in range(1, v+1):
while v != str('stop'):
    print('введите ключевое слово: ')
    productlookfor = str(input())  # переменная для поиска по неполным данным
    resultList_mpls_p = []  # результат
    resultList_ec_p = []

    for i in range(len(product1)):  # лёгкое заполнение списка оборудования// mpls
        if productlookfor in product1[i]:
            resultList_mpls_p.append(i)  # заполнение индексами оборудования, список содержит только индексы !!!!!

    for i in range(len(product2)):  # // ec
        if productlookfor in product2[i]:
            resultList_ec_p.append(i)

    '''ВЫВОД ПОИСКА'''

    for i in range(len(resultList_mpls_p)):
        g = resultList_mpls_p[i]
        print('\n')
        print('позиция оборудования в мплс: ', g)
        print('товар: ', product1[g])
        print('цена: ', price1[g])
        print('артикул: ', articuls1[g])
        i += 1

    print('\n')
    print('-----РАЗДЕЛ-----')

    for i in range(len(resultList_ec_p)):
        g = resultList_ec_p[i]
        print('\n')
        print('позиция оборудования в эц: ', g)
        print('товар: ', product2[g])
        print('цена: ', price2[g])
        print('артикул: ', articuls2[g])
        i += 1

    '''ДАЁМ ВЫБОР ПОЛЬЗОВАТЕЛЮ'''

    print('\n')
    print('НАПИШИТЕ НОМЕР ПОЗИЦИИ ВЫБРАННОГО ОБОРУДОВАНИЯ мплс: ')
    mpls_index = input()
    mpls_index = int(mpls_index)
    print('КОЛИЧЕСТВО: ')
    mpls_number = input()
    mpls_number = int(mpls_number)

    a1 = product1[mpls_index]
    b1 = price1[mpls_index] * mpls_number
    c1 = articuls1[mpls_index]

    print('\n')
    print('НАПИШИТЕ НОМЕР ПОЗИЦИИ ВЫБРАННОГО ОБОРУДОВАНИЯ эц: ')
    ec_index = input()
    ec_index = int(ec_index)
    print('КОЛИЧЕСТВО:')
    ec_number = input()
    ec_number = int(ec_number)

    a2 = product2[ec_index]
    b2 = price2[ec_index] * ec_number
    c2 = articuls2[ec_index]

    sheet_for_write.write(aa, 0, a1)
    sheet_for_write.write(aa, 1, b1)
    sheet_for_write.write(aa, 2, c1)

    sheet_for_write.write(aa, 4, a2)
    sheet_for_write.write(aa, 5, b2)
    sheet_for_write.write(aa, 6, c2)
    aa += 1
    print('напишите stop если вы закончили, любую букву или слово если хтите продолжить: ')
    v = input()
else:
    print('ВВЕДИТЕ НАЗВАНИЕ ФАЙЛА: ')

file_name = input()

book_for_write.save('{0}.xls'.format(file_name))  # сохраняем книгу
print('файл сохранён')