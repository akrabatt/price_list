import xlwt

book = xlwt.Workbook('utf8')  # создаем книгу

sheet = book.add_sheet('ЛИСТ_ПРОБНЫЙ')  # создаем лист

sheet.write(0, 0, 'text')  # заполняем ячейку

book.save('filename.xls')  # сохраняем книгу