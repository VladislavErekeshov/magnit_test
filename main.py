import sqlite3 as sql
import openpyxl
import xlsxwriter
import pdfminer.high_level
import sys
import fpdf

# создаём базу данных и подключаемся к ней
conn = sql.connect('database.db')
cur = conn.cursor()

# функция для создания таблиц в бд
def create_tables():
    # создаём таблицу users
    cur.execute("""CREATE TABLE IF NOT EXISTS users(
        id INTEGER NOT NULL PRIMARY KEY,
        second_name TEXT NOT NULL,
        first_name TEXT NOT NULL,
        patronymic TEXT,
        region_id INTEGER NOT NULL,
        city_id INTEGER NOT NULL,
        phone TEXT NOT NULL,
        email TEXT);
        """)

    # создаём таблицу regions
    cur.execute("""CREATE TABLE IF NOT EXISTS regions(
        id INTEGER PRIMARY KEY NOT NULL,
        region_name TEXT NOT NULL);
        """)

    # создаём таблицу cities
    cur.execute("""CREATE TABLE IF NOT EXISTS cities(
        id INTEGER PRIMARY KEY NOT NULL,
        region_id INTEGER NOT NULL,
        city_name TEXT NOT NULL);
        """)

    # заполняем первичные данные
    regions = [(0, "Краснодарский край"), (1, "Ростовская область"), (2, "Ставропольский край")]
    cur.executemany("INSERT OR IGNORE INTO regions VALUES(?, ?);", regions)

    cities = [(0, 0, "Краснодар"), (1, 0, "Кропоткин"), (2, 0, "Славянск"), (3, 1, "Ростов"), (4, 1, "Шахты"), (5, 1, "Батайск"), (6, 2, "Ставрополь"), (7, 2, "Пятигорск"), (8, 2, "Кисловодск")]
    cur.executemany("INSERT OR IGNORE INTO cities VALUES(?, ?, ?);", cities)

    conn.commit()


# функция, которая импортирует данные из Excel в таблицу users
def xlsx_import():
    wb = openpyxl.reader.excel.load_workbook(filename = "import.xlsx", data_only=True) # читаем Excel файл, из которово нам нужно импортировать данные
    # выбираем активный лист
    wb.active = 0
    sheet = wb.active

    # парсим xlsx документ и заполняем таблицу users
    for i in range(2, 1048577): # по версии Google в листе Excel содержится 1048576 строк
        if sheet['A' + str(i)].value == None: # завершаем цикл когда программа наткнется на пустую ячейку А
            break
        else:
            # находим id региона с помощью таблицы regions и сохраняем в переменную
            region_name = sheet['E' + str(i)].value # парсим название региона
            cur.execute('SELECT id FROM regions WHERE region_name = ?', (region_name,))
            region_id = cur.fetchall()

            # находим id города с помощью таблицы cities и сохраняем в переменную
            city_name = sheet['F' + str(i)].value  # парсим название города
            cur.execute('SELECT id FROM cities WHERE city_name = ?', (city_name,))
            city_id = cur.fetchall()

            # сохраняем спаршенные данные одной строки в кортеж и кидаем в таблицу users
            values = (sheet['A' + str(i)].value, sheet['B' + str(i)].value, sheet['C' + str(i)].value, sheet['D' + str(i)].value, region_id[0][0], city_id[0][0], sheet['G' + str(i)].value, sheet['H' + str(i)].value)
            cur.execute("INSERT OR IGNORE INTO users VALUES(?, ?, ?, ?, ?, ?, ?, ?);", values)
    
    conn.commit()


# функция для нахождения количества строк в таблице users
def users_len():
    cur.execute("SELECT * FROM users")
    rows = cur.fetchall()
    return len(rows)


# функция, которая экспортирует данные из таблицы users в Excel
def xlsx_export():
    # создаем Excel файл
    workbook = xlsxwriter.Workbook('export.xlsx')
    worksheet = workbook.add_worksheet()
    
    # заполняем первую строку с наименованиями
    worksheet.write(0, 0, "id")
    worksheet.write(0, 1, "Фамилия")
    worksheet.write(0, 2, "Имя")
    worksheet.write(0, 3, "Отчество")
    worksheet.write(0, 4, "Регион")
    worksheet.write(0, 5, "Город")
    worksheet.write(0, 6, "Контактный телефон")
    worksheet.write(0, 7, "E-mail")

    # берем данные из таблицы users
    cur.execute("SELECT * FROM users")
    data = cur.fetchall()

    # заполняем Excel файл
    len = users_len()
    for i in range(0, len):
        d = data[i]  # данные i-ой строки таблицы users

        # находим название региона из таблицы regions по его id и записываем в Excel файл
        region_id =  d[4]
        cur.execute('SELECT region_name FROM regions WHERE id = ?', (region_id,))
        region_name = cur.fetchall()
        worksheet.write(i+1, 4, region_name[0][0])

        # находим название города из таблицы cities по его id и записываем в Excel файл
        city_id = d[5]
        cur.execute('SELECT city_name FROM cities WHERE id = ?', (city_id,))
        city_name = cur.fetchall()
        worksheet.write(i+1, 5, city_name[0][0])
        
        # записываем остальные данные
        worksheet.write(i+1, 0, d[0])
        worksheet.write(i+1, 1, d[1])
        worksheet.write(i+1, 2, d[2])
        worksheet.write(i+1, 3, d[3])
        worksheet.write(i+1, 6, d[6])
        worksheet.write(i+1, 7, d[7])
           
    workbook.close()


# функция, которая импортирует данные из pdf в таблицу users
def pdf_import():
    with open('import.pdf', 'rb') as file: # открываем pdf файл
        # парсим текст и сохраняем в кортеж
        text = pdfminer.high_level.extract_text(file, sys.stdout)
        res = tuple(map(str, text.split(' ')))

        # находим id региона с помощью таблицы regions и сохраняем в переменную
        region_name = res[4] + " " + res[5]
        cur.execute('SELECT id FROM regions WHERE region_name = ?', (region_name,))
        region_id = cur.fetchall()

        # находим id города с помощью таблицы cities и сохраняем в переменную
        city_name = res[6]
        cur.execute('SELECT id FROM cities WHERE city_name = ?', (city_name,))
        city_id = cur.fetchall()

        # заполняем данными таблицу users
        values = (int(res[0]), res[1], res[2], res[3], region_id[0][0], city_id[0][0], res[7], res[8][1:])
        cur.execute("INSERT OR IGNORE INTO users VALUES(?, ?, ?, ?, ?, ?, ?, ?);", values)
        
        conn.commit()
        

# функция, которая экспортирует данные из таблицы users в pdf файл
def pdf_export():
    # создаем pdf файл
    pdf = fpdf.FPDF(format='letter')
    pdf.add_page() 

    # выбираем и настраиваем unicode шрифт
    font_path = 'DejaVuSansCondensed.ttf'
    font_family = 'family'
    pdf.add_font(family=font_family, fname=font_path, uni=True)
    pdf.set_font(family=font_family, size=8)

    # берем данные из таблицы users
    cur.execute("SELECT * FROM users")
    data = cur.fetchall()

    # заполняем pdf файл
    len = users_len()
    for i in range(0, len):
        d = data[i] # данные i-ой строки таблицы users

        # находим название региона из таблицы regions по его id и сохраняем в переменную
        region_id =  d[4]
        cur.execute('SELECT region_name FROM regions WHERE id = ?', (region_id,))
        region_name = cur.fetchall()
        
        # находим название города из таблицы cities по его id и сохраняем в переменную
        city_id = d[5]
        cur.execute('SELECT city_name FROM cities WHERE id = ?', (city_id,))
        city_name = cur.fetchall()

        # прописываем варианты переменных с пустым отчествои или E-mail
        if d[3] == None:
            patronymic = ''
        else:
            patronymic = d[3]
        if d[7] == None:
            email = ''
        else:
            email = d[7]

        # заполняем pdf файл
        pdf.cell(200, 10, txt="{0} {1} {2} {3} {4} {5} {6} {7}".format(d[0], d[1], d[2], patronymic, region_name[0][0], city_name[0][0], d[6], email), ln=i+1, align="L")

    pdf.output("export.pdf")


create_tables()
xlsx_import()
xlsx_export()

#pdf_import()
pdf_export()