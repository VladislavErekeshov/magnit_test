import sqlite3 as sql
import openpyxl
import xlsxwriter
import pdfminer.high_level
import sys
import fpdf


conn = sql.connect('database.db')
cur = conn.cursor()


def create_tables():
    cur.execute("""CREATE TABLE IF NOT EXISTS users(
        id INT PRIMARY KEY NOT NULL,
        second_name TEXT NOT NULL,
        first_name TEXT NOT NULL,
        patronymic TEXT,
        region_id INT NOT NULL,
        city_id INT NOT NULL,
        phone TEXT NOT NULL,
        email TEXT);
        """)

    cur.execute("""CREATE TABLE IF NOT EXISTS regions(
        id INT PRIMARY KEY NOT NULL,
        region_name TEXT NOT NULL);
        """)

    cur.execute("""CREATE TABLE IF NOT EXISTS cities(
        id INT PRIMARY KEY NOT NULL,
        region_id INT NOT NULL,
        city_name TEXT NOT NULL);
        """)

    regions = [(0, "Краснодарский край"), (1, "Ростовская область"), (2, "Ставропольский край")]
    cur.executemany("INSERT OR IGNORE INTO regions VALUES(?, ?);", regions)

    cities = [(0, 0, "Краснодар"), (1, 0, "Кропоткин"), (2, 0, "Славянск"), (3, 1, "Ростов"), (4, 1, "Шахты"), (5, 1, "Батайск"), (6, 2, "Ставрополь"), (7, 2, "Пятигорск"), (8, 2, "Кисловодск")]
    cur.executemany("INSERT OR IGNORE INTO cities VALUES(?, ?, ?);", cities)

    conn.commit()


def xlsx_import():
    wb = openpyxl.reader.excel.load_workbook(filename = "import.xlsx", data_only=True)
    wb.active = 0
    sheet = wb.active

    for i in range(2, 1048577):
        if sheet['A' + str(i)].value == None:
            break
        else:
            region_name = sheet['E' + str(i)].value
            cur.execute('SELECT id FROM regions WHERE region_name = ?', (region_name,))
            region = cur.fetchall()

            city_name = sheet['F' + str(i)].value
            cur.execute('SELECT id FROM cities WHERE city_name = ?', (city_name,))
            city = cur.fetchall()

            values = (sheet['A' + str(i)].value, sheet['B' + str(i)].value, sheet['C' + str(i)].value, sheet['D' + str(i)].value, region[0][0], city[0][0], sheet['G' + str(i)].value, sheet['H' + str(i)].value)
            cur.execute("INSERT OR IGNORE INTO users VALUES(?, ?, ?, ?, ?, ?, ?, ?);", values)
    
    conn.commit()


def users_len():
    cur.execute("SELECT * FROM users")
    rows = cur.fetchall()
    return len(rows)


def xlsx_export():
    len = users_len()
    workbook = xlsxwriter.Workbook('export.xlsx')
    worksheet = workbook.add_worksheet()
    
    worksheet.write(0, 0, "id")
    worksheet.write(0, 1, "Фамилия")
    worksheet.write(0, 2, "Имя")
    worksheet.write(0, 3, "Отчество")
    worksheet.write(0, 4, "Регион")
    worksheet.write(0, 5, "Город")
    worksheet.write(0, 6, "Контактный телефон")
    worksheet.write(0, 7, "E-mail")
    cur.execute("SELECT * FROM users")
    data = cur.fetchall()

    for i in range(0, len):
        d = data[i]
        region_id =  d[4]
        city_id = d[5]
        cur.execute('SELECT region_name FROM regions WHERE id = ?', (region_id,))
        region = cur.fetchall()
        worksheet.write(i+1, 4, region[0][0])
        cur.execute('SELECT city_name FROM cities WHERE id = ?', (city_id,))
        city = cur.fetchall()
        worksheet.write(i+1, 5, city[0][0])
        
        worksheet.write(i+1, 0, d[0])
        worksheet.write(i+1, 1, d[1])
        worksheet.write(i+1, 2, d[2])
        worksheet.write(i+1, 3, d[3])
        worksheet.write(i+1, 6, d[6])
        worksheet.write(i+1, 7, d[7])
           
    conn.commit()
    workbook.close()


def pdf_import():
    with open('import.pdf', 'rb') as file:
        text = pdfminer.high_level.extract_text(file, sys.stdout)
        res = tuple(map(str, text.split(' ')))

        region_name = res[4] + " " + res[5]
        cur.execute('SELECT id FROM regions WHERE region_name = ?', (region_name,))
        region = cur.fetchall()

        city_name = res[6]
        cur.execute('SELECT id FROM cities WHERE city_name = ?', (city_name,))
        city = cur.fetchall()

        values = (int(res[0]), res[1], res[2], res[3], region[0][0], city[0][0], res[7], res[8][1:])
        cur.execute("INSERT OR IGNORE INTO users VALUES(?, ?, ?, ?, ?, ?, ?, ?);", values)
        
        conn.commit()
        


def pdf_export():
    pdf = fpdf.FPDF(format='letter') #pdf format
    pdf.add_page() #create new page
    font_path = 'DejaVuSansCondensed.ttf'
    font_family = 'family'
    pdf.add_font(family=font_family, fname=font_path, uni=True)
    pdf.set_font(family=font_family, size=8)

    len = users_len()
    cur.execute("SELECT * FROM users")
    data = cur.fetchall()

    for i in range(0, len):
        d = data[i]
        region_id =  d[4]
        city_id = d[5]
        cur.execute('SELECT region_name FROM regions WHERE id = ?', (region_id,))
        region = cur.fetchall()
        
        cur.execute('SELECT city_name FROM cities WHERE id = ?', (city_id,))
        city = cur.fetchall()
        if d[3] == None:
            patronymic = ''
        else:
            patronymic = d[3]
        if d[7] == None:
            email = ''
        else:
            email = d[7]
        pdf.cell(200, 10, txt="{0} {1} {2} {3} {4} {5} {6} {7}".format(d[0], d[1], d[2], patronymic, region[0][0], city[0][0], d[6], email), ln=i+1, align="L")

    conn.commit()
    pdf.output("export.pdf")


create_tables()
#xlsx_import()
#xlsx_export()

pdf_import()
pdf_export()