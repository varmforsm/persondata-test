from faker import Faker
import random
import sys
import openpyxl
import urllib.request
import sqlite3
from sqlite3 import Error
import os.path
import getopt

create_firstname_table = """
CREATE TABLE IF NOT EXISTS firstname (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  firstname TEXT NOT NULL,
  holders INTEGER
);
"""

create_lastname_table = """
CREATE TABLE IF NOT EXISTS lastname (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  lastname TEXT NOT NULL,
  holders INTEGER
);
"""

def create_connection(path):
    connection = None

    try:
        connection = sqlite3.connect(path)
        print("Connection to SQLite DB successful")
    except Error as e:
        print(f"The error '{e}' occurred")
    return connection

def create_tables(connection, query):
    cursor = connection.cursor()
    try:
        cursor.execute(query)
        connection.commit()
    except Error as e:
        print(f"The error '{e}' occurred")

def retrieve_data(connection, query):
    cursor = connection.cursor()
    cursor.execute(query)    
    records = cursor.fetchall()
    return records

def setupDatabase(conn):
    create_tables(conn, create_firstname_table)
    create_tables(conn, create_lastname_table)

def createPersonnummer(count, max_age):
    fake = Faker()
    data = []

    for i in range(count):
        outdate = fake.date_between(start_date='today', end_date=max_age).strftime('%Y%m%d')
        checknum = random.randint(1111, 9999)
        data.append("{}-{}".format(outdate, checknum))
    return data

def downloadTestdata(url): 
    file_name="namndata.xlsx"
    if (os.path.isfile(file_name) == False):
        file_name, headers = urllib.request.urlretrieve(url, filename=file_name)
    wrkbk = openpyxl.load_workbook(file_name) 
    enamn = wrkbk["Efternamn"]
    fnamnk = wrkbk["Förnamn kvinnor"]
    fnamnm = wrkbk["Förnamn män"]
    return {"enamn": enamn, "fnamnk": fnamnk, "fnamnm": fnamnm }

def getRandomEntries(cursor, tablename, count):
    select_random_entry = """SELECT * FROM """ + tablename + """ ORDER BY RANDOM() LIMIT """ + str(count)
    cursor.execute(select_random_entry)
    record = cursor.fetchall()
    return record

def insert_names(conn, tdata, sheet, tablename):
    data = []
    for cells in tdata[sheet].iter_rows(min_row=6, min_col=1, max_col=2):
        data.append((cells[0].value, cells[1].value))
    conn.executemany("""insert into """ + tablename + """(""" + tablename + """, holders) values (?,?)""", data)

def print_help(url, count, age):
    print ('persondata-generator.py -w -m -u ' + url + ' -n ' + str(count) + ' -a ' + age)
    print ('-a : max age of testpersons - default 90 years.')
    print ('-w : women.')
    print ('-m : men.')
    print ('-u : testdata URL - default displayed above.')
    print ('-n : number of testpersons - default 100.')
    print ('If you only want one sex, choose "-w" _or_ "-m", otherwise both sexes will be loaded.')
    sys.exit()

def create_testdata(count, age, conn):
    cursor = conn.cursor()

    names = []
    firstname = getRandomEntries(cursor, "firstname", count)
    lastname = getRandomEntries(cursor, "lastname", count)
    personnummer = createPersonnummer(count, age)
    for i in range(count):
        names.append("{}, {} : {}".format(lastname[i][1], firstname[i][1], personnummer[i]))

    print(*names, sep = "\n")

def main(argv):
    url = "https://www.scb.se/globalassets/namn-med-minst-tva-barare-31-december-2022_20230228.xlsx"
    count = 100
    age = '-90y'
    women = False
    men = False

    opts, args = getopt.getopt(argv,"hwmu:n:a:",["help", "women", "men", "url=", "number=", "age="])
    for opt, arg in opts:
        if opt in ("-h", "--help"):
            print_help(url, count, age)
        elif opt in ("-u", "--url"):
            url = arg
        elif opt in ("-n", "--number"):
            count = int(arg)
        elif opt in ("-a", "--age"):
            age = arg
        elif opt in ("-w", "--women"):
            women=True
        elif opt in ("-m", "--men"):
            men=True

    if (women == True and men == True ) :
        print_help

    conn = sqlite3.connect(":memory:")
    setupDatabase(conn)

    tdata = downloadTestdata(url)
  
    insert_names(conn, tdata, "enamn", "lastname")
    if (men == False): 
        insert_names(conn, tdata, "fnamnk", "firstname")
    if (women == False):
        insert_names(conn, tdata, "fnamnm", "firstname")

    create_testdata(count, age, conn)

if __name__ == "__main__":
   main(sys.argv[1:])       