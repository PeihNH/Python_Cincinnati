import string
import random
import time
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import sqlite3
from sqlite3 import Error
from datetime import datetime
import xlrd
from xlsxwriter.workbook import Workbook


LETTERS = string.ascii_letters
NUMBERS = string.digits
DB_FILE = r"sqlite.db"
URL = "https://apply.cincinnatistate.edu/"
MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
DAYS = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28]
YEARS = [1997, 1998, 1999, 2000, 2001]
CITY = "Reynoldsburg"

STATE = "Ohio"
ZIPCODE = "43068"


def pause_chut():
    time.sleep(0.5)


# Generate Password function
def nickname_generator(length=8):
    printable = f'{LETTERS}{NUMBERS}'
    printable = list(printable)

    random.shuffle(printable)

    random_password = random.choices(printable, k=length)
    random_password = ''.join(random_password)

    return random_password


def auto_fill(first_name="first", last_name="last", email="email@gmail.com", ssn='123456789', phone='1234567890',
              address="addressne",nickname="nickname"):
    try:
        # maximize browser
        # options = webdriver.ChromeOptions()
        # options.add_argument("--start-maximized")

        # mo url
        # driver = webdriver.Chrome(options=options)
        # driver.get(URL)
        
        driver = webdriver.Firefox()
        driver.maximize_window()
        driver.get(URL)

        # click on Standard Application
        driver.find_element_by_xpath('/html/body/div[3]/div[1]/div[3]/div[1]/div[3]/button/span').click()

        # input
        first_name_ele = driver.find_element_by_id('FirstName')
        first_name_ele.send_keys(first_name.title())
        pause_chut()

        last_name_ele = driver.find_element_by_id('LastName')
        last_name_ele.send_keys(last_name.title())
        pause_chut()

        nickname_ele = driver.find_element_by_id('NickName')
        nickname_ele.send_keys(nickname)
        pause_chut()

        gender_select = Select(driver.find_element_by_id('GenderID'))
        gender_select.select_by_visible_text("Male")
        pause_chut()

        suffix_select = Select(driver.find_element_by_id('Suffix'))
        suffix_select.select_by_visible_text("Sr.")
        pause_chut()

        dobmonth_sel = Select(driver.find_element_by_id('DOBMonth'))
        mm = random.choice(MONTHS)
        dobmonth_sel.select_by_visible_text(mm)
        pause_chut()
        dobday_sel = Select(driver.find_element_by_id('DOBDay'))
        dd = str(random.choice(DAYS))
        dobday_sel.select_by_visible_text(dd)
        pause_chut()
        dobyear_sel = Select(driver.find_element_by_id('DOBYear'))
        yy = str(random.choice(YEARS))
        dobyear_sel.select_by_visible_text(yy)
        pause_chut()

        ssn_ele = driver.find_element_by_id('SSN1')
        ssn_ele.click()
        ssn_ele.send_keys(ssn)
        pause_chut()

        ssn_ele = driver.find_element_by_id('CellPhone_AreaCode')
        ssn_ele.click()
        ssn_ele.send_keys(phone)
        pause_chut()

        email_ele = driver.find_element_by_id('Email')
        email_ele.send_keys(email)
        pause_chut()
        email_ele2 = driver.find_element_by_id('EmailConfirmed')
        email_ele2.send_keys(email)
        pause_chut()

        driver.find_element_by_xpath(
            '/html/body/div[3]/div[2]/div[6]/div[2]/div[2]/div[2]/div[2]/div[3]/div[1]/input[1]').click()
        pause_chut()

        latino_select = Select(driver.find_element_by_id('Ethnicity_Hispanic'))
        latino_select.select_by_visible_text("No")
        pause_chut()

        driver.find_element_by_xpath(
            '/html/body/div[3]/div[2]/div[6]/div[2]/div[2]/div[2]/div[2]/div[6]/div/div[3]/div[3]/input[1]').click()
        pause_chut()

        driver.find_element_by_xpath('/html/body/div[3]/div[2]/div[7]/button[2]/span').click()
        pause_chut()

        time.sleep(2)

        state_select = Select(driver.find_element_by_id('Address_StateID'))
        state_select.select_by_visible_text(STATE)
        pause_chut()

        city_ele = driver.find_element_by_id('Address_City')
        city_ele.send_keys(CITY)
        pause_chut()

        zip_ele = driver.find_element_by_id('Address_Zip')
        zip_ele.send_keys(ZIPCODE)
        pause_chut()

        address_ele = driver.find_element_by_id('Address_Address1')
        address_ele.send_keys(address)
        pause_chut()

        years_select = Select(driver.find_element_by_id('Address_YearsInOhio'))
        years_select.select_by_visible_text("5")
        pause_chut()
        months_select = Select(driver.find_element_by_id('Address_MonthsInOhio'))
        months_select.select_by_visible_text("5")
        pause_chut()

        driver.find_element_by_xpath('/html/body/div[3]/div[4]/div[5]/button[2]/span').click()
        pause_chut()

        semester_select = Select(driver.find_element_by_id('SemesterID'))
        semester_select.select_by_visible_text("Spring Semester 2021")
        pause_chut()
        reason_select = Select(driver.find_element_by_id('ReasonForCS'))
        reason_select.select_by_visible_text("To obtain an associate degree - for transfer")
        pause_chut()
        driver.find_element_by_xpath('/html/body/div[3]/div[5]/div[4]/div/div[3]/div/div[2]/div[2]/div[2]/div[1]/div[1]/div[2]/label[2]').click()
        pause_chut()
        driver.find_element_by_xpath('/html/body/div[3]/div[5]/div[4]/div/div[3]/div/div[2]/div[2]/div[2]/div[1]/div[2]/div[2]/div/div[3]/div[2]/span[2]/span/a').click()
        time.sleep(5)
        driver.find_element_by_link_text("Business Management Technology (BM.AAB)").click()
        time.sleep(1)
        driver.find_element_by_xpath('//*[@id="prog-detail-BM_AAB"]/div[1]/button[1]/span').click()
        pause_chut()

        driver.find_element_by_xpath(
            # '/html/body/div[3]/div[5]/div[4]/div/div[3]/div/div[2]/div[2]/div[1]/div[3]/div[2]/input[1]'
            '/html/body/div[3]/div[5]/div[5]/button[2]/span'
            ).click()
        pause_chut()

        #driver.find_element_by_xpath('/html/body/div[3]/div[5]/div[5]/button[2]/span').click()
        #pause_chut()

        hs_select = Select(driver.find_element_by_id('HighSchoolStatusID'))
        hs_select.select_by_visible_text("High School Diploma")
        pause_chut()

        hsgm_sel = Select(driver.find_element_by_id('HighSchoolGraduationMonth'))
        hsgm_sel.select_by_visible_text("August")
        pause_chut()
        hsgy_sel = Select(driver.find_element_by_id('HighSchoolGraduationYear'))
        hsgy_sel.select_by_visible_text("2019")
        pause_chut()
        hsid_sel = Select(driver.find_element_by_id('HighSchoolID'))
        hsid_sel.select_by_visible_text("Brookville High School (OH)")
        pause_chut()

        driver.find_element_by_xpath('/html/body/div[3]/div[6]/div[5]/button[2]/span').click()
        pause_chut()

        wait_time = input("Nhap captcha xong thi enter")
        pause_chut()

        driver.find_element_by_xpath('/html/body/div[3]/div[10]/div[3]/button[2]/span').click()
        pause_chut()

        time.sleep(13)

        q1_sel = Select(driver.find_element_by_id('standard-q12'))
        q1_sel.select_by_visible_text("Yes")
        pause_chut()
        q2_sel = Select(driver.find_element_by_id('standard-q13'))
        q2_sel.select_by_visible_text("Yes")
        pause_chut()
        q3_sel = Select(driver.find_element_by_id('standard-q6'))
        q3_sel.select_by_visible_text("Web search/Internet")
        pause_chut()
        q4_sel = Select(driver.find_element_by_id('standard-q7'))
        q4_sel.select_by_visible_text("Part-time")
        pause_chut()
        q5_sel = Select(driver.find_element_by_id('standard-q8'))
        q5_sel.select_by_visible_text("Part-time (less than 12 credit hours)")
        pause_chut()
        q6_sel = Select(driver.find_element_by_id('standard-q9'))
        q6_sel.select_by_visible_text("Campus Tour")
        pause_chut()
        q7_sel = Select(driver.find_element_by_id('standard-q10'))
        q7_sel.select_by_visible_text("Partial high school")
        pause_chut()
        q8_sel = Select(driver.find_element_by_id('standard-q11'))
        q8_sel.select_by_visible_text("None")
        pause_chut()
        q9_sel = Select(driver.find_element_by_id('standard-q14'))
        q9_sel.select_by_visible_text("Afternoons")
        pause_chut()

        driver.find_element_by_xpath(
            '/html/body/div[3]/div[14]/div[4]/div[5]/div[2]/div[2]/div[1]/div[3]/button/span').click()
        pause_chut()

        driver.quit()

        return (1, yy, mm, dd)
    except Exception as e:
        print(e)
        return (0, 0, 0, 0)


# create a database connection to a SQLite database
def create_connection(db_file):
    connection = None
    try:
        connection = sqlite3.connect(db_file)
        c = connection.cursor()
        c.execute("""CREATE TABLE IF NOT EXISTS IMPORTED_DATA (
                          phone text,
                          ssn text PRIMARY KEY,
                          zipcode text ,
                          nickname text,
                          first_name text,
                          last_name text,
                          dob text,
                          email text,
                          address text,
                          creation_date text
                          )""")
        print("sqlite3 version:", sqlite3.version)
    except Error as e:
        print(e)

    return connection

    
# function check if ssn is exists
def check_ssn_exists(conn, ssn):
    cur = conn.cursor()
    cur.execute("SELECT SSN FROM IMPORTED_DATA WHERE SSN = ?", (ssn,))
    rows = cur.fetchall()

    for _ in rows:
        return 1

    return 0


def insert_data(conn, data):
    sql = """INSERT INTO IMPORTED_DATA (
                        phone,
                        ssn,
                        zipcode,
                        nickname,
                        first_name,
                        last_name,
                        dob,
                        email,
                        address,
                        creation_date)
            VALUES (?,?,?,?,?,?,?,?,?,?)"""
    c = conn.cursor()
    c.execute(sql, data)
    conn.commit()

    return c.lastrowid


def export_db():
    workbook = Workbook('output.xlsx')
    worksheet = workbook.add_worksheet()
    conn = create_connection(DB_FILE)
    c = conn.cursor()
    sel = c.execute("SELECT * FROM IMPORTED_DATA")
    for i, row in enumerate(sel):
        for j, value in enumerate(row):
            worksheet.write(i, j, value)

    workbook.close()


def get_email():
    email_user_input = input("Nhap email:")
    return email_user_input


def main():
    line = int(input("Nhap line muon xu ly: "))

    # Give the location of the file
    loc = "data.xlsx"

    # To open Workbook
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)

    conn = create_connection(DB_FILE)

    i = line - 1
    email = get_email()
    data = None
    creation_date = datetime.today().strftime('%Y-%m-%d-%H:%M:%S')

    first_name = sheet.cell_value(i, 0)
    last_name = sheet.cell_value(i, 1)
    address = sheet.cell_value(i, 2)
    phone = sheet.cell_value(i, 3)
    ssn = sheet.cell_value(i, 4)
    nickname = nickname_generator(16)

    is_exists = check_ssn_exists(conn, ssn)

    if is_exists == 0:
        is_ok = 0
        y = ''
        m = ''
        d = ''

        (is_ok, y, m, d) = auto_fill(first_name=first_name, last_name=last_name, email=email, ssn=ssn, phone=phone, address=address, nickname=nickname)

        dob = str(y) + '/' + str(m) + '/' + str(d)

        if is_ok == 1:
            data = (phone, ssn, ZIPCODE, nickname, first_name, last_name, dob, email, address, creation_date)
            insert_data(conn, data)
            print("Chay thanh cong SSN:", ssn)
    else:
        print("SSN da xu ly roi")

    if conn:
        conn.close()


if __name__ == "__main__":
    choice = int(input("Nhap 1 de chay chuong trinh, 2 de lay output"))
    if choice == 1:
        main()
    elif choice == 2:
        export_db()
        
