import time
from selenium import webdriver
import pandas as pd
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)

PATH = "./chromedriver.exe" #ต้องเช็คเวอร์ชั่น Chrome ทุกครั้งแล้ว dowload chromedriver
driver = webdriver.Chrome(PATH)
df = pd.DataFrame()

driver.get("https://web.reg.tu.ac.th/registrar/class_info.asp?lang=th")

for fac in range(8, 10) :
    for term in range(2):
        # drpo down
        driver.find_element(By.CSS_SELECTOR, "body > table > tbody > tr.ContentBody > td:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(2) > font:nth-child(2) > select").click()
        # select major
        driver.find_element(By.CSS_SELECTOR, "body > table > tbody > tr.ContentBody > td:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(2) > font:nth-child(2) > select > option:nth-child("+ str(fac) +")").click()
        # drpo down term
        driver.find_element(By.CSS_SELECTOR, "body > table > tbody > tr.ContentBody > td:nth-child(2) > table > tbody > tr:nth-child(6) > td:nth-child(2) > table > tbody > tr:nth-child(1) > td:nth-child(2) > font:nth-child(1) > select").click()
        # select term
        driver.find_element(By.CSS_SELECTOR, "body > table > tbody > tr.ContentBody > td:nth-child(2) > table > tbody > tr:nth-child(6) > td:nth-child(2) > table > tbody > tr:nth-child(1) > td:nth-child(2) > font:nth-child(1) > select > option:nth-child(" + str(term + 1) + " )").click()
        # search
        driver.find_element(By.XPATH, "/html/body/table/tbody/tr[1]/td[2]/table/tbody/tr[7]/td[2]/table/tbody/tr/td/font[3]/input").click()


        while (True):
            lenoftr = driver.find_elements(By.XPATH, "/html/body/table/tbody/tr[1]/td[2]/font/font/font/table/tbody/tr")
            for i in range(4, len(lenoftr) - 1):
                try:
                    teacher = " "
                    center = driver.find_element(By.CSS_SELECTOR, "body > table > tbody > tr.ContentBody > td:nth-child(2) > font > font > font > table > tbody > tr:nth-child(" + str(i) + ") > td:nth-child(2)").text
                    year = driver.find_element(By.CSS_SELECTOR, "body > table > tbody > tr.ContentBody > td:nth-child(2) > font > font > font > table > tbody > tr:nth-child(" + str(i) + ") > td:nth-child(3)").text
                    course = driver.find_element(By.CSS_SELECTOR, "body > table > tbody > tr.ContentBody > td:nth-child(2) > font > font > font > table > tbody > tr:nth-child(" + str(i) + ") > td:nth-child(5)").text
                    subject = driver.find_element(By.CSS_SELECTOR, "body > table > tbody > tr.ContentBody > td:nth-child(2) > font > font > font > table > tbody > tr:nth-child(" + str(i) + ") > td:nth-child(6)").text
                    for l in range(len(subject.splitlines())):
                        if l == 0:
                            subjects = subject.splitlines()[0]
                        elif subject.splitlines()[l][0] != '*':
                            teacher = teacher + subject.splitlines()[l] + ","

                    credit = driver.find_element(By.CSS_SELECTOR, "body > table > tbody > tr.ContentBody > td:nth-child(2) > font > font > font > table > tbody > tr:nth-child(" + str(i) + ") > td:nth-child(7)").text
                    section = driver.find_element(By.CSS_SELECTOR, "body > table > tbody > tr.ContentBody > td:nth-child(2) > font > font > font > table > tbody > tr:nth-child(" + str(i) + ") > td:nth-child(8)").text
                    times = driver.find_element(By.CSS_SELECTOR, "body > table > tbody > tr.ContentBody > td:nth-child(2) > font > font > font > table > tbody > tr:nth-child(" + str(i) + ") > td:nth-child(9)").text
                    times_exam = driver.find_element(By.CSS_SELECTOR, "body > table > tbody > tr.ContentBody > td:nth-child(2) > font > font > font > table > tbody > tr:nth-child(" + str(i) + ") > td:nth-child(10)").text
                    count = driver.find_element(By.CSS_SELECTOR, "body > table > tbody > tr.ContentBody > td:nth-child(2) > font > font > font > table > tbody > tr:nth-child(" + str(i) + ") > td:nth-child(11)").text
                    seat = driver.find_element(By.CSS_SELECTOR, "body > table > tbody > tr.ContentBody > td:nth-child(2) > font > font > font > table > tbody > tr:nth-child(" + str(i) + ") > td:nth-child(12)").text
                    status = driver.find_element(By.CSS_SELECTOR, "body > table > tbody > tr.ContentBody > td:nth-child(2) > font > font > font > table > tbody > tr:nth-child(" + str(i) + ") > td:nth-child(13)").text
                    faculty = driver.find_element(By.CSS_SELECTOR, "body > table > tbody > tr.ContentBody > td:nth-child(2) > div:nth-child(2) > font > b").text

                    df = df.append(
                        {
                            'ศูนย์': center,
                            'หลักสูตร': year,
                            'รหัสวิชา': course,
                            'ชื่อวิชา': subjects,
                            'อาจารย์': teacher,
                            'หน่วยกิต': credit,
                            'Section': section,
                            'เวลา': times,
                            'เวลาสอบ': times_exam,
                            'จำนวนรับ': count,
                            'เหลือ': seat,
                            'สถานะ': status,
                            'คณะ': faculty
                        }, ignore_index=True
                    )
                    print(teacher)

                except:
                    break
            try:
                href = driver.find_element(By.PARTIAL_LINK_TEXT, "[หน้าต่อไป]")
                href.click()
            except:
                break
        # back
        driver.find_element(By.CSS_SELECTOR, "#menu > ul > li > a").click()
df.to_excel('data.xlsx', sheet_name='Fact_Table')
time.sleep(5)
driver.quit()
