import os
import csv
import smtplib
import ssl
import bs4 as bs
from selenium import webdriver
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from pathlib import Path


class TucanGradeNotification:

    def __init__(self):
        # Insert TU-ID and TU-Password
        username = ""
        password = ""
        # Insert TU-Email-Address (...@stud.tu-darmstadt.de) and receiver-email-address
        sender_email = ""
        receiver_email = ""

        temp_path = self.shorten_path(os.path.abspath(__file__), 1)
        path_csv = os.path.join(temp_path, 'grades.csv')

        opts = webdriver.FirefoxOptions()
        opts.headless = True
        driver = webdriver.Firefox(options=opts)

        tucan_grade_view = self.open_tucan_grade_view(driver, username, password)

        grade_list = self.get_grade_list(tucan_grade_view)

        driver.quit()

        if not os.path.exists(path_csv):
            self.save_grades_to_csv(grade_list, path_csv)
        else:
            change_detected, data_from_csv = self.check_for_new_grades(grade_list, path_csv)
            if change_detected:
                change_list = [[x[1], x[4]] for x in grade_list if x not in data_from_csv]
                self.send_mail(change_list, sender_email, receiver_email, username, password)
                self.save_grades_to_csv(grade_list, path_csv)

    def shorten_path(self, file_path, length):
        return Path(*Path(file_path).parts[:-length])

    def open_tucan_grade_view(self, driver, username, password):
        driver.implicitly_wait(6)
        driver.get("https://www.tucan.tu-darmstadt.de")

        username_input = driver.find_element_by_name("usrname")
        password_input = driver.find_element_by_name("pass")
        username_input.send_keys(username)
        password_input.send_keys(password)

        driver.find_element_by_id("logIn_btn").click()
        driver.implicitly_wait(6)
        driver.find_element_by_link_text("Prüfungen").click()
        driver.implicitly_wait(6)
        driver.find_element_by_link_text("Leistungsspiegel").click()

        soup = bs.BeautifulSoup(driver.page_source, "html.parser")
        return soup

    def get_grade_list(self, soup):
        grade_list = []
        table_grades = soup.find('table', attrs={'class': "nb list students_results"})
        table_grades_body = table_grades.find('tbody')
        rows_grades = table_grades_body.find_all('tr')

        for row in rows_grades:
            relevant_row = row.find('a', href=True)
            if relevant_row is not None:
                cols = row.find_all('td')
                cols = [ele.text.strip() for ele in cols]
                grade_list.append([ele for ele in cols if ele])

        return grade_list

    def save_grades_to_csv(self, grade_list, path_csv):
        with open(path_csv, 'w', newline='') as file:
            file.write(str(len(grade_list)) + '\n')
            writer = csv.writer(file)
            for item in grade_list:
                writer.writerow(item)

    def check_for_new_grades(self, grade_list, path_csv):
        change_detected = False
        data_from_csv = []
        with open(path_csv, 'rt') as f:
            data = csv.reader(f)
            for row in data:
                if int(row[0]) == len(grade_list):
                    break
                else:
                    for row2 in data:
                        data_from_csv.append(row2)
                    change_detected = True
                    break
            f.close()
        return change_detected, data_from_csv

    def send_mail(self, grades, sender_email, receiver_email, username, password):
        message = MIMEMultipart()
        subject = "New Grade Online"
        message['Subject'] = subject
        message['From'] = sender_email
        message['To'] = receiver_email

        text = ""
        for item in grades:
            text += str(item[0]) + ": " + str(item[1]) + '\n'

        body = MIMEText(text, 'plain')
        message.attach(body)

        context = ssl.create_default_context()
        with smtplib.SMTP_SSL('smtp.tu-darmstadt.de', 465, context=context) as server:
            server.login(username, password)
            server.sendmail(sender_email, receiver_email, message.as_string())


tucan_notificator = TucanGradeNotification()
