import selenium
from selenium import webdriver
from selenium.webdriver.common.by import By

import ctypes
from time import sleep
import math
from bs4 import BeautifulSoup
import re
import xlsxwriter
import sys
from PyQt6.QtWidgets import *
from PyQt6 import QtCore, QtGui


class Trudvsem_parcer():
    def __init__(self, vacantion_search_url='https://trudvsem.ru/vacancy/search', num_pages=0):
        self.url = vacantion_search_url
        self.num_pages = num_pages

    def parcing(self):
        book = xlsxwriter.Workbook(r"./parced_data.xlsx")
        page = book.add_worksheet("данные")
        #from selenium_stealth import stealth
        # options = webdriver.ChromeOptions()
        # options.add_argument("start-maximized")
        # options.add_experimental_option("detach", True)
        #driver = webdriver.Chrome()#options=options)  #
        # stealth(driver,
        #         languages=["ru", "ru-RU"],
        #         vendor="Google Inc.",
        #         platform="Win32",
        #         webgl_vendor="Intel Inc.",
        #         renderer="Intel Iris OpenGL Engine",
        #         fix_hairline=True,
        #         )
        driver = webdriver.Firefox()


        try:
            driver.get(self.url)
            start_button = driver.find_element(By.XPATH, "//button[@class='search-content__button']")
            start_button.click()
            sleep(3)
            if self.num_pages == 0:
                try:
                    num_vacancy_text = driver.find_element(By.CLASS_NAME, 'ib-filter__result-counter').text
                    num_vacancy = int(''.join(num_vacancy_text[:num_vacancy_text.rfind(' ')].split()))
                    if num_vacancy > 300:
                        num_vacancy = 300
                    print(num_vacancy_text)
                    for _ in range(math.ceil(num_vacancy / 10)):
                        sleep(1)
                        element = driver.find_elements(By.CLASS_NAME, 'button_secondary')
                        for e in element:
                            if e.text == 'Загрузить ещё':
                                driver.execute_script("arguments[0].click();", e)
                    for i in range(num_vacancy+1):
                        elem = driver.find_element(By.XPATH, '//div[@class="search-results-simple-card mb-1"]')
                        info_div = elem.find_elements(By.XPATH,
                                                      '//div[@class="search-results-simple-card__wrapper search-results-simple-card__wrapper_column"]')

                        soup_vacancy_info = BeautifulSoup(info_div[i].get_attribute('innerHTML'), 'lxml')
                        employer, region = list(
                            map(lambda x: x.text,
                                soup_vacancy_info.find_all('div', {'class': 'content_small content_clip'})))

                        driver.execute_script("arguments[0].click();", elem)
                        soup = BeautifulSoup(driver.page_source, "lxml")
                        vacancy_name_html = soup.find('a', {'class': "link link_title"})
                        if not vacancy_name_html:
                            while not vacancy_name_html:
                                soup = BeautifulSoup(driver.page_source, "lxml")

                        vacancy_name = vacancy_name_html.text
                        print(vacancy_name)
                        salary = soup.find('span', {
                            'class': 'content__section-subtitle search-results-full-card__salary'}).text.strip()
                        print(salary)
                        date = soup.find('span', {'class': 'content_small content_pale'}).text
                        date = date[date.find(' ') + 1:]
                        print(date)
                        print(region)
                        print(employer)
                        vacancy_descr = soup.find('div', {'class': "tabs__content tabs_active",
                                                          'id': "vacancy-details"}).text.split()
                        vacancy_descr = ' '.join(list(map(lambda x: x.strip(), vacancy_descr)))
                        print(vacancy_descr)
                        requirements = vacancy_descr[
                                       vacancy_descr.find('Требования к кандидату'):vacancy_descr.find(
                                           'Данные по вакансии')]
                        print('requirements:')
                        print(requirements)
                        print('\n')
                        if 'Опыт работы' in requirements:
                            print('experience:')
                            experience = requirements[requirements.find('Опыт работы'):]
                            experience = experience[:[m.start() for m in re.finditer(' ', experience)][4]]
                            print(experience)
                        else:
                            experience = ''
                            print(experience)

                        if 'График работы' in vacancy_descr:
                            print('schedule:')
                            schedule = vacancy_descr[vacancy_descr.find('График работы'):]
                            schedule = schedule[:[m.start() for m in re.finditer(' ', schedule)][2]]
                            print(schedule)
                        else:
                            schedule = ''
                            print(schedule)

                        data = [vacancy_name, vacancy_descr, salary, date, region, employer, requirements, experience,
                                schedule]
                        column_num = 0
                        for col in data:
                            page.write(i, column_num, col)
                            column_num += 1

                        print(i)
                except selenium.common.exceptions.NoSuchElementException:
                    print('Завершено')
                book.close()
        except:
            book.close()
            print('Выход')


class ExampleApp(QWidget):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.setWindowTitle("Trudvsem-parcer")
        self.setWindowIcon(QtGui.QIcon('./icon.svg'))
        self.setMinimumSize(400,0)
        self.main_layout = QVBoxLayout()
        self.parsing_layout = QVBoxLayout()
        self.loading_layout = QVBoxLayout()

        self.label_main = QLabel()
        font = QtGui.QFont()
        font.setPointSize(16)
        self.label_main.setFont(font)
        self.label_main.setTextFormat(QtCore.Qt.TextFormat.AutoText)
        self.label_main.setObjectName("label_main")
        self.label_main.setText("Введите url с сайта trudvsem.ru")

        self.url = QLineEdit()
        self.url.setInputMask("")
        self.url.setObjectName("url")
        self.url.setText('https://trudvsem.ru/vacancy/search?_title=')

        self.parsing_button = QPushButton()
        self.parsing_button.setEnabled(True)
        self.parsing_button.setObjectName("parsing_button")
        self.parsing_button.setText('Спарсить')

        self.clear_switch = False
        self.clear_button = QPushButton('Очистить')
        self.clear_button.clicked.connect(lambda: self.clear_layer(self.loading_layout))

        self.parsing_layout.addWidget(self.label_main)
        self.parsing_layout.addWidget(self.url)
        self.parsing_layout.addWidget(self.parsing_button)

        self.main_layout.addLayout(self.parsing_layout)
        self.main_layout.addLayout(self.loading_layout)

        self.setLayout(self.main_layout)
        self.parsing_button.clicked.connect(self.parcing)

    def parcing(self):
        self.label1 = QLabel("Загрузка...")
        self.loading_layout.addWidget(self.label1)
        self.parsing_button.setEnabled(False)

        Trudvsem_parcer(self.url.text()).parcing()

        self.parsing_button.setEnabled(True)
        self.label2 = QLabel("Выполнено!")
        self.loading_layout.addWidget(self.label2)
        self.main_layout.addWidget(self.clear_button)
        self.clear_switch = True

    def clear_layer(self, layer):
        for i in reversed(range(layer.count())):
            layer.takeAt(i).widget().deleteLater()
        self.resize(400, 1)


def main():
    app = QApplication(sys.argv)
    window = ExampleApp()
    window.show()
    app.exec()


if __name__ == '__main__':
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("company.app.1")
    main()