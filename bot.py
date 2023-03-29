from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
from excelapp import ExcelApp
import json
import sys
import os
import threading
from pynput.keyboard import GlobalHotKeys
# from selenium.webdriver.support.wait import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC

class ParseBot():
    def __init__(self, config_data):
        options = webdriver.ChromeOptions()
        options.add_argument("log-level=3")
        options.add_argument('headless')
        self.driver = webdriver.Chrome(options=options)
        self.config_data = config_data
        self.login = config_data["login"]
        self.password = config_data["password"]
        self.sleeps = config_data["sleeps"]
        self.from_id = config_data["from_id"]
        self.last_needed_id = config_data["to_id"]
        self.decrease_count = config_data["decrease_count"]
        self.rows_in_excel = config_data["rows_in_excel"]
        self.last_saved_id = config_data["last_saved_id"]
        # self.last_site_id = self._get_last_site_id()

    def start_parsing(self):
        self.driver.get('http://webppo.zakazrf.ru/Logon/Participants/id/133159927760330491')
        login_field = self.driver.find_element('xpath', '/html/body/div[3]/div/form/div/table/tbody/tr[1]/td[2]/input')
        login_field.send_keys(self.login)
        pass_field = self.driver.find_element('xpath', '/html/body/div[3]/div/form/div/table/tbody/tr[2]/td[2]/input')
        pass_field.send_keys(self.password)
        login_in = self.driver.find_element('xpath', '/html/body/div[3]/div/form/div/button[1]/span')
        sleep(self.sleeps)
        login_in.click()
        print("Вход на сайте выполнен!")
        self._get_last_site_id()

    def parsing(self, event=None, xl=None):
        if self.last_needed_id == 0 and self.from_id == 0:
            from_id = self.last_saved_id + 1
            last_site_id = self.last_site_id
        else:
            from_id = self.from_id
            last_site_id = self.last_needed_id
        while from_id > self.last_site_id:
            from_id -= self.decrease_count
        rest_count = int(last_site_id) - int(from_id)
        xl.insert_rest_count(rest_count)
        for id in range(from_id, last_site_id+1):
            if event.is_set():
                xl.close()
                save_config_to_json('config.json', self.config_data)
                print(f"\nПрограмма прекращает работу, последняя сохраненная организация: {id-1}")
                sleep(1000)
                os._exit(1)
            else:
                participant = self._parsing_participant(id)
                xl.add_participant(participant)
                # self.config_data["from_id"] = id
                # save_config_to_json('config.json', self.config_data)
        if last_site_id > self.last_saved_id:
            self.config_data["last_saved_id"] = last_site_id
        save_config_to_json('config.json', self.config_data)

        print("Программа завершила работу, файлы Excel сохранены в папке расположения программы", end='\r')

    def _parsing_participant(self, id):
        participant = {}
        participant['id'] = id
        self.driver.get(f'http://webppo.zakazrf.ru/Participant/id/{id}')
        self._add_field(participant, 'Org', '/html/body/div[3]/div/form/div/div[4]/div[1]/table[1]/tbody/tr[3]/td[2]')
        if participant['Org'] == 'ЮЛ':
            self._add_field(participant, 'Organisation_name', '/html/body/div[3]/div/form/div/div[4]/div[1]/table[1]/tbody/tr[5]/td[2]')
            self._add_field(participant, 'INN', '/html/body/div[3]/div/form/div/div[4]/div/table[3]/tbody/tr[1]/td[2]')
            self._add_field(participant, 'KPP', '/html/body/div[3]/div/form/div/div[4]/div[1]/table[3]/tbody/tr[2]/td[2]')
            self._add_field(participant, 'Phone', '/html/body/div[3]/div/form/div/div[4]/div[1]/table[3]/tbody/tr[5]/td[2]')
            self._add_field(participant, 'Mail', '/html/body/div[3]/div/form/div/div[4]/div[1]/table[3]/tbody/tr[6]/td[2]/a')
            self._add_field(participant, 'Add_mail', '/html/body/div[3]/div/form/div/div[4]/div[1]/table[3]/tbody/tr[7]/td[2]')
            self._add_field(participant, 'Site', '/html/body/div[3]/div/form/div/div[4]/div[1]/table[3]/tbody/tr[8]/td[2]')
            self._add_field(participant, 'Fax', '//html/body/div[3]/div/form/div/div[4]/div[1]/table[3]/tbody/tr[9]/td[2]')
            self._add_field(participant, 'Lastname', '/html/body/div[3]/div/form/div/div[4]/div[1]/table[4]/tbody/tr[1]/td[2]')
            self._add_field(participant, 'Name', '/html/body/div[3]/div/form/div/div[4]/div[1]/table[4]/tbody/tr[2]/td[2]')
            self._add_field(participant, 'Patronymic', '/html/body/div[3]/div/form/div/div[4]/div[1]/table[4]/tbody/tr[3]/td[2]')
            self._add_field(participant, 'Position', '/html/body/div[3]/div/form/div/div[4]/div[1]/table[4]/tbody/tr[4]/td[2]')
            self._add_field(participant, 'Contact_info', '/html/body/div[3]/div/form/div/div[4]/div[1]/table[4]/tbody/tr[5]/td[2]')
        else:
            self._add_field(participant, 'Organisation_name', '/html/body/div[3]/div/form/div/div[4]/div[1]/table[1]/tbody/tr[4]/td[2]')
            self._add_field(participant, 'INN', '/html/body/div[3]/div/form/div/div[4]/div[1]/table[3]/tbody/tr[1]/td[2]')
            self._add_field(participant, 'Phone', '/html/body/div[3]/div/form/div/div[4]/div[1]/table[3]/tbody/tr[3]/td[2]')
            self._add_field(participant, 'Mail', '/html/body/div[3]/div/form/div/div[4]/div[1]/table[3]/tbody/tr[4]/td[2]')
            self._add_field(participant, 'Add_mail', '/html/body/div[3]/div/form/div/div[4]/div[1]/table[3]/tbody/tr[5]/td[2]')
            self._add_field(participant, 'Fax', '/html/body/div[3]/div/form/div/div[4]/div[1]/table[3]/tbody/tr[6]/td[2]')
            self._add_field(participant, 'Contact_info', '/html/body/div[3]/div/form/div/div[4]/div[1]/table[4]/tbody/tr[5]/td[2]')
        return participant

    def _get_last_site_id(self):
        self.driver.get(f'http://webppo.zakazrf.ru/Participant')
        sleep(self.sleeps)
        last_page = self.driver.find_elements('xpath', '/html/body/div[3]/div/form/div/a[4]/span')[0]
        last_page.click()
        self.last_site_id = 20
        while self.last_site_id == 20:
            try:
                last_site_link = self.driver.find_elements(By.CLASS_NAME, 'property-link')[-1].get_property('href')
                self.last_site_id = int(last_site_link.split('/')[-1])
            except:
                pass
        print(f"Последняя организация на сайте: {self.last_site_id}")

    def _add_field(self, participant, field_name, xpath):
        try:
            participant[field_name] = self.driver.find_element('xpath', xpath).text
        except:
            participant[field_name] = ''

def open_json(name):
    try:
        f = open(name)
        config_data = json.load(f)
        f.close()
    except:
        print("Проверьте файл config.json и запустите программу заново")
        sleep(1000)
        os._exit(1)
    return config_data

def save_config_to_json(name, config_data):
    try:
        fp = open(name, 'w')
        json.dump(config_data, fp)
        fp.close()
    except:
        print("Проверьте файл config.json и запустите программу заново")
        sleep(1000)
        os._exit(1)

def global_application(event):
    config_data = open_json('config.json')
    xl = ExcelApp(visible=False, config_data=config_data)
    PB = ParseBot(config_data)
    PB.start_parsing()
    # PB._get_last_site_id()
    PB.parsing(event=event, xl=xl)
    sleep(10)
    xl.close()
    # xl.workbook.close()
    # xl.workbook = xl.app.Workbooks.open(f"{xl.directory}\\Template.xlsx")
    # xl.workbook.close()

def toggle_event():
    global parsing_stop
    parsing_stop.set()

def main():
    global parsing_stop
    parsing_stop = threading.Event()
    parsing_thread = threading.Thread(target=global_application, args=(parsing_stop,))
    parsing_thread.start()
    # in threading of excel don't foget about pythoncom.CoInitializeEx(0)    
    with GlobalHotKeys({'<ctrl>+<alt>+s': toggle_event}) as listener:
        listener.join()

if __name__ == '__main__':
    main()
    sleep(1000)