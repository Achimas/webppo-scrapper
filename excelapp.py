from win32com.client import Dispatch
import os
from config import rows_in_excel
from time import sleep
import sys
import json
import pythoncom

class ExcelApp():
    def __init__(self, directory = os.getcwd(), visible = True, config_data = False):
        self.app = Dispatch("Excel.Application",pythoncom.CoInitialize())
        self.app.Visible = visible
        self.app.DisplayAlerts = False
        self.directory = directory
        try:
            self.workbook = self.app.Workbooks.open(f"{directory}\\Template.xlsx")
        except:
            print("Проверьте файл Template.xlsx и запустите программу заново")
            sleep(1000)
            os._exit(1)
        self.sheet_IP = self.app.Worksheets('2')
        self.sheet_UL = self.app.Worksheets('1')
        self.adding_count = 0
        self.file_count = 1
        self.IP_count = 2
        self.UL_count = 2
        self.rows_in_excel = config_data["rows_in_excel"]
        self.config_data = config_data

    def insert_rest_count(self, rest_count):
        self.rest_count = rest_count

    def add_participant(self, participant):
        if self.adding_count >= self.rows_in_excel:
            self._do_next_file(participant['id'])
        self.participant = participant
        if self.participant['Org'] == 'ЮЛ':
            self._add_row(self.sheet_UL, self.UL_count, self.participant)
            self.UL_count += 1
            self.adding_count += 1      
        else:
            self._add_row(self.sheet_IP, self.IP_count, self.participant)
            self.IP_count += 1
            self.adding_count += 1
        # print("                                                                                ", end='\r')
        print (f"Org{self.file_count}: \
            {self.adding_count} из {self.rows_in_excel}, осталось на сайте: {self.rest_count}      ", end='\r')

    def close(self):
        self.workbook.SaveCopyAs(f"{self.directory}\\Org{self.file_count}.xlsx")
        self.workbook.Close()
        self.app.quit()

    def _add_row (self, sheet, count, participant):
        participant_list = []
        if self.participant['Org'] == 'ЮЛ':
            participant_list.append(participant['id'])
            participant_list.append(participant['Organisation_name'])
            participant_list.append(participant['INN'])
            participant_list.append(participant['KPP'])
            participant_list.append(participant['Phone'])
            participant_list.append(participant['Mail'])
            participant_list.append(participant['Add_mail'])
            participant_list.append(participant['Site'])
            participant_list.append(participant['Fax'])
            participant_list.append(participant['Position'])
            participant_list.append(participant['Lastname'])
            participant_list.append(participant['Name'])
            participant_list.append(participant['Patronymic'])
            participant_list.append(participant['Contact_info'])
        else:
            participant_list.append(participant['id'])
            participant_list.append(participant['Organisation_name'])
            participant_list.append(participant['INN'])
            participant_list.append(participant['Phone'])
            participant_list.append(participant['Mail'])
            participant_list.append(participant['Add_mail'])
            participant_list.append(participant['Fax'])
            participant_list.append(participant['Contact_info'])
            participant_list.append('')
            participant_list.append('')
            participant_list.append('')
            participant_list.append('')
            participant_list.append('')
            participant_list.append('')
        try:
            sheet.Range(f'A{count}:N{count}').Value = participant_list
        except:
            pass

    def _do_next_file(self, id):
        if self.config_data["last_saved_id"] < int(id) :
            self.config_data["last_saved_id"] = int(id)-1
        self.workbook.SaveAs(f"{self.directory}\\Org{self.file_count}.xlsx")
        self.workbook.Close()
        self.save_config_to_json('config.json', self.config_data)
        self.workbook = self.app.Workbooks.open(f"{self.directory}\\Template.xlsx")
        # self.workbook = self.app.workbooks("Template.xlsx").Activate()
        self.sheet_IP = self.app.Worksheets('2')
        self.sheet_UL = self.app.Worksheets('1')
        self.adding_count = 0
        self.file_count += 1
        self.IP_count = 2
        self.UL_count = 2
        self.rest_count -= self.rows_in_excel
        # self.sheet_IP.Range('A2:N10001').ClearContents
        # self.sheet_UL.Range('A2:N10001').ClearContents

    def save_config_to_json(self, name, config_data):
        fp = open(name, 'w')
        json.dump(config_data, fp)
        fp.close()

if __name__ == '__main__':
    participant = [1, 'IP vasya pup', '14566743', '1', 1, 'IP vasya pup']
    # print (participant, end='\r')
    # sleep(2)
    # print(" "*100, end='\r')
    # sleep(2)
    # print (participant)
    xl = ExcelApp()
    count = 2
    print(xl.sheet_IP)
    print(xl.sheet_UL)
    xl.sheet_IP.Range(f'A{count}:N{count}').Value = participant
    xl.close()