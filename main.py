import os.path
import glob
from collections.abc import Iterable
from pprint import pprint
import win32com.client

import openpyxl
import requests
from bs4 import BeautifulSoup
import aspose.tasks as tasks

START_ROW_VEDOM = 5
END_ROW_VEDOM = 38

abspath = os.getcwd()
DIRECTORY = os.path.join(abspath, 'файлы')


def parce_vedom():
    vedom = glob.glob(os.path.join(DIRECTORY, '*.xlsx'))[0]
    file = openpyxl.load_workbook(vedom, data_only=True)
    sheet = file["ВОР"]

    name_work = [cell.value.strip() for cell in sheet['B'][START_ROW_VEDOM - 1:END_ROW_VEDOM] if cell.value]
    GSN = [cell.value.strip() for cell in sheet['E'][START_ROW_VEDOM - 1:END_ROW_VEDOM] if cell.value]
    volume = [cell.value for cell in sheet['D'][START_ROW_VEDOM - 1:END_ROW_VEDOM] if cell.value]

    file.close()

    return [{work: {'gsn': gsn, 'volume': vol}} for work, gsn, vol in zip(name_work, GSN, volume)]


def parce_defsmeta(gsn):
    resources = {g: [] for g in gsn}
    temp_url = r'https://www.defsmeta.com/rgsn/gsn_{}/giesn-{}.php'
    # https://www.defsmeta.com/rgsn/gsn_11/giesn-11-01-036-01.php

    for code in gsn:
        main_code = code.split('-')[0]
        url = temp_url.format(main_code, code)
        page = requests.get(url)
        soup = BeautifulSoup(page.text, 'lxml')

        start_element = soup.find('p', string='РАСХОД МАТЕРИАЛОВ')
        if start_element:
            elements_between = []
            current_element = start_element
            for _ in range(3):
                elements_between.append(current_element)
                current_element = current_element.find_next()

            table = elements_between[2]
            rows = table.find_all('tr')[1:-1]

            for material in rows:
                material = material.find_all('td')[1:-1]
                resources[code].append({
                    'name': material[1].text,
                    'consumption': float(material[3].text.replace(',', '.')),
                })

    return resources


class MSProject:
    def __init__(self, file):
        self.mpp = win32com.client.Dispatch("MSProject.Application")
        self.mpp.FileOpen(file)
        self.project = self.mpp.ActiveProject

    def delete_resources(self, resource_name: Iterable):
        for r in self.project.Resources:
            if r.name in resource_name:
                r.Delete()

    def delete_all_resources(self):
        self.delete_resources(self.get_name_of_resources())

    def add_resources(self, resource_name: Iterable):
        for name in resource_name:
            self.project.Resources.Add(Name=name)

    def get_name_of_resources(self):
        return [r.name for r in self.project.Resources]

    def get_resources_name_id(self):
        return {t.Name: t.Id for t in self.project.Resources}

    def get_tasks_id_name(self):
        return {t.Id: t.Name for t in self.project.Tasks}

    def get_resources_object(self):
        return self.project.Resources

    def get_tasks_object(self):
        return self.project.Tasks

    def _save_file(self):
        self.mpp.FileSave()

    def _close_file(self):
        self.mpp.Quit()

    def close(self):
        self._save_file()
        self._close_file()


def main():
    project = MSProject(glob.glob(os.path.join(DIRECTORY, '*.mpp'))[0])
    vedom = parce_vedom()[:]

    gsns = []
    for w in vedom:
        for g in w.values():
            gsns.append(g['gsn'])

    resources_defsmeta = parce_defsmeta(gsns)

    # pprint(vedom)
    # pprint(resources_defsmeta)

    # Добавление всех ресурсов в проект
    set_resources = set()
    for lst in resources_defsmeta.values():
        for r in lst:
            set_resources.add(r['name'])

    project.add_resources(set_resources)

    # Добавление ресурсов в задачу
    resources_obj = project.get_resources_object()
    tasks_obj = project.get_tasks_object()
    list_tasks_obj = list(tasks_obj)
    resources = project.get_resources_name_id()

    for task_vedom in vedom:
        for task_obj in list_tasks_obj:
            name_task_vedom = list(task_vedom.keys())[0].strip()
            if task_obj.Name.strip() == name_task_vedom:
                gsn = task_vedom[name_task_vedom]['gsn']
                gsn_def = resources_defsmeta[gsn]
                if not list(task_obj.Assignments):
                    for res_name in gsn_def:

                        res_id = resources[res_name['name']]
                        res_obj = resources_obj.Item(res_id)
                        res_obj.Code = "м³"
                        res_obj.MaxUnits = res_name['consumption']
                        try:
                            task_obj.Assignments.Add(task_obj.ID, res_id)
                        except Exception as e:
                            print(e)
                    break

    project.close()


if __name__ == '__main__':
    main()
