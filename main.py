import os.path
import glob
from collections.abc import Iterable
import win32com.client

import openpyxl
import requests
from bs4 import BeautifulSoup

START_ROW_VEDOM = 5
END_ROW_VEDOM = 38

abspath = os.getcwd()
DIRECTORY = os.path.join(abspath, 'файлы')


def parce_vedom():
    vedom = glob.glob(os.path.join(DIRECTORY, '*.xlsx'))[0]
    file = openpyxl.load_workbook(vedom, data_only=True)
    sheet = file["ВОР"]

    name_work = [cell.value.strip() for cell in sheet['B'][START_ROW_VEDOM - 1:END_ROW_VEDOM] if cell.value]
    GSN = [cell.value.strip().lower().replace('гэсн', '') for cell in sheet['E'][START_ROW_VEDOM - 1:END_ROW_VEDOM]
           if cell.value]
    volume = [cell.value for cell in sheet['D'][START_ROW_VEDOM - 1:END_ROW_VEDOM] if cell.value]

    file.close()

    return [{work: {'gsn': gsn, 'volume': vol}}
            for work, gsn, vol in zip(name_work, GSN, volume)]


def parce_defsmeta(gsn):
    resources = {g: [] for g in gsn}
    temp_url = r'https://www.defsmeta.com/rgsn/gsn_{}/giesn-{}.php'

    for code in gsn:
        main_code = code.split('-')[0]
        url = temp_url.format(main_code, code)
        page = requests.get(url)
        soup = BeautifulSoup(page.text, 'lxml')

        for label, type_res in zip(
                ('РАСХОД МАТЕРИАЛОВ', 'ЭКСПЛУАТАЦИЯ МАШИН И МЕХАНИЗМОВ', 'ТРУДОЗАТРАТЫ'),
                ('material', 'machine', 'people')
        ):
            start_element = soup.find('p', string=label)
            if start_element:
                elements_between = []
                current_element = start_element
                for _ in range(3):
                    elements_between.append(current_element)
                    current_element = current_element.find_next()
                table = elements_between[2]
                rows = table.find_all('tr')
                if label == 'ТРУДОЗАТРАТЫ':
                    rows = rows[1:-len(rows) // 2]
                else:
                    rows = rows[1:-1]

                for material in rows:
                    material = material.find_all('td')
                    if label == 'ТРУДОЗАТРАТЫ':
                        material = material[1:]
                        consumption = float(material[2].text.replace(',', '.'))
                        unit = material[1].text
                        name = material[0].text
                    else:
                        material = material[1:-1]
                        consumption = float(material[3].text.replace(',', '.'))
                        unit = material[2].text
                        name = material[1].text

                    resources[code].append({
                        'name': name,
                        'consumption': consumption,
                        'type': type_res,
                        'unit': unit.replace('\xa0', '')
                    })
                    if resources[code][-1]['unit'] == 'м':
                        resources[code][-1]['unit'] = 'метры'

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

from openpyxl import Workbook

def main():
    project = MSProject(glob.glob(os.path.join(DIRECTORY, '*.mpp'))[0])
    vedom = parce_vedom()[:6]

    gsns = []
    for w in vedom:
        for g in w.values():
            gsns.append(g['gsn'])

    resources_defsmeta = parce_defsmeta(set(gsns))

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
    count_vedom = len(vedom)
    new_vedom = [['Работа', 'ГЭСН', 'Материал', 'Объем']]
    for ind, task_vedom in enumerate(vedom, start=1):
        new_vedom.append([])
        print(f'{ind}/{count_vedom}')
        for task_obj in list_tasks_obj:
            name_task_vedom = list(task_vedom.keys())[0].strip()
            if task_obj.Name.strip() == name_task_vedom:
                gsn = task_vedom[name_task_vedom]['gsn']
                gsn_def = resources_defsmeta[gsn]
                if not list(task_obj.Assignments):
                    new_vedom[-1].append(name_task_vedom)
                    new_vedom[-1].append(gsn)
                    new_vedom[-1].append(None)
                    new_vedom[-1].append(None)
                    for res_name in gsn_def:
                        new_vedom.append([None, None])
                        new_vedom[-1].append(res_name['name'])
                        new_vedom[-1].append(res_name['consumption'] * task_vedom[name_task_vedom]['volume'])
                        continue
                        res_id = resources[res_name['name']]
                        res_obj = resources_obj.Item(res_id)

                        if res_name['type'] == 'material':
                            res_obj.Type = 1
                            res_obj.MaterialLabel = res_name['unit']
                        elif res_name['type'] == 'machine':
                            res_obj.Group = "Машины"

                        elif res_name['type'] == 'people':
                            res_obj.Group = "Рабочие"

                        try:
                            task_obj.FixedDuration = True
                            t = task_obj.Assignments.Add(task_obj.ID, res_id)
                            t.Units = res_name['consumption'] * task_vedom[name_task_vedom]['volume']

                        except Exception as e:
                            print(e)
                    project._save_file()
                    break

    wb = Workbook()
    ws1 = wb.create_sheet("Материалы")
    for row in new_vedom:
        ws1.append(row)
    wb.save('Материалы.xlsx')

    project.close()


if __name__ == '__main__':
    main()
