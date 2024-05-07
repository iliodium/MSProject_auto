## Требования

1. Установите requirements.txt в виртуальное окружение
2. Python 3.10.

> Примечание: Убедитесь, что ваше виртуальное окружение активировано перед выполнением этой команды.
```bash
pip install -r requirements.txt
```
## Установка проекта с использованием Git BASH (для новичков)

1. Установите [Git BASH](https://gitforwindows.org/).
2. Установите [Python 3.10](https://www.python.org/downloads/release/python-3100/).
3. Откройте Git BASH в каталоге, куда вы хотите установить проект, щелкнув правой кнопкой мыши и выбрав "Open Git BASH here".
4. Выполните следующие команды поочередно:

```bash
git clone https://github.com/iliodium/MSProject_auto.git
py -3.10 -m venv venv310
source venv310/Scripts/activate
pip install -r requirements.txt
```

## Запуск проекта
## Ведомость и MSProject

В папке "файлы" должны находиться 2 файла, один с раширением ".xlsx" (ведомость), другой с расширением
".mpp" (MSProject), называться они могут как угодно.
Формат ведомости должен совпадать с ведомостью из примера, а в MSProject
должны быть вставлены все работы.

## Настройка номеров строк

Для настройки номеров строк первой и последней работы, откройте файл `main.py` и измените следующие строки (10 и 11):

```python
START_ROW_VEDOM = 5  # Номер строки первой работы в файле "ведомость финал.xlsx"
END_ROW_VEDOM = 38   # Номер строки последней работы в файле "ведомость финал.xlsx"
```

Обратите внимание, что номер строки первой работы и номер строки последней работы должны быть установлены в соответствии с образцом в файле "ведомость финал.xlsx" в папке "файлы".
```

Пожалуйста, дайте мне знать, если нужно что-то еще или если есть другие вопросы!