<img src="https://r2cdn.perplexity.ai/pplx-full-logo-primary-dark%402x.png" class="logo" width="120"/>

# 🚀 Развертывание TDD-Suite на Windows

Подробная инструкция по установке и запуску системы TDD-Suite v3.0 на операционной системе Windows.

## 📋 Предварительные требования

### Системные требования

- **Windows 10/11** (или Windows 7 SP1+)
- **Python 3.8+**
- **Минимум 2 ГБ свободного места** на диске
- **Права администратора** для установки Python (при первой установке)


## 🔧 Пошаговая установка

### Шаг 1: Установка Python

#### 1.1 Скачивание Python

1. Перейдите на официальный сайт [python.org](https://www.python.org/downloads/)
2. Скачайте **Python 3.8+** для Windows
3. Выберите версию **x64** для 64-битной Windows

#### 1.2 Установка Python

1. **Запустите установщик** от имени администратора
2. **Обязательно поставьте галочку** "Add Python to PATH"[^1]
3. Выберите **"Install Now"**
4. Дождитесь завершения установки

#### 1.3 Проверка установки

Откройте **Командную строку** (cmd) и проверьте:

```cmd
python --version
```

Должно показать что-то вроде: `Python 3.8.10`[^2]

### Шаг 2: Создание виртуального окружения

#### 2.1 Откройте командную строку

- Нажмите **Win + R**
- Введите `cmd` и нажмите Enter[^3]
- Или найдите "Command Prompt" в меню Пуск


#### 2.2 Создайте папку проекта

```cmd
mkdir TDD-Suite
cd TDD-Suite
```


#### 2.3 Создайте виртуальное окружение

```cmd
python -m venv venv
```

Это создаст папку `venv` с изолированной средой Python[^4][^5].

#### 2.4 Активируйте виртуальное окружение

```cmd
venv\Scripts\activate
```

После активации в командной строке появится префикс `(venv)`[^5].

### Шаг 3: Установка зависимостей

#### 3.1 Обновите pip

```cmd
python -m pip install --upgrade pip
```


#### 3.2 Установите необходимые библиотеки

```cmd
pip install pandas>=2.1.0
pip install openpyxl>=3.1.0
pip install xlsxwriter>=3.0.0
```

Или создайте файл `requirements.txt` и установите все сразу:

**requirements.txt:**

```
pandas>=2.1.0
openpyxl>=3.1.0
xlsxwriter>=3.0.0
```

Затем выполните:

```cmd
pip install -r requirements.txt
```


### Шаг 4: Создание структуры проекта

#### 4.1 Создайте структуру папок

```cmd
mkdir config core processors cli data\input data\output data\intermediate data\archive tests
```


#### 4.2 Создайте __init__.py файлы

```cmd
type nul > config\__init__.py
type nul > core\__init__.py
type nul > processors\__init__.py
type nul > cli\__init__.py
type nul > tests\__init__.py
```


#### 4.3 Итоговая структура

```
TDD-Suite/
├── venv/                       # Виртуальное окружение
├── main.py                     # Главный файл запуска
├── requirements.txt            # Зависимости
├── config/
│   ├── __init__.py
│   └── settings.py
├── core/
│   ├── __init__.py
│   ├── base_processor.py
│   ├── excel_utils.py
│   └── file_manager.py
├── processors/
│   ├── __init__.py
│   ├── compressor.py
│   ├── cleaner.py
│   ├── merger.py
│   └── article_filler.py
├── cli/
│   ├── __init__.py
│   └── menu.py
└── data/
    ├── input/                  # Входные Excel файлы
    ├── output/                 # Обработанные файлы
    ├── intermediate/           # Промежуточные результаты
    └── archive/                # Архив
```


### Шаг 5: Размещение файлов кода

Скопируйте все файлы TDD-Suite (из предыдущего диалога) в соответствующие папки согласно структуре выше.

## 🚀 Запуск системы

### Способ 1: Через командную строку

#### 1. Активируйте виртуальное окружение

```cmd
cd TDD-Suite
venv\Scripts\activate
```


#### 2. Запустите главный скрипт

```cmd
python main.py
```


### Способ 2: Создание bat-файла для удобства

Создайте файл `run_tdd.bat` в папке TDD-Suite:

```batch
@echo off
cd /d "%~dp0"
call venv\Scripts\activate
python main.py
pause
```

Теперь можно запускать систему двойным кликом по файлу `run_tdd.bat`.

## 📂 Подготовка данных

### Размещение Excel файлов

1. Поместите ваши Excel файлы в папку `data\input\`
2. Система автоматически найдет их при запуске

### Пример использования

```cmd
TDD-Suite\data\input\
├── Спецификация_Дизель.xlsx
├── Replay_OF_72.xlsx
└── IP_Filimonova_weights.xlsx
```


## 🎯 Использование системы

После запуска появится меню:

```
🚀 TDD-Suite v3.0 - Комплексная обработка Excel файлов
======================================================================
1. 📊 Сжатие спецификаций
2. 🧹 Очистка файлов (апострофы, D→.)
3. 🔗 Объединение листов PL INV
4. 📝 Заполнение артикулов
5. 📦 Пакетная обработка
6. ⚙️  Настройки
7. 📊 Статистика
8. ❌ Выход
```


## 🔧 Устранение проблем

### Проблема: "python is not recognized"

**Решение:** Python не добавлен в PATH[^2]

1. Переустановите Python с галочкой "Add Python to PATH"
2. Или добавьте вручную путь к Python в переменные среды

### Проблема: Ошибки при установке пакетов

**Решение:**

```cmd
python -m pip install --upgrade pip
pip install --user pandas openpyxl xlsxwriter
```


### Проблема: Кодировка в командной строке

**Решение:** Установите UTF-8 кодировку[^6]

```cmd
chcp 65001
```


### Проблема: Не хватает прав доступа

**Решение:** Запустите командную строку от имени администратора[^7]

## 📋 Дополнительные рекомендации

### Для продвинутых пользователей

- Используйте **Windows Terminal** вместо стандартной cmd
- Установите **Git for Windows** для лучшего опыта работы с командной строкой[^8]
- Рассмотрите использование **PyCharm** или **VS Code** для разработки


### Автоматизация

Создайте ярлык на рабочем столе с командой:

```
C:\Windows\System32\cmd.exe /k "cd /d C:\TDD-Suite && venv\Scripts\activate"
```


### Обновление зависимостей

Периодически обновляйте библиотеки:

```cmd
pip install --upgrade pandas openpyxl xlsxwriter
```


## ✅ Готово к использованию!

После выполнения всех шагов система TDD-Suite полностью готова к работе на Windows. Система предоставляет:

- **Сжатие спецификаций** с опциональным расчетом долей в коробках
- **Очистку файлов** от апострофов и замену символов
- **Объединение листов** PL INV в сводные таблицы
- **Заполнение артикулов** на основе эталонных данных
- **Пакетную обработку** множества файлов одновременно

Просто поместите Excel файлы в папку `data\input\`, запустите `python main.py` и выберите нужную операцию в интерактивном меню!

<div style="text-align: center">⁂</div>

[^1]: https://docs.python.org/3/faq/windows.html

[^2]: https://stackoverflow.com/questions/4621255/how-do-i-run-a-python-program-in-the-command-prompt-in-windows-7

[^3]: https://www.madpenguin.org/how-to-run-a-python-script-in-windows/

[^4]: https://docs.python.org/3/library/venv.html

[^5]: https://python.plainenglish.io/setting-up-virtualenv-in-windows-996005c4fa6c?gi=a86c11e2a530

[^6]: https://www.geeksforgeeks.org/how-to-install-openpyxl-in-python-on-windows/

[^7]: https://www.geeksforgeeks.org/installation-guide/how-to-install-openpyxl-in-python-on-windows/

[^8]: https://stackoverflow.com/questions/38364404/how-to-install-openpyxl-with-pip

[^9]: https://www.cambridge.org/core/product/identifier/CBO9781139871754A011/type/book_part

[^10]: https://www.semanticscholar.org/paper/ede6d701188e18edb6951ed7ccc4ed25717a764e

[^11]: https://www.semanticscholar.org/paper/b11ecaadb4ffe8fbdea1752bc1bd6bee16d04cb1

[^12]: https://www.semanticscholar.org/paper/b08e48c58f4fe0c305d22466f14a4157737dc41e

[^13]: https://www.semanticscholar.org/paper/d8d864a520ce06377cf6a1259b585f6f13b628e0

[^14]: https://www.semanticscholar.org/paper/a146906469c3beb7b37b70e8b83241b3ceba58d1

[^15]: http://dpi-journals.com/index.php/dtmse/article/view/6924

[^16]: https://www.nature.com/articles/s41592-022-01442-1

[^17]: https://it.engineering.oregonstate.edu/setting-virtual-environments-python

[^18]: https://packaging.python.org/guides/installing-using-pip-and-virtual-environments/

[^19]: https://virtualenv.pypa.io/en/latest/installation.html

[^20]: https://stackoverflow.com/questions/17737203/python-and-virtualenv-on-windows

[^21]: https://mothergeo-py.readthedocs.io/en/latest/development/how-to/venv-win.html

[^22]: https://docs.python.org/id/3.10/faq/windows.html

[^23]: https://sdk.teamwire.eu/api_guide/virtualenv_setup/

[^24]: https://realpython.com/run-python-scripts/

[^25]: https://bmcbioinformatics.biomedcentral.com/articles/10.1186/s12859-019-2606-y

[^26]: https://www.semanticscholar.org/paper/4e4f311858573c1264e15dd00f6c90e04100bf7c

[^27]: https://pmc.ncbi.nlm.nih.gov/articles/PMC10850855/

[^28]: http://arxiv.org/pdf/2407.10317.pdf

[^29]: https://arxiv.org/pdf/2503.14443.pdf

[^30]: https://www.mdpi.com/1424-8220/24/12/3837/pdf?version=1718289543

[^31]: https://arxiv.org/pdf/1105.0061.pdf

[^32]: https://arxiv.org/pdf/2412.06294v1.pdf

[^33]: https://linkinghub.elsevier.com/retrieve/pii/S2666166724005598

[^34]: https://arxiv.org/pdf/2502.13681.pdf

[^35]: https://arxiv.org/pdf/2206.10558.pdf

[^36]: https://pmc.ncbi.nlm.nih.gov/articles/PMC9797609/

[^37]: https://www.reddit.com/r/learnpython/comments/p2b76c/installing_modules_pandas_openpyxl/

[^38]: https://docs.python-guide.org/dev/virtualenvs/

[^39]: https://pypi.org/project/openpyxl/

[^40]: https://www.w3schools.com/python/python_virtualenv.asp

[^41]: https://stackoverflow.com/questions/43452604/running-python-scripts-through-the-windows-command-line

