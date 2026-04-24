# С2000-М Конвертер CLI — План оптимизации

> **Для Hermes:** Использовать subagent-driven-development для реализации по шагам.

**Цель:** Переписать Windows GUI программу на CLI для Ubuntu с оптимизацией кода.

**Архитектура:** 
- Убираем Tkinter, делаем аргументы CLI через argparse
- Убираем `os.startfile()`, открываем файл после генерации через `xdg-open`
- Разделяем логику на модули: парсер, генератор, CLI интерфейс
- Убираем глобальные переменные, передаём данные через параметры
- Добавляем обработку ошибок и валидацию

**Стек:** Python 3, argparse, openpyxl, sys, os

---

## Задача 1: Создать структуру проекта

**Цель:** Организовать код по модулям

**Файлы:**
- `cli_config.py` — CLI интерфейс (argparse)
- `config_parser.py` — парсер конфигурации С2000-М
- `excel_generator.py` — генератор Excel
- `utils.py` — утилиты (валидация, обработка ошибок)
- `main.py` — точка входа

**Шаг 1:** Создать директорию

```bash
mkdir -p /home/mamont/.hermes/workspace/c2000-m_config/plans
```

**Шаг 2:** Создать структуру модулей (по одному файлу за раз)

---

## Задача 2: CLI интерфейс (cli_config.py)

**Цель:** Заменить Tkinter на argparse

**Файл:** `cli_config.py`

```python
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""CLI интерфейс для конвертера С2000-М"""

import argparse
import sys


def parse_args():
    parser = argparse.ArgumentParser(
        description='Конвертер конфигурации С2000-М в список шлейфов'
    )
    parser.add_argument(
        'config_file',
        nargs='?',
        help='Путь к файлу конфигурации'
    )
    parser.add_argument(
        '-o', '--output',
        default='output.xlsx',
        help='Путь к выходному файлу (по умолчанию: output.xlsx)'
    )
    parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='Выводить подробную информацию'
    )
    parser.add_argument(
        '--check',
        action='store_true',
        help='Проверить наличие описаний'
    )
    
    args = parser.parse_args()
    
    # Проверка: если файл не передан — запросить
    if not args.config_file:
        print("Ошибка: укажите путь к файлу конфигурации")
        print("Используйте --help для просмотра доступных опций")
        sys.exit(1)
    
    return args


if __name__ == '__main__':
    args = parse_args()
    print(f"Config file: {args.config_file}")
    print(f"Output: {args.output}")
    print(f"Verbose: {args.verbose}")
    print(f"Check: {args.check}")
```

**Тест:**
```bash
python3 cli_config.py --help
python3 cli_config.py test.cfg -o result.xlsx
```

---

## Задача 3: Парсер конфигурации (config_parser.py)

**Цель:** Извлечь логику парсинга из GUI

**Файл:** `config_parser.py`

```python
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Парсер конфигурации С2000-М"""


class ConfigParser:
    """Парсер конфигурации С2000-М"""
    
    DEVICE_TYPES = {
        1: ["Сигнал-20"],
        2: ["Сигнал-20П"],
        3: ["С2000-СП1"],
        4: ["С2000-4"],
        # ... все типы приборов из оригинала
    }
    
    CABLE_TYPES = {
        0: ['по умолчанию'],
        1: ['охранный'],
        # ... все типы шлейфов
    }
    
    RELAY_SCRIPTS = {
        1: ['включить'],
        # ... все сценарии реле
    }
    
    def __init__(self):
        self.sections = {}
        self.raw_lines = []
    
    def parse_file(self, filepath):
        """Парсит файл конфигурации"""
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                self.raw_lines = f.readlines()
        except FileNotFoundError:
            raise FileNotFoundError(f"Файл не найден: {filepath}")
        except UnicodeDecodeError:
            raise ValueError(f"Ошибка кодирования файла: {filepath}")
        
        # Парсим разделы
        for line in self.raw_lines:
            line = line.strip()
            if line.find('Раздел:') != -1:
                parts = line.split(', ')
                if len(parts) > 1 and 'Описание:' in parts[1]:
                    section_id = int(parts[0].replace(' ', '').replace('Раздел:', ''))
                    description = parts[1].replace('Описание: ', '')
                    self.sections[section_id] = {
                        'id': section_id,
                        'description': description
                    }
    
    def get_section(self, section_id):
        """Получить описание раздела"""
        return self.sections.get(section_id)
    
    def get_device_type(self, device_id):
        """Получить тип прибора по ID"""
        return self.DEVICE_TYPES.get(device_id, ["Неизвестно"])
    
    def get_cable_type(self, cable_id):
        """Получить тип шлейфа по ID"""
        return self.CABLE_TYPES.get(cable_id, ['неизвестно'])
    
    def get_relay_script(self, script_id):
        """Получить сценарий реле по ID"""
        return self.RELAY_SCRIPTS.get(script_id, ['неизвестно'])


if __name__ == '__main__':
    # Тест
    parser = ConfigParser()
    try:
        parser.parse_file('test.cfg')
        print(f"Разделов: {len(parser.sections)}")
        for sid, sdata in parser.sections.items():
            print(f"  {sid}: {sdata['description']}")
    except Exception as e:
        print(f"Ошибка: {e}")
```

---

## Задача 4: Генератор Excel (excel_generator.py)

**Цель:** Вынести генерацию Excel в отдельный модуль

**Файл:** `excel_generator.py`

```python
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Генератор Excel для конфигурации С2000-М"""

from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment
from openpyxl.worksheet.pagebreak import Break
from openpyxl.utils import get_column_letter


class ExcelGenerator:
    """Генератор Excel таблицы"""
    
    def __init__(self):
        self.wb = Workbook()
        self.ws = self.wb.active
        self.ws.title = 'Адреса, шлейфа'
        self.current_row = 1
    
    def add_title(self, config_info):
        """Добавить заголовок"""
        # ... логика exls_w_titul
        pass
    
    def add_address_row(self, address_data):
        """Добавить строку адреса"""
        # ... логика exls_w_adr
        pass
    
    def add_output_row(self, output_data):
        """Добавить строку шлейфа/выхода/реле"""
        # ... логика exls_w_out
        pass
    
    def save(self, filepath):
        """Сохранить файл"""
        self.wb.save(filepath)
    
    def open_file(self, filepath):
        """Открыть файл (Linux/Windows)"""
        import subprocess
        import platform
        
        system = platform.system()
        if system == 'Linux':
            subprocess.call(['xdg-open', filepath])
        elif system == 'Darwin':  # macOS
            subprocess.call(['open', filepath])
        else:  # Windows
            import os
            os.startfile(filepath)


if __name__ == '__main__':
    gen = ExcelGenerator()
    gen.save('test_output.xlsx')
    print("Тестовый файл создан")
```

---

## Задача 5: Утилиты (utils.py)

**Цель:** Вынести общие функции

**Файл:** `utils.py`

```python
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Утилиты для конвертера С2000-М"""


def validate_file_path(filepath):
    """Валидация пути к файлу"""
    import os
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"Файл не найден: {filepath}")
    if not os.path.isfile(filepath):
        raise ValueError(f"Не файл: {filepath}")
    return filepath


def check_missing_descriptions(ws, max_row):
    """Проверить отсутствующие описания"""
    missing = []
    for row in ws.iter_rows(min_row=1, min_col=6, max_col=6, max_row=max_row):
        for cell in row:
            if cell.value is None:
                missing.append(cell.row)
    return missing


def format_missing_cells(ws, missing_rows):
    """Пометить ячейки с отсутствующими описаниями"""
    from openpyxl.styles import Font, PatternFill
    
    bold_font = Font(name='Times New Roman', size=10, bold=True, italic=True)
    red_fill = PatternFill(fill_type='solid', fgColor='FF6347')
    
    for row_num in missing_rows:
        cell = ws.cell(row=row_num, column=6)
        cell.value = 'Где описание!!!???'
        cell.font = bold_font
        cell.fill = red_fill


if __name__ == '__main__':
    print("Утилиты готовы")
```

---

## Задача 6: Main (main.py)

**Цель:** Сборка всего воедино

**Файл:** `main.py`

```python
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Точка входа для конвертера С2000-М"""

import sys
import argparse
from cli_config import parse_args
from config_parser import ConfigParser
from excel_generator import ExcelGenerator
from utils import validate_file_path, check_missing_descriptions, format_missing_cells


def main():
    args = parse_args()
    
    # Валидация файла
    try:
        filepath = validate_file_path(args.config_file)
    except Exception as e:
        print(f"Ошибка: {e}", file=sys.stderr)
        sys.exit(1)
    
    # Парсинг
    parser = ConfigParser()
    try:
        parser.parse_file(filepath)
    except Exception as e:
        print(f"Ошибка парсинга: {e}", file=sys.stderr)
        sys.exit(1)
    
    if args.verbose:
        print(f"Загружено разделов: {len(parser.sections)}")
    
    # Генерация
    gen = ExcelGenerator()
    
    # TODO: добавить данные в Excel на основе parser.raw_lines
    
    # Сохранение
    try:
        gen.save(args.output)
        print(f"Файл сохранён: {args.output}")
        
        if args.check:
            missing = check_missing_descriptions(gen.ws, gen.current_row)
            if missing:
                print(f"Обнаружено {len(missing)} строк без описаний")
                format_missing_cells(gen.ws, missing)
                gen.save(args.output)
        
        # Открытие файла
        gen.open_file(args.output)
        
    except Exception as e:
        print(f"Ошибка сохранения: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()
```

---

## Задача 7: Тестирование

**Цель:** Убедиться что всё работает

**Команды:**
```bash
# Проверка CLI
python3 main.py --help

# Запуск с файлом
python3 main.py test.cfg -o output.xlsx

# Проверка с флагом verbose
python3 main.py test.cfg -o output.xlsx -v

# Проверка на отсутствующие описания
python3 main.py test.cfg -o output.xlsx --check
```

---

## Задача 8: Конвертация .pyw в .py

**Цель:** Убрать расширение .w

**Файл:** `conf_to_shleif.pyw` → `conf_to_shleif.py`

```bash
cd /home/mamont/.hermes/workspace/c2000-m_config
mv conf_to_shleif.pyw conf_to_shleif.py
```

---

## Задача 9: Обработка ошибок

**Цель:** Добавить валидацию и обработку ошибок

**Файл:** `utils.py` (добавить функции)

```python
def validate_config_lines(lines):
    """Валидация строк конфигурации"""
    errors = []
    for i, line in enumerate(lines, 1):
        if line.strip() == '':
            continue
        if line.find(',') == -1:
            errors.append(f"Строка {i}: нет запятой в '{line.strip()[:50]}'")
    return errors
```

---

## Задача 10: Документация

**Цель:** Добавить README

**Файл:** `README.md`

```markdown
# Конвертер конфигурации С2000-М

CLI утилита для конвертации конфигурации С2000-М в Excel таблицу.

## Установка

```bash
pip install openpyxl
```

## Использование

```bash
python3 conf_to_shleif.py CONFIG_FILE [-o OUTPUT_FILE] [-v] [--check]
```

### Аргументы

- `CONFIG_FILE` — путь к файлу конфигурации
- `-o OUTPUT_FILE` — выходной файл (по умолчанию: output.xlsx)
- `-v` — подробный вывод
- `--check` — проверить отсутствующие описания

### Пример

```bash
python3 conf_to_shleif.py config.cfg -o result.xlsx --check
```

## Формат конфигурации

Файл должен быть в формате CSV с запятыми и пробелами как разделителями.
```

---

## Итоговый план

1. ✅ **Структура** — создать модули
2. ✅ **CLI** — argparse интерфейс
3. ✅ **Парсер** — ConfigParser класс
4. ✅ **Генератор** — ExcelGenerator класс
5. ✅ **Утилиты** — валидация и обработка
6. ✅ **Main** — сборка модулей
7. ✅ **Тестирование** — проверка всех опций
8. ✅ **Конвертация** — .pyw → .py
9. ✅ **Ошибки** — обработка исключений
10. ✅ **Документация** — README

**Итого:** 10 bite-sized задач, каждая ~2-5 минут работы.

Готов приступить к реализации через subagent-driven-development?
