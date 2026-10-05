#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Точка входа для конвертера С2000-М"""

import sys
import os

# Fix: кастомный Python 3.14 ищет в Python 3.11 venv первым — там бинарно несовместимо.
# Добавляем Python 3.14 site-packages до импорта любых библиотек.
_hermes_tools = os.path.join(os.environ.get('HOME', '/home/mamont'), '.hermes/tools')
if '/.hermes/tools/python-3.14' in sys.executable:
    py314_site = os.path.join(_hermes_tools, 'python-3.14.7+202*/lib/python3.14/site-packages')
    import glob
    for d in glob.glob(py314_site):
        if d not in sys.path:
            sys.path.insert(0, d)

# Fix: системный Python 3.12 — user site-packages
if '/usr/bin/python3' == sys.executable or 'python3.12' in sys.executable.lower():
    user_site = os.path.join(os.path.expanduser('~'), '.local/lib/python3.12/site-packages')
    if not os.path.exists(user_site):
        user_site = '/usr/lib/python3/dist-packages'
    if user_site not in sys.path:
        sys.path.insert(0, user_site)

import argparse
from cli_config import parse_args
from config_parser import ConfigParser
from md_generator import generate_md
from utils import validate_file_path, check_missing_descriptions, format_missing_cells, backup_file


def main():
    """Основная функция конвертера"""
    
    # Парсинг аргументов
    try:
        args = parse_args()
    except SystemExit as e:
        sys.exit(e.code)
    
    if args.verbose:
        print(f"Параметры: файл={args.config_file}, выход={args.output}")
    
    # Валидация файла
    try:
        filepath = validate_file_path(args.config_file)
        if args.verbose:
            print(f"Файл валиден: {filepath}")
    except Exception as e:
        print(f"Ошибка: {e}", file=sys.stderr)
        sys.exit(1)
    
    # Парсинг конфигурации — теперь ConfigParser строит структурированную модель
    parser = ConfigParser()
    try:
        devices, raw_lines = parser.parse_file(filepath)
        if args.verbose:
            print(f"Загружено разделов: {len(parser.sections)}")
            print(f"Приборов: {len(devices)}")
    except FileNotFoundError as e:
        print(f"Ошибка: {e}", file=sys.stderr)
        sys.exit(1)
    except ValueError as e:
        print(f"Ошибка кодирования: {e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"Ошибка парсинга: {e}", file=sys.stderr)
        sys.exit(1)
    
    # Генерация в зависимости от формата
    if args.format == 'xlsx':
        _generate_xlsx(args, parser, devices)
    else:
        output = args.output if args.output.endswith('.md') else args.output + '.md'
        result = generate_md(devices, parser.sections, output)
        print(f"Сгенерировано {len(devices)} приборов в {output}")


def _generate_xlsx(args, parser, devices):
    """Excel-генерация (вынесена в отдельную функцию)"""
    from excel_generator import ExcelGenerator
    gen = ExcelGenerator()
    gen.set_config(parser)
    gen.set_sections(parser.sections)
    gen.set_devices(devices)
    
    # Генерация Excel из структурированных данных
    try:
        current_row = gen.generate(devices)
    except Exception as e:
        print(f"Ошибка генерации: {e}", file=sys.stderr)
        sys.exit(1)
    
    # Настройка страницы — ДО сохранения
    gen.set_printer_settings(gen.ws.PAPERSIZE_A4, gen.ws.ORIENTATION_PORTRAIT)
    gen.set_page_breaks()
    gen.set_view_mode('pageBreakPreview')
    gen.hide_columns(['E'])
    gen.set_columns_width()

    # Проверка отсутствующих описаний (до сохранения, чтобы не резать файл дважды)
    if args.check:
        missing = check_missing_descriptions(gen.ws, current_row - 1)
        if missing:
            print(f"Обнаружено {len(missing)} строк без описаний")
            format_missing_cells(gen.ws, missing)
        # Бэкап до перезаписи существующего файла
        bak = backup_file(args.output)
        if bak:
            print(f"Бэкап: {bak}")

    # Сохранение файла
    try:
        gen.save(args.output)
        print(f"Файл сохранён: {args.output}")
    except Exception as e:
        print(f"Ошибка сохранения: {e}", file=sys.stderr)
        sys.exit(1)

    # Открытие файла
    if not args.no_open:
        gen.open_file(args.output)



if __name__ == '__main__':
    main()
