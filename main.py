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
    
    # Парсинг конфигурации
    parser = ConfigParser()
    try:
        parser.parse_file(filepath)
        if args.verbose:
            print(f"Загружено разделов: {len(parser.sections)}")
    except FileNotFoundError as e:
        print(f"Ошибка: {e}", file=sys.stderr)
        sys.exit(1)
    except ValueError as e:
        print(f"Ошибка кодирования: {e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"Ошибка парсинга: {e}", file=sys.stderr)
        sys.exit(1)
    
    # Генерация Excel
    gen = ExcelGenerator()
    gen.set_config(parser)
    
    # Обработка строк конфигурации
    current_row = 1
    try:
        for string_f in parser.raw_lines:
            if args.verbose:
                print(f"Обрабатываем строку: {string_f[:50]}...")
            
            # Пропускаем пустые строки
            if string_f.find('\n') != -1:
                string_f = string_f.replace('\n', '').strip()
            
            if not string_f:
                continue
            
            # Преобразуем строку в массив
            str_array = string_f.split(', ')
            
            # Нормализация полей
            j = 0
            for h in str_array:
                if h.find('Описание:') != -1:
                    str_array[j] = h[h.find('Описание:'):].strip()
                elif h.find('Шлейф:') != -1:
                    str_array[j] = h[h.find('Шлейф:'):].strip()
                elif h.find('Раздел:') != -1:
                    str_array[j] = h[h.find('Раздел:'):].strip()
                elif h.find('Время') != -1:
                    str_array[j] = h[h.find('Время'):].strip()
                elif h.find('Выход:') != -1:
                    str_array[j] = h[h.find('Выход:'):].strip()
                elif h.find('Реле:') != -1:
                    str_array[j] = h[h.find('Реле:'):].strip()
                j += 1
            
            # Обработка типа строки
            if str_array and str_array[0].find('Конфигурация') != -1:
                current_row = gen.add_title(str_array, current_row)
            elif str_array and str_array[0].find('Версия:') != -1:
                current_row = gen.add_title(str_array, current_row)
            elif len(str_array) >= 3 and str_array[0].find('Адрес:') != -1:
                current_row = gen.add_address_row(str_array, current_row)
            elif len(str_array) >= 3 and str_array[0].find('Шлейф:') != -1:
                current_row = gen.add_output_row(str_array, current_row)
            elif len(str_array) >= 2 and str_array[0].find('Выход:') != -1:
                current_row = gen.add_output_row(str_array, current_row)
            elif len(str_array) >= 2 and str_array[0].find('Реле:') != -1:
                current_row = gen.add_output_row(str_array, current_row)
            
            if args.verbose:
                print(f"  Текущая строка: {current_row}")
        
    except Exception as e:
        print(f"Ошибка обработки: {e}", file=sys.stderr)
        sys.exit(1)
    
    # Сохранение файла
    try:
        gen.save(args.output)
        print(f"Файл сохранён: {args.output}")
        
        # Проверка отсутствующих описаний
        if args.check:
            missing = check_missing_descriptions(gen.ws, current_row - 1)
            if missing:
                print(f"Обнаружено {len(missing)} строк без описаний")
                format_missing_cells(gen.ws, missing)
                gen.save(args.output)
                print(f"Обновлённый файл: {args.output}")
        
        # Настройка страницы
        gen.set_printer_settings(gen.ws.PAPERSIZE_A4, gen.ws.ORIENTATION_PORTRAIT)
        gen.set_page_breaks()
        gen.set_view_mode('pageBreakPreview')
        gen.hide_columns(['E'])
        gen.set_columns_width()
        
        # Открытие файла
        gen.open_file(args.output)
        
    except Exception as e:
        print(f"Ошибка сохранения: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()
