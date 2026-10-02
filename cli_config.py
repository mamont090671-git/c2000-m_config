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
    
    parser.add_argument(
        '--format',
        choices=['xlsx', 'md'],
        default='xlsx',
        help='Формат вывода: xlsx или md (по умолчанию: xlsx)'
    )
    parser.add_argument(
        '--no-open',
        action='store_true',
        help='Не открывать выходной файл после генерации'
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
