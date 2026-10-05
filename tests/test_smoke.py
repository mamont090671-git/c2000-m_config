#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Smoke-тесты конвертера С2000-М. Запуск: python3 tests/test_smoke.py"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
TEST_CFG = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'test.cfg')

from config_parser import ConfigParser


def test_parse_test_cfg():
    parser = ConfigParser()
    devices, _ = parser.parse_file(TEST_CFG)

    # 3 прибора в [Приборы]; «Адрес: 9» после [Привязка управления] не прибор
    assert len(devices) == 3, f"Ожидалось 3 прибора, получено {len(devices)}"

    dev1, dev2, dev3 = devices

    # Типы приборов из [Типы_приборов] файла
    assert dev1.type_name == 'С2000-КПБ', dev1.type_name
    assert dev2.type_name == 'С2000-КДЛ', dev2.type_name
    # Неизвестный id — fallback из словаря/«Неизвестно»
    assert dev3.type_name != '', dev3.type_name

    # Поля dev1
    assert dev1.version == '2.01'
    assert dev1.description == 'Тестовый КПБ'
    assert len(dev1.shleifs) == 1
    assert len(dev1.outputs) == 2
    assert dev1.outputs[0]['description'] == 'Сирена'
    assert len(dev1.relays) == 1
    assert dev1.relays[0]['program_id'] == 1
    assert dev1.relays[0]['section_id'] == 1  # Раздел: на следующей строке

    # dev2: шлейфы с inline-полями
    assert len(dev2.shleifs) == 2
    assert dev2.shleifs[0]['section_id'] == 1
    assert dev2.shleifs[0]['description'] == 'Шлейф тест 1'
    assert len(dev2.readers) == 1

    # Разделы из [Разделы]
    assert parser.sections.get(1, {}).get('description') == 'Раздел один'
    assert parser.sections.get(2, {}).get('description') == 'Раздел два'


def test_sections_fallback_outside_block():
    """Файл без [Разделы]: строки 'Раздел: N, Описание: X' парсятся как fallback."""
    parser = ConfigParser()
    devices, _ = parser.parse_file(TEST_CFG)
    # Файл содержит строку 'Раздел: 7, Описание: ...' вне [Разделы] — fallback
    assert parser.sections.get(7, {}).get('description') == 'Свободный раздел'


def test_md_generation():
    from md_generator import generate_md
    parser = ConfigParser()
    devices, _ = parser.parse_file(TEST_CFG)
    md = generate_md(devices, parser.sections)
    assert '**1**' in md
    assert 'С2000-КПБ' in md
    assert 'Сирена' in md
    # Шлейф без раздела и описания не попадает в md
    assert 'Шл. 1 — ' in md


def test_excel_generation():
    from excel_generator import ExcelGenerator
    from utils import check_missing_descriptions

    parser = ConfigParser()
    devices, _ = parser.parse_file(TEST_CFG)

    gen = ExcelGenerator()
    gen.set_config(parser)
    gen.set_sections(parser.sections)
    gen.set_devices(devices)
    current_row = gen.generate(devices)

    # Стендаун-выходы: dev1 Выход:2 без описания/раздела — пропускается;
    # Выход:1 с описанием совпал с Реле:1 → consumed. Стендаун — dev2: нет.
    # Проверяем, что реле с программой отобразилось
    b_col = [gen.ws.cell(row=r, column=2).value for r in range(1, current_row)]
    assert 'Рел.1' in b_col, f"Реле не найдено: {b_col}"

    # check_missing_descriptions: только строки данных с пустым F
    missing = check_missing_descriptions(gen.ws, current_row - 1)
    for r in missing:
        a = gen.ws.cell(row=r, column=1).value
        b = gen.ws.cell(row=r, column=2).value
        f = gen.ws.cell(row=r, column=6).value
        assert not isinstance(a, int), f"Красится строка-заголовок {r}"
        assert b not in (None, ''), f"Красится разделитель {r}"
        assert f in (None, ''), f"Строка {r} не пустая в F"

    # Граница — чёрная непрозрачная
    c = gen.ws.cell(row=1, column=1)
    rgb = c.border.top.color.rgb if c.border.top and c.border.top.color else None
    assert rgb == 'FF000000', f"Цвет границы: {rgb}"


if __name__ == '__main__':
    failures = 0
    for name, fn in sorted(globals().items()):
        if name.startswith('test_') and callable(fn):
            try:
                fn()
                print(f"PASS {name}")
            except AssertionError as e:
                failures += 1
                print(f"FAIL {name}: {e}")
    sys.exit(1 if failures else 0)
