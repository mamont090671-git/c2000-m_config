#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Парсер конфигурации С2000-М"""

import os
import re
from dataclasses import dataclass, field


@dataclass
class Device:
    """Структурированный прибор"""
    addr: int = 0
    type_id: int = 0
    type_name: str = "Неизвестно"
    version: str = ""
    description: str = ""
    shleifs: list = field(default_factory=list)
    relays: list = field(default_factory=list)
    outputs: list = field(default_factory=list)
    readers: list = field(default_factory=list)


class ConfigParser:
    """Парсер конфигурации С2000-М"""

    DEVICE_TYPES = {
        1: "Сигнал-20", 2: "Сигнал-20П", 3: "С2000-СП1", 4: "С2000-4",
        7: "С2000-К", 8: "С2000-ИТ", 9: "С2000-КДЛ", 10: "С2000-БИ/БКИ",
        11: "Сигнал-20(вер. 02)", 13: "С2000-КС", 14: "С2000-АСПТ",
        15: "С2000-КПБ", 16: "С2000-2", 19: "УО-ОРИОН", 20: "Рупор",
        21: "Рупор-Диспетчер исп.01", 22: "С2000-ПТ", 24: "УО-4С",
        25: "Поток-3Н", 26: "Сигнал-20М", 28: "С2000-БИ-01",
        30: "Рупор-01", 31: "С2000-Adem",
        32: "РИП-12 исп.50, исп.51, без исполнения",
        33: "РИП-12 исп.50, исп.51, без исполнения", 34: "Сигнал-10",
        35: "РИП-12 исп.54", 36: "С2000-ПП",
        37: "РИП-24 исп.50, исп.51", 38: "РИП-12 исп.54",
        39: "РИП-24 исп.50, исп.51", 41: "С2000-КДЛ-2И",
        43: "С2000-PGE", 44: "С2000-БКИ", 45: "Поток-БКИ",
        46: "Рупор-200", 47: "С2000-Периметр", 48: "МИП-12",
        49: "МИП-24", 53: "РИП-48 исп.01", 54: "РИП-12 исп.56",
        55: "РИП-24 исп.56", 59: "Рупор исп.02", 61: "С2000-КДЛ-Modbus",
        66: "Рупор исп.03", 67: "Рупор-300", 76: "С2000-PGE исп.01"
    }
    UNKNOWN_DEVICE = "Неизвестно"

    CABLE_TYPES = {
        0: "по умолчанию", 1: "охранный", 2: "пожарный", 3: "тревожный",
        4: "технологический", 5: "входной",
        6: "адресно-аналоговый дымовой", 7: "адресно-аналоговый тепловой",
        8: "8 тип - хз", 9: "цепь ДС дверей", 10: "ручной пуск",
        11: "дистанционный пуск", 12: "состояние автоматики"
    }
    UNKNOWN_CABLE = "неизвестно"

    RELAY_SCRIPTS = {
        1: "включить", 2: "выключить", 3: "включить на время",
        4: "выключить на время",
        5: "мигать из состояния выключено", 6: "мигать из состояния включено",
        7: "мигать из состояния выключено на время",
        8: "мигать из состояния включено на время",
        9: "лампа", 10: "пцн", 11: "аспт", 12: "нет"
    }
    UNKNOWN_RELAY_SCRIPT = "неизвестно"

    CABLE_SCRIPTS = {
        1: "Нет", 2: "Снять шлейф", 3: "Взять шлейф", 4: "Сбросить тревогу",
        5: "Откл. Автоматику", 6: "Вкл. Автоматику",
        7: "Отменить пуск АУП", 8: "Запустить АУП",
        9: "Вкл. режим тестирования", 10: "Откл. режим тестирования"
    }
    UNKNOWN_CABLE_SCRIPT = "неизвестно"

    RELAY = {
        1: "включить", 2: "выключить", 3: "включить на время",
        4: "выключить на время",
        5: "мигать из состояния выключено", 6: "мигать из состояния включено",
        7: "мигать из состояния выключено", 8: "мигать из состояния включено",
        9: "лампа", 10: "пцн", 11: "аспт", 12: "сирена", 13: "пожарный пцн",
        14: "выход неисправность", 15: "пожарная лампа", 16: "старая тактика пцн",
        17: "включить на время перед взятием",
        18: "выключить на время перед взятием",
        19: "включить на время при взятии",
        20: "выключить на время при взятии",
        21: "включить на время при снятии",
        22: "включить на время при снятии",
        23: "включить на время при невзятии",
        24: "выключить на время при невзятии",
        25: "включить на время при нарушении",
        26: "включить на время при нарушении",
        27: "включить при снятии", 28: "выключить при снятии",
        29: "включить при взятии", 30: "включить при взятии",
        31: "включить при нарушении тех. шлейфа",
        32: "включить при нарушении тех. шлейфа",
        33: "аспт-1", 34: "аспт-а", 35: "аспт-а1",
        36: "включить при повышении температуры",
        37: "выключить при повышении температуры",
        38: "включить при задержке пуска", 39: "включить при пуске",
        40: "включить при тушении", 41: "включить при неудачном пуске",
        42: "включить при включении автоматики",
        43: "выключить при включении автоматики",
        44: "включить при выключении автоматики",
        45: "выключить при выключении автоматики",
        46: "вкл. если исп. устройство в рабочем состоянии",
        47: "вык. если исп. устройство в рабочем состоянии",
        48: "вкл. если исп. устройство в исходном состоянии",
        49: "вык. если исп. устройство в исходном состоянии",
        50: "включить при пожар2", 51: "выключить при пожар2",
        52: "мигать при пожар2; исх. сост. выключено",
        53: "мигать при пожар2; исх. сост. включено",
        54: "включить при нападении", 55: "выключить при нападении",
        56: "лампа 2", 57: "сирена 2"
    }
    UNKNOWN_RELAY = "неизвестно"

    CABLE_MASKS = {
        1: "****************", 2: "----------------",
        3: "********--------", 4: "--------********",
        5: "****----****----", 6: "----****----****",
        7: "****------------", 8: "----************",
        9: "**--**--**--**--", 10: "--**--**--**--**",
        11: "**----**----**--", 12: "--****--****--**",
        13: "**--------------", 14: "--**************",
        15: "**--**__********", 16: "--**--**--------",
        17: "*-*-*-*-*-*-*-*-", 18: "*-*-*-*-*-*-*-*",
        19: "*--*--*--*--*--*", 20: "-**-**-**-**-**-",
        21: "*-------*-------", 22: "-*******-*******",
        23: "*---------------", 24: "-***************",
        25: "*-*-----*-*-----", 26: "-*-**-**-**-**-",
        27: "*-*-------------", 28: "-*-*************",
        29: "*-*-*-----------", 30: "-*-*-***********"
    }
    UNKNOWN_MASK = "неизвестно"

    def __init__(self):
        self.sections = {}
        self.raw_lines = []

    def parse_file(self, filepath):
        """Парсит файл конфигурации, возвращая (devices, raw_lines)"""
        if not os.path.exists(filepath):
            raise FileNotFoundError(f"Файл не найден: {filepath}")
        if not os.path.isfile(filepath):
            raise ValueError(f"Не файл: {filepath}")

        encodings = ['utf-8', 'cp1251', 'iso-8859-1', 'windows-1251']
        encoding_found = None

        for encoding in encodings:
            try:
                with open(filepath, 'r', encoding=encoding) as f:
                    self.raw_lines = f.readlines()
                encoding_found = encoding
                if encoding != 'utf-8':
                    print(f"Внимание: файл открыт с кодировкой {encoding}")
                break
            except UnicodeDecodeError:
                continue

        if not encoding_found:
            raise UnicodeDecodeError(f"Не удалось определить кодировку файла: {filepath}")

        # Парсим разделы:
        # 1) [Разделы] — нормальный источник (Описание в кавычках)
        # 2) остальные строки с 'Раздел:' — fallback для аномальных файлов
        in_sections = False
        for line in self.raw_lines:
            line = line.strip()
            if not line:
                continue
            if line == '[Разделы]':
                in_sections = True
                continue
            if line.startswith('[') and in_sections:
                in_sections = False
                continue
            if in_sections and 'Раздел:' in line:
                parts = line.split(',', 1)
                if len(parts) >= 2:
                    try:
                        section_id = int(parts[0].replace(' ', '').replace('Раздел:', ''))
                        description = parts[1].replace('Описание: ', '').replace('Описание:', '').strip()
                        if description.startswith('"') and description.endswith('"'):
                            description = description[1:-1]
                        self.sections[section_id] = {
                            'id': section_id,
                            'description': description
                        }
                    except (ValueError, IndexError):
                        pass
            elif not in_sections and 'Раздел:' in line and 'Описание:' in line:
                parts = line.split(',', 1)
                if len(parts) >= 2:
                    try:
                        section_id = int(parts[0].replace(' ', '').replace('Раздел:', ''))
                        description = parts[1].split('Описание:', 1)[1].strip().strip('"')
                        self.sections[section_id] = {
                            'id': section_id,
                            'description': description
                        }
                    except (ValueError, IndexError):
                        pass

        # Типы приборов из файла ([Типы_приборов]) — источник правды;
        # словарь DEVICE_TYPES остаётся fallback для неизвестных id
        self._parse_device_types()

        # Строим структурированную модель
        devices = self._parse_devices()
        return devices, self.raw_lines

    def _parse_device_types(self):
        """Парсит [Типы_приборов]: 'Тип_прибора: N, ..., Название: "..."'"""
        in_types = False
        for line in self.raw_lines:
            line = line.strip()
            if line == '[Типы_приборов]':
                in_types = True
                continue
            if line.startswith('[') and in_types:
                in_types = False
                continue
            if not in_types or 'Тип_прибора:' not in line or 'Название:' not in line:
                continue
            try:
                type_id = int(line.split('Тип_прибора:', 1)[1].split(',', 1)[0].strip())
                name = line.split('Название:', 1)[1].strip().strip('"')
                if type_id > 0:
                    self.DEVICE_TYPES[type_id] = name
            except (ValueError, IndexError):
                pass

    def _parse_devices(self):
        """Парсит raw_lines и возвращает список приборов с подчинёнными элементами"""
        devices = []
        current_dev = None
        in_devices_section = False
        pending = None  # 'shleif', 'relay', 'output', 'reader', None
        stop_at = ('[Уровни', '[Привязка')  # секции, закрывающие [Приборы]

        for raw_line in self.raw_lines:
            line = raw_line.strip()
            if not line:
                continue

            if '[Приборы]' in line and 'Адрес' not in line:
                in_devices_section = True
                pending = None
                continue
            if line.startswith(stop_at):
                # Закрывающая секция: добавляем последний прибор и
                # продолжаем (без break) — в аномальном порядке секций
                # данные после неё не теряются молча
                if current_dev:
                    devices.append(current_dev)
                    current_dev = None
                in_devices_section = False
                continue
            if not in_devices_section:
                continue

            # Новый прибор
            addr_match = re.match(r'Адрес:\s*(\d+)', line)
            if addr_match:
                if current_dev:
                    devices.append(current_dev)
                current_dev = Device(addr=int(addr_match.group(1)))
                pending = None

                # Тип_прибора и Версия могут быть на той же строке
                type_m = re.search(r'Тип_прибора:\s*(\d+)', line)
                if type_m:
                    current_dev.type_id = int(type_m.group(1))
                    current_dev.type_name = self.DEVICE_TYPES.get(current_dev.type_id, self.UNKNOWN_DEVICE)

                ver_m = re.search(r'Версия:\s*(\d+\.\d+)', line)
                if ver_m:
                    current_dev.version = ver_m.group(1)

                desc_m = re.search(r'Описание:\s*(.+)', line)
                if desc_m:
                    current_dev.description = desc_m.group(1).strip().strip('"').strip('"')
                continue

            if current_dev is None:
                continue

            # Шлейф (сбрасывает pending)
            shleif_match = re.match(r'Шлейф:\s*(\d+)', line)
            if shleif_match:
                current_dev.shleifs.append({
                    'num': int(shleif_match.group(1)),
                    'section_id': None,
                    'type_id': None,
                    'type_name': self.UNKNOWN_CABLE,
                    'description': ''
                })
                pending = 'shleif'
                # Поле Раздел: и Тип_шлейфа: могут быть в той же строке
                sec_m = re.search(r'Раздел:\s*(\d+)', line)
                if sec_m:
                    current_dev.shleifs[-1]['section_id'] = int(sec_m.group(1))
                type_m = re.search(r'Тип_шлейфа:\s*(\d+)', line)
                if type_m:
                    t = int(type_m.group(1))
                    current_dev.shleifs[-1]['type_id'] = t
                    current_dev.shleifs[-1]['type_name'] = self.CABLE_TYPES.get(t, self.UNKNOWN_CABLE)
                # Описание в той же строке
                if 'Описание:' in line:
                    desc = line.split('Описание:', 1)[1].strip().strip('"').strip('"')
                    current_dev.shleifs[-1]['description'] = desc
                continue

            # Реле (сбрасывает pending)
            relay_match = re.match(r'Реле:\s*(\d+)', line)
            if relay_match:
                current_dev.relays.append({
                    'num': int(relay_match.group(1)),
                    'program_id': None,
                    'program_name': '',
                    'description': ''
                })
                pending = 'relay'
                # Поле Программа: и Описание: могут быть в той же строке
                prog_m = re.search(r'Программа:\s*(\d+)', line)
                if prog_m:
                    relay_id = int(prog_m.group(1))
                    current_dev.relays[-1]['program_id'] = relay_id
                    current_dev.relays[-1]['program_name'] = self.RELAY.get(relay_id, self.UNKNOWN_RELAY)
                # Описание в той же строке
                if 'Описание:' in line:
                    desc = line.split('Описание:', 1)[1].strip().strip('"').strip('"')
                    current_dev.relays[-1]['description'] = desc
                continue

            # Выход (сбрасывает pending)
            output_match = re.match(r'Выход:\s*(\d+)', line)
            if output_match:
                current_dev.outputs.append({
                    'num': int(output_match.group(1)),
                    'section_id': None,
                    'description': ''
                })
                pending = 'output'
                # Поле Раздел: может быть в той же строке
                sec_m = re.search(r'Раздел:\s*(\d+)', line)
                if sec_m:
                    current_dev.outputs[-1]['section_id'] = int(sec_m.group(1))
                # Описание в той же строке
                if 'Описание:' in line:
                    desc = line.split('Описание:', 1)[1].strip().strip('"').strip('"')
                    current_dev.outputs[-1]['description'] = desc
                continue

            # Считыватель (сбрасывает pending)
            reader_match = re.match(r'Считыватель:\s*(\d+)', line)
            if reader_match:
                current_dev.readers.append({
                    'num': int(reader_match.group(1)),
                    'section_id': None,
                    'description': ''
                })
                pending = 'reader'
                continue

            # Тип_прибора (вне адреса)
            if 'Тип_прибора:' in line:
                tm = re.search(r'Тип_прибора:\s*(\d+)', line)
                if tm:
                    current_dev.type_id = int(tm.group(1))
                    current_dev.type_name = self.DEVICE_TYPES.get(current_dev.type_id, self.UNKNOWN_DEVICE)

            # Версия (вне адреса)
            if 'Версия:' in line:
                vm = re.search(r'Версия:\s*(\d+\.\d+)', line)
                if vm:
                    current_dev.version = vm.group(1)

            # Описание прибора (вне адреса)
            if 'Описание:' in line and 'Тип_прибора:' not in line and 'Версия:' not in line:
                desc = line.split('Описание:', 1)[1].strip().strip('"').strip('"')
                current_dev.description = desc

            # Тип шлейфа
            if 'Тип_шлейфа:' in line:
                type_m = re.search(r'Тип_шлейфа:\s*(\d+)', line)
                if type_m:
                    t = int(type_m.group(1))
                    if current_dev.shleifs:
                        current_dev.shleifs[-1]['type_id'] = t
                        current_dev.shleifs[-1]['type_name'] = self.CABLE_TYPES.get(t, self.UNKNOWN_CABLE)

            # Раздел для pending элемента
            if pending == 'shleif' and 'Раздел:' in line and 'Описание:' not in line:
                sec_m = re.search(r'Раздел:\s*(\d+)', line)
                if sec_m:
                    current_dev.shleifs[-1]['section_id'] = int(sec_m.group(1))

            # Раздел для реле/выхода/считывателя
            if pending != 'shleif' and 'Раздел:' in line:
                sec_m = re.search(r'Раздел:\s*(\d+)', line)
                if sec_m:
                    sid = int(sec_m.group(1))
                    if pending == 'relay' and current_dev.relays:
                        current_dev.relays[-1]['section_id'] = sid
                    elif pending == 'output' and current_dev.outputs:
                        current_dev.outputs[-1]['section_id'] = sid
                    elif pending == 'reader' and current_dev.readers:
                        current_dev.readers[-1]['section_id'] = sid

            # Программа для реле
            if 'Программа:' in line:
                prog_m = re.search(r'Программа:\s*(\d+)', line)
                if prog_m:
                    relay_id = int(prog_m.group(1))
                    current_dev.relays[-1]['program_id'] = relay_id
                    current_dev.relays[-1]['program_name'] = self.RELAY.get(relay_id, self.UNKNOWN_RELAY)

            # Описание элемента
            if 'Описание:' in line and 'Тип_шлейфа:' not in line and 'Задержка' not in line:
                desc = line.split('Описание:', 1)[1].strip().strip('"').strip('"')
                if pending == 'shleif' and current_dev.shleifs:
                    current_dev.shleifs[-1]['description'] = desc
                elif pending == 'relay' and current_dev.relays:
                    current_dev.relays[-1]['description'] = desc
                elif pending == 'output' and current_dev.outputs:
                    current_dev.outputs[-1]['description'] = desc
                elif pending == 'reader' and current_dev.readers:
                    current_dev.readers[-1]['description'] = desc

        # Последний прибор добавляется при входе в [Уровни/[Привязка
        return devices

    def get_section(self, section_id):
        """Получить описание раздела"""
        return self.sections.get(section_id)

    def get_device_type(self, device_id: int) -> str:
        """Получить тип прибора по ID"""
        return self.DEVICE_TYPES.get(device_id, self.UNKNOWN_DEVICE)

    def get_cable_type(self, cable_id: int) -> str:
        """Получить тип шлейфа по ID"""
        return self.CABLE_TYPES.get(cable_id, self.UNKNOWN_CABLE)

    def get_relay_script(self, script_id: int) -> str:
        """Получить сценарий реле по ID"""
        return self.RELAY_SCRIPTS.get(script_id, self.UNKNOWN_RELAY_SCRIPT)

    def get_relay(self, relay_id: int) -> str:
        """Получить действие реле по ID"""
        return self.RELAY.get(relay_id, self.UNKNOWN_RELAY)

    def get_cable_script(self, script_id: int) -> str:
        """Получить сценарий шлейфа по ID"""
        return self.CABLE_SCRIPTS.get(script_id, self.UNKNOWN_CABLE_SCRIPT)

    def get_mask(self, mask_id: int) -> str:
        """Получить маску мигания по ID"""
        return self.CABLE_MASKS.get(mask_id, self.UNKNOWN_MASK)


if __name__ == '__main__':
    # Тест
    parser = ConfigParser()
    try:
        devices, raw_lines = parser.parse_file('test.cfg')
        print(f"Разделов: {len(parser.sections)}")
        print(f"Приборов: {len(devices)}")
        for dev in devices:
            print(f"  {dev.addr}: {dev.type_name} v{dev.version} — {len(dev.shleifs)} шл., {len(dev.relays)} реле, {len(dev.outputs)} вых.")
    except Exception as e:
        print(f"Ошибка: {e}")
