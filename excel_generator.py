#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Генератор Excel для конфигурации С2000-М"""

import os
import platform
import subprocess
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment
from openpyxl.worksheet.pagebreak import Break


class ExcelGenerator:
    """Генератор Excel таблицы"""
    
    def __init__(self):
        self.wb = Workbook()
        self.ws = self.wb.active
        self.ws.title = 'Адреса, шлейфа'
        self.current_row = 1
        self._setup_styles()
    
    def _setup_styles(self):
        """Настроить стили"""
        self.thin = Side(border_style="thin", color="000000")
        self.bold_font = Font(name='Times New Roman', size=10, bold=True)
        self.regular_font = Font(name='Times New Roman', size=10, italic=True)
        self.wrap = Alignment(wrap_text=True)
        self.light_fill = PatternFill(fill_type='solid', fgColor='F0F8FF')
        self.white_fill = PatternFill(fill_type='solid', fgColor='FFFAF0')
        self.header_fill = PatternFill(fill_type='solid', fgColor='FFFFE0')
        self.address_fill = PatternFill(fill_type='solid', fgColor='E6E6FA')
        self.error_fill = PatternFill(fill_type='solid', fgColor='FF6347')
    
    def set_border(self, cell_range='A1:F1'):
        """Установить границы ячеек"""
        for row in self.ws[cell_range]:
            for cell in row:
                cell.border = Border(top=self.thin, bottom=self.thin, left=self.thin, right=self.thin)
    
    def add_title(self, list_str, row=1, column=1):
        """Добавить заголовок"""
        self.ws.merge_cells('A1:C1')
        for cel in list_str:
            if cel.find('Конфигурация ') != -1:
                column = 1
            if cel.find('Версия:') != -1:
                column = 4
            cell = self.ws.cell(row=row, column=column, value=cel)
            cell.alignment = self.wrap
            cell.font = self.bold_font
            cell.fill = self.light_fill if row % 2 == 0 else self.white_fill
            column += 1
        row += 1
        return row
    
    def add_address_row(self, list_str, row=1, column=1):
        """Добавить строку адреса"""
        title_sh = ['', 'Шлейф', 'Раздел', 'Тип шлейфа', '', 'Описание']
        
        for cel in list_str:
            if cel.find('Адрес:') != -1:
                column = 1
            if cel.find('Тип_прибора: ') != -1:
                cell_value = cel[cel.find('Тип_прибора: '):].replace('Тип_прибора: ', '')
                try:
                    device_id = int(cell_value)
                    cel = str(self._config.get_device_type(device_id)).strip('[]').strip('\'\'')
                except (ValueError, AttributeError):
                    pass
            if column == 2 and cel.find('Сценарий_упр:') == -1:
                t_cell = ' '
                cell = self.ws.cell(row=row, column=3, value=t_cell)
                cell.border = Border(bottom=self.thin, top=self.thin, left=self.thin, right=self.thin)
                cell.fill = self.address_fill
            if cel.find('Сценарий_упр:') != -1:
                column = 3
            if cel.find('Версия:') != -1:
                column = 4
            if cel.find('Описание:') != -1:
                cel = cel.replace('Описание:', '').replace('\"', '')
                column = 6
            cell = self.ws.cell(row=row, column=column, value=cel)
            cell.font = self.bold_font
            cell.border = Border(bottom=self.thin, top=self.thin, left=self.thin, right=self.thin)
            cell.fill = self.address_fill
            column += 1
        
        row += 1
        column = 1
        # Заполняем заголовок
        for it in title_sh:
            cell = self.ws.cell(row=row, column=column, value=it)
            cell.font = self.bold_font
            cell.border = Border(bottom=self.thin, top=self.thin)
            cell.fill = self.header_fill
            column += 1
        row += 1
        return row
    
    def add_output_row(self, list_str, row=1, column=1):
        """Добавить строку шлейфа/выхода/реле"""
        group_start = row
        level = 1
        
        for cel in list_str:
            if cel.find('Шлейф:') != -1:
                column = 2
                level = 1
            elif cel.find('Выход:') != -1:
                column = 2
                level = 2
            elif cel.find('Реле:') != -1:
                column = 2
                level = 3
                self.ws.row_dimensions[row].height = 25
            elif cel.find('Раздел:') != -1:
                try:
                    int_r = int(cel.replace(' ', '').replace('Раздел:', ''))
                    if self._config:
                        section = self._config.get_section(int_r)
                        if section:
                            cel = section['description']
                        else:
                            cel = ''
                except (ValueError, AttributeError):
                    pass
                self.ws.row_dimensions[row].height = 25
                column = 3
            elif cel.find('Программа:') != -1 and str(self.ws.cell(row=row, column=2).value).find('Реле:') != -1:
                try:
                    int_relay = int(cel.replace(' ', '').replace('Программа:', ''))
                    cel = 'Пр. упр: ' + str(self._config.get_relay(int_relay)).strip('[]').strip('\'\'')
                except (ValueError, AttributeError):
                    pass
                if 20 < len(cel) < 30:
                    self.ws.row_dimensions[row].height = 25
                if len(cel) > 30:
                    self.ws.row_dimensions[row].height = 37
                column = 3
            elif cel.find('Тип_шлейфа:') != -1:
                try:
                    int_cable = int(cel.replace(' ', '').replace('Тип_шлейфа:', ''))
                    cel = str(self._config.get_cable_type(int_cable)).strip('[]')
                except (ValueError, AttributeError):
                    pass
                self.ws.row_dimensions[row].height = 25
                column = 4
            elif cel.find('Описание:') != -1:
                cel = cel.replace('Описание:', '')
                column = 6
            elif cel.find('Время') != -1:
                column = 4
            
            cell = self.ws.cell(row=row, column=column, value=cel)
            cell.font = self.regular_font
            cell.alignment = self.wrap
            cell.fill = self.light_fill if row % 2 == 0 else self.white_fill
            column += 1
        
        self.ws.row_dimensions.group(start=group_start, end=row, outline_level=level, hidden=False)
        row += 1
        return row
    
    def set_config(self, config):
        """Установить конфигурацию для доступа к методам"""
        self._config = config
    
    def save(self, filepath):
        """Сохранить файл"""
        self.wb.save(filepath)
    
    def open_file(self, filepath):
        """Открыть файл (Linux/Windows/macOS)"""
        system = platform.system()
        try:
            if system == 'Linux':
                subprocess.call(['xdg-open', filepath])
            elif system == 'Darwin':  # macOS
                subprocess.call(['open', filepath])
            else:  # Windows
                os.startfile(filepath)
        except Exception as e:
            print(f"Не удалось открыть файл: {e}", file=sys.stderr)
    
    def mark_missing_descriptions(self, max_row):
        """Пометить ячейки с отсутствующими описаниями"""
        for row in self.ws.iter_rows(min_row=1, min_col=6, max_col=6, max_row=max_row):
            for cell in row:
                if cell.value is None:
                    cell.fill = self.error_fill
                    cell.value = 'Где описание!!!???'
                    cell.font = self.bold_font
    
    def set_columns_width(self):
        """Установить ширину столбцов"""
        self.ws.column_dimensions['A'].width = 10
        self.ws.column_dimensions['B'].width = 13
        self.ws.column_dimensions['C'].width = 18
        self.ws.column_dimensions['D'].width = 20
        self.ws.column_dimensions['E'].width = 18
        self.ws.column_dimensions['F'].width = 25
    
    def hide_columns(self, hidden_cols):
        """Скрыть столбцы"""
        for col in hidden_cols:
            self.ws.column_dimensions[col].hidden = True
    
    def set_printer_settings(self, paper_size, orientation):
        """Установить настройки принтера"""
        self.ws.set_printer_settings(paper_size, orientation)
    
    def set_page_breaks(self):
        """Установить разрывы страниц"""
        self.ws.col_breaks.brk = [Break(6)]
    
    def set_view_mode(self, mode):
        """Установить режим просмотра"""
        self.ws.sheet_view.view = mode


if __name__ == '__main__':
    # Тест
    gen = ExcelGenerator()
    gen.set_printer_settings(gen.wb.PAPERSIZE_A4, gen.wb.ORIENTATION_PORTRAIT)
    gen.set_page_breaks()
    gen.set_view_mode('pageBreakPreview')
    gen.hide_columns(['E'])
    gen.set_columns_width()
    gen.save('test_output.xlsx')
    print("Тестовый файл создан")
