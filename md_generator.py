#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Генерация Markdown-списка из структурированных данных (без таблиц)"""


def generate_md(devices, sections, output_path=None):
    """Генерация Markdown: адрес -> строки шлейфов/реле/выходов"""
    lines = []
    lines.append("# Подключённые адреса\n")

    for dev in devices:
        # Строка адреса
        lines.append(f"**{dev.addr}** — {dev.type_name} — {dev.version}")
        if dev.description:
            lines.append(f"   Описание: {dev.description}")
        lines.append("")

        # Шлейфы
        for shleif in dev.shleifs:
            has_section = bool(shleif.get('section_id'))
            has_desc = bool(shleif.get('description'))
            if not has_section and not has_desc:
                continue

            section_desc = ''
            if shleif.get('section_id'):
                section = sections.get(shleif['section_id'])
                if section:
                    section_desc = section['description']

            type_name = shleif.get('type_name', '')
            desc = shleif.get('description', '')
            lines.append(f"  Шл. {shleif['num']} — {section_desc} — {type_name} — {desc}")

        # Реле/выходы объединены по номеру
        merged_relays = []
        output_by_num = {}
        for o in dev.outputs:
            output_by_num[o['num']] = o
        for r in dev.relays:
            if r.get('program_id') and r['program_id'] != 0:
                merged = dict(r)
                if r['num'] in output_by_num:
                    out = output_by_num[r['num']]
                    if out.get('description'):
                        merged['output_desc'] = out['description']
                    if out.get('section_id') and not merged.get('section_id'):
                        merged['section_id'] = out['section_id']
                merged_relays.append(merged)

        for relay in merged_relays:
            if not relay.get('program_id') or relay['program_id'] == 0:
                continue

            section_desc = ''
            if relay.get('section_id'):
                section = sections.get(relay['section_id'])
                if section:
                    section_desc = section['description']

            program_name = relay.get('program_name', '')
            program_val = f"Пр. упр: {program_name}" if program_name else relay.get('description', '')

            output_desc = relay.get('output_desc', '')
            relay_desc = relay.get('description', '')
            desc_val = output_desc if output_desc else relay_desc

            lines.append(f"  Рел.{relay['num']} — {section_desc} — {program_val} — {desc_val}")

        # Считыватели
        for reader in dev.readers:
            has_section = bool(reader.get('section_id'))
            has_desc = bool(reader.get('description'))
            if not has_section and not has_desc:
                continue

            section_desc = ''
            if reader.get('section_id'):
                section = sections.get(reader['section_id'])
                if section:
                    section_desc = section['description']

            desc = reader.get('description', '')
            lines.append(f"  Чит.{reader['num']} — {section_desc} — — {desc}")

        lines.append("")  # пустая строка между приборами

    result = '\n'.join(lines)

    if output_path:
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(result)

    return result


if __name__ == '__main__':
    from config_parser import ConfigParser
    parser = ConfigParser()
    devices, raw_lines = parser.parse_file('/media/mamont/4060-4999/020926.md')
    result = generate_md(devices, parser.sections)
    print(result)
