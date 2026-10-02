# Конвертер конфигурации С2000-М

CLI утилита для конвертации конфигурации С2000-М в Excel или Markdown.

## Установка

```bash
pip install openpyxl
```

## Использование

```bash
python3 main.py CONFIG_FILE [-o OUTPUT] [-v] [--check] [--format xlsx|md] [--no-open]
```

### Аргументы

- `CONFIG_FILE` — путь к файлу конфигурации С2000-М (обязательно)
- `-o OUTPUT` — выходной файл (по умолчанию: output)
- `-v` — подробный вывод
- `--check` — пометить отсутствующие описания
- `--format xlsx|md` — формат вывода (по умолчанию: xlsx)
- `--no-open` — не открывать файл после генерации

### Примеры

```bash
# Excel по умолчанию
python3 main.py 020926.TXT -o out.xlsx

# Markdown
python3 main.py 020926.TXT --format md -o out.md

# С проверкой описаний
python3 main.py 020926.TXT --check

# Без открытия файла
python3 main.py 020926.TXT --format xlsx --no-open
```

## Формат конфигурации

Файл конфигурации С2000-М в кодировке cp1251.

### Структура проекта

```
c2000-m_config/
├── main.py              # Точка входа
├── cli_config.py        # CLI интерфейс (argparse)
├── config_parser.py     # Парсер конфигурации
├── excel_generator.py   # Генератор Excel (openpyxl)
├── md_generator.py      # Генератор Markdown
├── utils.py             # Утилиты и функции валидации
└── README.md
```

## Модули

### cli_config.py

CLI интерфейс на основе argparse. Аргументы: config_file, output, verbose, format, no_open, check.

### config_parser.py

Класс `ConfigParser` для парсинга файлов конфигурации С2000-М (cp1251). Строит структурированную модель: devices, sections, raw_lines.

### excel_generator.py

Класс `ExcelGenerator` для создания Excel через openpyxl. Поддерживает:
- Заголовки приборов, шлейфов, реле/выходов
- Объединение реле и выходов по номеру
- Границы ячеек (thin, color="000000")
- Проверку отсутствующих описаний

### md_generator.py

Класс `generate_md()` для вывода списка:
- **Адрес** — тип прибора — версия
- `Шл.N` — раздел — тип шлейфа — описание
- `Рел.N` — раздел — программа — описание
- `Чит.N` — раздел — описание

### utils.py

Валидация, проверка описаний, форматирование.

## Требования

- Python 3.6+
- openpyxl 3.0+
- Linux/Windows/macOS

## License

MIT License
