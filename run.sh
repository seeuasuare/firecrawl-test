#!/bin/bash
# Один скрипт — делает всё сам: клонирует репо, ставит библиотеки, запускает парсинг.
# Использование:  curl -sL URL | bash   или просто:  bash run.sh

set -e

echo "=== Скачиваю репозиторий ==="
cd ~/Desktop 2>/dev/null || cd ~
rm -rf firecrawl-test
git clone https://github.com/seeuasuare/firecrawl-test.git
cd firecrawl-test
git checkout claude/scrape-company-emails-BxCZm

echo ""
echo "=== Устанавливаю зависимости ==="
pip install openpyxl aiohttp beautifulsoup4

echo ""
echo "=== Запускаю парсинг email-адресов ==="
python scrape_emails.py

echo ""
echo "=== Готово! ==="
echo "Файл с результатами: $(pwd)/companies_with_emails.xlsx"
open companies_with_emails.xlsx 2>/dev/null || true
