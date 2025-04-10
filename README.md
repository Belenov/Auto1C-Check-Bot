# Auto1C-Check-Bot
Telegram-бот на Python для автоматической проверки обновлений 1С, сбора писем и генерации отчётов в Excel
# 🧠 1С AutoUpdate & Mail Checker Bot

Telegram-бот, который:
- Авторизуется на портале 1С
- Проверяет актуальные версии конфигураций
- Сравнивает их с локальным Excel
- Генерирует отчёты
- Чекает входящие письма (Mail.ru, IMAP)
- Уведомляет пользователей через Telegram

---

## 📦 Возможности

- ✅ Авторизация и подписка через Telegram
- ✅ Поддержка Excel-отчётов (через openpyxl)
- ✅ Выявление ошибок и предупреждений в письмах
- ✅ Подсветка данных (green/red) в Excel
- ✅ Игра “Змейка” прямо из консоли (а как без этого)

---

## 🧰 Используемые технологии

- Python 3.10+
- `selenium`, `openpyxl`, `pandas`
- `telegram.ext`
- `imaplib`, `bs4`, `threading`
- `msvcrt` + `colorama` для Windows UI

---

## 🚀 Запуск

```bash
python main.py
