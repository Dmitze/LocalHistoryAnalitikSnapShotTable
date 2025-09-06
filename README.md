# 📊 Google Sheets Auto History & Comparison Script

## 🇺🇦 Опис (Ukrainian)

Цей Google Apps Script автоматизує копіювання кількох таблиць з одного аркуша Google Sheets до відповідних аркушів-історій в іншій таблиці, з автоматичним порівнянням змін між версіями.

### ✨ Можливості
- Копіює дані з кількох діапазонів (`A5:O16`, `Q5:AE16`, `AG5:AU16`, тощо) з аркуша **"Загальна АТ"**.
- Вставляє дані у відповідні аркуші-історії (наприклад, `History Загальна АТ`, `History АТ 1 Бат`, …).
- Додає мітку часу до кожного блоку даних.
- Порівнює з попереднім збереженим блоком:
  - 📈 Збільшення значення — зелений фон та стрілка вгору.
  - 📉 Зменшення значення — червоний фон та стрілка вниз.
  - ➖ Без змін — просто число без підсвітки.
- Меню в інтерфейсі для ручного запуску та пересоздання тригера.
- Підтримка запуску за розкладом (тригери).

### 🛠 Як налаштувати
1. Відкрийте редактор скриптів у Google Sheets (`Extensions` → `Apps Script`).
2. Вставте код із цього репозиторію.
3. Задайте **ID вихідної таблиці** та **ID цільової таблиці** у змінних `sourceSpreadsheetId` та `targetSpreadsheetId`.
4. Запустіть функцію `createOrUpdateTimeTrigger()` для створення тригера (наприклад, щохвилини).
5. Використовуйте меню **📊 Аналітика** у Google Sheets для ручного запуску.

### 📂 Структура коду
- `copyATDataToHistory()` — обгортка для тригера.
- `createOrUpdateTimeTrigger()` — пересоздає тригер.
- `onOpen()` — додає меню в інтерфейс.
- `copyAllATTables()` — основний цикл по всіх діапазонах і аркушах-історіях.
- `copyTableWithDiff()` — копіювання та порівняння з попереднім блоком.
- Допоміжні функції: `findPrevBlockTop`, `isDateStamp`, `toNum`, `formatVal`, `ensureSheet`.

---

## 🇬🇧 Description (English)

This Google Apps Script automates copying multiple tables from one Google Sheets tab to corresponding history tabs in another spreadsheet, with automatic change comparison between versions.

### ✨ Features
- Copies data from multiple ranges (`A5:O16`, `Q5:AE16`, `AG5:AU16`, etc.) from the **"Загальна АТ"** sheet.
- Inserts data into corresponding history sheets (e.g., `History Загальна АТ`, `History АТ 1 Бат`, …).
- Adds a timestamp to each data block.
- Compares with the previous saved block:
  - 📈 Increase — green background and upward arrow.
  - 📉 Decrease — red background and downward arrow.
  - ➖ No change — plain number without highlight.
- Adds a custom menu for manual run and trigger recreation.
- Supports scheduled runs via time-based triggers.

### 🛠 Setup
1. Open the Script Editor in Google Sheets (`Extensions` → `Apps Script`).
2. Paste the code from this repository.
3. Set the **source spreadsheet ID** and **target spreadsheet ID** in `sourceSpreadsheetId` and `targetSpreadsheetId`.
4. Run `createOrUpdateTimeTrigger()` to create a trigger (e.g., every minute).
5. Use the **📊 Аналітика** menu in Google Sheets for manual execution.

### 📂 Code Structure
- `copyATDataToHistory()` — trigger wrapper.
- `createOrUpdateTimeTrigger()` — recreates the trigger.
- `onOpen()` — adds a menu to the UI.
- `copyAllATTables()` — main loop through all ranges and history sheets.
- `copyTableWithDiff()` — copies and compares with the previous block.
- Helper functions: `findPrevBlockTop`, `isDateStamp`, `toNum`, `formatVal`, `ensureSheet`.

---

## 📜 License
This project is licensed under the MIT License — see the [LICENSE](LICENSE) file for details.
