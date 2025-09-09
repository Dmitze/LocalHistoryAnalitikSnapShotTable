// Обёртка для триггера: без аргументов
function copyATDataToHistory() {
  copyAllATTables();
}

// Пересоздание триггера (пример: каждую минуту)
function createOrUpdateTimeTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === 'copyATDataToHistory') {
      ScriptApp.deleteTrigger(t);
    }
    if (t.getHandlerFunction() === 'copyTableWithDiff') {
      // На всякий случай удалим неверный
      ScriptApp.deleteTrigger(t);
    }
  }
  ScriptApp.newTrigger('copyATDataToHistory')
    .timeBased()
    .everyMinutes(1)
    .create();
}

// Меню для ручного запуска и ремонта триггера
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Аналітика')
    .addItem('Запустити парсер зараз', 'copyATDataToHistory')
    .addItem('Пересоздать триггер (щохвилини)', 'createOrUpdateTimeTrigger')
    .addToUi();
}

// Основной запуск: обходит все диапазоны и листы-истории
function copyAllATTables() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    console.log('ℹ️ Пропуск: уже идёт выполнение.');
    return;
  }
  try {
    const sourceSpreadsheetId = '';
    const targetSpreadsheetId = '';
    const sourceSheetName = 'Загальна АТ';

    const now = new Date();
    const dateTime = Utilities.formatDate(now, 'Europe/Kyiv', 'dd.MM.yy HH:mm');

    const sourceSS = SpreadsheetApp.openById(sourceSpreadsheetId);
    const sourceSheet = sourceSS.getSheetByName(sourceSheetName);
    const targetSS = SpreadsheetApp.openById(targetSpreadsheetId);

    if (!sourceSheet) throw new Error(`Лист "${sourceSheetName}" не знайдено.`);

    const tasks = [
      { range: 'A5:O16',  target: 'History Загальна АТ' },
      { range: 'Q5:AE16', target: 'History АТ 1 Бат' },
      { range: 'AG5:AU16', target: 'History АТ 2 Бат' },
      { range: 'AW5:BK16', target: 'History АТ 3 Бат' },
      { range: 'BM5:CA16', target: 'History РТР' },
      { range: 'CC5:CQ16', target: 'History РМТЗ' },
      { range: 'CS5:DG16', target: 'History ІСР' }
    ];

    for (const t of tasks) {
      const targetSheet = ensureSheet(targetSS, t.target);
      copyTableWithDiff(sourceSheet, targetSheet, t.range, dateTime);
    }
  } catch (e) {
    console.error('❌ Помилка у copyAllATTables:', e.message);
    throw e;
  } finally {
    lock.releaseLock();
  }
}

// Универсальная копия + сравнение с предыдущим блоком
function copyTableWithDiff(sourceSheet, targetSheet, sourceRangeA1, dateTime) {
  if (!sourceSheet || typeof sourceSheet.getRange !== 'function') {
    throw new Error('sourceSheet не є листом. Перевір триггер: він має викликати обгортку без аргументів.');
  }
  if (!targetSheet) throw new Error('targetSheet не знайдено.');

  const sourceRange = sourceSheet.getRange(sourceRangeA1);
  const values = sourceRange.getValues();
  const numRows = values.length;
  const numCols = values[0].length;

  const lastRow = targetSheet.getLastRow();
  const nextRow = lastRow === 0 ? 4 : lastRow + 2;

  // Метка часу в A
  for (let i = 0; i < numRows; i++) {
    targetSheet.getRange(nextRow + i, 1)
      .setValue(dateTime)
      .setNumberFormat('@')
      .setFontStyle('italic')
      .setFontColor('gray')
      .setHorizontalAlignment('center');
  }

  // Вставка даних в B:...
  const dataRange = targetSheet.getRange(nextRow, 2, numRows, numCols);
  dataRange.setValues(values);

  // Стилі (без фону)
  dataRange.setFontColors(sourceRange.getFontColors());
  dataRange.setFontSizes(sourceRange.getFontSizes());
  dataRange.setFontWeights(sourceRange.getFontWeights());
  dataRange.setFontStyles(sourceRange.getFontStyles());
  dataRange.setHorizontalAlignment(sourceRange.getHorizontalAlignment());
  dataRange.setVerticalAlignment(sourceRange.getVerticalAlignment());
  dataRange.setWrapStrategy(sourceRange.getWrapStrategy());

  // Пошук попереднього блоку
  const prevTopRow = findPrevBlockTop(targetSheet, nextRow, numRows);
  const hasPrev = prevTopRow > 0;

  if (hasPrev) {
    const prevValues = targetSheet.getRange(prevTopRow, 2, numRows, numCols).getValues();
    for (let i = 0; i < numRows; i++) {
      for (let j = 2; j < numCols; j++) {
        const cell = dataRange.getCell(i + 1, j + 1);
        const isPercentCol = (j === numCols - 1);
        const newVal = toNum(values[i][j], isPercentCol);
        const oldVal = toNum(prevValues[i][j], isPercentCol);
        const diff = newVal - oldVal;

        if (diff > 0) {
          cell.setValue(formatVal(newVal, isPercentCol) + ' ↑ +' + formatVal(diff, isPercentCol))
              .setBackground('#d4edda');
        } else if (diff < 0) {
          cell.setValue(formatVal(newVal, isPercentCol) + ' ↓ ' + formatVal(Math.abs(diff), isPercentCol))
              .setBackground('#f8d7da');
        } else {
          cell.setValue(formatVal(newVal, isPercentCol)).setBackground(null);
        }
      }
    }
  } else {
    // Перша база без підсвітки
    for (let i = 0; i < numRows; i++) {
      for (let j = 2; j < numCols; j++) {
        const isPercentCol = (j === numCols - 1);
        dataRange.getCell(i + 1, j + 1)
          .setValue(formatVal(toNum(values[i][j], isPercentCol), isPercentCol))
          .setBackground(null);
      }
    }
  }
}

// Пошук верхньої строки попереднього блоку тієї ж висоти
function findPrevBlockTop(sheet, nextRow, numRows) {
  let r = nextRow - 1;
  if (r < 2) return 0;
  const colA = sheet.getRange(2, 1, r - 1, 1).getValues();
  while (r >= 2) {
    const val = colA[r - 2][0];
    if (isDateStamp(val)) break;
    r--;
  }
  if (r < 2) return 0;
  const top = r - (numRows - 1);
  return top >= 2 ? top : 0;
}

function isDateStamp(v) {
  return (typeof v === 'string') && /^\d{2}\.\d{2}\.\d{2} \d{2}:\d{2}$/.test(v);
}

function toNum(v, isPercent) {
  if (v === null || v === '' || v === undefined) return 0;
  if (typeof v === 'number') return isNaN(v) ? 0 : v;
  let s = String(v)
    .replace(/↑.*$/, '')
    .replace(/↓.*$/, '')
    .replace(/%/g, '')
    .replace(/[^\d,\-\.]/g, '')
    .replace(/,/g, '.')
    .trim();
  let num = parseFloat(s);
  return isNaN(num) ? 0 : num;
}

function formatVal(x, isPercent) {
  if (isPercent) return (Math.round(x * 100) / 100).toFixed(2) + '%';
  return String(Math.round(x));
}

// Создаёт лист, если его нет
function ensureSheet(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}
