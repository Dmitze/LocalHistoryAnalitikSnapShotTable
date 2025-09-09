function copyATDataToHistory() {
  copyAllATTables();
}

function createOrUpdateTimeTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'copyATDataToHistory') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Создаём новый — каждый день в 8 утра по Киеву
  ScriptApp.newTrigger('copyATDataToHistory')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .inTimezone("Europe/Kyiv")
    .create();
}



function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Аналітика')
    .addItem('Запустити парсер зараз', 'copyATDataToHistory')
    .addItem('Пересоздать триггер (щохвилини)', 'createOrUpdateTimeTrigger')
    .addToUi();
}

function copyAllATTables() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    console.log('ℹ️ Пропуск: уже идёт выполнение.');
    return;
  }
  try {
    const sourceSpreadsheetId = '1CcdvLUU5V9DgyJllttNTyLVoxIwJXILFvEOFU31W7lY';
    const targetSpreadsheetId = '10AI7S1fXCbE6kZdGCW4shgzDjK75d24z3K0xQXgHDKA';
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

function copyTableWithDiff(sourceSheet, targetSheet, sourceRangeA1, dateTime) {
  if (!sourceSheet || typeof sourceSheet.getRange !== 'function') {
    throw new Error('sourceSheet не є листом.');
  }
  if (!targetSheet) throw new Error('targetSheet не знайдено.');

  const sourceRange = sourceSheet.getRange(sourceRangeA1);
  const values = sourceRange.getValues();
  const numRows = values.length;
  const numCols = values[0].length;

  const lastRow = targetSheet.getLastRow();
  const nextRow = lastRow === 0 ? 4 : lastRow + 2;

  // Дата/час в колонку A
  for (let i = 0; i < numRows; i++) {
    targetSheet.getRange(nextRow + i, 1)
      .setValue(dateTime)
      .setNumberFormat('@')
      .setFontStyle('italic')
      .setFontColor('gray')
      .setHorizontalAlignment('center');
  }

  // Копируем значения
  const dataRange = targetSheet.getRange(nextRow, 2, numRows, numCols);
  dataRange.setValues(values);

  // Копируем стили (без фона)
  dataRange.setFontColors(sourceRange.getFontColors());
  dataRange.setFontSizes(sourceRange.getFontSizes());
  dataRange.setFontWeights(sourceRange.getFontWeights());
  dataRange.setFontStyles(sourceRange.getFontStyles());
  dataRange.setHorizontalAlignment(sourceRange.getHorizontalAlignment());
  dataRange.setVerticalAlignment(sourceRange.getVerticalAlignment());
  dataRange.setWrapStrategy(sourceRange.getWrapStrategy());

  // Поиск предыдущего блока
  const prevTopRow = findPrevBlockTop(targetSheet, nextRow, numRows);
  const hasPrev = prevTopRow > 0;

  if (!hasPrev) {
    // Первый снимок — без окраски
    for (let i = 0; i < numRows; i++) {
      for (let j = 2; j < numCols; j++) {
        dataRange.getCell(i + 1, j + 1).setBackground(null);
      }
    }
    return;
  }

  const prevValues = targetSheet.getRange(prevTopRow, 2, numRows, numCols).getValues();

  // Красим только изменения
  for (let i = 0; i < numRows; i++) {
    for (let j = 2; j < numCols; j++) {
      const newVal = toNum(values[i][j]);
      const oldVal = toNum(prevValues[i][j]);
      const cell = dataRange.getCell(i + 1, j + 1);

      if (newVal > oldVal) {
        cell.setBackground('#d4edda'); // зелёный
      } else if (newVal < oldVal) {
        cell.setBackground('#f8d7da'); // красный
      } else {
        cell.setBackground(null);      // без фона
      }
    }
  }
}

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

function toNum(v) {
  if (v === null || v === '' || v === undefined) return 0;
  if (typeof v === 'number') return isNaN(v) ? 0 : v;
  let s = String(v)
    .replace(/%/g, '')
    .replace(/[^\d,\-\.]/g, '')
    .replace(/,/g, '.')
    .trim();
  let num = parseFloat(s);
  return isNaN(num) ? 0 : num;
}

function ensureSheet(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}
