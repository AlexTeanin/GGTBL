/** Ежедневный журнал учёта металла — Apps Script
 *  Порт исходного.txt (CustomTkinter) в Google Sheets.
 *  Пароль изменения справочника/порогов/структуры: 1234
 */

var PASSWORD = '1234';
var TEMPLATE = 'Шаблон_месяца';
var REF = 'Справочник';
var CFG = {
  days: 31,           // максимум колонок дней
  rowsPerItem: 4,     // Приход / Расход / Лом / Баланс
  inputRows: ['Приход', 'Расход', 'Лом', null] // null = строка Баланс (формулы)
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Журнал')
    .addItem('Создать месяц…', 'createMonthPrompt')
    .addSeparator()
    .addItem('Добавить артикул', 'addItem')
    .addItem('Удалить артикул', 'removeItem')
    .addItem('Настроить пороги', 'setThresholds')
    .addSeparator()
    .addItem('Исправить формулы месяца', 'repairCurrentMonthFormulas')
    .addSeparator()
    .addItem('Снять защиту (пароль)…', 'unprotectWithPassword')
    .addItem('Восстановить защиту', 'protectAllSheets')
    .addToUi();
}

function askPassword() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt('Ввод пароля', 'Введите пароль администратора:', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() != ui.Button.OK) return false;
  return resp.getResponseText().trim() === PASSWORD;
}

/* ---------- Создание листа месяца ---------- */

function createMonthPrompt() {
  var ui = SpreadsheetApp.getUi();
  var def = Utilities.formatDate(new Date(), 'GMT+3', 'yyyy-MM');
  var resp = ui.prompt('Создать месяц', 'Месяц в формате ГГГГ-ММ (напр. ' + def + '):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() != ui.Button.OK) return;
  createMonth(resp.getResponseText().trim());
}

function createMonth(ym) {
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(ym)) {
    SpreadsheetApp.getUi().alert('Неверный формат месяца. Используйте ГГГГ-ММ.');
    return;
  }
  var ss = SpreadsheetApp.getActive();
  if (ss.getSheetByName(ym)) {
    SpreadsheetApp.getUi().alert('Лист ' + ym + ' уже существует. Существующие листы никогда не перезаписываются.');
    return;
  }
  buildTemplateIfNeeded(ss);
  var tmpl = ss.getSheetByName(TEMPLATE);
  var sheet = tmpl.copyTo(ss).setName(ym);
  // заголовок месяца
  sheet.getRange('A1').setValue('Журнал учёта металла — ' + ym);
  // перенос остатка из прошлого месяца (исправление бага .py: никакой перезаписи)
  var prev = prevMonth(ym);
  var prevSheet = ss.getSheetByName(prev);
  if (prevSheet) {
    var items = getItems();
    var start = 3; // первая строка данных (совпадает с applyMonthFormulas/protectSheet)
    // колонку «Итог месяца» прошлого месяца вычисляем по его реальной раскладке
    var py = parseInt(prev.slice(0, 4)), pm = parseInt(prev.slice(5, 7));
    var prevLayout = computeMonthLayout(py, pm);
    for (var i = 0; i < items.length; i++) {
      // в шаблоне "Начальный остаток" в колонке B на строке Баланса
      var balRow = start + i * CFG.rowsPerItem + 3;
      var prevFinal = prevSheet.getRange(balRow, prevLayout.monthTotalCol).getValue();
      // копируем только число: текст-ошибку ('#ERROR!' и т.п.) не переносим
      if (typeof prevFinal === 'number') {
        sheet.getRange(balRow, 2).setValue(prevFinal);
      }
    }
  }
  applyMonthFormulas(sheet, ym, getItems());
  applyIndicators(sheet);
  protectSheet(sheet);
  ss.setActiveSheet(sheet);
}

function ensureColumns(sheet, minCols) {
  if (sheet.getMaxColumns() < minCols) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), minCols - sheet.getMaxColumns());
  }
}

function prevMonth(ym) {
  var y = parseInt(ym.slice(0, 4)), m = parseInt(ym.slice(5, 7));
  if (m == 1) { y--; m = 12; } else { m--; }
  return y + '-' + ('0' + m).slice(-2);
}

/* ---------- Справочник ---------- */

function getItems() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(REF);
  if (!sh) return [];
  var vals = sh.getDataRange().getValues();
  var items = [];
  for (var r = 1; r < vals.length; r++) {
    if (vals[r][0]) {
      // прикладываем номер строки справочника (для порогов в applyIndicators):
      // пустые строки пропускаются, поэтому i+2 может не совпадать с реальной строкой
      vals[r].srcRow = r + 1;
      items.push(vals[r]);
    }
  }
  return items;
}

function addItem() {
  if (!askPassword()) { SpreadsheetApp.getUi().alert('Неверный пароль.'); return; }
  var ui = SpreadsheetApp.getUi();
  var art = ui.prompt('Новый артикул', 'Артикул F12:', ui.ButtonSet.OK_CANCEL);
  if (art.getSelectedButton() != ui.Button.OK) return;
  var name = ui.prompt('Наименование 1С', 'Наименование:', ui.ButtonSet.OK_CANCEL);
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(REF);
  if (!sh) { buildReferenceSheet(); sh = ss.getSheetByName(REF); }
  sh.appendRow([art.getResponseText().trim(), art.getResponseText().trim(),
                name.getSelectedButton() == ui.Button.OK ? name.getResponseText() : '',
                '', 0, 0, 0, 0, 0, '', '', 0]);
  rebuildAllMonths();
}

function removeItem() {
  if (!askPassword()) { SpreadsheetApp.getUi().alert('Неверный пароль.'); return; }
  var ui = SpreadsheetApp.getUi();
  var art = ui.prompt('Удалить артикул', 'Артикул F12:', ui.ButtonSet.OK_CANCEL);
  if (art.getSelectedButton() != ui.Button.OK) return;
  var a = art.getResponseText().trim();
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(REF);
  var vals = sh.getDataRange().getValues();
  for (var r = 1; r < vals.length; r++) {
    if (vals[r][0] == a) { sh.deleteRow(r + 1); break; }
  }
  rebuildAllMonths();
}

function setThresholds() {
  if (!askPassword()) { SpreadsheetApp.getUi().alert('Неверный пароль.'); return; }
  SpreadsheetApp.getUi().alert('Пороги (min_balance) правятся в листе «Справочник», колонка L. Снятие защиты уже выполнено.');
  unprotectSheet(SpreadsheetApp.getActive().getSheetByName(REF));
}

/* ---------- Шаблон и формулы ---------- */

/** Строит лист «Шаблон_месяца»: строки статичны, колонки дней 1–31 + недельные итоги + итог месяца.
 *  Раскладка: A=Артикул, B=Начальный остаток (строка Баланса), C..=дни/итоги.
 *  Для недели/месяца: колонка «Итог N» (формула) и «Пересчёт N» (ввод).
 */
function buildTemplateIfNeeded(ss) {
  if (ss.getSheetByName(TEMPLATE)) return;
  var sh = ss.insertSheet(TEMPLATE);
  buildReferenceSheet();
  ensureColumns(sh, 60);

  sh.getRange('A1').setValue('Шаблон месяца');
  sh.getRange('A2').setValue('Артикул').setFontWeight('bold');
  sh.getRange('B2').setValue('Нач. ост.').setFontWeight('bold');
  // колонки дней
  var col = 3;
  for (var d = 1; d <= CFG.days; d++) {
    sh.getRange(2, col).setValue(d).setFontWeight('bold');
    col++;
    // после воскресенья вставляем Итог/Пересчёт недели — реальная раскладка
    // зависит от месяца, поэтому формула-генерация делается в applyMonthFormulas
  }
  // заголовки «Итог месяца»/«Пересчёт» не задаются жёстко:
  // реальную раскладку под конкретный месяц пишет applyMonthFormulas
  sh.setFrozenRows(2);
  sh.setFrozenColumns(2);
}

/** Раскладка колонок конкретного месяца: дни с колонки 3, после каждого воскресенья
 *  +2 колонки («Итог N», «Пересчёт N»), последняя неделя = «Итог месяца»/«Пересчёт».
 *  Используется и в applyMonthFormulas, и при переносе остатка в createMonth.
 */
function computeMonthLayout(y, m) {
  var daysInMonth = new Date(y, m, 0).getDate();
  var dayCols = {};   // day -> column
  var weekEnd = {};   // weekNumber -> column of 'Итог N'
  var weekRecalc = {};// weekNumber -> column of 'Пересчёт N'
  var col = 3, week = 1;
  for (var d = 1; d <= daysInMonth; d++) {
    dayCols[d] = col; col++;
    var dow = new Date(y, m - 1, d).getDay();
    if (dow === 0 && d < daysInMonth) { // воскресенье — конец недели
      weekEnd[week] = col; col++;
      weekRecalc[week] = col; col++;
      week++;
    }
  }
  weekEnd[week] = col; col++;      // итог последней неполной недели
  weekRecalc[week] = col; col++;   // = «Итог месяца»/«Пересчёт»
  return {
    daysInMonth: daysInMonth,
    dayCols: dayCols,
    weekEnd: weekEnd,
    weekRecalc: weekRecalc,
    week: week,
    monthTotalCol: weekEnd[week],
    monthRecalcCol: weekRecalc[week]
  };
}

/** Защита от «сломанных» формул: строка должна полностью соответствовать шаблону
 *  (например, мусорная цифра после последней ')' вызывает исключение).
 */
function assertCleanFormula(f) {
  // недопустимы посторонние символы и мусорные цифры после ')' (напр. '...SUM(F5))2')
  if (!/^=[()A-Z0-9+\-<>,."$ ]+$/.test(f) || /\)\d/.test(f)) {
    throw new Error('Проверка формулы не пройдена: ' + f);
  }
}

/** Генерирует раскладку колонок под конкретный месяц и формулы баланса.
 *  Дни, чей номер больше числа дней в месяце, скрываются.
 */
function applyMonthFormulas(sheet, ym, items) {
  ensureColumns(sheet, 60);
  var y = parseInt(ym.slice(0, 4)), m = parseInt(ym.slice(5, 7));
  var layout = computeMonthLayout(y, m);
  var daysInMonth = layout.daysInMonth;
  var dayCols = layout.dayCols;
  var weekEnd = layout.weekEnd;
  var weekRecalc = layout.weekRecalc;
  var week = layout.week;
  var monthTotalCol = layout.monthTotalCol;
  var monthRecalcCol = layout.monthRecalcCol;

  // заголовки
  var sh = sheet;
  // строка заголовков строится по реальным колонкам месяца:
  // после воскресенья идут +2 колонки («Итог N», «Пересчёт N»),
  // поэтому дни пишутся не подряд, а в свои dayCols
  var headerRow = [];
  for (var hc = 3; hc <= monthRecalcCol; hc++) headerRow[hc - 3] = '';
  var names = {0:'Вс',1:'Пн',2:'Вт',3:'Ср',4:'Чт',5:'Пт',6:'Сб'};
  for (var d = 1; d <= daysInMonth; d++) {
    var dow = new Date(y, m - 1, d).getDay();
    headerRow[dayCols[d] - 3] = d + ' ' + names[dow];
  }
  for (var w = 1; w <= week; w++) {
    headerRow[weekEnd[w] - 3] = (w == week ? 'Итог месяца' : 'Итог ' + w);
    headerRow[weekRecalc[w] - 3] = (w == week ? 'Пересчёт' : 'Пересчёт ' + w);
  }
  // стираем старые заголовки прошлых раскладок и пишем новую строку целиком
  sh.getRange(2, 3, 1, monthRecalcCol - 2).clearContent();
  sh.getRange(2, 3, 1, monthRecalcCol - 2).setValues([headerRow]).setFontWeight('bold');
  // скрыть неиспользуемые колонки (после последней)
  var maxCols = sh.getMaxColumns();
  if (monthRecalcCol + 1 <= maxCols) {
    sh.hideColumns(monthRecalcCol + 1, maxCols - monthRecalcCol);
  }

  // строки данных
  var start = 3;
  for (var i = 0; i < items.length; i++) {
    var r = start + i * CFG.rowsPerItem;
    var art = items[i][0];
    // числовой начальный остаток (перенос из прошлого месяца) сохраняем,
    // чтобы метка «Баланс» его не затёрла
    var savedBal = sh.getRange(r + 3, 2).getValue();
    sh.getRange(r, 1, 4, 1).mergeVertically().setValue(art).setVerticalAlignment('middle');
    sh.getRange(r, 2).setValue('Приход');
    sh.getRange(r + 1, 2).setValue('Расход');
    sh.getRange(r + 2, 2).setValue('Лом');
    sh.getRange(r + 3, 2).setValue('Баланс');
    if (typeof savedBal === 'number') {
      sh.getRange(r + 3, 2).setValue(savedBal);
    }
    sh.getRange(r, 2, 4, 1).setFontStyle('italic').setFontSize(9);
    sh.getRange(r, 1, 1, 2).setFontWeight('bold');

    // Баланс дня: N(предыдущий баланс) + SUM(приход) − SUM(расход) − SUM(лом).
    // С понедельника, если заполнен «Пересчёт N» прошлой недели, база = пересчёт.
    var balRow = r + 3;
    var incRow = r, outRow = r + 1, scrapRow = r + 2;
    for (var d = 1; d <= daysInMonth; d++) {
      var c = dayCols[d];
      var base;
      if (d === 1) {
        base = 'N($B$' + balRow + ')';
      } else {
        base = 'N(' + sh.getRange(balRow, dayCols[d - 1]).getA1Notation() + ')';
      }
      var dowD = new Date(y, m - 1, d).getDay();
      if (d > 1 && dowD === 1) {
        for (var w = 1; w <= week; w++) {
          if (weekEnd[w] < c) {
            var rcA1 = sh.getRange(balRow, weekRecalc[w]).getA1Notation();
            base = 'IF(' + rcA1 + '<>"",N(' + rcA1 + '),' + base + ')';
          }
        }
      }
      var inc = sh.getRange(incRow, c).getA1Notation();
      var out = sh.getRange(outRow, c).getA1Notation();
      var scr = sh.getRange(scrapRow, c).getA1Notation();
      var f = '=' + base + '+SUM(' + inc + ')-SUM(' + out + ')-SUM(' + scr + ')';
      assertCleanFormula(f);
      // очистить ячейку от старой мусорной формулы и записать чистую строку
      var cell = sh.getRange(balRow, c);
      cell.setValue('');
      cell.setFormula(f);
    }
    // недельные итоги = баланс воскресенья
    for (var w = 1; w <= week; w++) {
      var lastDay = getWeekLastDay(y, m, w, daysInMonth);
      var wf = '=' + sh.getRange(balRow, dayCols[lastDay]).getA1Notation();
      assertCleanFormula(wf);
      var wcell = sh.getRange(balRow, weekEnd[w]);
      wcell.setValue('');
      wcell.setFormula(wf);
    }
    // начальный остаток на строке баланса (переносится из прошлого месяца при создании)
  }
  sheet.setFrozenRows(2);
  sheet.setFrozenColumns(2);
}

function getWeekLastDay(y, m, w, daysInMonth) {
  var cnt = 0, last = 1;
  for (var d = 1; d <= daysInMonth; d++) {
    var dow = new Date(y, m - 1, d).getDay();
    if (dow === 0 && d < daysInMonth) { cnt++; if (cnt == w) return d; last = d; }
    if (d == daysInMonth) { if (cnt < w) return daysInMonth; }
  }
  return last;
}

/* ---------- Ремонт формул текущего месяца ---------- */

/** Чинит формулы на АКТИВНОМ листе месяца: снимает защиту, сохраняет ввод,
 *  очищает строки Баланса и колонки «Итог N» от мусорных формул, заново
 *  применяет раскладку/формулы, возвращает ввод и восстанавливает защиту.
 */
function repairCurrentMonthFormulas() {
  var sheet = SpreadsheetApp.getActive().getActiveSheet();
  var name = sheet.getName();
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(name)) {
    SpreadsheetApp.getUi().alert('Откройте лист месяца (ГГГГ-ММ)');
    return;
  }
  var items = getItems();
  if (!items.length) return;
  unprotectSheet(sheet);
  // сохранить пользовательский ввод
  var saved = saveInputs(sheet, items);
  var start = 3;
  var lastCol = Math.max(3, sheet.getLastColumn());
  // полностью очистить строки Баланса (формулы + мусор вида "...SUM(F5))2")
  for (var i = 0; i < items.length; i++) {
    var balRow = start + i * CFG.rowsPerItem + 3;
    sheet.getRange(balRow, 3, 1, lastCol - 2).clearContent();
  }
  // полностью очистить колонки «Итог N» (формулы/мусор во всех строках)
  var headers = sheet.getRange(2, 3, 1, lastCol - 2).getValues()[0];
  for (var c = 0; c < headers.length; c++) {
    if (String(headers[c]).indexOf('Итог') === 0) {
      sheet.getRange(start, c + 3, items.length * CFG.rowsPerItem, 1).clearContent();
    }
  }
  // заново применить раскладку и формулы
  applyMonthFormulas(sheet, name, items);
  // вернуть ввод (кроме ячеек Баланса вне колонок «Пересчёт»)
  restoreInputs(sheet, saved, items);
  applyIndicators(sheet);
  protectSheet(sheet);
  SpreadsheetApp.getUi().alert('Формулы листа «' + name + '» восстановлены.');
}

/* ---------- Справочник (лист) ---------- */

function buildReferenceSheet() {
  var ss = SpreadsheetApp.getActive();
  if (ss.getSheetByName(REF)) return;
  var sh = ss.insertSheet(REF);
  sh.getRange(1, 1, 1, 12).setValues([[
    'article_f12', 'article_A', 'name_1c', 'article_1c', 'length', 'width',
    'thickness', 'weight', 'density', 'scrap_material', 'group', 'min_balance'
  ]]).setFontWeight('bold');
  sh.setFrozenRows(1);
  // 23 артикула из исходного кода
  var data = [
['A3042B10N30X15','A3042B10N30X15','СТАЛЬ НЖ AISI-304 1,0*1500*3000 (БЕЗ ПЛЕНКИ), 2В','60438728231d',3000,1500,1.0,35.685,7.93,'ЛОМ НЕРЖАВЕЮЩЕЙ СТАЛИ AISI 304','AISI304',0],
['A3042B12N30X15','A3042B12N30X15','СТАЛЬ НЖ AISI-304 1,20*1500*3000 (БЕЗ ПЛЕНКИ), 2В','60438173931d',3000,1500,1.2,42.822,7.93,'ЛОМ НЕРЖАВЕЮЩЕЙ СТАЛИ AISI 304','AISI304',0],
['A3042B15N25X12','A3042B15P25X12','СТАЛЬ НЖ AISI-304 1,5*1250*2500 (БЕЗ ПЛЕНКИ), 2В','60438783931d',2500,1250,1.5,37.171,7.93,'ЛОМ НЕРЖАВЕЮЩЕЙ СТАЛИ AISI 304','AISI304',0],
['A3042B30N30X15','A3042B30P30X15','СТАЛЬ НЖ AISI-304 3*1500*3000 (БЕЗ ПЛЕНКИ), 2В','60438728531d',3000,1500,3.0,107.055,7.93,'ЛОМ НЕРЖАВЕЮЩЕЙ СТАЛИ AISI 304','AISI304',0],
['A304SB50N30X15','A304SB50N30X15','СТАЛЬ НЖ AISI-304 4N+PE 5,0*1500*3000 (С ПЛЕНКОЙ)','60442728631d',3000,1500,5.0,178.994,7.93,'ЛОМ НЕРЖАВЕЮЩЕЙ СТАЛИ AISI 304','AISI304',0],
['A304BA10P30X15','A304BA10P30X15','СТАЛЬ НЖ AISI-304 BA+PE 1,0*1500*3000 (зерк.)','60440728231d',3000,1500,1.0,35.685,7.93,'ЛОМ НЕРЖАВЕЮЩЕЙ СТАЛИ AISI 304','AISI304',0],
['A430BA08P30X15','A430BA08P30X15','СТАЛЬ НЖ AISI-430 ВА+РЕ 0,8*1500*3000 (зерк.)','60441728830d',3000,1500,0.8,28.469,7.75,'ЛОМ НЕРЖАВЕЮЩЕЙ СТАЛИ AISI 430','AISI430',0],
['A430SB08P30X10','A430SB08P30X10','СТАЛЬ НЖ AISI-430 4N+LРЕ 0,8*1000*3000','60444720831d',3000,1000,0.8,18.6,7.75,'ЛОМ НЕРЖАВЕЮЩЕЙ СТАЛИ AISI 430','AISI430',0],
['A430SB08P30X15','A430SB08P30X15','СТАЛЬ НЖ AISI-430 4N+РЕ 0,8*1500*3000','60444728830d',3000,1500,0.8,28.469,7.75,'ЛОМ НЕРЖАВЕЮЩЕЙ СТАЛИ AISI 430','AISI430',0],
['A430SB10P30X10','A430SB10P30X10','СТАЛЬ НЖ AISI-430 4N+РЕ 1,0*1000*3000','60444720231d',3000,1000,1.0,23.63,7.75,'ЛОМ НЕРЖАВЕЮЩЕЙ СТАЛИ AISI 430','AISI430',0],
['A430SB10P30X15','A430SB10P30X15','СТАЛЬ НЖ AISI-430 4N+РЕ 1,0*1500*3000','60444728230d',3000,1500,1.0,35.444,7.75,'ЛОМ НЕРЖАВЕЮЩЕЙ СТАЛИ AISI 430','AISI430',0],
['A430SB12P30X10','A430SB12P30X10','ЗАГОТОВКА 3000*1000*1,2(430)','2515361d',3000,1000,1.2,28.28,7.75,'ЛОМ НЕРЖАВЕЮЩЕЙ СТАЛИ AISI 430','AISI430',0],
['A430SB15P30X15','A430SB15P30X15','СТАЛЬ НЖ AISI-430 4N+РЕ 1,5*1500*3000','60444728930d',3000,1500,1.5,52.882,7.75,'ЛОМ НЕРЖАВЕЮЩЕЙ СТАЛИ AISI 430','AISI430',0],
['A430SB30P30X15','A430SB30P30X15','СТАЛЬ НЖ AISI-430 4N+РЕ 3,0*1500*3000','60444728531d',3000,1500,3.0,105.194,7.75,'ЛОМ НЕРЖАВЕЮЩЕЙ СТАЛИ AISI 430','AISI430',0],
['AL575410N30X15','AL575410N30X15','ЛИСТ АЛЮМИНИЕВ. 1,0*1500*3000 АМГ3М','60304053031d',3000,1500,1.0,11.97,2.66,'ЛОМ АЛЮМИНИЯ АМГ (отходы ленты)','AL',0],
['AL575412N30X15','AL575412N30X15','ЗАГОТОВКА 3000*1500*1,2(алюм)','2515362d',3000,1500,1.2,14.364,2.66,'ЛОМ АЛЮМИНИЯ АМГ (отходы ленты)','AL',0],
['FEZ14010N30X12','FEZ14010N30X12','ЗАГОТОВКА 3000*1250*1,0(оц)','2519414d',3000,1250,1.0,29.475,7.86,'ЛОМ ОЦИНКОВАННОЙ СТАЛИ','ZINC',0],
['FEZ14015N25X12','FEZ14015N25X12','ЗАГОТОВКА 2500*1250*1,5 (оц.сталь.)','2516071d',2500,1250,1.5,36.844,7.86,'ЛОМ ОЦИНКОВАННОЙ СТАЛИ','ZINC',0],
['FEZ14020N25X12','FEZ14020N25X12','ЛИСТ ОЦИНК. 2,00х1250х2500 ГОСТ 14918','60139310301d',2500,1250,2.0,49.125,7.86,'ЛОМ ОЦИНКОВАННОЙ СТАЛИ','ZINC',0],
['A304BA10P30X12','A304BA10P30X12','','',0,0,0,0,0,'ЛОМ НЕРЖАВЕЮЩЕЙ СТАЛИ AISI 304','AISI304',0],
['A430BA10P30X15','A430BA10P30X15','','',0,0,0,0,0,'ЛОМ НЕРЖАВЕЮЩЕЙ СТАЛИ AISI 430','AISI430',0],
['A430SB20P25X12','A430SB20P25X12','','',0,0,0,0,0,'ЛОМ НЕРЖАВЕЮЩЕЙ СТАЛИ AISI 430','AISI430',0],
['A304SB15P30X15','A304SB15P30X15','','',0,0,0,0,0,'ЛОМ НЕРЖАВЕЮЩЕЙ СТАЛИ AISI 304','AISI304',0]
  ];
  sh.getRange(2, 1, data.length, 12).setValues(data);
  protectSheet(sh);
}

/* ---------- Индикаторы ---------- */

function applyIndicators(sheet) {
  // Красный: баланс < порога; жёлтый: < порога*1.2; зелёный — иначе.
  // Порог берётся из Справочник!L (min_balance) по артикулу строки.
  var items = getItems();
  var rules = [];
  var start = 3;
  var lastCol = sheet.getLastColumn();
  for (var i = 0; i < items.length; i++) {
    var r = start + i * CFG.rowsPerItem + 3; // строка Баланс
    var refRow = items[i].srcRow || (i + 2); // реальная строка справочника
    var rng = sheet.getRange(r, 3, 1, lastCol - 2);
    var thr = "Справочник!$L$" + refRow;
    var yellow = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER($B$' + r + '),C' + r + '<>' + '"",C' + r + '<' + thr + '*1.2,C' + r + '>=' + thr + ')')
      .setBackground('#FFF2CC')
      .setRanges([rng])
      .build();
    var red = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER(C' + r + '),C' + r + '<' + thr + ')')
      .setBackground('#F4C7C3')
      .setRanges([rng])
      .build();
    rules.push(yellow, red);
  }
  sheet.setConditionalFormatRules(rules);
}

/* ---------- Защита ---------- */

function protectSheet(sheet) {
  var p = sheet.protect().setDescription('Журнал — формулы/итоги защищены');
  // разрешить редактирование только строк ввода (Приход/Расход/Лом) и «Пересчёт»
  var items = getItems();
  var start = 3;
  var unprotected = [];
  for (var i = 0; i < items.length; i++) {
    var r = start + i * CFG.rowsPerItem;
    unprotected.push(sheet.getRange(r, 3, 3, Math.max(1, sheet.getLastColumn() - 2))); // приход/расход/лом
  }
  // колонки «Пересчёт N» остаются редактируемыми: их адреса вычислим по заголовкам
  var headers = sheet.getRange(2, 3, 1, sheet.getLastColumn() - 2).getValues()[0];
  for (var c = 0; c < headers.length; c++) {
    if (String(headers[c]).indexOf('Пересчёт') === 0) {
      unprotected.push(sheet.getRange(3, c + 3, Math.max(1, items.length * CFG.rowsPerItem), 1));
    }
  }
  unprotected.push(sheet.getRange('B:B')); // начальные остатки
  p.setUnprotectedRanges(unprotected);
  try { p.removeEditors(p.getEditors()); } catch (e) {}
}

function unprotectSheet(sheet) {
  var p = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0];
  if (p) p.remove();
}

function unprotectWithPassword() {
  if (!askPassword()) { SpreadsheetApp.getUi().alert('Неверный пароль.'); return; }
  var sh = SpreadsheetApp.getActive().getActiveSheet();
  unprotectSheet(sh);
  SpreadsheetApp.getUi().alert('Защита снята с листа «' + sh.getName() + '». Восстановите её через меню после правок.');
}

function protectAllSheets() {
  var ss = SpreadsheetApp.getActive();
  ss.getSheets().forEach(function (sh) {
    if (/^\d{4}-\d{2}$/.test(sh.getName()) || sh.getName() == REF) protectSheet(sh);
  });
  SpreadsheetApp.getUi().alert('Защита восстановлена.');
}

/* ---------- Обновление всех листов месяцев (после правки справочника) ---------- */

function rebuildAllMonths() {
  var ss = SpreadsheetApp.getActive();
  var items = getItems();
  ss.getSheets().forEach(function (sh) {
    var nm = sh.getName();
    if (/^\d{4}-\d{2}$/.test(nm)) {
      // сохранить введённые операции, пересобрать формулы
      var saved = saveInputs(sh, items);
      applyMonthFormulas(sh, nm, items);
      restoreInputs(sh, saved, items);
      applyIndicators(sh);
      protectSheet(sh);
    }
  });
}

function saveInputs(sh, items) {
  var vals = {};
  var start = 3, lastCol = sh.getLastColumn();
  for (var i = 0; i < items.length * CFG.rowsPerItem; i++) {
    for (var c = 3; c <= lastCol; c++) {
      var v = sh.getRange(start + i, c).getValue();
      if (v !== '') vals[(start + i) + '_' + c] = v;
    }
  }
  return vals;
}

/** Возвращает сохранённый ввод: только строки Приход/Расход/Лом и колонки «Пересчёт».
 *  Ячейки строк Баланса вне колонок «Пересчёт» (формулы) никогда не перезаписываются.
 */
function restoreInputs(sh, vals, items) {
  var start = 3, lastCol = sh.getLastColumn();
  var headers = sh.getRange(2, 3, 1, Math.max(1, lastCol - 2)).getValues()[0];
  for (var key in vals) {
    var parts = key.split('_');
    var r = parseInt(parts[0]), c = parseInt(parts[1]);
    var rowOffset = r - start;
    if (rowOffset >= 0 && rowOffset % CFG.rowsPerItem === 3) {
      // строка Баланса: восстанавливаем только колонки «Пересчёт»
      var h = String(headers[c - 3]);
      if (h.indexOf('Пересчёт') !== 0) continue;
    }
    sh.getRange(r, c).setValue(vals[key]);
  }
}

/* ---------- Триггер нового месяца ---------- */

function ensureCurrentMonth() {
  var ss = SpreadsheetApp.getActive();
  var ym = Utilities.formatDate(new Date(), 'GMT+3', 'yyyy-MM');
  if (!ss.getSheetByName(ym)) createMonth(ym);
}