// ============================================
// СИСТЕМА УПРАВЛЕНИЯ ЗАДАЧАМИ — БЭКЕНД
// Google Apps Script
// ============================================

const SHEET_ID = '1TVKt3b_YTLQ0e1OQPconeRF94zP8UFOBeWuhtFAV2f0';
const ADMIN_PASSWORD = 'R*JIdyV45K3x';
const TZ = 'Europe/Moscow';

const SHEETS = {
  EMPLOYEES: 'employees',
  TASKS: 'tasks',
  CHECKINS: 'checkins'
};

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const action = data.action;
  try {
    switch (action) {
      case 'adminLogin':         return respond(adminLogin(data));
      case 'employeeLogin':      return respond(employeeLogin(data));
      case 'getEmployees':       return respond(getEmployees());
      case 'addEmployee':        return respond(addEmployee(data));
      case 'updateEmployee':     return respond(updateEmployee(data));
      case 'deactivateEmployee': return respond(deactivateEmployee(data));
      case 'getTasks':           return respond(getTasks(data));
      case 'addTask':            return respond(addTask(data));
      case 'completeTask':       return respond(completeTask(data));
      case 'getCheckins':        return respond(getCheckins(data));
      case 'addCheckin':         return respond(addCheckin(data));
      case 'getDashboard':       return respond(getDashboard());
      default:                   return respond({ ok: false, error: 'Неизвестное действие' });
    }
  } catch (err) {
    return respond({ ok: false, error: err.message });
  }
}

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function toDateStr(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, TZ, 'yyyy-MM-dd');
  }
  const s = String(val).trim();
  // Если это "Sun May 31 2026 ..." — парсим через Date
  if (s.length > 10 && !s.startsWith('20')) {
    try {
      return Utilities.formatDate(new Date(s), TZ, 'yyyy-MM-dd');
    } catch(e) {}
  }
  if (s.length >= 10 && s[4] === '-') return s.substring(0, 10);
  return s;
}

function toMonthStr(val) {
  const d = toDateStr(val);
  return d ? d.substring(0, 7) : '';
}

function getSheet(name) {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
}

function getCurrentMonth() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM');
}

function adminLogin(data) {
  if (data.password === ADMIN_PASSWORD) return { ok: true, role: 'admin' };
  return { ok: false, error: 'Неверный пароль' };
}

function employeeLogin(data) {
  const rows = getSheet(SHEETS.EMPLOYEES).getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const [id, name, password, object, active] = rows[i];
    if (
      String(name).trim() === String(data.name).trim() &&
      String(password).trim() === String(data.password).trim() &&
      active
    ) {
      return { ok: true, employee: { id: String(id), name: String(name), object: String(object) } };
    }
  }
  return { ok: false, error: 'Неверное имя или пароль' };
}

function getEmployees() {
  const rows = getSheet(SHEETS.EMPLOYEES).getDataRange().getValues();
  const employees = [];
  for (let i = 1; i < rows.length; i++) {
    const [id, name, password, object, active] = rows[i];
    if (active) employees.push({ id: String(id), name: String(name), object: String(object) });
  }
  return { ok: true, employees };
}

function addEmployee(data) {
  const id = 'emp_' + Date.now();
  getSheet(SHEETS.EMPLOYEES).appendRow([id, data.name, data.password, data.object, true]);
  return { ok: true, id };
}

function updateEmployee(data) {
  const sheet = getSheet(SHEETS.EMPLOYEES);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      if (data.name)     sheet.getRange(i + 1, 2).setValue(data.name);
      if (data.password) sheet.getRange(i + 1, 3).setValue(data.password);
      if (data.object)   sheet.getRange(i + 1, 4).setValue(data.object);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Сотрудник не найден' };
}

function deactivateEmployee(data) {
  const sheet = getSheet(SHEETS.EMPLOYEES);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      sheet.getRange(i + 1, 5).setValue(false);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Сотрудник не найден' };
}

function getTasks(data) {
  const rows = getSheet(SHEETS.TASKS).getDataRange().getValues();
  const tasks = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const id = String(row[0] || '').trim();
    if (!id) continue;

    const employeeId = String(row[1] || '').trim();
    const employeeName = String(row[2] || '');
    const title = String(row[3] || '');
    const deadline = toDateStr(row[4]);
    const month = deadline ? deadline.substring(0, 7) : toMonthStr(row[5]);
    const status = String(row[6] || 'pending');
    const comment = String(row[7] || '');
    const completedAt = row[8] instanceof Date ? row[8].toISOString() : String(row[8] || '');

    if (data.employeeId && employeeId !== String(data.employeeId)) continue;
    if (data.month && month !== data.month) continue;

    tasks.push({ id, employeeId, employeeName, title, deadline, month, status, comment, completedAt });
  }

  return { ok: true, tasks };
}

function addTask(data) {
  const id = 'task_' + Date.now();
  const month = data.deadline ? data.deadline.substring(0, 7) : getCurrentMonth();
  getSheet(SHEETS.TASKS).appendRow([
    id, data.employeeId, data.employeeName,
    data.title, data.deadline, month,
    'pending', '', ''
  ]);
  return { ok: true, id };
}

function completeTask(data) {
  const sheet = getSheet(SHEETS.TASKS);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.taskId)) {
      sheet.getRange(i + 1, 7).setValue('done');
      sheet.getRange(i + 1, 8).setValue(data.comment || '');
      sheet.getRange(i + 1, 9).setValue(new Date().toISOString());
      return { ok: true };
    }
  }
  return { ok: false, error: 'Задача не найдена' };
}

function getCheckins(data) {
  const rows = getSheet(SHEETS.CHECKINS).getDataRange().getValues();
  const checkins = [];
  for (let i = 1; i < rows.length; i++) {
    const [id, employeeId, employeeName, weekDate, done, willDo, needHelp, createdAt] = rows[i];
    if (!id) continue;
    if (data.employeeId && String(employeeId) !== String(data.employeeId)) continue;
    checkins.push({
      id: String(id), employeeId: String(employeeId),
      employeeName: String(employeeName),
      weekDate: String(weekDate), done: String(done),
      willDo: String(willDo), needHelp: String(needHelp || ''),
      createdAt: createdAt instanceof Date ? createdAt.toISOString() : String(createdAt || '')
    });
  }
  return { ok: true, checkins };
}

function addCheckin(data) {
  const id = 'ci_' + Date.now();
  getSheet(SHEETS.CHECKINS).appendRow([
    id, data.employeeId, data.employeeName,
    data.weekDate, data.done, data.willDo,
    data.needHelp || '', new Date().toISOString()
  ]);
  return { ok: true, id };
}

function getDashboard() {
  const employees = getSheet(SHEETS.EMPLOYEES).getDataRange().getValues().slice(1)
    .filter(r => r[4])
    .map(r => ({ id: String(r[0]), name: String(r[1]), object: String(r[3]) }));

  const taskRows = getSheet(SHEETS.TASKS).getDataRange().getValues().slice(1);
  const currentMonth = getCurrentMonth();
  const now = new Date();

  const stats = employees.map(emp => {
    const empTasks = taskRows
      .filter(r => {
        if (!r[0]) return false;
        const deadline = toDateStr(r[4]);
        const month = deadline ? deadline.substring(0, 7) : toMonthStr(r[5]);
        return String(r[1]) === emp.id && month === currentMonth;
      })
      .map(r => ({
        status: String(r[6] || 'pending'),
        deadline: toDateStr(r[4])
      }));

    return {
      ...emp,
      total: empTasks.length,
      done: empTasks.filter(t => t.status === 'done').length,
      pending: empTasks.filter(t => t.status !== 'done').length,
      overdue: empTasks.filter(t => t.status !== 'done' && t.deadline && new Date(t.deadline) < now).length
    };
  });

  return { ok: true, stats, currentMonth };
}

function initSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  let empSheet = ss.getSheetByName(SHEETS.EMPLOYEES) || ss.insertSheet(SHEETS.EMPLOYEES);
  if (empSheet.getLastRow() === 0) {
    empSheet.appendRow(['id', 'name', 'password', 'object', 'active']);
    empSheet.appendRow(['emp_1', 'Константин', '1234', 'Княже', true]);
    empSheet.appendRow(['emp_2', 'Андрей', '1234', 'Таёжный', true]);
    empSheet.appendRow(['emp_3', 'Валерия', '1234', 'Оба', true]);
  }

  let taskSheet = ss.getSheetByName(SHEETS.TASKS) || ss.insertSheet(SHEETS.TASKS);
  if (taskSheet.getLastRow() === 0) {
    taskSheet.appendRow(['id', 'employeeId', 'employeeName', 'title', 'deadline', 'month', 'status', 'comment', 'completedAt']);
  }

  let ciSheet = ss.getSheetByName(SHEETS.CHECKINS) || ss.insertSheet(SHEETS.CHECKINS);
  if (ciSheet.getLastRow() === 0) {
    ciSheet.appendRow(['id', 'employeeId', 'employeeName', 'weekDate', 'done', 'willDo', 'needHelp', 'createdAt']);
  }

  return 'OK';
}

function testGetTasks2() {
  const result = getTasks({});
  Logger.log(JSON.stringify(result));
}
