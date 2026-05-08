// ============================================
// СИСТЕМА УПРАВЛЕНИЯ ЗАДАЧАМИ — БЭКЕНД
// Google Apps Script
// ============================================

const SHEET_ID = 'ВСТАВЬ_ID_СВОЕЙ_ТАБЛИЦЫ'; // Заменить на ID твоей Google Sheets
const ADMIN_PASSWORD = 'admin2024'; // Пароль руководителя — поменяй на свой

// ---- Названия листов ----
const SHEETS = {
  EMPLOYEES: 'employees',
  TASKS: 'tasks',
  CHECKINS: 'checkins'
};

// ============================================
// ТОЧКА ВХОДА — обработка всех запросов
// ============================================

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const action = data.action;

  try {
    switch (action) {
      case 'adminLogin':       return respond(adminLogin(data));
      case 'employeeLogin':    return respond(employeeLogin(data));
      case 'getEmployees':     return respond(getEmployees());
      case 'addEmployee':      return respond(addEmployee(data));
      case 'updateEmployee':   return respond(updateEmployee(data));
      case 'deactivateEmployee': return respond(deactivateEmployee(data));
      case 'getTasks':         return respond(getTasks(data));
      case 'addTask':          return respond(addTask(data));
      case 'completeTask':     return respond(completeTask(data));
      case 'getCheckins':      return respond(getCheckins(data));
      case 'addCheckin':       return respond(addCheckin(data));
      case 'getDashboard':     return respond(getDashboard());
      default:                 return respond({ ok: false, error: 'Неизвестное действие' });
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

// ============================================
// АВТОРИЗАЦИЯ
// ============================================

function adminLogin(data) {
  if (data.password === ADMIN_PASSWORD) {
    return { ok: true, role: 'admin' };
  }
  return { ok: false, error: 'Неверный пароль' };
}

function employeeLogin(data) {
  const sheet = getSheet(SHEETS.EMPLOYEES);
  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    const [id, name, password, object, active] = rows[i];
    if (name === data.name && password === data.password && active === true) {
      return { ok: true, employee: { id, name, object } };
    }
  }
  return { ok: false, error: 'Неверное имя или пароль' };
}

// ============================================
// СОТРУДНИКИ
// ============================================

function getEmployees() {
  const sheet = getSheet(SHEETS.EMPLOYEES);
  const rows = sheet.getDataRange().getValues();
  const employees = [];

  for (let i = 1; i < rows.length; i++) {
    const [id, name, password, object, active] = rows[i];
    if (active) {
      employees.push({ id, name, object, active });
    }
  }
  return { ok: true, employees };
}

function addEmployee(data) {
  const sheet = getSheet(SHEETS.EMPLOYEES);
  const id = 'emp_' + Date.now();
  sheet.appendRow([id, data.name, data.password, data.object, true]);
  return { ok: true, id };
}

function updateEmployee(data) {
  const sheet = getSheet(SHEETS.EMPLOYEES);
  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.id) {
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
    if (rows[i][0] === data.id) {
      sheet.getRange(i + 1, 5).setValue(false);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Сотрудник не найден' };
}

// ============================================
// ЗАДАЧИ
// ============================================

function getTasks(data) {
  const sheet = getSheet(SHEETS.TASKS);
  const rows = sheet.getDataRange().getValues();
  let tasks = [];

  for (let i = 1; i < rows.length; i++) {
    const [id, employeeId, employeeName, title, deadline, month, status, comment, completedAt] = rows[i];
    const task = { id, employeeId, employeeName, title, deadline, month, status, comment, completedAt };

    // Фильтр по сотруднику (для страницы сотрудника)
    if (data.employeeId && employeeId !== data.employeeId) continue;

    // Фильтр по месяцу
    if (data.month && month !== data.month) continue;

    tasks.push(task);
  }

  return { ok: true, tasks };
}

function addTask(data) {
  const sheet = getSheet(SHEETS.TASKS);
  const id = 'task_' + Date.now();
  const month = data.deadline ? data.deadline.substring(0, 7) : getCurrentMonth();

  sheet.appendRow([
    id,
    data.employeeId,
    data.employeeName,
    data.title,
    data.deadline,
    month,
    'pending',   // статус: pending / in_progress / done
    '',          // комментарий
    ''           // дата выполнения
  ]);

  return { ok: true, id };
}

function completeTask(data) {
  const sheet = getSheet(SHEETS.TASKS);
  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.taskId) {
      sheet.getRange(i + 1, 7).setValue('done');
      sheet.getRange(i + 1, 8).setValue(data.comment || '');
      sheet.getRange(i + 1, 9).setValue(new Date().toISOString());
      return { ok: true };
    }
  }
  return { ok: false, error: 'Задача не найдена' };
}

// ============================================
// ЧЕКПОИНТЫ (еженедельные)
// ============================================

function getCheckins(data) {
  const sheet = getSheet(SHEETS.CHECKINS);
  const rows = sheet.getDataRange().getValues();
  let checkins = [];

  for (let i = 1; i < rows.length; i++) {
    const [id, employeeId, employeeName, weekDate, done, willDo, needHelp, createdAt] = rows[i];
    if (data.employeeId && employeeId !== data.employeeId) continue;
    checkins.push({ id, employeeId, employeeName, weekDate, done, willDo, needHelp, createdAt });
  }

  return { ok: true, checkins };
}

function addCheckin(data) {
  const sheet = getSheet(SHEETS.CHECKINS);
  const id = 'ci_' + Date.now();

  sheet.appendRow([
    id,
    data.employeeId,
    data.employeeName,
    data.weekDate,
    data.done,
    data.willDo,
    data.needHelp || '',
    new Date().toISOString()
  ]);

  return { ok: true, id };
}

// ============================================
// ДАШБОРД — сводные данные для руководителя
// ============================================

function getDashboard() {
  const empSheet = getSheet(SHEETS.EMPLOYEES);
  const taskSheet = getSheet(SHEETS.TASKS);

  const employees = empSheet.getDataRange().getValues().slice(1)
    .filter(r => r[4] === true)
    .map(r => ({ id: r[0], name: r[1], object: r[3] }));

  const tasks = taskSheet.getDataRange().getValues().slice(1)
    .map(r => ({
      id: r[0], employeeId: r[1], employeeName: r[2],
      title: r[3], deadline: r[4], month: r[5],
      status: r[6], comment: r[7], completedAt: r[8]
    }));

  const currentMonth = getCurrentMonth();
  const currentTasks = tasks.filter(t => t.month === currentMonth);

  const stats = employees.map(emp => {
    const empTasks = currentTasks.filter(t => t.employeeId === emp.id);
    return {
      ...emp,
      total: empTasks.length,
      done: empTasks.filter(t => t.status === 'done').length,
      pending: empTasks.filter(t => t.status === 'pending').length,
      overdue: empTasks.filter(t => t.status !== 'done' && t.deadline && new Date(t.deadline) < new Date()).length
    };
  });

  return { ok: true, stats, currentMonth };
}

// ============================================
// ИНИЦИАЛИЗАЦИЯ ТАБЛИЦЫ (запусти один раз!)
// ============================================

function initSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // Employees
  let empSheet = ss.getSheetByName(SHEETS.EMPLOYEES) || ss.insertSheet(SHEETS.EMPLOYEES);
  if (empSheet.getLastRow() === 0) {
    empSheet.appendRow(['id', 'name', 'password', 'object', 'active']);
    // Добавим тестового сотрудника
    empSheet.appendRow(['emp_1', 'Константин', '1234', 'Княже', true]);
    empSheet.appendRow(['emp_2', 'Андрей', '1234', 'Таёжный', true]);
    empSheet.appendRow(['emp_3', 'Валерия', '1234', 'Оба', true]);
  }

  // Tasks
  let taskSheet = ss.getSheetByName(SHEETS.TASKS) || ss.insertSheet(SHEETS.TASKS);
  if (taskSheet.getLastRow() === 0) {
    taskSheet.appendRow(['id', 'employeeId', 'employeeName', 'title', 'deadline', 'month', 'status', 'comment', 'completedAt']);
  }

  // Checkins
  let ciSheet = ss.getSheetByName(SHEETS.CHECKINS) || ss.insertSheet(SHEETS.CHECKINS);
  if (ciSheet.getLastRow() === 0) {
    ciSheet.appendRow(['id', 'employeeId', 'employeeName', 'weekDate', 'done', 'willDo', 'needHelp', 'createdAt']);
  }

  return 'Таблица инициализирована успешно!';
}

// ============================================
// УТИЛИТЫ
// ============================================

function getSheet(name) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  return ss.getSheetByName(name);
}

function getCurrentMonth() {
  const now = new Date();
  return now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0');
}
