// ========================================
// 📋 社員マスタ読込（共通マスタ）
// ========================================

var EMPLOYEE_COLUMN_ALIASES = {
  id: ['社員ID', 'QRコード', 'ID'],
  name: ['氏名', '作業者名', '名前'],
  email: ['Email', 'メール', 'メールアドレス'],
  office: ['事業所', '工場', '営業所'],
  department: ['部署'],
  role: ['ロール', '権限'],
  active: ['有効', '稼働', 'ステータス']
};

function loadEmployeesFromSheet() {
  if (!MASTER_SS_ID) return [];
  var cacheKey = 'employees_' + MASTER_SS_ID;
  var cached = getCachedJson_(cacheKey);
  if (cached) return cached;

  try {
    var masterSs = SpreadsheetApp.openById(MASTER_SS_ID);
    var sheet = masterSs.getSheetByName(EMPLOYEE_MASTER_SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) {
      putCachedJson_(cacheKey, []);
      return [];
    }
    var list = parseEmployeeRowsFromSheet_(sheet);
    putCachedJson_(cacheKey, list);
    return list;
  } catch (e) {
    Logger.log('社員マスタ読込エラー: ' + e.message);
    return [];
  }
}

function parseEmployeeRowsFromSheet_(sheet) {
  var lastCol = Math.max(sheet.getLastColumn(), 9);
  var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function(h) { return String(h || '').trim(); });
  var colMap = mapEmployeeColumns_(headerRow);
  if (colMap.email === -1) return [];

  var data = sheet.getRange(2, 1, sheet.getLastRow(), lastCol).getValues();
  var list = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (colMap.active !== -1 && !isActiveEmployeeRow_(row[colMap.active])) continue;
    var email = String(row[colMap.email] || '').trim().toLowerCase();
    if (!email || email.indexOf('@') === -1) continue;
    var deptRaw = colMap.department !== -1 ? String(row[colMap.department] || '').trim() : '';
    var roleRaw = colMap.role !== -1 ? String(row[colMap.role] || '').trim() : '';
    list.push({
      id: colMap.id !== -1 ? String(row[colMap.id] || '').trim() : '',
      name: colMap.name !== -1 ? String(row[colMap.name] || '').trim() : email.split('@')[0],
      email: email,
      office: colMap.office !== -1 ? String(row[colMap.office] || '').trim() : '',
      department: deptRaw,
      departments: parseMultiValues_(deptRaw),
      role: roleRaw,
      roles: parseMultiValues_(roleRaw)
    });
  }
  return list;
}

function mapEmployeeColumns_(headers) {
  var map = {};
  Object.keys(EMPLOYEE_COLUMN_ALIASES).forEach(function(key) {
    map[key] = findColumnIndex_(headers, EMPLOYEE_COLUMN_ALIASES[key]);
  });
  return map;
}

function findColumnIndex_(headers, aliases) {
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i] || '').replace(/\s/g, '');
    for (var j = 0; j < aliases.length; j++) {
      if (h === aliases[j] || h.indexOf(aliases[j]) !== -1) return i;
    }
  }
  return -1;
}

function isActiveEmployeeRow_(val) {
  var s = String(val || '有効').trim();
  return s !== '無効' && s !== 'FALSE' && s !== '0' && s !== '停止';
}

function findEmployeeByEmail(email) {
  var key = String(email || '').trim().toLowerCase();
  if (!key) return null;
  var employees = loadEmployeesFromSheet();
  for (var i = 0; i < employees.length; i++) {
    if (employees[i].email === key) return employees[i];
  }
  return null;
}

function displayNameForEmail_(email) {
  var emp = findEmployeeByEmail(email);
  if (emp && emp.name) return emp.name;
  var key = String(email || '').trim();
  return key.indexOf('@') > -1 ? key.split('@')[0] : key;
}

function displayNamesForEmails_(emails) {
  return (emails || []).map(displayNameForEmail_);
}

function parseMultiValues_(value) {
  var raw = String(value || '').trim();
  if (!raw) return [];
  return raw.split(/[,、\/\n;|｜]+/).map(function(v) { return v.trim(); }).filter(function(v) { return !!v; });
}

function getEmployeeRoles_(employee) {
  if (!employee) return [];
  if (employee.roles && employee.roles.length) return employee.roles;
  return parseMultiValues_(employee.role);
}

function getEmployeeDepartments_(employee) {
  if (!employee) return [];
  if (employee.departments && employee.departments.length) return employee.departments;
  return parseMultiValues_(employee.department);
}

function employeeHasRole_(employee, targetRole) {
  targetRole = String(targetRole || '').trim();
  if (!targetRole || !employee) return false;
  var roles = getEmployeeRoles_(employee);
  for (var i = 0; i < roles.length; i++) {
    if (roles[i] === targetRole) return true;
  }
  return false;
}

function employeeMatchesDepartmentFilter_(employee, departmentFilter) {
  departmentFilter = String(departmentFilter || '').trim();
  if (!departmentFilter || !employee) return true;
  var filterDepts = parseMultiValues_(departmentFilter);
  var empDepts = getEmployeeDepartments_(employee);
  if (!filterDepts.length) return true;
  if (!empDepts.length) return false;
  for (var i = 0; i < filterDepts.length; i++) {
    for (var j = 0; j < empDepts.length; j++) {
      if (filterDepts[i] === empDepts[j]) return true;
    }
  }
  return false;
}

function employeeHasOrgApproverRole_(employee) {
  var roles = getEmployeeRoles_(employee);
  for (var i = 0; i < roles.length; i++) {
    if (wfIsOrgApproverRole_(roles[i])) return true;
  }
  return false;
}
