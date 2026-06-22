/**
 * ระบบลาออนไลน์ผ่าน LINE
 * Google Apps Script + Google Sheets + LINE Messaging API + LIFF
 * ------------------------------------------------------------
 * Phase 0 — Foundation
 *   - CONFIG / SHEETS
 *   - initializeSheets() : 11 ชีต (JSON-per-row) + seed users + leave types
 *   - Auth / Session (UUID token)
 *   - doGet router (web backend + liff)
 * ------------------------------------------------------------
 */

// ============================================================
// CONFIG
// ============================================================
const CONFIG = {
  APP_NAME: 'ระบบลาออนไลน์',
  APP_VERSION: '1.0',
  SESSION_TIMEOUT: 28800,   // วินาที (8 ชั่วโมง)
  FOLDER_ID: '',            // โฟลเดอร์ Drive สำหรับอัพโหลดโลโก้/ไฟล์แนบ (ตั้งใน Settings)

  // บัญชีเริ่มต้นฝั่ง Web  (username : password)
  ADMIN_USERS: {
    admin: 'admin123',
    hr: 'hr123',
    supervisor: 'sup123'
  },

  // บทบาท + สิทธิ์การเข้าถึงเมนู
  USER_ROLES: {
    admin: {
      name: 'ผู้ดูแลระบบ',
      permissions: ['dashboard', 'approval', 'leave', 'calendar', 'employee', 'master', 'report', 'settings', 'users', 'help']
    },
    hr: {
      name: 'ฝ่ายบุคคล (HR)',
      permissions: ['dashboard', 'approval', 'leave', 'calendar', 'employee', 'master', 'report', 'help']
    },
    supervisor: {
      name: 'หัวหน้างาน',
      permissions: ['dashboard', 'approval', 'leave', 'calendar', 'help']
    }
  },

  TELEGRAM_BOT_TOKEN: '',
  TELEGRAM_CHAT_ID: ''
};

// ชื่อชีต -> header column (1 JSON object ต่อ 1 แถว)
const SHEETS = {
  Config:        'config_json',
  Users:         'user_json',
  Employees:     'employee_json',
  Departments:   'department_json',
  LeaveTypes:    'leavetype_json',
  LeaveBalances: 'balance_json',
  LeaveRequests: 'leave_json',
  Holidays:      'holiday_json',
  Sessions:      'session_json',
  Errors:        'error_json',
  Notifications: 'notification_json'
};

// ============================================================
// WEB ENTRY
// ============================================================
function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'app';

  // ฝั่ง LIFF (พนักงาน) — เนื้อหาเต็มจะเติมใน Phase 5
  if (page === 'liff') {
    return HtmlService.createTemplateFromFile('liff')
      .evaluate()
      .setTitle(CONFIG.APP_NAME)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // ฝั่ง Web backend
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle(CONFIG.APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================
// SHEET HELPERS
// ============================================================
function _ss() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function _sheet(name) {
  var sh = _ss().getSheetByName(name);
  if (!sh) throw new Error('ไม่พบชีต: ' + name);
  return sh;
}

/** อ่านทุกแถว -> [{rowIndex, data}] */
function _readAll(name) {
  var sh = _ss().getSheetByName(name);
  if (!sh) return [];
  var values = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var cell = values[i][0];
    if (cell === '' || cell === null) continue;
    try { out.push({ rowIndex: i + 1, data: JSON.parse(cell) }); } catch (err) {}
  }
  return out;
}

/** เพิ่ม 1 แถว */
function _append(name, obj) {
  _sheet(name).appendRow([JSON.stringify(obj)]);
  return obj;
}

/** อัพเดทแถวตาม rowIndex */
function _update(name, rowIndex, obj) {
  _sheet(name).getRange(rowIndex, 1).setValue(JSON.stringify(obj));
  return obj;
}

/** ลบแถว */
function _delete(name, rowIndex) {
  _sheet(name).deleteRow(rowIndex);
}

function _now() { return new Date().toISOString(); }
function _uuid() { return Utilities.getUuid(); }

// ============================================================
// INITIALIZE SHEETS
// ============================================================
function initializeSheets() {
  try {
    var ss = _ss();
    var existing = ss.getSheets().map(function (s) { return s.getName(); });

    // 1) สร้างทุกชีตที่ยังไม่มี พร้อม header column
    Object.keys(SHEETS).forEach(function (name) {
      if (existing.indexOf(name) === -1) {
        var sh = ss.insertSheet(name);
        sh.appendRow([SHEETS[name]]);
        sh.setFrozenRows(1);
        sh.getRange(1, 1).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
        sh.setColumnWidth(1, 720);
      }
    });

    // ลบ Sheet1 default ถ้ายังว่างอยู่
    var s1 = ss.getSheetByName('Sheet1');
    if (s1 && ss.getSheets().length > 1) {
      try { ss.deleteSheet(s1); } catch (e) {}
    }

    // 2) Config (1 แถว)
    if (_readAll('Config').length === 0) {
      _append('Config', {
        app_name: CONFIG.APP_NAME,
        app_version: CONFIG.APP_VERSION,
        line_channel_access_token: '',
        line_channel_secret: '',
        line_liff_id: '',
        folder_id: CONFIG.FOLDER_ID,
        logo_file_id: '',
        approval_levels: 2,
        level1_role: 'supervisor',
        level2_role: 'hr',
        work_days: ['MO', 'TU', 'WE', 'TH', 'FR'],
        work_hours: '08:00-17:00',
        half_day_enabled: true,
        fiscal_start_month: 1,
        buddhist_era: true,
        notification_enabled: true,
        email_notifications: false,
        telegram_bot_token: CONFIG.TELEGRAM_BOT_TOKEN,
        telegram_chat_id: CONFIG.TELEGRAM_CHAT_ID,
        session_timeout: CONFIG.SESSION_TIMEOUT,
        maintenance_mode: false,
        created_at: _now(),
        updated_at: _now()
      });
    }

    // 3) Users (seed จาก CONFIG.ADMIN_USERS)
    if (_readAll('Users').length === 0) {
      Object.keys(CONFIG.ADMIN_USERS).forEach(function (username) {
        var role = CONFIG.USER_ROLES[username] || { name: username, permissions: [] };
        _append('Users', {
          id: _uuid(),
          username: username,
          password: CONFIG.ADMIN_USERS[username],
          role: username,
          name: role.name,
          email: '',
          line_user_id: '',
          employee_id: '',
          permissions: role.permissions,
          active: true,
          last_login: '',
          created_at: _now(),
          updated_at: _now()
        });
      });
    } else {
      try { migrateOldUsers(); } catch (err) { logError(err.toString(), 'migrateOldUsers'); }
    }

    // 4) LeaveTypes (ค่าตั้งต้น)
    if (_readAll('LeaveTypes').length === 0) {
      var defaults = [
        { name: 'ลาป่วย',             code: 'SICK',       color: '#ef4444', quota: 30,  attach: true  },
        { name: 'ลากิจส่วนตัว',        code: 'PERSONAL',   color: '#f59e0b', quota: 10,  attach: false },
        { name: 'ลาพักผ่อนประจำปี',    code: 'ANNUAL',     color: '#10b981', quota: 10,  attach: false },
        { name: 'ลาคลอดบุตร',         code: 'MATERNITY',  color: '#ec4899', quota: 90,  attach: true  },
        { name: 'ลาอุปสมบท',          code: 'ORDINATION', color: '#8b5cf6', quota: 120, attach: true  },
        { name: 'ลาอื่นๆ',            code: 'OTHER',      color: '#64748b', quota: 0,   attach: false }
      ];
      defaults.forEach(function (d, i) {
        _append('LeaveTypes', {
          id: _uuid(),
          name: d.name,
          code: d.code,
          color: d.color,
          default_quota_days: d.quota,
          max_per_request: 0,
          requires_attachment: d.attach,
          allow_half_day: true,
          paid: true,
          gender_restrict: '',
          carry_over: false,
          order: i + 1,
          active: true,
          created_at: _now()
        });
      });
    }

    return { status: 'success', message: 'สร้างชีตและข้อมูลเริ่มต้นเรียบร้อยแล้ว' };
  } catch (err) {
    return { status: 'error', message: 'เกิดข้อผิดพลาด: ' + err.toString() };
  }
}

function migrateOldUsers() {
  var users = _readAll('Users');
  users.forEach(function (u) {
    var d = u.data, changed = false;
    if (d.permissions === undefined) {
      var r = CONFIG.USER_ROLES[d.role];
      d.permissions = r ? r.permissions : [];
      changed = true;
    }
    if (d.line_user_id === undefined) { d.line_user_id = ''; changed = true; }
    if (d.employee_id === undefined) { d.employee_id = ''; changed = true; }
    if (changed) { d.updated_at = _now(); _update('Users', u.rowIndex, d); }
  });
  return { status: 'success', message: 'migrate users สำเร็จ' };
}

// ============================================================
// AUTH / SESSION
// ============================================================
function login(username, password) {
  try {
    if (!username || !password) return { status: 'error', message: 'กรุณากรอกข้อมูลให้ครบ' };

    var users = _readAll('Users');
    var found = null;
    var uname = String(username).trim();
    for (var i = 0; i < users.length; i++) {
      if (users[i].data.username === uname) { found = users[i]; break; }
    }
    if (!found) return { status: 'error', message: 'ไม่พบบัญชีผู้ใช้นี้' };
    if (!found.data.active) return { status: 'error', message: 'บัญชีนี้ถูกระงับการใช้งาน' };
    if (found.data.password !== password) return { status: 'error', message: 'รหัสผ่านไม่ถูกต้อง' };

    var rolePerms = (CONFIG.USER_ROLES[found.data.role] || {}).permissions || found.data.permissions || [];
    var now = new Date();
    var token = _uuid();
    _append('Sessions', {
      id: _uuid(),
      token: token,
      user_id: found.data.id,
      username: found.data.username,
      role: found.data.role,
      name: found.data.name,
      permissions: rolePerms,
      created_at: now.toISOString(),
      expires_at: new Date(now.getTime() + CONFIG.SESSION_TIMEOUT * 1000).toISOString()
    });

    found.data.permissions = rolePerms;
    found.data.last_login = now.toISOString();
    _update('Users', found.rowIndex, found.data);

    return {
      status: 'success',
      message: 'เข้าสู่ระบบสำเร็จ',
      token: token,
      user: {
        id: found.data.id,
        name: found.data.name,
        username: found.data.username,
        role: found.data.role,
        permissions: rolePerms
      }
    };
  } catch (err) {
    logError(err.toString(), 'login');
    return { status: 'error', message: 'เกิดข้อผิดพลาด: ' + err.toString() };
  }
}

function checkSession(token) {
  try {
    if (!token) return { status: 'error', message: 'ไม่พบ token' };
    var sessions = _readAll('Sessions');
    for (var i = 0; i < sessions.length; i++) {
      if (sessions[i].data.token === token) {
        var s = sessions[i].data;
        if (new Date(s.expires_at) < new Date()) {
          _delete('Sessions', sessions[i].rowIndex);
          return { status: 'error', message: 'เซสชันหมดอายุ' };
        }
        return {
          status: 'success',
          user: { id: s.user_id, name: s.name, username: s.username, role: s.role, permissions: s.permissions }
        };
      }
    }
    return { status: 'error', message: 'เซสชันไม่ถูกต้อง' };
  } catch (err) {
    logError(err.toString(), 'checkSession');
    return { status: 'error', message: 'เกิดข้อผิดพลาด' };
  }
}

function logout(token) {
  try {
    var sessions = _readAll('Sessions');
    for (var i = 0; i < sessions.length; i++) {
      if (sessions[i].data.token === token) { _delete('Sessions', sessions[i].rowIndex); break; }
    }
    return { status: 'success', message: 'ออกจากระบบแล้ว' };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

/** ตรวจ token แล้วคืน user — โยน error ถ้าไม่ผ่าน (ใช้ guard ฟังก์ชันที่ต้องล็อกอิน) */
function _auth(token) {
  var r = checkSession(token);
  if (r.status !== 'success') throw new Error('UNAUTHORIZED');
  return r.user;
}

// ============================================================
// APP INFO / DASHBOARD (Phase 0)
// ============================================================
function getAppInfo() {
  try {
    var cfg = _readAll('Config');
    var c = cfg.length ? cfg[0].data : {};
    var logoUrl = '';
    if (c.logo_file_id) {
      logoUrl = 'https://drive.google.com/thumbnail?id=' + c.logo_file_id + '&sz=s200';
    }
    return {
      status: 'success',
      app_name: c.app_name || CONFIG.APP_NAME,
      app_version: c.app_version || CONFIG.APP_VERSION,
      logo_url: logoUrl,
      initialized: cfg.length > 0,
      buddhist_era: c.buddhist_era !== false
    };
  } catch (err) {
    return {
      status: 'success',
      app_name: CONFIG.APP_NAME,
      app_version: CONFIG.APP_VERSION,
      logo_url: '',
      initialized: false
    };
  }
}

function getDashboardSummary(token) {
  var user = _auth(token);
  var today = new Date().toISOString().substring(0, 10);
  var thisMonth = today.substring(0, 7);
  var cfg = _getConfigRow().data;
  var deptFilter = _supervisorDeptIds(user);

  var leaves = _readAll('LeaveRequests').map(function(r){ return r.data; });
  if (deptFilter) leaves = leaves.filter(function(d){ return deptFilter.indexOf(d.department_id) !== -1; });

  var pending = leaves.filter(function(d){
    if (d.status === 'SUBMITTED') return _canApprove(user.role, 1, cfg);
    if (d.status === 'L1_APPROVED') return _canApprove(user.role, 2, cfg);
    return false;
  }).length;

  var onLeaveToday = leaves.filter(function(d){
    return d.status === 'APPROVED' && d.start_date <= today && d.end_date >= today;
  }).length;

  var thisMonthLeaves = leaves.filter(function(d){
    return d.status === 'APPROVED' && d.start_date.substring(0, 7) === thisMonth;
  }).length;

  var totalEmps = _readAll('Employees').filter(function(r){ return r.data.status !== 'inactive'; }).length;

  return {
    status: 'success',
    pending_approval: pending,
    on_leave_today: onLeaveToday,
    leave_this_month: thisMonthLeaves,
    total_employees: totalEmps
  };
}

/** คืน array ของ department_id ที่ supervisor คนนี้ดูแล (null = ไม่กรอง) */
function _supervisorDeptIds(user) {
  if (user.role !== 'supervisor') return null;
  var ids = [];
  _readAll('Departments').forEach(function(r){
    var d = r.data;
    var heads = Array.isArray(d.head_user_ids) ? d.head_user_ids : (d.head_user_id ? [d.head_user_id] : []);
    if (heads.indexOf(user.id) !== -1) ids.push(d.id);
  });
  return ids.length ? ids : null;
}

// ============================================================
// ERROR LOG
// ============================================================
function logError(message, context) {
  try {
    _append('Errors', {
      id: _uuid(),
      message: message,
      context: context || '',
      user: '',
      timestamp: _now()
    });
  } catch (e) {}
}

// ============================================================
// SETUP — รันครั้งเดียวจาก Apps Script editor (เลือกฟังก์ชัน setup แล้วกด Run)
// ============================================================
function setup() {
  var r = initializeSheets();
  Logger.log(r.message);
  return r;
}

// ============================================================
// ============================================================
// CHANGE PASSWORD (self)
// ============================================================
function changePassword(token, currentPw, newPw) {
  var user = _auth(token);
  try {
    if (!currentPw || !newPw) return { status: 'error', message: 'กรุณากรอกข้อมูลให้ครบ' };
    if (newPw.length < 6) return { status: 'error', message: 'รหัสผ่านต้องมีอย่างน้อย 6 ตัวอักษร' };
    var users = _readAll('Users');
    for (var i = 0; i < users.length; i++) {
      if (users[i].data.username === user.username) {
        if (users[i].data.password !== currentPw) return { status: 'error', message: 'รหัสผ่านปัจจุบันไม่ถูกต้อง' };
        users[i].data.password = newPw;
        users[i].data.updated_at = _now();
        _update('Users', users[i].rowIndex, users[i].data);
        return { status: 'success', message: 'เปลี่ยนรหัสผ่านเรียบร้อยแล้ว' };
      }
    }
    return { status: 'error', message: 'ไม่พบบัญชีผู้ใช้' };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

// ============================================================
// REMIND PENDING APPROVALS (admin/hr trigger)
// ============================================================
function remindPendingApprovals() {
  try {
    var cfg = _getConfigRow().data;
    if (!cfg.notification_enabled) return;
    var cutoff = new Date(Date.now() - 2 * 24 * 60 * 60 * 1000).toISOString();
    var pending = _readAll('LeaveRequests').map(function(r){ return r.data; })
      .filter(function(d){ return (d.status === 'SUBMITTED' || d.status === 'L1_APPROVED') && d.submitted_at < cutoff; });
    if (!pending.length) return;
    var ids = [];
    _readAll('Users').forEach(function(r){
      var u = r.data;
      if (u.active && u.line_user_id && (u.role === 'admin' || u.role === 'hr' || u.role === 'supervisor')) {
        ids.push(u.line_user_id);
      }
    });
    if (!ids.length) return;
    var msg = {
      type: 'flex', altText: 'มีใบลารออนุมัติ ' + pending.length + ' รายการ',
      contents: {
        type: 'bubble',
        header: { type: 'box', layout: 'vertical', backgroundColor: '#f59e0b', paddingAll: '14px',
          contents: [{ type: 'text', text: 'แจ้งเตือน: ใบลาค้างอนุมัติ', color: '#ffffff', weight: 'bold' }] },
        body: { type: 'box', layout: 'vertical', spacing: 'sm',
          contents: [{ type: 'text', text: 'มีใบลาที่รออนุมัติเกิน 2 วัน จำนวน ' + pending.length + ' รายการ กรุณาตรวจสอบ', wrap: true }] }
      }
    };
    ids.forEach(function(id){ _linePush(id, [msg]); });
  } catch(err) { logError(err.toString(), 'remindPendingApprovals'); }
}

function setupRemindTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'remindPendingApprovals') ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger('remindPendingApprovals').timeBased().everyDays(1).atHour(9).create();
  return { status: 'success', message: 'ตั้งค่าการแจ้งเตือนอัตโนมัติทุกวัน 9:00 น.' };
}

// ============================================================
// IMPORT EMPLOYEES FROM CSV
// ============================================================
function importEmployeesCSV(token, rows) {
  _auth(token);
  try {
    var depts = {};
    _readAll('Departments').forEach(function(r){ depts[r.data.name] = r.data.id; });
    var ok = 0, skip = 0, errors = [];
    rows.forEach(function(row, idx) {
      try {
        var firstName = (row.first_name || row['ชื่อ'] || '').trim();
        var lastName  = (row.last_name  || row['นามสกุล'] || '').trim();
        if (!firstName || !lastName) { skip++; return; }
        var deptName = (row.department || row['แผนก'] || '').trim();
        _append('Employees', {
          id: _uuid(),
          prefix: row.prefix || row['คำนำหน้า'] || 'นาย',
          first_name: firstName, last_name: lastName,
          nickname: row.nickname || row['ชื่อเล่น'] || '',
          emp_code: row.emp_code || row['รหัสพนักงาน'] || '',
          department_id: depts[deptName] || '',
          position: row.position || row['ตำแหน่ง'] || '',
          employment_type: row.employment_type || row['ประเภท'] || 'ประจำ',
          phone: row.phone || row['เบอร์โทร'] || '',
          email: row.email || row['อีเมล'] || '',
          start_date: row.start_date || row['วันเริ่มงาน'] || '',
          status: 'active', line_user_id: '', line_linked: false,
          photo_file_id: '', supervisor_id: '',
          created_at: _now(), updated_at: _now()
        });
        ok++;
      } catch(e) { errors.push('แถว ' + (idx+2) + ': ' + e.toString()); }
    });
    return { status: 'success', message: 'นำเข้าสำเร็จ ' + ok + ' คน' + (skip ? ' (ข้าม ' + skip + ' แถว)' : ''), errors: errors };
  } catch(err) {
    return { status: 'error', message: err.toString() };
  }
}

// ============================================================
// FULL MONTHLY CALENDAR
// ============================================================
function getMonthlyCalendar(token, year, month) {
  var user = _auth(token);
  var deptFilter = _supervisorDeptIds(user);
  var y = parseInt(year, 10), m = parseInt(month, 10);
  var from = y + '-' + (m<10?'0'+m:m) + '-01';
  var lastDay = new Date(y, m, 0).getDate();
  var to = y + '-' + (m<10?'0'+m:m) + '-' + lastDay;

  var empMap = {};
  _readAll('Employees').forEach(function(r){ empMap[r.data.id] = r.data; });

  var data = _readAll('LeaveRequests').map(function(r){ return r.data; })
    .filter(function(d){
      if (d.status !== 'APPROVED') return false;
      if (d.end_date < from || d.start_date > to) return false;
      if (deptFilter && deptFilter.indexOf(d.department_id) === -1) return false;
      return true;
    })
    .map(function(d){
      var emp = empMap[d.employee_id] || {};
      return {
        id: d.id, request_no: d.request_no,
        employee_name: d.employee_name,
        department_id: d.department_id,
        leave_type_name: d.leave_type_name,
        start_date: d.start_date, end_date: d.end_date,
        total_days: d.total_days
      };
    });

  var holidays = _readAll('Holidays').map(function(r){ return r.data; })
    .filter(function(h){ return h.date >= from && h.date <= to; });

  return { status: 'success', year: y, month: m, last_day: lastDay, leaves: data, holidays: holidays };
}

// ============================================================
// LIFF — MY STATS
// ============================================================
function liffGetMyStats(lineUserId, year) {
  var emp = _empByLine(lineUserId);
  if (!emp) return { status: 'error', message: 'ยังไม่ได้ผูกบัญชี' };
  var y = parseInt(year, 10) || new Date().getFullYear();
  var types = {};
  _readAll('LeaveTypes').forEach(function(r){ types[r.data.id] = r.data; });

  var stats = {};
  _readAll('LeaveRequests').map(function(r){ return r.data; })
    .filter(function(d){
      return d.employee_id === emp.id && d.status === 'APPROVED' && d.start_date.substring(0,4) === String(y);
    })
    .forEach(function(d){
      if (!stats[d.leave_type_id]) stats[d.leave_type_id] = { name: d.leave_type_name, days: 0, count: 0, color: (types[d.leave_type_id] || {}).color || '#64748b' };
      stats[d.leave_type_id].days += d.total_days;
      stats[d.leave_type_id].count++;
    });

  return { status: 'success', year: y, data: Object.keys(stats).map(function(k){ return stats[k]; }) };
}

// ============================================================
// USER MANAGEMENT (admin only)
// ============================================================
function getUsers(token) {
  var user = _auth(token);
  if (user.role !== 'admin') return { status: 'error', message: 'ไม่มีสิทธิ์' };
  try {
    var rows = _readAll('Users');
    var data = rows.map(function(r) {
      var d = Object.assign({}, r.data);
      delete d.password;
      return d;
    });
    return { status: 'success', data: data };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function saveUser(token, payload) {
  var actor = _auth(token);
  if (actor.role !== 'admin') return { status: 'error', message: 'ไม่มีสิทธิ์' };
  try {
    var id = payload.id || '';
    var users = _readAll('Users');

    // ตรวจ username ซ้ำ
    for (var i = 0; i < users.length; i++) {
      var u = users[i].data;
      if (u.username === payload.username && u.id !== id) {
        return { status: 'error', message: 'ชื่อผู้ใช้นี้มีอยู่แล้ว' };
      }
    }

    var rolePerms = (CONFIG.USER_ROLES[payload.role] || {}).permissions || [];

    if (id) {
      // แก้ไข
      var found = null;
      for (var j = 0; j < users.length; j++) {
        if (users[j].data.id === id) { found = users[j]; break; }
      }
      if (!found) return { status: 'error', message: 'ไม่พบผู้ใช้' };
      var d = found.data;
      d.username = payload.username;
      d.name = payload.name;
      d.email = payload.email || '';
      d.role = payload.role;
      d.permissions = rolePerms;
      d.active = payload.active !== false;
      if (payload.password) d.password = payload.password;
      d.updated_at = _now();
      _update('Users', found.rowIndex, d);
      return { status: 'success', message: 'แก้ไขผู้ใช้เรียบร้อยแล้ว' };
    } else {
      // เพิ่มใหม่
      if (!payload.password) return { status: 'error', message: 'กรุณาระบุรหัสผ่าน' };
      _append('Users', {
        id: _uuid(),
        username: payload.username,
        password: payload.password,
        role: payload.role,
        name: payload.name,
        email: payload.email || '',
        line_user_id: '',
        employee_id: '',
        permissions: rolePerms,
        active: payload.active !== false,
        last_login: '',
        created_at: _now(),
        updated_at: _now()
      });
      return { status: 'success', message: 'เพิ่มผู้ใช้เรียบร้อยแล้ว' };
    }
  } catch (err) {
    logError(err.toString(), 'saveUser');
    return { status: 'error', message: err.toString() };
  }
}

function deleteUser(token, id) {
  var actor = _auth(token);
  if (actor.role !== 'admin') return { status: 'error', message: 'ไม่มีสิทธิ์' };
  try {
    var users = _readAll('Users');
    var found = null;
    for (var i = 0; i < users.length; i++) {
      if (users[i].data.id === id) { found = users[i]; break; }
    }
    if (!found) return { status: 'error', message: 'ไม่พบผู้ใช้' };
    if (found.data.username === actor.username) return { status: 'error', message: 'ไม่สามารถลบบัญชีตัวเองได้' };
    _delete('Users', found.rowIndex);
    return { status: 'success', message: 'ลบผู้ใช้เรียบร้อยแล้ว' };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

// ============================================================
// PHASE 1 — SETTINGS / CONFIG
// ============================================================
function _getConfigRow() {
  var c = _readAll('Config');
  return c.length ? c[0] : null;
}

function getConfig(token) {
  _auth(token);
  var c = _getConfigRow();
  if (!c) return { status: 'error', message: 'ยังไม่ได้ติดตั้งระบบ' };
  return { status: 'success', config: c.data };
}

function saveConfig(token, payload) {
  _auth(token);
  try {
    var c = _getConfigRow();
    if (!c) return { status: 'error', message: 'ไม่พบ Config' };
    var d = c.data;
    var allow = ['app_name', 'line_channel_access_token', 'line_channel_secret', 'line_liff_id',
      'folder_id', 'approval_levels', 'level1_role', 'level2_role', 'work_days', 'work_hours',
      'half_day_enabled', 'fiscal_start_month', 'buddhist_era', 'notification_enabled',
      'email_notifications', 'telegram_bot_token', 'telegram_chat_id', 'maintenance_mode'];
    allow.forEach(function (k) { if (payload[k] !== undefined) d[k] = payload[k]; });
    if (d.approval_levels !== undefined) d.approval_levels = Number(d.approval_levels) || 1;
    d.updated_at = _now();
    _update('Config', c.rowIndex, d);
    return { status: 'success', message: 'บันทึกการตั้งค่าเรียบร้อยแล้ว', config: d };
  } catch (err) {
    logError(err.toString(), 'saveConfig');
    return { status: 'error', message: err.toString() };
  }
}

/** คืน folder_id ที่ใช้งานได้จริง — ถ้ายังไม่มีจะสร้างโฟลเดอร์ใหม่ให้ */
function _ensureFolder() {
  var c = _getConfigRow();
  var d = c.data;
  if (d.folder_id) {
    try { DriveApp.getFolderById(d.folder_id); return d.folder_id; } catch (e) {}
  }
  var folder = DriveApp.createFolder((d.app_name || 'ระบบลาออนไลน์') + ' - Files');
  d.folder_id = folder.getId();
  d.updated_at = _now();
  _update('Config', c.rowIndex, d);
  return d.folder_id;
}

function createDriveFolder(token) {
  _auth(token);
  try {
    var id = _ensureFolder();
    return { status: 'success', message: 'พร้อมใช้งานโฟลเดอร์แล้ว', folder_id: id };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

/** อัพโหลดโลโก้ (base64) เข้า Drive FOLDER_ID */
function uploadLogo(token, base64Data, filename, mimeType) {
  _auth(token);
  try {
    var folderId = _ensureFolder();
    var folder = DriveApp.getFolderById(folderId);
    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType || 'image/png', filename || 'logo.png');
    var file = folder.createFile(blob);
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}

    var c = _getConfigRow();
    var d = c.data;
    if (d.logo_file_id) { try { DriveApp.getFileById(d.logo_file_id).setTrashed(true); } catch (e) {} }
    d.logo_file_id = file.getId();
    d.updated_at = _now();
    _update('Config', c.rowIndex, d);

    return {
      status: 'success',
      message: 'อัพโหลดโลโก้เรียบร้อยแล้ว',
      logo_file_id: file.getId(),
      logo_url: 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=s200'
    };
  } catch (err) {
    logError(err.toString(), 'uploadLogo');
    return { status: 'error', message: err.toString() };
  }
}

// ============================================================
// PHASE 1 — MASTER : LEAVE TYPES
// ============================================================
function getLeaveTypes(token) {
  _auth(token);
  var rows = _readAll('LeaveTypes').map(function (r) { return r.data; });
  rows.sort(function (a, b) { return (a.order || 0) - (b.order || 0); });
  return { status: 'success', data: rows };
}

function saveLeaveType(token, obj) {
  _auth(token);
  try {
    if (!obj.name) return { status: 'error', message: 'กรุณากรอกชื่อประเภทลา' };
    var rows = _readAll('LeaveTypes');
    if (obj.id) {
      for (var i = 0; i < rows.length; i++) {
        if (rows[i].data.id === obj.id) {
          var d = rows[i].data;
          ['name', 'code', 'color', 'default_quota_days', 'max_per_request', 'requires_attachment',
            'allow_half_day', 'paid', 'gender_restrict', 'carry_over', 'order', 'active', 'seniority_tiers'].forEach(function (k) {
            if (obj[k] !== undefined) d[k] = obj[k];
          });
          d.default_quota_days = Number(d.default_quota_days) || 0;
          d.max_per_request = Number(d.max_per_request) || 0;
          d.order = Number(d.order) || 0;
          _update('LeaveTypes', rows[i].rowIndex, d);
          return { status: 'success', message: 'บันทึกประเภทลาแล้ว' };
        }
      }
      return { status: 'error', message: 'ไม่พบรายการ' };
    }
    _append('LeaveTypes', {
      id: _uuid(), name: obj.name, code: obj.code || '', color: obj.color || '#64748b',
      default_quota_days: Number(obj.default_quota_days) || 0, max_per_request: Number(obj.max_per_request) || 0,
      requires_attachment: !!obj.requires_attachment, allow_half_day: obj.allow_half_day !== false,
      paid: obj.paid !== false, gender_restrict: obj.gender_restrict || '', carry_over: !!obj.carry_over,
      order: Number(obj.order) || (rows.length + 1), active: obj.active !== false,
      seniority_tiers: obj.seniority_tiers || [], created_at: _now()
    });
    return { status: 'success', message: 'เพิ่มประเภทลาแล้ว' };
  } catch (err) {
    logError(err.toString(), 'saveLeaveType');
    return { status: 'error', message: err.toString() };
  }
}

function deleteLeaveType(token, id) {
  _auth(token);
  var rows = _readAll('LeaveTypes');
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].data.id === id) { _delete('LeaveTypes', rows[i].rowIndex); return { status: 'success', message: 'ลบแล้ว' }; }
  }
  return { status: 'error', message: 'ไม่พบรายการ' };
}

// ============================================================
// PHASE 1 — MASTER : DEPARTMENTS
// ============================================================
function getDepartments(token) {
  _auth(token);
  var userMap = {};
  _readAll('Users').forEach(function(r){ userMap[r.data.id] = r.data.name; });
  var data = _readAll('Departments').map(function (r) {
    var d = r.data;
    // รองรับทั้ง head_user_ids (array ใหม่) และ head_user_id (เดิม)
    var ids = Array.isArray(d.head_user_ids) ? d.head_user_ids : (d.head_user_id ? [d.head_user_id] : []);
    d.head_user_ids = ids;
    d.head_user_names = ids.map(function(id){ return userMap[id] || ''; }).filter(Boolean).join(', ');
    return d;
  });
  return { status: 'success', data: data };
}

function saveDepartment(token, obj) {
  _auth(token);
  try {
    if (!obj.name) return { status: 'error', message: 'กรุณากรอกชื่อแผนก' };
    var rows = _readAll('Departments');
    if (obj.id) {
      for (var i = 0; i < rows.length; i++) {
        if (rows[i].data.id === obj.id) {
          var d = rows[i].data;
          ['name', 'code', 'head_user_ids', 'parent_id', 'active'].forEach(function (k) {
            if (obj[k] !== undefined) d[k] = obj[k];
          });
          d.updated_at = _now();
          _update('Departments', rows[i].rowIndex, d);
          return { status: 'success', message: 'บันทึกแผนกแล้ว' };
        }
      }
      return { status: 'error', message: 'ไม่พบรายการ' };
    }
    _append('Departments', {
      id: _uuid(), name: obj.name, code: obj.code || '',
      head_user_ids: obj.head_user_ids || [], parent_id: obj.parent_id || '',
      active: obj.active !== false, created_at: _now(), updated_at: _now()
    });
    return { status: 'success', message: 'เพิ่มแผนกแล้ว' };
  } catch (err) {
    logError(err.toString(), 'saveDepartment');
    return { status: 'error', message: err.toString() };
  }
}

function deleteDepartment(token, id) {
  _auth(token);
  var rows = _readAll('Departments');
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].data.id === id) { _delete('Departments', rows[i].rowIndex); return { status: 'success', message: 'ลบแล้ว' }; }
  }
  return { status: 'error', message: 'ไม่พบรายการ' };
}

// ============================================================
// PHASE 1 — MASTER : HOLIDAYS
// ============================================================
function getHolidays(token, year) {
  _auth(token);
  var rows = _readAll('Holidays').map(function (r) { return r.data; });
  if (year) rows = rows.filter(function (h) { return String(h.year) === String(year); });
  rows.sort(function (a, b) { return a.date < b.date ? -1 : 1; });
  return { status: 'success', data: rows };
}

function saveHoliday(token, obj) {
  _auth(token);
  try {
    if (!obj.name || !obj.date) return { status: 'error', message: 'กรุณากรอกชื่อและวันที่' };
    var yr = parseInt(String(obj.date).substring(0, 4), 10);
    var rows = _readAll('Holidays');
    if (obj.id) {
      for (var i = 0; i < rows.length; i++) {
        if (rows[i].data.id === obj.id) {
          var d = rows[i].data;
          d.name = obj.name; d.date = obj.date; d.type = obj.type || d.type || 'public';
          d.year = yr; d.active = obj.active !== false; d.updated_at = _now();
          _update('Holidays', rows[i].rowIndex, d);
          return { status: 'success', message: 'บันทึกวันหยุดแล้ว' };
        }
      }
      return { status: 'error', message: 'ไม่พบรายการ' };
    }
    _append('Holidays', {
      id: _uuid(), name: obj.name, date: obj.date, type: obj.type || 'public',
      year: yr, active: obj.active !== false, created_at: _now()
    });
    return { status: 'success', message: 'เพิ่มวันหยุดแล้ว' };
  } catch (err) {
    logError(err.toString(), 'saveHoliday');
    return { status: 'error', message: err.toString() };
  }
}

function deleteHoliday(token, id) {
  _auth(token);
  var rows = _readAll('Holidays');
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].data.id === id) { _delete('Holidays', rows[i].rowIndex); return { status: 'success', message: 'ลบแล้ว' }; }
  }
  return { status: 'error', message: 'ไม่พบรายการ' };
}

/** เพิ่มวันหยุดราชการไทย (เฉพาะวันที่ตายตัวตามปฏิทินสากล) ของปี ค.ศ. ที่กำหนด */
function seedThaiHolidays(token, gregorianYear) {
  _auth(token);
  try {
    var y = parseInt(gregorianYear, 10) || new Date().getFullYear();
    var fixed = [
      ['01-01', 'วันขึ้นปีใหม่'],
      ['04-06', 'วันจักรี'],
      ['04-13', 'วันสงกรานต์'],
      ['04-14', 'วันสงกรานต์'],
      ['04-15', 'วันสงกรานต์'],
      ['05-01', 'วันแรงงานแห่งชาติ'],
      ['05-04', 'วันฉัตรมงคล'],
      ['06-03', 'วันเฉลิมพระชนมพรรษา สมเด็จพระนางเจ้าฯ พระบรมราชินี'],
      ['07-28', 'วันเฉลิมพระชนมพรรษา พระบาทสมเด็จพระเจ้าอยู่หัว'],
      ['08-12', 'วันแม่แห่งชาติ'],
      ['10-13', 'วันคล้ายวันสวรรคต ร.9'],
      ['10-23', 'วันปิยมหาราช'],
      ['12-05', 'วันชาติ / วันพ่อแห่งชาติ'],
      ['12-10', 'วันรัฐธรรมนูญ'],
      ['12-31', 'วันสิ้นปี']
    ];
    var existing = _readAll('Holidays').map(function (r) { return r.data.date; });
    var added = 0;
    fixed.forEach(function (h) {
      var date = y + '-' + h[0];
      if (existing.indexOf(date) === -1) {
        _append('Holidays', { id: _uuid(), name: h[1], date: date, type: 'public', year: y, active: true, created_at: _now() });
        added++;
      }
    });
    return { status: 'success', message: 'เพิ่มวันหยุด ' + added + ' รายการ (หมายเหตุ: วันพระ/วันหยุดตามจันทรคติต้องเพิ่มเอง)', added: added };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

// PHASE 3 — LEAVE REQUESTS
// ============================================================
function _parseDate(str) { return new Date(str + 'T00:00:00'); }
function _isoDate(d) {
  var m = ('0' + (d.getMonth() + 1)).slice(-2);
  var dd = ('0' + d.getDate()).slice(-2);
  return d.getFullYear() + '-' + m + '-' + dd;
}

/** คำนวณจำนวนวันลา (ตัดวันหยุดสุดสัปดาห์ + วันหยุด) รองรับครึ่งวัน */
function _calcLeaveDays(startDate, endDate, startHalf, endHalf) {
  if (!startDate || !endDate) return 0;
  var cfg = _getConfigRow().data;
  var workDays = cfg.work_days || ['MO', 'TU', 'WE', 'TH', 'FR'];
  var dayMap = ['SU', 'MO', 'TU', 'WE', 'TH', 'FR', 'SA'];
  var holidays = {};
  _readAll('Holidays').forEach(function (h) { if (h.data.active) holidays[h.data.date] = true; });

  var s = _parseDate(startDate), e = _parseDate(endDate);
  if (e < s) return 0;

  var count = 0;
  var cur = new Date(s.getTime());
  while (cur <= e) {
    var dow = dayMap[cur.getDay()];
    if (workDays.indexOf(dow) !== -1 && !holidays[_isoDate(cur)]) count++;
    cur.setDate(cur.getDate() + 1);
  }

  if (startDate === endDate) {
    if (startHalf && startHalf !== 'full') count = count > 0 ? 0.5 : 0;
  } else {
    if (startHalf === 'afternoon' && count > 0) count -= 0.5;
    if (endHalf === 'morning' && count > 0) count -= 0.5;
  }
  return count;
}

function previewLeaveDays(token, startDate, endDate, startHalf, endHalf) {
  _auth(token);
  return { status: 'success', days: _calcLeaveDays(startDate, endDate, startHalf, endHalf) };
}

function _empById(id) {
  var rows = _readAll('Employees');
  for (var i = 0; i < rows.length; i++) if (rows[i].data.id === id) return rows[i].data;
  return null;
}
function _empName(id) {
  if (!id) return '';
  var e = _empById(id);
  return e ? ((e.prefix || '') + e.first_name + ' ' + e.last_name).trim() : '';
}
function _typeById(id) {
  var rows = _readAll('LeaveTypes');
  for (var i = 0; i < rows.length; i++) if (rows[i].data.id === id) return rows[i].data;
  return null;
}

/** ปรับ pending_days (และ used_days) ใน LeaveBalances */
function _adjustBalance(empId, typeId, year, dPending, dUsed) {
  var rows = _readAll('LeaveBalances');
  for (var i = 0; i < rows.length; i++) {
    var b = rows[i].data;
    if (b.employee_id === empId && b.leave_type_id === typeId && String(b.year) === String(year)) {
      b.pending_days = Math.max(0, Number(b.pending_days) + (dPending || 0));
      b.used_days = Math.max(0, Number(b.used_days) + (dUsed || 0));
      b.updated_at = _now();
      _update('LeaveBalances', rows[i].rowIndex, b);
      return b;
    }
  }
  // ยังไม่มี -> สร้างใหม่จาก default ของประเภท
  var t = _typeById(typeId) || {};
  var nb = {
    id: _uuid(), employee_id: empId, year: year, leave_type_id: typeId,
    entitled_days: Number(t.default_quota_days) || 0,
    used_days: Math.max(0, dUsed || 0), pending_days: Math.max(0, dPending || 0),
    carried_over: 0, updated_at: _now()
  };
  _append('LeaveBalances', nb);
  return nb;
}

function _remaining(empId, typeId, year) {
  var rows = _readAll('LeaveBalances');
  for (var i = 0; i < rows.length; i++) {
    var b = rows[i].data;
    if (b.employee_id === empId && b.leave_type_id === typeId && String(b.year) === String(year)) {
      return Number(b.entitled_days) + Number(b.carried_over) - Number(b.used_days) - Number(b.pending_days);
    }
  }
  var t = _typeById(typeId);
  return t ? Number(t.default_quota_days) : 0;
}

function _genRequestNo(beYear) {
  var rows = _readAll('LeaveRequests');
  var n = 0;
  var prefix = 'LV' + beYear + '-';
  rows.forEach(function (r) { if ((r.data.request_no || '').indexOf(prefix) === 0) n++; });
  return prefix + ('0000' + (n + 1)).slice(-4);
}

function createLeaveRequest(token, obj) {
  _auth(token);
  try {
    if (!obj.employee_id || !obj.leave_type_id || !obj.start_date || !obj.end_date) {
      return { status: 'error', message: 'กรอกข้อมูลให้ครบ (พนักงาน/ประเภท/วันที่)' };
    }
    var emp = _empById(obj.employee_id);
    if (!emp) return { status: 'error', message: 'ไม่พบพนักงาน' };
    var type = _typeById(obj.leave_type_id);
    if (!type) return { status: 'error', message: 'ไม่พบประเภทการลา' };

    var startHalf = obj.start_half || 'full';
    var endHalf = obj.end_half || 'full';
    var total = _calcLeaveDays(obj.start_date, obj.end_date, startHalf, endHalf);
    if (total <= 0) return { status: 'error', message: 'ช่วงวันที่เลือกไม่มีวันทำงาน' };

    var year = parseInt(obj.start_date.substring(0, 4), 10);

    // ตรวจโควตา (เฉพาะประเภทที่มีโควตาจำกัด)
    if (Number(type.default_quota_days) > 0) {
      var remain = _remaining(emp.id, type.id, year);
      if (total > remain) {
        return { status: 'error', message: 'วันลาคงเหลือไม่พอ (เหลือ ' + remain + ' วัน ขอลา ' + total + ' วัน)' };
      }
    }

    var cfg = _getConfigRow().data;
    var beYear = year + 543;
    var req = {
      id: _uuid(),
      request_no: _genRequestNo(beYear),
      employee_id: emp.id,
      employee_name: (emp.prefix || '') + emp.first_name + ' ' + emp.last_name,
      department_id: emp.department_id || '',
      leave_type_id: type.id,
      leave_type_name: type.name,
      start_date: obj.start_date,
      end_date: obj.end_date,
      start_half: startHalf,
      end_half: endHalf,
      total_days: total,
      reason: obj.reason || '',
      contact_during_leave: obj.contact_during_leave || '',
      attachment_file_ids: obj.attachment_file_ids || [],
      status: 'SUBMITTED',
      current_level: 1,
      approvals: [],
      submitted_at: _now(),
      decided_at: '',
      created_at: _now(),
      updated_at: _now()
    };
    _append('LeaveRequests', req);

    // หักเข้า pending
    _adjustBalance(emp.id, type.id, year, total, 0);

    _notifyApprovers(req, 1);
    _notifyEmployeeSubmit(req);
    return { status: 'success', message: 'ยื่นใบลาเรียบร้อย (' + req.request_no + ')', id: req.id, request_no: req.request_no };
  } catch (err) {
    logError(err.toString(), 'createLeaveRequest');
    return { status: 'error', message: err.toString() };
  }
}

function getLeaveRequests(token, filters) {
  var user = _auth(token);
  filters = filters || {};
  var depts = {};
  _readAll('Departments').forEach(function (r) { depts[r.data.id] = r.data.name; });

  var deptFilter = _supervisorDeptIds(user);
  var data = _readAll('LeaveRequests').map(function (r) {
    var d = r.data;
    d.department_name = depts[d.department_id] || '';
    return d;
  }).filter(function(d){
    if (deptFilter && deptFilter.indexOf(d.department_id) === -1) return false;
    return true;
  });

  if (filters.status) data = data.filter(function (d) { return d.status === filters.status; });
  if (filters.employee_id) data = data.filter(function (d) { return d.employee_id === filters.employee_id; });
  if (filters.department_id) data = data.filter(function (d) { return d.department_id === filters.department_id; });
  if (filters.dateFrom) data = data.filter(function (d) { return d.start_date >= filters.dateFrom; });
  if (filters.dateTo) data = data.filter(function (d) { return d.start_date <= filters.dateTo; });
  if (filters.q) {
    var q = String(filters.q).toLowerCase();
    data = data.filter(function (d) {
      return (d.employee_name + ' ' + d.request_no + ' ' + d.leave_type_name).toLowerCase().indexOf(q) !== -1;
    });
  }

  data.sort(function (a, b) { return a.created_at < b.created_at ? 1 : -1; });
  return { status: 'success', data: data };
}

function getLeaveRequest(token, id) {
  _auth(token);
  var rows = _readAll('LeaveRequests');
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].data.id === id) {
      var d = rows[i].data;
      d.department_name = _deptName(d.department_id);
      var emp = _empById(d.employee_id);
      d.employee_phone = emp ? emp.phone : '';
      d.employee_position = emp ? emp.position : '';
      d.attachments = (d.attachment_file_ids || []).map(function (fid) {
        return { id: fid, url: 'https://drive.google.com/thumbnail?id=' + fid };
      });
      return { status: 'success', data: d };
    }
  }
  return { status: 'error', message: 'ไม่พบใบลา' };
}

function cancelLeaveRequest(token, id, byName) {
  _auth(token);
  try {
    var rows = _readAll('LeaveRequests');
    for (var i = 0; i < rows.length; i++) {
      if (rows[i].data.id === id) {
        var d = rows[i].data;
        if (d.status === 'APPROVED') return { status: 'error', message: 'ใบลาที่อนุมัติแล้วยกเลิกไม่ได้ (ติดต่อ HR)' };
        if (d.status === 'CANCELLED' || d.status === 'REJECTED') return { status: 'error', message: 'ใบลานี้สิ้นสุดแล้ว' };

        var year = parseInt(d.start_date.substring(0, 4), 10);
        // คืน pending
        _adjustBalance(d.employee_id, d.leave_type_id, year, -d.total_days, 0);

        d.status = 'CANCELLED';
        d.decided_at = _now();
        d.updated_at = _now();
        d.approvals.push({ level: d.current_level, approver_name: byName || 'ผู้ดูแล', action: 'cancel', comment: 'ยกเลิกใบลา', timestamp: _now() });
        _update('LeaveRequests', rows[i].rowIndex, d);
        return { status: 'success', message: 'ยกเลิกใบลาแล้ว' };
      }
    }
    return { status: 'error', message: 'ไม่พบใบลา' };
  } catch (err) {
    logError(err.toString(), 'cancelLeaveRequest');
    return { status: 'error', message: err.toString() };
  }
}

/** อัพโหลดไฟล์แนบใบลา -> file_id */
function uploadAttachment(token, base64Data, filename, mimeType) {
  return uploadEmployeePhoto(token, base64Data, filename, mimeType);
}

// ============================================================
// PHASE 2 — EMPLOYEES  (canonical, single copy)
// ============================================================
function _deptName(id) {
  if (!id) return '';
  var rows = _readAll('Departments');
  for (var i = 0; i < rows.length; i++) if (rows[i].data.id === id) return rows[i].data.name;
  return '';
}

function getEmployees(token) {
  _auth(token);
  var depts = {};
  _readAll('Departments').forEach(function (r) { depts[r.data.id] = r.data.name; });
  var data = _readAll('Employees').map(function (r) {
    var e = r.data;
    e.department_name = depts[e.department_id] || '';
    e.full_name = (e.prefix || '') + e.first_name + ' ' + e.last_name;
    e.photo_url = e.photo_file_id ? ('https://drive.google.com/thumbnail?id=' + e.photo_file_id) : '';
    return e;
  });
  data.sort(function (a, b) { return (a.emp_code || '') < (b.emp_code || '') ? -1 : 1; });
  return { status: 'success', data: data };
}

function getEmployeesSimple(token) {
  _auth(token);
  var data = _readAll('Employees').map(function (r) {
    return { id: r.data.id, name: (r.data.prefix || '') + r.data.first_name + ' ' + r.data.last_name, emp_code: r.data.emp_code };
  });
  return { status: 'success', data: data };
}

function saveEmployee(token, obj) {
  _auth(token);
  try {
    if (!obj.first_name || !obj.last_name) return { status: 'error', message: 'กรุณากรอกชื่อ-นามสกุล' };
    var rows = _readAll('Employees');
    if (obj.emp_code) {
      for (var i = 0; i < rows.length; i++) {
        if (rows[i].data.emp_code === obj.emp_code && rows[i].data.id !== obj.id) {
          return { status: 'error', message: 'รหัสพนักงาน ' + obj.emp_code + ' มีอยู่แล้ว' };
        }
      }
    }
    var fields = ['emp_code', 'prefix', 'first_name', 'last_name', 'nickname', 'position',
      'department_id', 'supervisor_id', 'employment_type', 'start_date', 'phone', 'email',
      'photo_file_id', 'status'];
    if (obj.id) {
      for (var j = 0; j < rows.length; j++) {
        if (rows[j].data.id === obj.id) {
          var d = rows[j].data;
          fields.forEach(function (k) { if (obj[k] !== undefined) d[k] = obj[k]; });
          d.updated_at = _now();
          _update('Employees', rows[j].rowIndex, d);
          return { status: 'success', message: 'บันทึกข้อมูลพนักงานแล้ว', id: d.id };
        }
      }
      return { status: 'error', message: 'ไม่พบพนักงาน' };
    }
    var emp = {
      id: _uuid(), emp_code: obj.emp_code || '', prefix: obj.prefix || '',
      first_name: obj.first_name, last_name: obj.last_name, nickname: obj.nickname || '',
      position: obj.position || '', department_id: obj.department_id || '',
      supervisor_id: obj.supervisor_id || '', employment_type: obj.employment_type || 'ประจำ',
      start_date: obj.start_date || '', phone: obj.phone || '', email: obj.email || '',
      line_user_id: '', line_linked: false, photo_file_id: obj.photo_file_id || '',
      status: obj.status || 'active', created_at: _now(), updated_at: _now()
    };
    _append('Employees', emp);
    _generateBalances(emp.id, new Date().getFullYear());
    return { status: 'success', message: 'เพิ่มพนักงานและสร้างโควตาการลาเรียบร้อยแล้ว', id: emp.id };
  } catch (err) {
    logError(err.toString(), 'saveEmployee');
    return { status: 'error', message: err.toString() };
  }
}

function deleteEmployee(token, id) {
  _auth(token);
  var rows = _readAll('Employees');
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].data.id === id) {
      var ph = rows[i].data.photo_file_id;
      if (ph) { try { DriveApp.getFileById(ph).setTrashed(true); } catch (e) {} }
      _delete('Employees', rows[i].rowIndex);
      return { status: 'success', message: 'ลบพนักงานแล้ว' };
    }
  }
  return { status: 'error', message: 'ไม่พบพนักงาน' };
}

function uploadEmployeePhoto(token, base64Data, filename, mimeType) {
  _auth(token);
  try {
    var folder = DriveApp.getFolderById(_ensureFolder());
    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType || 'image/jpeg', filename || 'emp.jpg');
    var file = folder.createFile(blob);
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    return { status: 'success', file_id: file.getId(), url: 'https://drive.google.com/thumbnail?id=' + file.getId() };
  } catch (err) {
    logError(err.toString(), 'uploadEmployeePhoto');
    return { status: 'error', message: err.toString() };
  }
}

function unlinkLine(token, employeeId) {
  _auth(token);
  var rows = _readAll('Employees');
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].data.id === employeeId) {
      var d = rows[i].data;
      d.line_user_id = ''; d.line_linked = false; d.updated_at = _now();
      _update('Employees', rows[i].rowIndex, d);
      return { status: 'success', message: 'ปลดการเชื่อมต่อ LINE แล้ว' };
    }
  }
  return { status: 'error', message: 'ไม่พบพนักงาน' };
}

function getEmployeeByLineId(lineUserId) {
  try {
    var rows = _readAll('Employees');
    for (var i = 0; i < rows.length; i++) {
      if (rows[i].data.line_user_id === lineUserId && rows[i].data.line_linked) {
        var e = rows[i].data;
        return {
          status: 'success', linked: true,
          employee: { id: e.id, emp_code: e.emp_code, name: (e.prefix || '') + e.first_name + ' ' + e.last_name, department_id: e.department_id, position: e.position }
        };
      }
    }
    return { status: 'success', linked: false };
  } catch (err) { return { status: 'error', message: err.toString() }; }
}

function linkLineAccount(empCode, phone, lineUserId) {
  try {
    if (!empCode || !phone || !lineUserId) return { status: 'error', message: 'ข้อมูลไม่ครบ' };
    var rows = _readAll('Employees');
    var code = String(empCode).trim();
    var ph = String(phone).trim().replace(/[-\s]/g, '');
    for (var i = 0; i < rows.length; i++) {
      var e = rows[i].data;
      var ephone = String(e.phone || '').replace(/[-\s]/g, '');
      if (e.emp_code === code && ephone === ph) {
        if (e.line_linked && e.line_user_id && e.line_user_id !== lineUserId) {
          return { status: 'error', message: 'พนักงานนี้ผูกบัญชี LINE อื่นไว้แล้ว' };
        }
        e.line_user_id = lineUserId; e.line_linked = true; e.updated_at = _now();
        _update('Employees', rows[i].rowIndex, e);
        return { status: 'success', message: 'เชื่อมต่อบัญชีสำเร็จ', employee: { id: e.id, name: (e.prefix || '') + e.first_name + ' ' + e.last_name } };
      }
    }
    return { status: 'error', message: 'ไม่พบข้อมูลที่ตรงกัน ตรวจสอบรหัสพนักงานและเบอร์โทร' };
  } catch (err) { logError(err.toString(), 'linkLineAccount'); return { status: 'error', message: err.toString() }; }
}

// ---------- LEAVE BALANCES ----------
function _generateBalances(employeeId, year) {
  var types = _readAll('LeaveTypes').map(function (r) { return r.data; }).filter(function (t) { return t.active; });
  var existing = _readAll('LeaveBalances').map(function (r) { return r.data; });
  var created = 0;
  types.forEach(function (t) {
    var has = existing.some(function (b) { return b.employee_id === employeeId && b.leave_type_id === t.id && String(b.year) === String(year); });
    if (!has) {
      _append('LeaveBalances', {
        id: _uuid(), employee_id: employeeId, year: year, leave_type_id: t.id,
        entitled_days: Number(t.default_quota_days) || 0, used_days: 0, pending_days: 0, carried_over: 0, updated_at: _now()
      });
      created++;
    }
  });
  return created;
}

function getLeaveBalances(token, employeeId, year) {
  _auth(token);
  var y = year || new Date().getFullYear();
  var types = {};
  _readAll('LeaveTypes').forEach(function (r) { types[r.data.id] = r.data; });
  var data = _readAll('LeaveBalances').map(function (r) { return r.data; })
    .filter(function (b) { return b.employee_id === employeeId && String(b.year) === String(y); })
    .map(function (b) {
      var t = types[b.leave_type_id] || {};
      var remaining = (Number(b.entitled_days) + Number(b.carried_over) - Number(b.used_days) - Number(b.pending_days));
      return {
        leave_type_id: b.leave_type_id, leave_type_name: t.name || '', color: t.color || '#64748b',
        entitled: Number(b.entitled_days), carried_over: Number(b.carried_over),
        used: Number(b.used_days), pending: Number(b.pending_days), remaining: remaining
      };
    });
  data.sort(function (a, b) { return (types[a.leave_type_id] ? types[a.leave_type_id].order : 0) - (types[b.leave_type_id] ? types[b.leave_type_id].order : 0); });
  return { status: 'success', data: data, year: y };
}

function generateBalancesForEmployee(token, employeeId, year) {
  _auth(token);
  var y = year || new Date().getFullYear();
  var n = _generateBalances(employeeId, y);
  return { status: 'success', message: 'สร้างโควตาเพิ่ม ' + n + ' ประเภท (ปี ' + (y + 543) + ')' };
}

function updateBalance(token, employeeId, leaveTypeId, year, entitled, carriedOver) {
  _auth(token);
  var rows = _readAll('LeaveBalances');
  for (var i = 0; i < rows.length; i++) {
    var b = rows[i].data;
    if (b.employee_id === employeeId && b.leave_type_id === leaveTypeId && String(b.year) === String(year)) {
      if (entitled !== undefined) b.entitled_days = Number(entitled) || 0;
      if (carriedOver !== undefined) b.carried_over = Number(carriedOver) || 0;
      b.updated_at = _now();
      _update('LeaveBalances', rows[i].rowIndex, b);
      return { status: 'success', message: 'ปรับโควตาแล้ว' };
    }
  }
  return { status: 'error', message: 'ไม่พบรายการโควตา' };
}

// ============================================================
// PHASE 4 — APPROVAL WORKFLOW (2 ระดับ)
// ============================================================
/** role ที่อนุมัติได้ในระดับนั้น (admin อนุมัติได้ทุกระดับ) */
function _canApprove(userRole, level, cfg) {
  if (userRole === 'admin') return true;
  if (level === 1) return userRole === (cfg.level1_role || 'supervisor');
  if (level === 2) return userRole === (cfg.level2_role || 'hr');
  return false;
}

/** คิวที่ผู้ใช้คนนี้ต้องอนุมัติ */
function getApprovalQueue(token) {
  var user = _auth(token);
  var cfg = _getConfigRow().data;
  var depts = {};
  _readAll('Departments').forEach(function (r) { depts[r.data.id] = r.data.name; });
  var deptFilter = _supervisorDeptIds(user);

  var data = _readAll('LeaveRequests').map(function (r) { return r.data; })
    .filter(function (d) {
      if (deptFilter && deptFilter.indexOf(d.department_id) === -1) return false;
      if (d.status === 'SUBMITTED') return _canApprove(user.role, 1, cfg);
      if (d.status === 'L1_APPROVED') return _canApprove(user.role, 2, cfg);
      return false;
    })
    .map(function (d) {
      d.department_name = depts[d.department_id] || '';
      d.waiting_level = d.status === 'SUBMITTED' ? 1 : 2;
      return d;
    });

  data.sort(function (a, b) { return a.submitted_at < b.submitted_at ? -1 : 1; });
  return { status: 'success', data: data };
}

function getApprovalCount(token) {
  var r = getApprovalQueue(token);
  return { status: 'success', count: r.data ? r.data.length : 0 };
}

function approveLeave(token, id, comment) {
  var user = _auth(token);
  return _processDecision(user.name, user.role, id, 'approve', comment);
}

function rejectLeave(token, id, comment) {
  var user = _auth(token);
  return _processDecision(user.name, user.role, id, 'reject', comment);
}

/** core การตัดสินใจ ใช้ร่วมทั้งเว็บและ LINE postback (มีการแจ้งเตือน) */
function _processDecision(actorName, actorRole, id, decision, comment) {
  try {
    var cfg = _getConfigRow().data;
    var levels = Number(cfg.approval_levels) || 1;
    var rows = _readAll('LeaveRequests');
    for (var i = 0; i < rows.length; i++) {
      if (rows[i].data.id === id) {
        var d = rows[i].data;
        if (d.status !== 'SUBMITTED' && d.status !== 'L1_APPROVED') {
          return { status: 'error', message: 'ใบลานี้ไม่อยู่ในสถานะรออนุมัติ' };
        }
        var level = d.current_level;
        if (!_canApprove(actorRole, level, cfg)) {
          return { status: 'error', message: 'คุณไม่มีสิทธิ์ดำเนินการในระดับนี้' };
        }
        var year = parseInt(d.start_date.substring(0, 4), 10);

        if (decision === 'reject') {
          if (!comment) return { status: 'error', message: 'กรุณาระบุเหตุผลการไม่อนุมัติ' };
          d.approvals.push({ level: level, approver_name: actorName, action: 'reject', comment: comment, timestamp: _now() });
          d.status = 'REJECTED'; d.decided_at = _now(); d.updated_at = _now();
          _update('LeaveRequests', rows[i].rowIndex, d);
          _adjustBalance(d.employee_id, d.leave_type_id, year, -d.total_days, 0);
          _notifyEmployeeDecision(d, 'reject', actorName, comment);
          return { status: 'success', message: 'บันทึกการไม่อนุมัติแล้ว' };
        }

        // approve
        d.approvals.push({ level: level, approver_name: actorName, action: 'approve', comment: comment || '', timestamp: _now() });
        if (level < levels) {
          d.status = 'L1_APPROVED'; d.current_level = level + 1; d.updated_at = _now();
          _update('LeaveRequests', rows[i].rowIndex, d);
          _notifyApprovers(d, level + 1);   // แจ้งผู้อนุมัติระดับถัดไป
          return { status: 'success', message: 'อนุมัติระดับ ' + level + ' แล้ว ส่งต่อระดับ ' + (level + 1), next: true };
        } else {
          d.status = 'APPROVED'; d.decided_at = _now(); d.updated_at = _now();
          _update('LeaveRequests', rows[i].rowIndex, d);
          _adjustBalance(d.employee_id, d.leave_type_id, year, -d.total_days, d.total_days);
          _notifyEmployeeDecision(d, 'approve', actorName, comment);
          return { status: 'success', message: 'อนุมัติใบลาเรียบร้อยแล้ว', approved: true };
        }
      }
    }
    return { status: 'error', message: 'ไม่พบใบลา' };
  } catch (err) {
    logError(err.toString(), '_processDecision');
    return { status: 'error', message: err.toString() };
  }
}

// ============================================================
// PHASE 5 — LIFF (ฝั่งพนักงาน) — เรียกด้วย lineUserId ไม่ใช้ session
// ============================================================
function liffGetConfig() {
  var c = _getConfigRow();
  var d = c ? c.data : {};
  var execUrl = '';
  try { execUrl = ScriptApp.getService().getUrl(); } catch (e) {}
  return {
    status: 'success',
    liff_id: d.line_liff_id || '',
    app_name: d.app_name || CONFIG.APP_NAME,
    logo_url: d.logo_file_id ? ('https://drive.google.com/thumbnail?id=' + d.logo_file_id + '&sz=s200') : '',
    buddhist_era: d.buddhist_era !== false,
    exec_url: execUrl ? (execUrl + '?page=liff') : ''
  };
}

function liffGetLeaveTypes() {
  var rows = _readAll('LeaveTypes').map(function (r) { return r.data; }).filter(function (t) { return t.active; });
  rows.sort(function (a, b) { return (a.order || 0) - (b.order || 0); });
  return {
    status: 'success',
    data: rows.map(function (t) {
      return { id: t.id, name: t.name, color: t.color, requires_attachment: t.requires_attachment, allow_half_day: t.allow_half_day, default_quota_days: t.default_quota_days };
    })
  };
}

function liffPreviewDays(startDate, endDate, startHalf, endHalf) {
  return { status: 'success', days: _calcLeaveDays(startDate, endDate, startHalf, endHalf) };
}

function _empByLine(lineUserId) {
  var rows = _readAll('Employees');
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].data.line_user_id === lineUserId && rows[i].data.line_linked) return rows[i].data;
  }
  return null;
}

function liffGetBalances(lineUserId) {
  var emp = _empByLine(lineUserId);
  if (!emp) return { status: 'error', message: 'ยังไม่ได้ผูกบัญชี', linked: false };
  var y = new Date().getFullYear();
  var types = {};
  _readAll('LeaveTypes').forEach(function (r) { types[r.data.id] = r.data; });
  var data = _readAll('LeaveBalances').map(function (r) { return r.data; })
    .filter(function (b) { return b.employee_id === emp.id && String(b.year) === String(y); })
    .map(function (b) {
      var t = types[b.leave_type_id] || {};
      return {
        name: t.name || '', color: t.color || '#64748b',
        entitled: Number(b.entitled_days) + Number(b.carried_over),
        used: Number(b.used_days), pending: Number(b.pending_days),
        remaining: Number(b.entitled_days) + Number(b.carried_over) - Number(b.used_days) - Number(b.pending_days),
        order: t.order || 0
      };
    });
  data.sort(function (a, b) { return a.order - b.order; });
  return { status: 'success', data: data, year: y, employee_name: (emp.prefix || '') + emp.first_name + ' ' + emp.last_name };
}

function liffGetHistory(lineUserId) {
  var emp = _empByLine(lineUserId);
  if (!emp) return { status: 'error', message: 'ยังไม่ได้ผูกบัญชี', linked: false };
  var data = _readAll('LeaveRequests').map(function (r) { return r.data; })
    .filter(function (d) { return d.employee_id === emp.id; })
    .map(function (d) {
      return {
        id: d.id, request_no: d.request_no, leave_type_name: d.leave_type_name,
        start_date: d.start_date, end_date: d.end_date, start_half: d.start_half,
        total_days: d.total_days, status: d.status, reason: d.reason,
        created_at: d.created_at,
        last_comment: (d.approvals && d.approvals.length) ? d.approvals[d.approvals.length - 1].comment : ''
      };
    });
  data.sort(function (a, b) { return a.created_at < b.created_at ? 1 : -1; });
  return { status: 'success', data: data };
}

function liffCancelLeave(lineUserId, leaveId, reason) {
  try {
    var emp = _empByLine(lineUserId);
    if (!emp) return { status: 'error', message: 'ยังไม่ได้ผูกบัญชี' };
    var rows = _readAll('LeaveRequests');
    for (var i = 0; i < rows.length; i++) {
      var d = rows[i].data;
      if (d.id === leaveId && d.employee_id === emp.id) {
        if (d.status !== 'SUBMITTED' && d.status !== 'L1_APPROVED') {
          return { status: 'error', message: 'ไม่สามารถยกเลิกใบลาที่ ' + d.status + ' แล้ว' };
        }
        var year = parseInt(d.start_date.substring(0, 4), 10);
        d.status = 'CANCELLED';
        if (reason) d.cancel_reason = reason;
        d.updated_at = _now();
        _update('LeaveRequests', rows[i].rowIndex, d);
        _adjustBalance(emp.id, d.leave_type_id, year, -d.total_days, 0);
        _notifyCancelToApprovers(d);
        return { status: 'success', message: 'ยกเลิกใบลาแล้ว' };
      }
    }
    return { status: 'error', message: 'ไม่พบใบลา' };
  } catch (err) {
    logError(err.toString(), 'liffCancelLeave');
    return { status: 'error', message: err.toString() };
  }
}

function _notifyCancelToApprovers(req) {
  try {
    var cfg = _getConfigRow().data;
    if (!cfg.notification_enabled) return;
    var ids = _approverLineIds(req, req.current_level, cfg);
    var msg = {
      type: 'flex', altText: 'ยกเลิกใบลา: ' + req.request_no,
      contents: {
        type: 'bubble',
        header: { type: 'box', layout: 'vertical', backgroundColor: '#94a3b8', paddingAll: '14px',
          contents: [{ type: 'text', text: 'ยกเลิกใบลา', color: '#ffffff', weight: 'bold' },
                     { type: 'text', text: req.request_no, color: '#e2e8f0', size: 'xs', margin: 'sm' }] },
        body: { type: 'box', layout: 'vertical', spacing: 'sm',
          contents: [
            { type: 'text', text: req.employee_name, weight: 'bold' },
            _flexRow('ประเภท', req.leave_type_name),
            _flexRow('ช่วงวัน', _rangeText(req)),
            _flexRow('จำนวน', req.total_days + ' วัน')
          ] }
      }
    };
    ids.forEach(function(id){ _linePush(id, [msg]); });
  } catch(err) { logError(err.toString(), '_notifyCancelToApprovers'); }
}

function liffGetLeaveDetail(lineUserId, leaveId) {
  try {
    var emp = _empByLine(lineUserId);
    if (!emp) return { status: 'error', message: 'ยังไม่ได้ผูกบัญชี' };
    var rows = _readAll('LeaveRequests');
    for (var i = 0; i < rows.length; i++) {
      var d = rows[i].data;
      if (d.id !== leaveId) continue;
      if (d.employee_id !== emp.id) return { status: 'error', message: 'ไม่มีสิทธิ์ดูใบลานี้' };
      var approvals = d.approvals || [];
      var lastApproval = approvals.length ? approvals[approvals.length - 1] : {};
      var approvedByName = '';
      if (lastApproval.approver_id) {
        var approver = _empById(lastApproval.approver_id);
        if (approver) approvedByName = approver.name;
      }
      var attachUrl = '';
      if (d.attachment_file_ids && d.attachment_file_ids.length) {
        attachUrl = 'https://drive.google.com/file/d/' + d.attachment_file_ids[0] + '/view';
      }
      return {
        status: 'success',
        data: {
          id: d.id, request_no: d.request_no, leave_type_name: d.leave_type_name,
          status: d.status, start_date: d.start_date, end_date: d.end_date,
          start_half: d.start_half, total_days: d.total_days,
          reason: d.reason || '', contact_during_leave: d.contact_during_leave || '',
          last_comment: lastApproval.comment || '',
          approved_by_name: approvedByName,
          created_at: d.created_at || '',
          attachment_url: attachUrl
        }
      };
    }
    return { status: 'error', message: 'ไม่พบใบลา' };
  } catch (err) {
    logError(err.toString(), 'liffGetLeaveDetail');
    return { status: 'error', message: err.toString() };
  }
}

function liffGetHolidays(year) {
  try {
    year = parseInt(year, 10) || new Date().getFullYear();
    var rows = _readAll('Holidays');
    var list = rows.map(function(r){ return r.data; })
      .filter(function(h){ return h.date && String(h.date).substring(0,4) === String(year); })
      .sort(function(a, b){ return a.date < b.date ? -1 : 1; })
      .map(function(h){ return { date: h.date, name: h.name, type: h.type || 'public' }; });
    return { status: 'success', data: list };
  } catch (err) {
    logError(err.toString(), 'liffGetHolidays');
    return { status: 'error', message: err.toString() };
  }
}

function liffUploadAttachment(lineUserId, base64Data, filename, mimeType) {
  var emp = _empByLine(lineUserId);
  if (!emp) return { status: 'error', message: 'ยังไม่ได้ผูกบัญชี' };
  try {
    var folder = DriveApp.getFolderById(_ensureFolder());
    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType || 'image/jpeg', filename || 'attach.jpg');
    var file = folder.createFile(blob);
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    return { status: 'success', file_id: file.getId() };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function liffSubmitLeave(lineUserId, obj) {
  try {
    var emp = _empByLine(lineUserId);
    if (!emp) return { status: 'error', message: 'ยังไม่ได้ผูกบัญชี กรุณาผูกบัญชีก่อน' };

    var type = _typeById(obj.leave_type_id);
    if (!type) return { status: 'error', message: 'ไม่พบประเภทการลา' };
    if (!obj.start_date || !obj.end_date) return { status: 'error', message: 'กรุณาเลือกวันที่' };

    var startHalf = obj.start_half || 'full';
    var endHalf = obj.end_half || 'full';
    var total = _calcLeaveDays(obj.start_date, obj.end_date, startHalf, endHalf);
    if (total <= 0) return { status: 'error', message: 'ช่วงวันที่เลือกไม่มีวันทำงาน' };

    var year = parseInt(obj.start_date.substring(0, 4), 10);
    if (Number(type.default_quota_days) > 0) {
      var remain = _remaining(emp.id, type.id, year);
      if (total > remain) return { status: 'error', message: 'วันลาคงเหลือไม่พอ (เหลือ ' + remain + ' วัน)' };
    }
    if (type.requires_attachment && (!obj.attachment_file_ids || !obj.attachment_file_ids.length)) {
      return { status: 'error', message: 'ประเภทนี้ต้องแนบเอกสาร' };
    }

    var beYear = year + 543;
    var req = {
      id: _uuid(), request_no: _genRequestNo(beYear),
      employee_id: emp.id,
      employee_name: (emp.prefix || '') + emp.first_name + ' ' + emp.last_name,
      department_id: emp.department_id || '',
      leave_type_id: type.id, leave_type_name: type.name,
      start_date: obj.start_date, end_date: obj.end_date,
      start_half: startHalf, end_half: endHalf, total_days: total,
      reason: obj.reason || '', contact_during_leave: obj.contact_during_leave || '',
      attachment_file_ids: obj.attachment_file_ids || [],
      status: 'SUBMITTED', current_level: 1, approvals: [],
      submitted_at: _now(), decided_at: '', created_at: _now(), updated_at: _now()
    };
    _append('LeaveRequests', req);
    _adjustBalance(emp.id, type.id, year, total, 0);

    _notifyApprovers(req, 1);
    _notifyEmployeeSubmit(req);
    return { status: 'success', message: 'ยื่นใบลาเรียบร้อย', request_no: req.request_no };
  } catch (err) {
    logError(err.toString(), 'liffSubmitLeave');
    return { status: 'error', message: err.toString() };
  }
}

// ============================================================
// PHASE 6 — LINE MESSAGING (webhook, push, flex, notify)
// ============================================================
function _lineToken() {
  var c = _getConfigRow();
  return c ? (c.data.line_channel_access_token || '') : '';
}

function _linePush(to, messages) {
  var token = _lineToken();
  if (!token || !to) return;
  try {
    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + token },
      payload: JSON.stringify({ to: to, messages: messages }),
      muteHttpExceptions: true
    });
  } catch (err) { logError(err.toString(), '_linePush'); }
}

function _lineReply(replyToken, messages) {
  var token = _lineToken();
  if (!token || !replyToken) return;
  try {
    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + token },
      payload: JSON.stringify({ replyToken: replyToken, messages: messages }),
      muteHttpExceptions: true
    });
  } catch (err) { logError(err.toString(), '_lineReply'); }
}

/** รายชื่อ lineUserId ของผู้อนุมัติในระดับที่กำหนด */
function _approverLineIds(req, level, cfg) {
  var ids = [];
  var role = level === 1 ? (cfg.level1_role || 'supervisor') : (cfg.level2_role || 'hr');

  // ถ้าระดับนี้เป็น supervisor -> หาหัวหน้าแผนก (Users) ของแผนกที่พนักงานสังกัด
  if (role === 'supervisor' && req.department_id) {
    var depts = _readAll('Departments');
    for (var i = 0; i < depts.length; i++) {
      if (depts[i].data.id === req.department_id) {
        var headIds = Array.isArray(depts[i].data.head_user_ids) ? depts[i].data.head_user_ids
          : (depts[i].data.head_user_id ? [depts[i].data.head_user_id] : []);
        if (headIds.length) {
          var userRows = _readAll('Users');
          userRows.forEach(function(ur){
            var u = ur.data;
            if (headIds.indexOf(u.id) !== -1 && u.line_user_id) ids.push(u.line_user_id);
          });
        }
        break;
      }
    }
  }

  // Users ที่ role ตรง + ผูก LINE (และ admin เสมอ)
  _readAll('Users').forEach(function (r) {
    var u = r.data;
    if (!u.active || !u.line_user_id) return;
    if (u.role === role || u.role === 'admin') {
      if (ids.indexOf(u.line_user_id) === -1) ids.push(u.line_user_id);
    }
  });
  return ids;
}

function _fmtDateTH(iso) {
  if (!iso) return '-';
  var p = iso.split('-');
  var m = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
  return Number(p[2]) + ' ' + m[Number(p[1]) - 1] + ' ' + (Number(p[0]) + 543);
}

function _rangeText(req) {
  var t = _fmtDateTH(req.start_date);
  if (req.start_date !== req.end_date) t += ' - ' + _fmtDateTH(req.end_date);
  if (req.start_half === 'morning') t += ' (ครึ่งเช้า)';
  else if (req.start_half === 'afternoon') t += ' (ครึ่งบ่าย)';
  return t;
}

/** Flex card สำหรับผู้อนุมัติ (มีปุ่ม postback) */
function _flexApproval(req, level) {
  return {
    type: 'flex',
    altText: 'ใบลารออนุมัติ: ' + req.employee_name,
    contents: {
      type: 'bubble',
      header: {
        type: 'box', layout: 'vertical', backgroundColor: '#3b3f9e', paddingAll: '16px',
        contents: [
          { type: 'text', text: 'คำขอลา · รออนุมัติระดับ ' + level, color: '#ffffff', size: 'sm' },
          { type: 'text', text: req.request_no, color: '#cdd0ff', size: 'xs', margin: 'sm' }
        ]
      },
      body: {
        type: 'box', layout: 'vertical', spacing: 'sm',
        contents: [
          { type: 'text', text: req.employee_name, weight: 'bold', size: 'md' },
          _flexRow('แผนก', req.department_name || '-'),
          _flexRow('ประเภท', req.leave_type_name),
          _flexRow('ช่วงวัน', _rangeText(req)),
          _flexRow('จำนวน', req.total_days + ' วัน'),
          _flexRow('เหตุผล', req.reason || '-')
        ]
      },
      footer: {
        type: 'box', layout: 'horizontal', spacing: 'sm',
        contents: [
          { type: 'button', style: 'secondary', height: 'sm',
            action: { type: 'postback', label: 'ไม่อนุมัติ', data: 'action=reject&id=' + req.id, displayText: 'ไม่อนุมัติ ' + req.request_no } },
          { type: 'button', style: 'primary', height: 'sm', color: '#3b3f9e',
            action: { type: 'postback', label: 'อนุมัติ', data: 'action=approve&id=' + req.id, displayText: 'อนุมัติ ' + req.request_no } }
        ]
      }
    }
  };
}

function _flexRow(label, value) {
  return {
    type: 'box', layout: 'baseline', spacing: 'sm',
    contents: [
      { type: 'text', text: label, color: '#94a3b8', size: 'sm', flex: 2 },
      { type: 'text', text: String(value), color: '#1e293b', size: 'sm', flex: 5, wrap: true }
    ]
  };
}

function _notifyApprovers(req, level) {
  try {
    var cfg = _getConfigRow().data;
    if (!cfg.notification_enabled) return;
    var ids = _approverLineIds(req, level, cfg);
    var msg = _flexApproval(req, level);
    ids.forEach(function (id) { _linePush(id, [msg]); });
  } catch (err) { logError(err.toString(), '_notifyApprovers'); }
}

function _notifyEmployeeDecision(req, decision, actorName, comment) {
  try {
    var cfg = _getConfigRow().data;
    if (!cfg.notification_enabled) return;
    var emp = _empById(req.employee_id);
    if (!emp || !emp.line_user_id) return;

    var approved = decision === 'approve';
    var color = approved ? '#10b981' : '#ef4444';
    var head = approved ? 'ใบลาได้รับการอนุมัติ' : 'ใบลาไม่ได้รับการอนุมัติ';
    var body = [
      { type: 'text', text: req.leave_type_name, weight: 'bold', size: 'md' },
      _flexRow('เลขที่', req.request_no),
      _flexRow('ช่วงวัน', _rangeText(req)),
      _flexRow('จำนวน', req.total_days + ' วัน'),
      _flexRow('โดย', actorName || '-')
    ];
    if (comment) body.push(_flexRow('หมายเหตุ', comment));

    _linePush(emp.line_user_id, [{
      type: 'flex', altText: head + ': ' + req.request_no,
      contents: {
        type: 'bubble',
        header: { type: 'box', layout: 'vertical', backgroundColor: color, paddingAll: '16px',
          contents: [{ type: 'text', text: head, color: '#ffffff', weight: 'bold' }] },
        body: { type: 'box', layout: 'vertical', spacing: 'sm', contents: body }
      }
    }]);
  } catch (err) { logError(err.toString(), '_notifyEmployeeDecision'); }
}

function _notifyEmployeeSubmit(req) {
  try {
    var cfg = _getConfigRow().data;
    if (!cfg.notification_enabled) return;
    var emp = _empById(req.employee_id);
    if (!emp || !emp.line_user_id) return;
    _linePush(emp.line_user_id, [{
      type: 'flex', altText: 'รับใบลาแล้ว: ' + req.request_no,
      contents: {
        type: 'bubble',
        header: { type: 'box', layout: 'vertical', backgroundColor: '#0ea5e9', paddingAll: '14px',
          contents: [{ type: 'text', text: 'รับใบลาแล้ว · รออนุมัติ', color: '#ffffff', weight: 'bold' },
                     { type: 'text', text: req.request_no, color: '#e0f2fe', size: 'xs', margin: 'sm' }] },
        body: { type: 'box', layout: 'vertical', spacing: 'sm',
          contents: [
            { type: 'text', text: req.leave_type_name, weight: 'bold' },
            _flexRow('ช่วงวัน', _rangeText(req)),
            _flexRow('จำนวน', req.total_days + ' วัน'),
            _flexRow('เหตุผล', req.reason || '-')
          ] }
      }
    }]);
  } catch(err) { logError(err.toString(), '_notifyEmployeeSubmit'); }
}

// ---------- WEBHOOK ----------
// ฟังก์ชันที่อนุญาตให้เรียกผ่าน JSON API (จากหน้า LIFF บน static host)
var LIFF_API_WHITELIST = {
  liffGetConfig: true, getEmployeeByLineId: true, linkLineAccount: true,
  liffGetLeaveTypes: true, liffGetBalances: true, liffPreviewDays: true,
  liffUploadAttachment: true, liffSubmitLeave: true, liffGetHistory: true,
  liffCancelLeave: true, liffGetLeaveDetail: true, liffGetHolidays: true,
  liffGetMyStats: true
};

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);

    // ---- JSON API (เรียกจากหน้า LIFF บน GitHub Pages ผ่าน fetch) ----
    if (body && body.fn) {
      var out;
      if (LIFF_API_WHITELIST[body.fn]) {
        out = { ok: true, data: globalThis[body.fn].apply(null, body.args || []) };
      } else {
        out = { ok: false, error: 'ฟังก์ชันไม่ได้รับอนุญาต: ' + body.fn };
      }
      return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
    }

    // ---- LINE webhook ----
    (body.events || []).forEach(function (ev) {
      if (ev.type === 'postback') _handlePostback(ev);
      else if (ev.type === 'message' && ev.message.type === 'text') _handleText(ev);
      else if (ev.type === 'follow') _handleFollow(ev);
    });
  } catch (err) {
    logError(err.toString(), 'doPost');
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' })).setMimeType(ContentService.MimeType.JSON);
}

function _parseQuery(s) {
  var o = {};
  (s || '').split('&').forEach(function (kv) {
    var p = kv.split('=');
    o[decodeURIComponent(p[0])] = decodeURIComponent(p[1] || '');
  });
  return o;
}

function _handlePostback(ev) {
  var q = _parseQuery(ev.postback.data);
  var lineUserId = ev.source.userId;

  if (q.action === 'approve' || q.action === 'reject') {
    // หา User ที่ผูก LINE นี้
    var actor = null;
    _readAll('Users').forEach(function (r) {
      if (r.data.line_user_id === lineUserId && r.data.active) actor = r.data;
    });
    if (!actor) {
      _lineReply(ev.replyToken, [{ type: 'text', text: 'บัญชี LINE นี้ไม่มีสิทธิ์อนุมัติ (ต้องเป็นผู้อนุมัติที่ผูก LINE ในระบบ)' }]);
      return;
    }
    if (q.action === 'reject') {
      // ปฏิเสธผ่าน LINE ใช้เหตุผลมาตรฐาน
      var rr = _processDecision(actor.name, actor.role, q.id, 'reject', 'ไม่อนุมัติผ่าน LINE');
      _lineReply(ev.replyToken, [{ type: 'text', text: rr.message }]);
    } else {
      var ar = _processDecision(actor.name, actor.role, q.id, 'approve', 'อนุมัติผ่าน LINE');
      _lineReply(ev.replyToken, [{ type: 'text', text: ar.message }]);
    }
  }
}

function _handleText(ev) {
  var lineUserId = ev.source.userId;
  var emp = _empByLine(lineUserId);
  var cfg = _getConfigRow().data;
  var liffUrl = cfg.line_liff_id ? ('https://liff.line.me/' + cfg.line_liff_id) : '';

  if (!emp) {
    _lineReply(ev.replyToken, [{ type: 'text', text: 'ยินดีต้อนรับ กรุณาผูกบัญชีก่อนใช้งาน' + (liffUrl ? '\n' + liffUrl : '') }]);
    return;
  }
  _lineReply(ev.replyToken, [{ type: 'text', text: 'สวัสดีคุณ ' + (emp.prefix || '') + emp.first_name + '\nเปิดเมนูเพื่อยื่นลา/ดูสถานะ' + (liffUrl ? '\n' + liffUrl : '') }]);
}

function _handleFollow(ev) {
  var cfg = _getConfigRow().data;
  var liffUrl = cfg.line_liff_id ? ('https://liff.line.me/' + cfg.line_liff_id) : '';
  _lineReply(ev.replyToken, [{ type: 'text', text: 'ขอบคุณที่เพิ่มเพื่อน ' + (cfg.app_name || 'ระบบลาออนไลน์') + '\nเริ่มต้นใช้งานได้ที่' + (liffUrl ? '\n' + liffUrl : 'เมนูด้านล่าง') }]);
}

/** ทดสอบส่ง push หา lineUserId (รันจาก editor หรือเรียกจากตั้งค่า) */
function testLinePush(token, lineUserId) {
  _auth(token);
  _linePush(lineUserId, [{ type: 'text', text: 'ทดสอบการแจ้งเตือนจากระบบลาออนไลน์ ✓' }]);
  return { status: 'success', message: 'ส่งข้อความทดสอบแล้ว' };
}

// ============================================================
// PHASE 7 — DASHBOARD (real-time)
// ============================================================
function getDashboardData(token) {
  _auth(token);
  var today = _isoDate(new Date());
  var now = new Date();
  var year = now.getFullYear();
  var month = now.getMonth() + 1;

  var typesMap = {};
  _readAll('LeaveTypes').forEach(function (r) { typesMap[r.data.id] = r.data; });

  var reqs = _readAll('LeaveRequests').map(function (r) { return r.data; });
  var emps = _readAll('Employees').map(function (r) { return r.data; });

  var pending = 0, onLeaveToday = 0, leaveThisMonth = 0;
  var byType = {};        // typeId -> days (approved ปีนี้)
  var monthly = [0,0,0,0,0,0,0,0,0,0,0,0]; // approved ปีนี้ ตามเดือนเริ่มลา
  var todayList = [];

  // คำนวณช่วงสัปดาห์นี้ (จันทร์-อาทิตย์)
  var nowDay = now.getDay(); // 0=Sun, 1=Mon, ...
  var diffToMon = (nowDay === 0) ? -6 : 1 - nowDay;
  var monday = new Date(now.getFullYear(), now.getMonth(), now.getDate() + diffToMon);
  var sunday = new Date(monday.getTime() + 6 * 24 * 60 * 60 * 1000);
  var mondayStr = _isoDate(monday);
  var sundayStr = _isoDate(sunday);
  var weekList = [];

  reqs.forEach(function (d) {
    if (d.status === 'SUBMITTED' || d.status === 'L1_APPROVED') pending++;

    if (d.status === 'APPROVED') {
      var sy = parseInt(d.start_date.substring(0, 4), 10);
      var sm = parseInt(d.start_date.substring(5, 7), 10);

      // อยู่ระหว่างลา วันนี้
      if (d.start_date <= today && today <= d.end_date) {
        onLeaveToday++;
        todayList.push({
          employee_name: d.employee_name,
          leave_type_name: d.leave_type_name,
          color: (typesMap[d.leave_type_id] || {}).color || '#64748b',
          start_date: d.start_date, end_date: d.end_date, total_days: d.total_days
        });
      }
      // อยู่ระหว่างลาสัปดาห์นี้
      if (d.start_date <= sundayStr && d.end_date >= mondayStr) {
        weekList.push({
          employee_name: d.employee_name,
          leave_type_name: d.leave_type_name,
          color: (typesMap[d.leave_type_id] || {}).color || '#64748b',
          start_date: d.start_date, end_date: d.end_date, total_days: d.total_days
        });
      }
      // ลาเดือนนี้
      if (sy === year && sm === month) leaveThisMonth++;
      // donut + monthly (ปีนี้)
      if (sy === year) {
        byType[d.leave_type_id] = (byType[d.leave_type_id] || 0) + Number(d.total_days);
        monthly[sm - 1] += Number(d.total_days);
      }
    }
  });

  var donut = [];
  Object.keys(byType).forEach(function (tid) {
    var t = typesMap[tid] || {};
    donut.push({ name: t.name || 'อื่นๆ', color: t.color || '#64748b', days: byType[tid] });
  });
  donut.sort(function (a, b) { return b.days - a.days; });

  return {
    status: 'success',
    year: year,
    kpis: {
      pending_approval: pending,
      on_leave_today: onLeaveToday,
      leave_this_month: leaveThisMonth,
      total_employees: emps.filter(function (e) { return e.status === 'active'; }).length
    },
    donut: donut,
    monthly: monthly,
    today_list: todayList,
    week_list: weekList
  };
}

function carryForwardBalances(token, fromYear) {
  _auth(token);
  try {
    fromYear = Number(fromYear) || (new Date().getFullYear() - 1);
    var toYear = fromYear + 1;
    var types = {};
    _readAll('LeaveTypes').forEach(function(r){ types[r.data.id] = r.data; });
    var fromBals = _readAll('LeaveBalances').map(function(r){ return r.data; })
      .filter(function(b){ return String(b.year) === String(fromYear); });
    var count = 0;
    fromBals.forEach(function(b) {
      var t = types[b.leave_type_id];
      if (!t || !t.carry_over) return;
      var remaining = Number(b.entitled_days) + Number(b.carried_over||0) - Number(b.used_days) - Number(b.pending_days);
      if (remaining <= 0) return;
      var maxCarry = Number(t.max_carry_over_days) || remaining;
      var carryDays = Math.min(remaining, maxCarry);
      var toBals = _readAll('LeaveBalances').map(function(r){ return r; })
        .filter(function(r){ return r.data.employee_id === b.employee_id && r.data.leave_type_id === b.leave_type_id && String(r.data.year) === String(toYear); });
      if (toBals.length > 0) {
        var tb = toBals[0].data;
        tb.carried_over = (Number(tb.carried_over)||0) + carryDays;
        tb.updated_at = _now();
        _update('LeaveBalances', toBals[0].rowIndex, tb);
      } else {
        _append('LeaveBalances', {
          id: _uuid(), employee_id: b.employee_id, year: toYear,
          leave_type_id: b.leave_type_id,
          entitled_days: Number(t.default_quota_days)||0,
          used_days: 0, pending_days: 0, carried_over: carryDays, updated_at: _now()
        });
      }
      count++;
    });
    return { status: 'success', message: 'ยกยอดวันลาแล้ว ' + count + ' รายการ (ปี ' + (fromYear+543) + ' → ' + (toYear+543) + ')' };
  } catch(err) {
    logError(err.toString(), 'carryForwardBalances');
    return { status: 'error', message: err.toString() };
  }
}

function _getUserByToken(token) {
  var rows = _readAll('Sessions');
  for (var i = 0; i < rows.length; i++) {
    var s = rows[i].data;
    if (s.token === token && new Date(s.expires_at) >= new Date()) {
      var users = _readAll('Users');
      for (var j = 0; j < users.length; j++) {
        if (users[j].data.username === s.username) return users[j].data;
      }
    }
  }
  return null;
}

function bulkApproveLeave(token, ids, comment) {
  _auth(token);
  var user = _getUserByToken(token);
  if (!user) return { status: 'error', message: 'ไม่พบผู้ใช้' };
  var results = [];
  (ids || []).forEach(function(id) {
    var r = _processDecision(user.name, user.role, id, 'approve', comment || '');
    results.push({ id: id, status: r.status, message: r.message });
  });
  var ok = results.filter(function(r){ return r.status === 'success'; }).length;
  return { status: 'success', message: 'อนุมัติสำเร็จ ' + ok + '/' + results.length + ' รายการ', results: results };
}

function bulkRejectLeave(token, ids, comment) {
  _auth(token);
  if (!comment) return { status: 'error', message: 'กรุณาระบุเหตุผล' };
  var user = _getUserByToken(token);
  if (!user) return { status: 'error', message: 'ไม่พบผู้ใช้' };
  var results = [];
  (ids || []).forEach(function(id) {
    var r = _processDecision(user.name, user.role, id, 'reject', comment);
    results.push({ id: id, status: r.status, message: r.message });
  });
  var ok = results.filter(function(r){ return r.status === 'success'; }).length;
  return { status: 'success', message: 'ไม่อนุมัติ ' + ok + '/' + results.length + ' รายการ', results: results };
}

function setupRichMenu(token) {
  _auth(token);
  try {
    var cfg = _getConfigRow().data;
    var lineToken = cfg.line_channel_access_token || '';
    if (!lineToken) return { status: 'error', message: 'ยังไม่ได้ตั้งค่า Channel Access Token' };
    var liffId = cfg.line_liff_id || '';
    var liffUrl = liffId ? ('https://liff.line.me/' + liffId) : '';
    if (!liffUrl) return { status: 'error', message: 'ยังไม่ได้ตั้งค่า LIFF ID' };

    var richMenu = {
      size: { width: 2500, height: 843 },
      selected: true,
      name: 'ระบบลาออนไลน์',
      chatBarText: 'เมนูการลา',
      areas: [
        { bounds: { x: 0, y: 0, width: 833, height: 843 },
          action: { type: 'uri', label: 'ยื่นใบลา', uri: liffUrl + '?tab=new' } },
        { bounds: { x: 833, y: 0, width: 834, height: 843 },
          action: { type: 'uri', label: 'ประวัติการลา', uri: liffUrl + '?tab=history' } },
        { bounds: { x: 1667, y: 0, width: 833, height: 843 },
          action: { type: 'uri', label: 'วันคงเหลือ', uri: liffUrl + '?tab=balance' } }
      ]
    };

    var headers = { 'Authorization': 'Bearer ' + lineToken, 'Content-Type': 'application/json' };

    // Delete existing rich menus first
    try {
      var listRes = UrlFetchApp.fetch('https://api.line.me/v2/bot/richmenu/list', { headers: headers, muteHttpExceptions: true });
      var listData = JSON.parse(listRes.getContentText());
      (listData.richmenus || []).forEach(function(rm) {
        UrlFetchApp.fetch('https://api.line.me/v2/bot/richmenu/' + rm.richMenuId, { method: 'delete', headers: headers, muteHttpExceptions: true });
      });
    } catch(e) {}

    var res = UrlFetchApp.fetch('https://api.line.me/v2/bot/richmenu', {
      method: 'post', contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + lineToken },
      payload: JSON.stringify(richMenu), muteHttpExceptions: true
    });
    var data = JSON.parse(res.getContentText());
    if (!data.richMenuId) return { status: 'error', message: 'สร้าง Rich Menu ไม่สำเร็จ: ' + res.getContentText() };

    // Set as default
    UrlFetchApp.fetch('https://api.line.me/v2/bot/user/all/richmenu/' + data.richMenuId, {
      method: 'post', headers: headers, muteHttpExceptions: true
    });

    return { status: 'success', message: 'ตั้งค่า Rich Menu สำเร็จ (richMenuId: ' + data.richMenuId + ') — อัพโหลดรูปภาพเมนูใน LINE Console เพื่อให้แสดงผล', richMenuId: data.richMenuId };
  } catch(err) {
    logError(err.toString(), 'setupRichMenu');
    return { status: 'error', message: err.toString() };
  }
}

// ============================================================
// PHASE 8 — REPORTS / EXPORT
// ============================================================
var _ST_TH = { SUBMITTED: 'รออนุมัติ', L1_APPROVED: 'ผ่านระดับ 1', APPROVED: 'อนุมัติ', REJECTED: 'ไม่อนุมัติ', CANCELLED: 'ยกเลิก' };

function reportData(token, filters) {
  _auth(token);
  filters = filters || {};
  var year = filters.year || new Date().getFullYear();
  var statusFilter = filters.status || 'APPROVED';
  var groupBy = filters.groupBy || 'employee';

  var depts = {};
  _readAll('Departments').forEach(function (r) { depts[r.data.id] = r.data.name; });

  var rows = _readAll('LeaveRequests').map(function (r) { return r.data; }).filter(function (d) {
    var sy = parseInt(d.start_date.substring(0, 4), 10);
    if (String(sy) !== String(year)) return false;
    if (statusFilter !== 'ALL' && d.status !== statusFilter) return false;
    if (filters.department_id && d.department_id !== filters.department_id) return false;
    if (filters.dateFrom && d.start_date < filters.dateFrom) return false;
    if (filters.dateTo && d.start_date > filters.dateTo) return false;
    return true;
  });

  var detail = rows.map(function (d) {
    return {
      request_no: d.request_no, employee_name: d.employee_name,
      department_name: depts[d.department_id] || '', leave_type_name: d.leave_type_name,
      start_date: d.start_date, end_date: d.end_date, total_days: d.total_days,
      status: d.status, status_th: _ST_TH[d.status] || d.status, reason: d.reason || ''
    };
  });
  detail.sort(function (a, b) { return a.start_date < b.start_date ? -1 : 1; });

  var map = {};
  rows.forEach(function (d) {
    var key = groupBy === 'department' ? (depts[d.department_id] || 'ไม่ระบุแผนก')
            : groupBy === 'type' ? d.leave_type_name
            : d.employee_name;
    if (!map[key]) map[key] = { key: key, count: 0, days: 0 };
    map[key].count++;
    map[key].days += Number(d.total_days);
  });
  var summary = Object.keys(map).map(function (k) { return map[k]; });
  summary.sort(function (a, b) { return b.days - a.days; });

  var totalDays = 0;
  rows.forEach(function (d) { totalDays += Number(d.total_days); });

  return {
    status: 'success', year: year, groupBy: groupBy, statusFilter: statusFilter,
    summary: summary, detail: detail,
    totals: { count: rows.length, days: totalDays }
  };
}

function _csvCell(v) {
  v = String(v == null ? '' : v);
  if (/[",\n\r]/.test(v)) return '"' + v.replace(/"/g, '""') + '"';
  return v;
}

function exportLeaveCSV(token, filters) {
  var rep = reportData(token, filters);
  var lines = [];
  lines.push(['เลขที่', 'พนักงาน', 'แผนก', 'ประเภท', 'วันเริ่ม', 'วันสิ้นสุด', 'จำนวนวัน', 'สถานะ', 'เหตุผล'].join(','));
  rep.detail.forEach(function (d) {
    lines.push([d.request_no, d.employee_name, d.department_name, d.leave_type_name,
      d.start_date, d.end_date, d.total_days, d.status_th, d.reason].map(_csvCell).join(','));
  });
  var text = '\uFEFF' + lines.join('\r\n');
  var b64 = Utilities.base64Encode(Utilities.newBlob(text, 'text/csv').getBytes());
  return { status: 'success', filename: 'leave-report-' + (rep.year + 543) + '.csv', base64: b64, mime: 'text/csv' };
}
