/**
 * Maintenance Tracker - Google Apps Script Backend
 * Valeo belső használatra
 */

// ============ KONFIGURÁCIÓ ============
const SPREADSHEET_ID = '1HQsXLyUbwnoF6AAtmtwUoyAXVWIQ4LBn4xqgJDFE7v0'; // <- IDE ÍROD A SHEET ID-T
const SHEET_NAMES = {
  ASSETS: 'Assets',
  TASKS: 'Tasks',
  USERS: 'Users',
  CONFIG: 'Config'
};

// ============ ESD SEGÉDFÜGGVÉNYEK ============
function calcEsdNextDate(baseDate, cycleMonths) {
  const d = new Date(baseDate);
  d.setMonth(d.getMonth() + parseInt(cycleMonths));
  return new Date(d.getFullYear(), d.getMonth() + 1, 0); // hónap utolsó napja
}

function calcEsdStatus(nextDate, today) {
  const ty = today.getFullYear(), tm = today.getMonth();
  const ny = nextDate.getFullYear(), nm = nextDate.getMonth();
  if (ty > ny || (ty === ny && tm > nm)) return 'overdue';
  if (ty === ny && tm === nm) return 'due_soon';
  return 'ok';
}

// ============ WEB APP ENTRY POINTS ============
function doGet(e) {
  const page = e.parameter.page || 'index';
  const template = HtmlService.createTemplateFromFile(page);
  return template.evaluate()
    .setTitle('Maintenance Tracker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============ INICIALIZÁLÁS ============
/**
 * Létrehozza a szükséges sheeteket ha nem léteznek
 * Ezt egyszer kell futtatni manuálisan
 */
function initializeSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // Assets sheet
  let assetsSheet = ss.getSheetByName(SHEET_NAMES.ASSETS);
  if (!assetsSheet) {
    assetsSheet = ss.insertSheet(SHEET_NAMES.ASSETS);
    assetsSheet.getRange(1, 1, 1, 7).setValues([[
      'asset_id', 'name', 'type', 'zone', 'cycle_days', 'drive_link', 'created_at'
    ]]);
    assetsSheet.setFrozenRows(1);
    assetsSheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#f3f3f3');
  }
  
  // Tasks sheet
  let tasksSheet = ss.getSheetByName(SHEET_NAMES.TASKS);
  if (!tasksSheet) {
    tasksSheet = ss.insertSheet(SHEET_NAMES.TASKS);
    tasksSheet.getRange(1, 1, 1, 8).setValues([[
      'task_id', 'asset_id', 'task_type', 'category', 'description', 'completed_at', 'completed_by', 'photo_link'
    ]]);
    tasksSheet.setFrozenRows(1);
    tasksSheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#f3f3f3');
  }
  
  // Users sheet
  let usersSheet = ss.getSheetByName(SHEET_NAMES.USERS);
  if (!usersSheet) {
    usersSheet = ss.insertSheet(SHEET_NAMES.USERS);
    usersSheet.getRange(1, 1, 1, 4).setValues([[
      'email', 'name', 'role', 'notify'
    ]]);
    usersSheet.setFrozenRows(1);
    usersSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#f3f3f3');
    
    // Alapértelmezett admin hozzáadása (a te emailed)
    usersSheet.appendRow([Session.getActiveUser().getEmail(), 'Admin', 'admin', true]);
  }
  
  // Config sheet
  let configSheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
  if (!configSheet) {
    configSheet = ss.insertSheet(SHEET_NAMES.CONFIG);
    configSheet.getRange(1, 1, 1, 3).setValues([['key', 'value', 'description']]);
    configSheet.setFrozenRows(1);
    configSheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#f3f3f3');
    
    // Alapértelmezett konfiguráció
    configSheet.appendRow(['zones', 'Z1 - Összeszerelés,Z2 - Minőségellenőrzés,Z3 - Raktár,Z4 - Csomagolás', 'Elérhető területek']);
    configSheet.appendRow(['asset_types', 'Workstation,Szék,Gép,Ionizátor', 'Eszköz típusok']);
    configSheet.appendRow(['task_categories', 'Preventív karbantartás,Javítás,Csere,Ellenőrzés', 'Feladat kategóriák']);
    configSheet.appendRow(['notification_day', '23', 'Havi értesítés napja']);
  }
  
  Logger.log('Sheets inicializálva!');
}

// ============ FELHASZNÁLÓ KEZELÉS ============
function getCurrentUser() {
  const email = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const usersSheet = ss.getSheetByName(SHEET_NAMES.USERS);
  const data = usersSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email) {
      return {
        email: data[i][0],
        name: data[i][1],
        role: data[i][2],
        notify: data[i][3]
      };
    }
  }
  
  // Ha nincs a listában, readonly user
  return {
    email: email,
    name: email.split('@')[0],
    role: 'readonly',
    notify: false
  };
}

function isAdmin() {
  const user = getCurrentUser();
  return user.role === 'admin';
}

// ============ CONFIG LEKÉRDEZÉSEK ============
function getConfig(key) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const configSheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
  const data = configSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      return data[i][1];
    }
  }
  return null;
}

function getZones() {
  const zones = getConfig('zones');
  return zones ? zones.split(',').map(z => z.trim()) : [];
}

function getAssetTypes() {
  const types = getConfig('asset_types');
  return types ? types.split(',').map(t => t.trim()) : [];
}

function getTaskCategories() {
  const cats = getConfig('task_categories');
  return cats ? cats.split(',').map(c => c.trim()) : [];
}

// ============ ASSETS CRUD ============
function getAssets(filtersInput) {
  // Handle both object and JSON string input
  let filters = {};
  if (typeof filtersInput === 'string') {
    try {
      filters = JSON.parse(filtersInput);
    } catch(e) {
      filters = {};
    }
  } else if (filtersInput && typeof filtersInput === 'object') {
    filters = filtersInput;
  }
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const assetsSheet = ss.getSheetByName(SHEET_NAMES.ASSETS);
  const tasksSheet = ss.getSheetByName(SHEET_NAMES.TASKS);
  
  if (!assetsSheet || !tasksSheet) {
    throw new Error('Hiányzó sheet! Futtasd az initializeSheets() függvényt.');
  }
  
  const assetsData = assetsSheet.getDataRange().getValues();
  const tasksData = tasksSheet.getDataRange().getValues() || [];
  
  // Utolsó preventív feladat dátumának megkeresése minden eszközhöz
  const lastPreventive = {};
  for (let i = 1; i < tasksData.length; i++) {
    const assetId = tasksData[i][1];
    const taskType = tasksData[i][2];
    const completedAt = tasksData[i][5];
    
    if (taskType === 'preventív' && completedAt) {
      const date = new Date(completedAt);
      if (!lastPreventive[assetId] || date > lastPreventive[assetId]) {
        lastPreventive[assetId] = date;
      }
    }
  }
  
  const assets = [];
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  for (let i = 1; i < assetsData.length; i++) {
    const row = assetsData[i];
    const assetId = row[0];
    const cycleDays = row[4] || 30;
    
    // Következő karbantartás számítása
    let nextMaintenance = null;
    let status = 'ok';
    let daysUntil = null;
    
    if (lastPreventive[assetId]) {
      nextMaintenance = new Date(lastPreventive[assetId]);
      nextMaintenance.setDate(nextMaintenance.getDate() + cycleDays);
      
      const diffTime = nextMaintenance.getTime() - today.getTime();
      daysUntil = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
      
      if (daysUntil < 0) {
        status = 'overdue';
      } else if (daysUntil <= 7) {
        status = 'due_soon';
      }
    } else {
      // Ha még nem volt preventív, a létrehozás dátumától számoljuk
      const createdAt = row[6] ? new Date(row[6]) : new Date();
      nextMaintenance = new Date(createdAt);
      nextMaintenance.setDate(nextMaintenance.getDate() + cycleDays);
      
      const diffTime = nextMaintenance.getTime() - today.getTime();
      daysUntil = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
      
      if (daysUntil < 0) {
        status = 'overdue';
      } else if (daysUntil <= 7) {
        status = 'due_soon';
      }
    }
    
    const asset = {
      asset_id: assetId,
      name: row[1],
      type: row[2],
      zone: row[3],
      cycle_days: cycleDays,
      drive_link: row[5],
      created_at: row[6],
      next_maintenance: nextMaintenance ? Utilities.formatDate(nextMaintenance, Session.getScriptTimeZone(), 'yyyy.MM.dd') : null,
      status: status,
      days_until: daysUntil
    };
    
    // Szűrések
    if (filters.zone && asset.zone !== filters.zone) continue;
    if (filters.type && asset.type !== filters.type) continue;
    if (filters.status) {
      if (filters.status === 'due' && status === 'ok') continue;
      if (filters.status === 'ok' && status !== 'ok') continue;
    }
    if (filters.search) {
      const search = filters.search.toLowerCase();
      if (!asset.asset_id.toLowerCase().includes(search) && 
          !asset.name.toLowerCase().includes(search)) continue;
    }
    
    assets.push(asset);
  }
  
  // Rendezés: először késésben, aztán lejáró, aztán OK
  assets.sort((a, b) => {
    const statusOrder = { overdue: 0, due_soon: 1, ok: 2 };
    if (statusOrder[a.status] !== statusOrder[b.status]) {
      return statusOrder[a.status] - statusOrder[b.status];
    }
    return (a.days_until || 999) - (b.days_until || 999);
  });
  
  return assets;
}

function getAssetById(assetId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const assetsSheet = ss.getSheetByName(SHEET_NAMES.ASSETS);
  
  if (!assetsSheet) {
    return null;
  }
  
  const data = assetsSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === assetId) {
      let createdAt = data[i][6];
      if (createdAt instanceof Date) {
        createdAt = Utilities.formatDate(createdAt, Session.getScriptTimeZone(), 'yyyy.MM.dd');
      }
      
      return {
        asset_id:     data[i][0],
        name:         data[i][1],
        type:         data[i][2],
        zone:         data[i][3],
        cycle_months: parseInt(data[i][4]) || 0,
        drive_link:   data[i][5] || '',
        created_at:   createdAt || '',
        has_esd_cycle: (parseInt(data[i][4]) || 0) > 0
      };
    }
  }
  return null;
}

function createAsset(asset) {
  if (!isAdmin()) {
    throw new Error('Nincs jogosultságod ehhez a művelethez');
  }
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const assetsSheet = ss.getSheetByName(SHEET_NAMES.ASSETS);
  
  // Ellenőrzés: létezik-e már ilyen ID
  const existing = getAssetById(asset.asset_id);
  if (existing) {
    throw new Error('Ez az eszköz ID már létezik: ' + asset.asset_id);
  }
  
  assetsSheet.appendRow([
    asset.asset_id,
    asset.name,
    asset.type,
    asset.zone,
    asset.cycle_days || 30,
    asset.drive_link || '',
    new Date()
  ]);
  
  clearCache();  
  return { success: true, asset_id: asset.asset_id };
}

function updateAsset(assetId, updates) {
  if (!isAdmin()) {
    throw new Error('Nincs jogosultságod ehhez a művelethez');
  }
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const assetsSheet = ss.getSheetByName(SHEET_NAMES.ASSETS);
  const data = assetsSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === assetId) {
      if (updates.name !== undefined) assetsSheet.getRange(i + 1, 2).setValue(updates.name);
      if (updates.type !== undefined) assetsSheet.getRange(i + 1, 3).setValue(updates.type);
      if (updates.zone !== undefined) assetsSheet.getRange(i + 1, 4).setValue(updates.zone);
      if (updates.cycle_days !== undefined) assetsSheet.getRange(i + 1, 5).setValue(updates.cycle_days);
      if (updates.drive_link !== undefined) assetsSheet.getRange(i + 1, 6).setValue(updates.drive_link);
      return { success: true };
    }
  }
  
  throw new Error('Eszköz nem található: ' + assetId);
}

function deleteAsset(assetId) {
  if (!isAdmin()) {
    throw new Error('Nincs jogosultságod ehhez a művelethez');
  }
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const assetsSheet = ss.getSheetByName(SHEET_NAMES.ASSETS);
  const data = assetsSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === assetId) {
      assetsSheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  
  throw new Error('Eszköz nem található: ' + assetId);
}

// ============ TASKS CRUD ============
function getTasks(filters = {}) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const tasksSheet = ss.getSheetByName(SHEET_NAMES.TASKS);
  const assetsSheet = ss.getSheetByName(SHEET_NAMES.ASSETS);
  
  if (!tasksSheet || !assetsSheet) {
    throw new Error('Hiányzó sheet! Futtasd az initializeSheets() függvényt.');
  }
  
  const tasksData = tasksSheet.getDataRange().getValues() || [];
  const assetsData = assetsSheet.getDataRange().getValues() || [];
  
  // Asset lookup
  const assetLookup = {};
  for (let i = 1; i < assetsData.length; i++) {
    assetLookup[assetsData[i][0]] = {
      name: assetsData[i][1],
      type: assetsData[i][2],
      zone: assetsData[i][3]
    };
  }
  
  const tasks = [];
  
  for (let i = 1; i < tasksData.length; i++) {
    const row = tasksData[i];
    const assetId = row[1];
    const asset = assetLookup[assetId] || {};
    
    let completedAtFormatted = '';
    let completedAtRaw = row[5];
    
    if (completedAtRaw) {
      try {
        const dateObj = new Date(completedAtRaw);
        if (!isNaN(dateObj.getTime())) {
          completedAtFormatted = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy.MM.dd');
        }
      } catch (e) {
        completedAtFormatted = '';
      }
    }
    
    const task = {
      task_id: row[0],
      asset_id: assetId,
      asset_name: asset.name || assetId,
      asset_type: asset.type || '',
      asset_zone: asset.zone || '',
      task_type: row[2] || '',
      category: row[3] || '',
      description: row[4] || '',
      completed_at: completedAtFormatted,
      completed_at_raw: completedAtRaw,
      completed_by: row[6] || '',
      photo_link: row[7] || ''
    };
    
    // Szűrések
    if (filters.asset_id && task.asset_id !== filters.asset_id) continue;
    if (filters.task_type && task.task_type !== filters.task_type) continue;
    if (filters.zone && task.asset_zone !== filters.zone) continue;
    
    // Dátum szűrés (hónap/év)
    if (filters.year && filters.month && task.completed_at_raw) {
      const taskDate = new Date(task.completed_at_raw);
      if (taskDate.getFullYear() !== filters.year || taskDate.getMonth() !== filters.month - 1) {
        continue;
      }
    }
    
    tasks.push(task);
  }
  
  // Rendezés: legújabb elöl
  tasks.sort((a, b) => {
    if (!a.completed_at_raw) return 1;
    if (!b.completed_at_raw) return -1;
    return new Date(b.completed_at_raw) - new Date(a.completed_at_raw);
  });
  
  return tasks;
}

function getTasksByAsset(assetId) {
  return getTasks({ asset_id: assetId });
}

function createTask(task) {
  if (!isAdmin()) {
    throw new Error('Nincs jogosultságod ehhez a művelethez');
  }
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const tasksSheet = ss.getSheetByName(SHEET_NAMES.TASKS);
  
  // Következő task_id
  const data = tasksSheet.getDataRange().getValues();
  let maxId = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] > maxId) maxId = data[i][0];
  }
  const newId = maxId + 1;
  
  const user = getCurrentUser();
  
  tasksSheet.appendRow([
    newId,
    task.asset_id,
    task.task_type,
    task.category,
    task.description,
    task.completed_at ? new Date(task.completed_at) : new Date(),
    user.email,
    task.photo_link || ''
  ]);

  clearCache(); 
  
  return { success: true, task_id: newId };
}

function deleteTask(taskId) {
  if (!isAdmin()) {
    throw new Error('Nincs jogosultságod ehhez a művelethez');
  }
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const tasksSheet = ss.getSheetByName(SHEET_NAMES.TASKS);
  const data = tasksSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === taskId) {
      tasksSheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  
  throw new Error('Feladat nem található: ' + taskId);
}

// ============ NAPTÁR ADATOK ============
function getCalendarData(year, month) {
  const tasks = getTasks({ year: year, month: month });
  
  // Csoportosítás napok szerint
  const calendar = {};
  
  tasks.forEach(task => {
    if (task.completed_at_raw) {
      try {
        const date = new Date(task.completed_at_raw);
        if (!isNaN(date.getTime())) {
          const day = date.getDate();
          
          if (!calendar[day]) {
            calendar[day] = [];
          }
          
          calendar[day].push({
            task_id: task.task_id,
            asset_id: task.asset_id,
            asset_name: task.asset_name,
            asset_type: task.asset_type,
            asset_zone: task.asset_zone,
            category: task.category,
            description: task.description ? task.description.substring(0, 50) + (task.description.length > 50 ? '...' : '') : ''
          });
        }
      } catch (e) {
        // Skip invalid dates
      }
    }
  });
  
  return calendar;
}

// ============ STATISZTIKÁK ============
function getStats() {
  const assets = getAssets();
  
  const stats = {
    total: assets.length,
    due_soon: 0,
    overdue: 0,
    ok: 0
  };
  
  assets.forEach(asset => {
    if (asset.status === 'overdue') stats.overdue++;
    else if (asset.status === 'due_soon') stats.due_soon++;
    else stats.ok++;
  });
  
  // Ebben a hónapban elvégzett preventív feladatok
  const today = new Date();
  const tasks = getTasks({ 
    year: today.getFullYear(), 
    month: today.getMonth() + 1 
  });
  stats.completed_this_month = tasks.filter(t => t.task_type === 'preventív').length;
  
  return stats;
}

// ============ EMAIL ÉRTESÍTÉS ============
function sendMonthlyReminder() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const usersSheet = ss.getSheetByName(SHEET_NAMES.USERS);
  const usersData = usersSheet.getDataRange().getValues();
  
  // Lejáró és késésben lévő eszközök
  const assets = getAssets({ status: 'due' });
  const overdue = assets.filter(a => a.status === 'overdue');
  const dueSoon = assets.filter(a => a.status === 'due_soon');
  
  if (overdue.length === 0 && dueSoon.length === 0) {
    Logger.log('Nincs lejáró feladat, email nem szükséges.');
    return;
  }
  
  // Email összeállítása
  let html = '<h2>Maintenance Tracker - Havi összefoglaló</h2>';
  
  if (overdue.length > 0) {
    html += '<h3 style="color: #dc2626;">Késésben lévő eszközök (' + overdue.length + ')</h3><ul>';
    overdue.forEach(a => {
      html += '<li><strong>' + a.asset_id + '</strong> - ' + a.name + ' (' + a.zone + ') - Határidő: ' + a.next_maintenance + '</li>';
    });
    html += '</ul>';
  }
  
  if (dueSoon.length > 0) {
    html += '<h3 style="color: #d97706;">Lejáró eszközök - 7 napon belül (' + dueSoon.length + ')</h3><ul>';
    dueSoon.forEach(a => {
      html += '<li><strong>' + a.asset_id + '</strong> - ' + a.name + ' (' + a.zone + ') - Határidő: ' + a.next_maintenance + '</li>';
    });
    html += '</ul>';
  }
  
  // Értesítendő userek
  for (let i = 1; i < usersData.length; i++) {
    const email = usersData[i][0];
    const notify = usersData[i][3];
    
    if (notify === true || notify === 'TRUE' || notify === 'true') {
      try {
        MailApp.sendEmail({
          to: email,
          subject: 'Maintenance Tracker - ' + overdue.length + ' késésben, ' + dueSoon.length + ' lejáró',
          htmlBody: html
        });
        Logger.log('Email elküldve: ' + email);
      } catch (e) {
        Logger.log('Email hiba (' + email + '): ' + e.message);
      }
    }
  }
}

/**
 * Trigger beállítása a havi értesítéshez
 * Ezt egyszer kell futtatni manuálisan
 */
function setupMonthlyTrigger() {
  // Töröljük a régi triggereket
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'sendMonthlyReminder') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Új trigger: minden hónap 23-án reggel 8-kor
  ScriptApp.newTrigger('sendMonthlyReminder')
    .timeBased()
    .onMonthDay(23)
    .atHour(8)
    .create();
  
  Logger.log('Havi trigger beállítva (23-án 8:00)');
}

// ============ USERS KEZELÉS (admin) ============
function getUsers() {
  if (!isAdmin()) {
    throw new Error('Nincs jogosultságod ehhez a művelethez');
  }
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const usersSheet = ss.getSheetByName(SHEET_NAMES.USERS);
  const data = usersSheet.getDataRange().getValues();
  
  const users = [];
  for (let i = 1; i < data.length; i++) {
    users.push({
      email: data[i][0],
      name: data[i][1],
      role: data[i][2],
      notify: data[i][3]
    });
  }
  
  return users;
}

function createUser(user) {
  if (!isAdmin()) {
    throw new Error('Nincs jogosultságod ehhez a művelethez');
  }
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const usersSheet = ss.getSheetByName(SHEET_NAMES.USERS);
  
  usersSheet.appendRow([
    user.email,
    user.name,
    user.role || 'readonly',
    user.notify || false
  ]);
  
  return { success: true };
}

function updateUser(email, updates) {
  if (!isAdmin()) {
    throw new Error('Nincs jogosultságod ehhez a művelethez');
  }
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const usersSheet = ss.getSheetByName(SHEET_NAMES.USERS);
  const data = usersSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email) {
      if (updates.name !== undefined) usersSheet.getRange(i + 1, 2).setValue(updates.name);
      if (updates.role !== undefined) usersSheet.getRange(i + 1, 3).setValue(updates.role);
      if (updates.notify !== undefined) usersSheet.getRange(i + 1, 4).setValue(updates.notify);
      return { success: true };
    }
  }
  
  throw new Error('Felhasználó nem található: ' + email);
}

function deleteUser(email) {
  if (!isAdmin()) {
    throw new Error('Nincs jogosultságod ehhez a művelethez');
  }
  
  // Ne lehessen saját magát törölni
  if (email === Session.getActiveUser().getEmail()) {
    throw new Error('Saját magadat nem törölheted');
  }
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const usersSheet = ss.getSheetByName(SHEET_NAMES.USERS);
  const data = usersSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email) {
      usersSheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  
  throw new Error('Felhasználó nem található: ' + email);
}

// ============ ESZKÖZLISTA EXPORTÁLÁS ============
function getAllAssetsForDropdown() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const assetsSheet = ss.getSheetByName(SHEET_NAMES.ASSETS);
  const data = assetsSheet.getDataRange().getValues();
  
  const assets = [];
  for (let i = 1; i < data.length; i++) {
    assets.push({
      asset_id: data[i][0],
      name: data[i][1],
      type: data[i][2]
    });
  }
  
  return assets;
}

function testAssetsJson() {
  Logger.log('Test 1 - JSON string:');
  const assets1 = getAssets('{}');
  Logger.log(JSON.stringify(assets1));
  
  Logger.log('Test 2 - empty object:');
  const assets2 = getAssets({});
  Logger.log(JSON.stringify(assets2));
  
  Logger.log('Test 3 - no param:');
  const assets3 = getAssets();
  Logger.log(JSON.stringify(assets3));
}

function getAllAssets() {
  const ss          = SpreadsheetApp.openById(SPREADSHEET_ID);
  const assetsSheet = ss.getSheetByName(SHEET_NAMES.ASSETS);
  const tasksSheet  = ss.getSheetByName(SHEET_NAMES.TASKS);

  const assetsData = assetsSheet.getDataRange().getValues();
  const tasksData  = tasksSheet.getDataRange().getValues();

  // Utolsó ESD mérés dátuma eszközönként
  const lastEsdTask = {};
  for (let i = 1; i < tasksData.length; i++) {
    const assetId    = tasksData[i][1];
    const taskType   = tasksData[i][2];
    const completedAt = tasksData[i][5];
    if (taskType === 'ESD mérés' && completedAt) {
      const date = new Date(completedAt);
      if (!lastEsdTask[assetId] || date > lastEsdTask[assetId]) {
        lastEsdTask[assetId] = date;
      }
    }
  }

  const assets = [];
  const today  = new Date();
  today.setHours(0, 0, 0, 0);

  for (let i = 1; i < assetsData.length; i++) {
    const row = assetsData[i];
    if (!row[0]) continue;

    const assetId     = row[0];
    const cycleMonths = parseInt(row[4]) || 0; // 0 = nincs ESD ciklus
    const hasEsdCycle = cycleMonths > 0;

    let nextMaintenance = null;
    let status          = 'ok';
    let daysUntil       = null;

    if (hasEsdCycle) {
      const baseDate  = lastEsdTask[assetId] || (row[6] ? new Date(row[6]) : new Date());
      nextMaintenance = calcEsdNextDate(baseDate, cycleMonths);
      status          = calcEsdStatus(nextMaintenance, today);
      const diffTime  = nextMaintenance.getTime() - today.getTime();
      daysUntil       = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    }

    assets.push({
      asset_id:        assetId,
      name:            row[1] || '',
      type:            row[2] || '',
      zone:            row[3] || '',
      cycle_months:    cycleMonths,
      drive_link:      row[5] || '',
      created_at:      row[6] || '',
      next_maintenance: nextMaintenance
        ? Utilities.formatDate(nextMaintenance, Session.getScriptTimeZone(), 'yyyy.MM.dd')
        : null,
      status:          status,
      days_until:      daysUntil,
      has_esd_cycle:   hasEsdCycle
    });
  }

  return assets;
}

function testGetAllAssets() {
  const assets = getAllAssets();
  Logger.log('Assets: ' + JSON.stringify(assets));
}

// ============ OPTIMALIZÁLT BETÖLTÉS ============
function getInitialData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // Config
  const zones = getZones();
  const assetTypes = getAssetTypes();
  const taskCategories = getTaskCategories();
  
  // User
  const user = getCurrentUser();
  
  // Assets
  const assets = getAllAssets();
  
  // Stats
  const stats = {
    total:    assets.length,
    due_soon: assets.filter(a => a.has_esd_cycle && a.status === 'due_soon').length,
    overdue:  assets.filter(a => a.has_esd_cycle && a.status === 'overdue').length,
    ok:       assets.filter(a => !a.has_esd_cycle || a.status === 'ok').length,
    completed_this_month: 0
  };

  // ESD mérések ebben a hónapban
  const tasksSheet  = ss.getSheetByName(SHEET_NAMES.TASKS);
  const tasksData   = tasksSheet.getDataRange().getValues();
  const today       = new Date();
  const currentMonth = today.getMonth();
  const currentYear  = today.getFullYear();

  for (let i = 1; i < tasksData.length; i++) {
    const taskType    = tasksData[i][2];
    const completedAt = tasksData[i][5];
    if (taskType === 'ESD mérés' && completedAt) {
      const date = new Date(completedAt);
      if (date.getMonth() === currentMonth && date.getFullYear() === currentYear) {
        stats.completed_this_month++;
      }
    }
  }
  
  return {
    user: user,
    zones: zones,
    assetTypes: assetTypes,
    taskCategories: taskCategories,
    assets: assets,
    stats: stats
  };
}

function updateTask(taskId, updates) {
  if (!isAdmin()) throw new Error('Nincs jogosultságod ehhez a művelethez');

  const ss         = SpreadsheetApp.openById(SPREADSHEET_ID);
  const tasksSheet = ss.getSheetByName(SHEET_NAMES.TASKS);
  const data       = tasksSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == taskId) {
      if (updates.task_type   !== undefined) tasksSheet.getRange(i + 1, 3).setValue(updates.task_type);
      if (updates.category    !== undefined) tasksSheet.getRange(i + 1, 4).setValue(updates.category);
      if (updates.description !== undefined) tasksSheet.getRange(i + 1, 5).setValue(updates.description);
      if (updates.completed_at !== undefined) tasksSheet.getRange(i + 1, 6).setValue(new Date(updates.completed_at));
      if (updates.photo_link  !== undefined) tasksSheet.getRange(i + 1, 8).setValue(updates.photo_link);
      clearCache();
      return { success: true };
    }
  }
  throw new Error('Feladat nem található: ' + taskId);
}

function testInitialData() {
  const data = getInitialData();
  Logger.log('Data: ' + JSON.stringify(data));
}

// ============ CACHE ============
const CACHE_KEY = 'initialData';
const CACHE_DURATION = 21600; // 6 óra másodpercben

function getInitialDataJson() {
  const cache = CacheService.getScriptCache();
  let cached = cache.get(CACHE_KEY);
  
  if (cached) {
    Logger.log('Returning cached data');
    return cached;
  }
  
  Logger.log('Building fresh data');
  const data = getInitialData();
  const json = JSON.stringify(data);
  
  try {
    cache.put(CACHE_KEY, json, CACHE_DURATION);
  } catch (e) {
    Logger.log('Cache error: ' + e.message);
  }
  
  return json;
}

function clearCache() {
  const cache = CacheService.getScriptCache();
  cache.remove(CACHE_KEY);
  Logger.log('Cache cleared');
}

function getTasksByAssetJson(assetId) {
  const tasks = getTasksByAsset(assetId);
  return JSON.stringify(tasks);
}

function getCalendarDataJson(year, month) {
  const data = getCalendarData(year, month);
  return JSON.stringify(data);
}
