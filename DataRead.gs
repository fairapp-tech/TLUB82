/**
 * Plik: DataRead.gs
 * Wersja: TURBO MODE V24 (External Vacation Source)
 * Zmiany:
 * - getVacations: Przekierowanie do DataVacations.gs (getVacationsExternal).
 */

const PLANNING_CACHE_VERSION = "6.6-Date-Warning-Fix"; 

// === POMOCNICZE ===
function _formatDate(dateObj, format = "dd.MM.yyyy HH:mm") {
  if (!dateObj) return '';
  if (typeof dateObj === 'string') return dateObj.replace(/'/g, '');
  try {
    return Utilities.formatDate(new Date(dateObj), Session.getScriptTimeZone(), format);
  } catch (e) { return ''; }
}

function _getStartOfDay(date) {
    const d = new Date(date);
    d.setHours(0,0,0,0);
    return d;
}

// === HELPER PARSOWANIA DAT ===
function _parseAnyDate(cell) {
    if (cell instanceof Date) return cell;
    if (!cell || typeof cell !== 'string') return null;
    const c = cell.trim();
    if (!c) return null;

    // 1. Format ISO: YYYY-MM-DD
    if (/^\d{4}-\d{2}-\d{2}$/.test(c)) {
        return new Date(c);
    }
    
    // 2. Format Polski: DD.MM.YYYY
    if (/^\d{1,2}\.\d{1,2}\.\d{4}$/.test(c)) {
        const parts = c.split('.');
        return new Date(parts[2], parts[1] - 1, parts[0]);
    }

    // 3. Format Polski Tekstowy: "20 maj", "1 gru"
    return parsePolishDate(c);
}

// === GŁÓWNA FUNKCJA POBIERAJĄCA ===
function getInitialData() {
  const timer = new Date();
  console.log("🚀 Start getInitialData (Turbo V24)");
  
  const response = {
    employees: [],
    employeeMeta: { source: 'NONE', date: null, timestamp: null },
    tasks: [],
    currentUserEmail: Session.getActiveUser().getEmail(),
    widgets: {
      stow: { today: [], recent: [], recentDateLabel: '' },
      carts: { today: [], recent: [], recentDateLabel: '' },
      missions: { today: [], recent: [], recentDateLabel: '' },
      support: { today: [], recent: [], recentDateLabel: '' },
      beauty: { today: [], recent: [], recentDateLabel: '' }
    },
    settings: null,
    planningData: [],
    groupStats: null
  };

  try {
    const ss = SpreadsheetApp.openById(MAIN_SHEET_ID);

    // 1. Cache & Settings
    const cacheData = _fetchAllCache(ss); 
    response.settings = _parseSettingsFromCache(cacheData.values) || _getDefaultSettings();
    
    // 2. Pracownicy
    const empResult = _fetchEmployeesWithFallback(ss, cacheData);
    response.employees = empResult.list;
    response.employeeMeta = empResult.meta;

    // 3. Rotacja
    const allRotation = _fetchRotationalDB(ss); 
    response.widgets = _processWidgetsFromDB(allRotation);

    // 4. Zadania
    try { response.tasks = _fetchTasks(ss); } catch(e) { console.error(e); }

    // 5. Planowanie
    const now = new Date();
    const h = now.getHours();
    const m = now.getMinutes();
    
    const isWindow1 = (h === 13 && m >= 30) || (h === 14 && m <= 50);
    const isWindow2 = (h === 21 && m >= 30 && m <= 50);

    if (isWindow1 || isWindow2) {
        try {
            console.log("🕒 Okno czasowe aktywne - pobieranie świeżego planu...");
            response.planningData = getPlanningData(0, true, ss) || [];
        } catch(e) {
            console.error("Błąd pobierania planu w oknie czasowym:", e);
            response.planningData = _parsePlanningFromCache(cacheData.values, 0) || []; 
        }
    } else {
        response.planningData = _parsePlanningFromCache(cacheData.values, 0) || []; 
    }

    // 6. Statystyki Grupy
    try {
        response.groupStats = getGroupData(response.employees, response.settings, ss);
    } catch (e) {
        console.error("Błąd GroupStats:", e);
        response.groupStats = { total: 0, present: 0, trained: 0, availableStow: 0, plan: [], subgroups: { A: {total:0, present:0}, B: {total:0, present:0} } };
    }

    console.log(`🏁 Koniec getInitialData. Źródło: ${response.employeeMeta.source}. Czas: ${new Date() - timer}ms`);
    return response;

  } catch (e) {
    console.error("CRITICAL ERROR:", e);
    throw new Error("Błąd ładowania danych: " + e.message);
  }
}

// =========================================================
// === SILNIKI ODCZYTU (DB & CACHE) ===
// =========================================================

function _fetchAllCache(ss) {
    const result = { values: {}, timestamps: {} };
    const sheet = ss.getSheetByName(globalCacheSheetName);
    if (!sheet) return result;
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return result;

    const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    for (let i = 0; i < data.length; i++) {
        const key = String(data[i][0]);
        const val = data[i][1];
        const ts = data[i][2];
        
        if (key) {
            result.values[key] = val;
            result.timestamps[key] = ts;
        }
    }
    return result;
}

function _parseSettingsFromCache(cacheValues) {
    if (cacheValues['SETTINGS']) {
        try { return JSON.parse(cacheValues['SETTINGS']); } catch(e) {}
    }
    return null;
}

function _parsePlanningFromCache(cacheValues, offset) {
    const todayStr = Utilities.formatDate(new Date(new Date().getTime() + offset * 86400000), Session.getScriptTimeZone(), "dd.MM.yyyy");
    const key = `PLAN_${todayStr}`;
    if (cacheValues[key]) {
        try { return JSON.parse(cacheValues[key]); } catch(e) {}
    }
    return null;
}

function _getDefaultSettings() {
    return { 
        defaultMode: 'tasks', userGroup: '', 
        customRoles: { support: ['Unloading', 'Receive', 'Pack', 'Shipping', 'Zwroty', 'Jakość'], missions: ['Kartony', 'X-info', 'Sprzątanie'] } 
    };
}

function _fetchRotationalDB(ss) {
    const sheet = ss.getSheetByName(dbRotacjaSheetName);
    if (!sheet) return [];
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    
    return data.map(row => ({
        date: row[0],
        mode: row[1],
        name: String(row[2]),
        p1: row[3],
        p2: row[4],
        timestamp: row[5]
    }));
}

// === ZMODYFIKOWANA FUNKCJA URLOPÓW (EXTERNAL LINK) ===
function getVacations(year) {
    // Delegacja do DataVacations.gs
    return getVacationsExternal(year);
}

function _processWidgetsFromDB(rows) {
    const widgets = {
        stow: { today: [], recent: [], recentDateLabel: '' },
        carts: { today: [], recent: [], recentDateLabel: '' },
        missions: { today: [], recent: [], recentDateLabel: '' },
        support: { today: [], recent: [], recentDateLabel: '' },
        beauty: { today: [], recent: [], recentDateLabel: '' }
    };

    const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    rows.forEach(row => {
        if (!row.date || !row.mode || !widgets[row.mode]) return;
        let dStr = "";
        try { dStr = Utilities.formatDate(new Date(row.date), Session.getScriptTimeZone(), "yyyy-MM-dd"); } catch(e){ return; }
        if (dStr === todayStr) {
            widgets[row.mode].today.push({ name: row.name, break: row.p1, floor: row.p2 });
        }
    });

    for (let i = rows.length - 1; i >= 0; i--) {
        const row = rows[i];
        if (!row.date || !row.mode || !widgets[row.mode]) continue;
        let dStr = "";
        try { dStr = Utilities.formatDate(new Date(row.date), Session.getScriptTimeZone(), "yyyy-MM-dd"); } catch(e){ continue; }
        
        if (dStr === todayStr) continue;

        if (!widgets[row.mode].recentDateLabel) {
            widgets[row.mode].recentDateLabel = Utilities.formatDate(new Date(row.date), Session.getScriptTimeZone(), "dd.MM");
            widgets[row.mode]._internalRecentDate = dStr; 
        }
        
        if (widgets[row.mode]._internalRecentDate === dStr) {
            const exists = widgets[row.mode].recent.some(r => r.name === row.name);
            if (!exists) {
                widgets[row.mode].recent.push({ name: row.name, break: row.p1, floor: row.p2 });
            }
        }
    }
    return widgets;
}

function _fetchEmployeesWithFallback(ss, cacheData) {
    let liveList = [];
    let isLiveValid = false;
    
    try {
        const sheet = ss.getSheetByName(employeesSheetName);
        if (sheet) {
            const lastRow = sheet.getLastRow();
            if (lastRow >= 4) {
                const data = sheet.getRange(4, 7, lastRow - 3, 20).getValues();
                for (let i = 0; i < data.length; i++) {
                    const name = String(data[i][0]).trim();
                    if (!name) continue;
                    const statusRaw = data[i][1];
                    const isPresent = (statusRaw === 8 || statusRaw === '8' || statusRaw === 7 || statusRaw === '7');
                    const subgroup = data[i][19] ? String(data[i][19]).trim().toUpperCase() : '';
                    const skills = {
                        unloading: data[i][4] == 1, receive: data[i][6] == 1, stow: data[i][7] == 1, sort: data[i][9] == 1,
                        pack: data[i][10] == 1, carts: data[i][11] == 1, shipping: data[i][12] == 1, returns: data[i][14] == 1,
                        quality: data[i][16] == 1, beauty: data[i][18] == 1
                    };
                    liveList.push({ name: name, isAtWork: isPresent, hasStowTraining: skills.stow, skills: skills, subgroup: subgroup });
                }
            }
        }
        const presentCount = liveList.filter(e => e.isAtWork).length;
        if (presentCount >= 2) isLiveValid = true;
    } catch(e) { console.error("Live Fetch Error", e); }

    const todayTs = new Date();

    if (isLiveValid) {
        try { if (typeof saveAttendanceToCache === 'function') saveAttendanceToCache(liveList); } catch(e) {}
        return { list: liveList, meta: { source: 'LIVE', date: _formatDate(todayTs, "dd.MM.yyyy"), timestamp: todayTs.getTime() } };
    } else {
        const cacheKey = 'ATTENDANCE_LATEST';
        if (cacheData.values[cacheKey]) {
            try {
                const cachedList = JSON.parse(cacheData.values[cacheKey]);
                const cachedTs = cacheData.timestamps[cacheKey] ? new Date(cacheData.timestamps[cacheKey]) : null;
                const cachedDateStr = cachedTs ? _formatDate(cachedTs, "dd.MM.yyyy") : "Nieznana";
                return { list: cachedList, meta: { source: 'CACHE', date: cachedDateStr, timestamp: cachedTs ? cachedTs.getTime() : 0 } };
            } catch(e) {}
        }
        return { list: liveList, meta: { source: 'NONE', date: null, timestamp: 0 } }; 
    }
}

function _fetchTasks(ss) {
    const sheet = ss.getSheetByName(tasksSheetName);
    if (!sheet) return [];
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    const data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
    return data.map(row => {
        let replies = []; try { replies = row[10] ? JSON.parse(row[10]) : []; } catch(e) {}
        return {
            employee: String(row[0]), description: String(row[1]),
            creationTime: _formatDate(row[2]), deadline: _formatDate(row[3], "dd.MM.yyyy"),
            status: row[4], taskId: row[5], taskGroupId: row[6], category: row[7],
            priority: row[8], creator: row[9], replies: replies
        };
    }).filter(t => t.taskId);
}

// =========================================================
// === FUNKCJE PUBLICZNE ===
// =========================================================

function getTasks() { const ss = SpreadsheetApp.openById(MAIN_SHEET_ID); return _fetchTasks(ss); }
function getAppSettings() { const ss = SpreadsheetApp.openById(MAIN_SHEET_ID); const cd = _fetchAllCache(ss); return _parseSettingsFromCache(cd.values) || _getDefaultSettings(); }

function getRotationalData(mode) {
    const ss = SpreadsheetApp.openById(MAIN_SHEET_ID);
    const cacheData = _fetchAllCache(ss);
    const allRows = _fetchRotationalDB(ss);
    const modeRows = allRows.filter(r => r.mode === mode);
    const stats = {};
    const todayTime = _getStartOfDay(new Date()).getTime();
    const today = new Date();
    const day = today.getDay();
    const diff = today.getDate() - day + (day == 0 ? -6:1);
    const startOfWeek = new Date(today.setDate(diff)).setHours(0,0,0,0);
    const startOfMonth = new Date(new Date().getFullYear(), new Date().getMonth(), 1).getTime();
    const startOfYear = new Date(new Date().getFullYear(), 0, 1).getTime();
    
    modeRows.forEach(row => {
        const d = new Date(row.date); d.setHours(0,0,0,0);
        const t = d.getTime();
        const n = row.name;
        if (!stats[n]) stats[n] = { week: 0, month: 0, year: 0, lastDate: 0 };
        if (t >= startOfWeek) stats[n].week++;
        if (t >= startOfMonth) stats[n].month++;
        if (t >= startOfYear) stats[n].year++;
        if (t > stats[n].lastDate) stats[n].lastDate = t;
    });

    const empResult = _fetchEmployeesWithFallback(ss, cacheData);
    const employees = empResult.list;
    const exclKey = `EXCLUSIONS_${mode}`;
    const excludedSet = new Set();
    if (cacheData.values[exclKey]) {
        try { const list = JSON.parse(cacheData.values[exclKey]); if (Array.isArray(list)) list.forEach(n => excludedSet.add(String(n))); } catch(e) {}
    }
    
    return employees.filter(emp => {
        let trainingCondition = true;
        if (mode === 'stow') trainingCondition = emp.hasStowTraining;
        return trainingCondition;
    }).map(emp => {
        const s = stats[emp.name] || { week: 0, month: 0, year: 0, lastDate: 0 };
        return {
            name: emp.name,
            weekCount: s.week,
            monthCount: s.month,
            yearCount: s.year,
            lastDate: s.lastDate,
            wasToday: s.lastDate === todayTime,
            isExcluded: excludedSet.has(emp.name),
            skills: emp.skills,
            subgroup: emp.subgroup,
            isAtWork: emp.isAtWork
        };
    }).sort((a, b) => a.lastDate - b.lastDate);
}

function getWidgetData(mode) { const ss = SpreadsheetApp.openById(MAIN_SHEET_ID); const ar = _fetchRotationalDB(ss); const pr = _processWidgetsFromDB(ar); return pr[mode] || { today: [], recent: [], recentDateLabel: '' }; }
function getEmployeeYearlyHistory(n, m) { const ss = SpreadsheetApp.openById(MAIN_SHEET_ID); return _fetchRotationalDB(ss).filter(r => r.mode === m && r.name === n).map(r => ({ date: _formatDate(r.date, "dd.MM.yyyy"), breakVal: r.p1, floor: r.p2 })).reverse(); }

function getPlanningData(dayOffset = 0, forceUpdate = false, existingSS = null) {
    try {
        const ss = existingSS || SpreadsheetApp.openById(MAIN_SHEET_ID);
        const targetDate = new Date();
        targetDate.setDate(targetDate.getDate() + dayOffset);
        const targetDateStr = Utilities.formatDate(targetDate, Session.getScriptTimeZone(), "dd.MM.yyyy");
        const cacheKey = `PLAN_${targetDateStr}`;
        if (!forceUpdate) {
            const cacheData = _fetchAllCache(ss);
            if (cacheData.values[cacheKey]) return JSON.parse(cacheData.values[cacheKey]);
            return ["⏳ Dane ładują się automatycznie:", "13:30 - 13:50 (Poranna)", "21:30 - 21:50 (Popołudniowa)"];
        }
        const planningSheet = ss.getSheetByName(planningSheetName);
        if (!planningSheet) return [];
        const dateCell = planningSheet.getRange("G2").getValue();
        let sheetDateStr = "";
        if (dateCell) sheetDateStr = (Object.prototype.toString.call(dateCell) === '[object Date]') ? Utilities.formatDate(dateCell, Session.getScriptTimeZone(), "dd.MM.yyyy") : String(dateCell);
        
        const resultList = [];
        if (sheetDateStr !== targetDateStr) {
             resultList.push(`⚠️ Data w planie (${sheetDateStr}) nie pasuje do wybranego dnia (${targetDateStr}).`);
        }
        
        const rangeCage = planningSheet.getRange("H3:H5").getValues().flat();
        const stowList = rangeCage.map(String).map(n => n.trim()).filter(n => n !== "").map(n => n + " (STOW)");
        const rangeMain = planningSheet.getRange("F3:F29").getValues().flat();
        const mainList = rangeMain.map(String).map(n => n.trim()).filter(n => n !== "");
        const sheetNames = [...stowList, ...mainList];
        
        sheetNames.forEach(n => resultList.push(n));

        try { if (typeof savePlanningToCache === 'function') savePlanningToCache(targetDateStr, resultList); } catch(e) {}
        return resultList;
    } catch (e) { return ["⚠️ Błąd systemu."]; }
}

function getGroupData(employeesList, settings, ss) {
    if (!ss) ss = SpreadsheetApp.openById(MAIN_SHEET_ID);
    
    let subStats = { A: { total: 0, present: 0 }, B: { total: 0, present: 0 } };
    employeesList.forEach(e => {
        if (e.subgroup === 'A') { subStats.A.total++; if (e.isAtWork) subStats.A.present++; }
        else if (e.subgroup === 'B') { subStats.B.total++; if (e.isAtWork) subStats.B.present++; }
    });
    
    const total = employeesList.length;
    const present = employeesList.filter(e => e.isAtWork).length;
    const trained = employeesList.filter(e => e.hasStowTraining).length;
    const availableStow = employeesList.filter(e => e.isAtWork && e.hasStowTraining).length;

    const now = new Date();
    const currentWeek = getWeekNumber(now);
    const day = now.getDay(); 
    let currentShift = "";

    if (settings && settings.userGroup) {
        const grafikShift = getShiftForGroup(settings.userGroup, ss);
        if (grafikShift) currentShift = `${grafikShift} (${settings.userGroup})`;
    }
    if (!currentShift) {
        if (currentWeek % 2 === 0) currentShift = (day === 0) ? "Popołudniowa" : "Poranna";
        else currentShift = "Popołudniowa";
    }

    const sheet = ss.getSheetByName(planningSheetName);
    const plan = [];

    if (sheet) {
        const rangeValues = sheet.getRange("J1:S50").getValues();
        
        const dayRanges = [
            { label: "Poniedziałek", start: 3, end: 7 },
            { label: "Wtorek",  start: 10, end: 14 },
            { label: "Środa",  start: 17, end: 21 },
            { label: "Czwartek", start: 24, end: 28 },
            { label: "Piątek",  start: 31, end: 35 },
            { label: "Sobota", start: 38, end: 42 },
            { label: "Niedziela",  start: 45, end: 49 }
        ];

        const targetPhrases = new Set();
        targetPhrases.add("UB8");
        targetPhrases.add("MATEUSZ 634"); 

        employeesList.forEach(e => {
            if (e.name) targetPhrases.add(e.name.trim().toUpperCase());
        });
        
        const targetList = Array.from(targetPhrases);

        dayRanges.forEach(d => {
            const locations = new Set();
            for (let r = d.start; r <= d.end; r++) {
                const rowIndex = r - 1; 
                if (rowIndex >= rangeValues.length) break;
                const row = rangeValues[rowIndex];
                const leftTexts = [String(row[1]), String(row[2]), String(row[3])];
                for (const text of leftTexts) {
                    const cleanText = text.trim().toUpperCase();
                    if (cleanText) {
                        const found = targetList.some(target => cleanText.includes(target));
                        if (found) {
                            const floor = String(row[0]).trim(); 
                            if (floor) locations.add(`${floor} (J)`);
                        }
                    }
                }
                const rightTexts = [String(row[7]), String(row[8]), String(row[9])];
                for (const text of rightTexts) {
                    const cleanText = text.trim().toUpperCase();
                    if (cleanText) {
                        const found = targetList.some(target => cleanText.includes(target));
                        if (found) {
                            const floor = String(row[6]).trim(); 
                            if (floor) locations.add(`${floor} (P)`);
                        }
                    }
                }
            }
            if (locations.size > 0) {
                const locArray = Array.from(locations).sort();
                plan.push({ day: d.label, info: locArray.join(', ') });
            }
        });
    }

    return { total, present, trained, availableStow, plan, currentWeek, currentShift, subgroups: subStats };
}

function getShiftForGroup(group, ss) {
    if(!ss) ss = SpreadsheetApp.openById(MAIN_SHEET_ID);
    const sheet = ss.getSheetByName(grafikSheetName); 
    if (!sheet) return null;
    const structure = _detectScheduleStructure(sheet);
    if (!structure) return null;
    const groupOffset = { 'A':0, 'B':1, 'C':2, 'D':3, 'E':4, 'F':5, 'G':6, 'H':7 };
    const targetGroupIdx = structure.groupStartIdx + groupOffset[group];
    const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    const lastRow = sheet.getLastRow();
    const data = sheet.getRange(structure.firstDataRow, structure.dateColIdx+1, lastRow - structure.firstDataRow + 1, 20).getValues();
    const relDate = 0; const relShift = structure.shiftColIdx - structure.dateColIdx; const relGroup = targetGroupIdx - structure.dateColIdx;
    
    let lastDateStr = null;
    
    for (let i = 0; i < data.length; i++) {
        const cell = data[i][relDate];
        const parsed = _parseAnyDate(cell); 
        
        if (parsed) {
             lastDateStr = Utilities.formatDate(parsed, Session.getScriptTimeZone(), "yyyy-MM-dd");
        }
        
        if (lastDateStr === todayStr) {
            const rawVal = String(data[i][relGroup]).trim().toUpperCase();
            if (rawVal && rawVal !== "") {
                 let shiftVal = String(data[i][relShift]).trim();
                 if(!shiftVal) shiftVal = "1";
                 return "Zmiana " + shiftVal; 
            }
            const nextRowDateCell = (i+1 < data.length) ? data[i+1][relDate] : "END";
            const nextParsed = _parseAnyDate(nextRowDateCell);
            if (nextParsed || i+1 >= data.length) return "Wolne";
        }
    }
    return null;
}

function getMonthlySchedule(group, monthOffset = 0) {
    const ss = SpreadsheetApp.openById(MAIN_SHEET_ID);
    const sheet = ss.getSheetByName(grafikSheetName); 
    if (!sheet) return { error: "Brak arkusza Grafik" };
    const structure = _detectScheduleStructure(sheet);
    if (!structure) return { error: "Nie wykryto struktury." };
    
    const groupOffset = { 'A':0, 'B':1, 'C':2, 'D':3, 'E':4, 'F':5, 'G':6, 'H':7 };
    const targetGroupIdx = structure.groupStartIdx + groupOffset[group];
    
    const targetDate = new Date(); 
    targetDate.setDate(1); 
    targetDate.setMonth(targetDate.getMonth() + monthOffset);
    
    const targetMonth = targetDate.getMonth(); 
    const targetYear = targetDate.getFullYear();
    const monthLabel = Utilities.formatDate(targetDate, Session.getScriptTimeZone(), "MMMM yyyy");
    
    const daysInMonth = new Date(targetYear, targetMonth + 1, 0).getDate();
    const daysMap = {}; 
    for (let d = 1; d <= daysInMonth; d++) daysMap[d] = "Wolne";
    
    const lastRow = sheet.getLastRow();
    const data = sheet.getRange(structure.firstDataRow, structure.dateColIdx+1, lastRow - structure.firstDataRow + 1, 20).getValues();
    const relDate = 0; 
    const relShift = structure.shiftColIdx - structure.dateColIdx; 
    const relGroup = targetGroupIdx - structure.dateColIdx;
    
    let lastValid = null;
    
    for(let i=0; i<data.length; i++) {
        const cell = data[i][relDate];
        const parsed = _parseAnyDate(cell); 
        
        if(parsed) {
            if (parsed.getFullYear() === new Date().getFullYear() && parsed.getMonth() === targetMonth && targetYear !== parsed.getFullYear()) {
                 parsed.setFullYear(targetYear);
            }
            lastValid = parsed;
        }
        
        if(!lastValid || lastValid.getMonth() !== targetMonth || lastValid.getFullYear() !== targetYear) continue;
        
        const val = String(data[i][relGroup]).trim();
        if(val && val !== "") { 
            let shiftVal = String(data[i][relShift]).trim();
            if(!shiftVal) shiftVal = "1";
            daysMap[lastValid.getDate()] = "Zmiana " + shiftVal;
        }
    }
    
    const finalDays = []; 
    for(let d=1; d<=daysInMonth; d++) finalDays.push({day:d, shift:daysMap[d]});
    
    const startDayOfWeek = (new Date(targetYear, targetMonth, 1).getDay() + 6) % 7; 
    
    return { monthLabel, days: finalDays, startDayOfWeek };
}

function parsePolishDate(dateStr) {
    if (!dateStr || typeof dateStr !== 'string') return null;
    const months = {'sty':0,'lut':1,'mar':2,'kwi':3,'maj':4,'cze':5,'lip':6,'sie':7,'wrz':8,'paź':9,'paz':9,'lis':10,'gru':11};
    const match = dateStr.trim().toLowerCase().match(/^(\d{1,2})\s*([a-ząśżźćńółę]{3,})/);
    if(match && months[match[2].substring(0,3)] !== undefined) {
        return new Date(new Date().getFullYear(), months[match[2].substring(0,3)], parseInt(match[1]));
    }
    return null;
}

function _detectScheduleStructure(sheet) {
    const range = sheet.getRange(1, 1, 100, 20);
    const values = range.getValues();
    for (let r = 0; r < values.length; r++) {
        for (let c = 0; c < values[r].length; c++) {
            const cell = values[r][c];
            if (_parseAnyDate(cell)) {
                 if (c+2 < values[r].length) {
                     const s = values[r][c+2];
                     if([1,2,3,'1','2','3'].includes(s)) return { firstDataRow: r+1, dateColIdx: c, shiftColIdx: c+2, groupStartIdx: c+3 };
                 }
            }
        }
    }
    return null;
}

function triggerAutoUpdate() {
   const now = new Date();
   const hour = now.getHours(); const minute = now.getMinutes();
   const isWindow1 = (hour === 13 && minute >= 30) || (hour === 14 && minute <= 50);
   const isWindow2 = (hour === 21 && minute >= 30 && minute <= 50);
   if (isWindow1 || isWindow2) {
       try { getPlanningData(1, true); } catch(e) {}
   }
}

function getWeekNumber(d) {
    d = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
    d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay()||7));
    var yearStart = new Date(Date.UTC(d.getUTCFullYear(),0,1));
    return Math.ceil((((d-yearStart)/86400000)+1)/7);
}