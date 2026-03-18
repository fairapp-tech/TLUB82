/**
 * Plik: DataWrite.gs
 * Wersja: TURBO MODE V44 (External Vacation Write)
 * Zmiany:
 * - addVacationRequest, updateVacationStatus, deleteVacationRequest: Przekierowanie do DataVacations.gs.
 * - approveVacationSplit: Implementacja dla zewnętrznego arkusza (zapis 'U' w zakresie komórek).
 */

// === FUNKCJE POMOCNICZE ===

function _findRowIndexById(sheet, taskId) {
  const ids = sheet.getRange("F2:F" + sheet.getLastRow()).getValues().flat();
  const index = ids.indexOf(taskId);
  return index !== -1 ? index + 2 : -1;
}

// === UNIWERSALNY ZAPIS DO DB_CACHE ===
function _saveToGlobalCache(key, data) {
    const lock = LockService.getScriptLock();
    lock.waitLock(5000);
    
    try {
        const ss = SpreadsheetApp.openById(MAIN_SHEET_ID);
        let sheet = ss.getSheetByName(globalCacheSheetName); // Zdefiniowane w Config.gs ('DB_Cache')
        
        if (!sheet) {
            sheet = ss.insertSheet(globalCacheSheetName);
            sheet.appendRow(["Key", "Value", "Timestamp"]);
            sheet.setFrozenRows(1);
        }
        
        const json = JSON.stringify(data);
        const timestamp = new Date();
        
        const lastRow = sheet.getLastRow();
        let foundRow = -1;
        
        if (lastRow > 1) {
            const keys = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
            const keyIndex = keys.indexOf(key);
            if (keyIndex !== -1) {
                foundRow = keyIndex + 2; 
            }
        }
        
        if (foundRow > 0) {
            sheet.getRange(foundRow, 2).setValue(json);
            sheet.getRange(foundRow, 3).setValue(timestamp);
        } else {
            sheet.appendRow([key, json, timestamp]);
        }
        
        return true;
    } catch(e) {
        console.error(`Błąd zapisu do Cache (${key}):`, e);
        throw e;
    } finally {
        lock.releaseLock();
    }
}

// === HELPER ODCZYTU Z CACHE ===
function _readFromGlobalCache(key) {
    const ss = SpreadsheetApp.openById(MAIN_SHEET_ID);
    const sheet = ss.getSheetByName(globalCacheSheetName);
    if (!sheet) return null;
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return null;
    
    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for(let i=0; i<data.length; i++) {
        if (data[i][0] === key) {
            try { return JSON.parse(data[i][1]); } catch(e) { return null; }
        }
    }
    return null;
}

// === ZAPIS USTAWIEŃ ===
function saveAppSettings(settings) {
    if (!settings) throw new Error("Brak danych ustawień.");
    try {
        _saveToGlobalCache('SETTINGS', settings);
        return { success: true, message: "Ustawienia zapisane." };
    } catch (e) {
        throw new Error("Nie udało się zapisać ustawień: " + e.message);
    }
}

// === ZAPIS PLANOWANIA ===
function savePlanningToCache(dateStr, data) {
    const key = `PLAN_${dateStr}`;
    try {
        _saveToGlobalCache(key, data);
        console.log(`Zapisano plan do cache: ${key}`);
    } catch (e) {
        console.error("Błąd zapisu planu do cache", e);
    }
}

// === ZAPIS OBECNOŚCI ===
function saveAttendanceToCache(data) {
    try {
        _saveToGlobalCache('ATTENDANCE_LATEST', data);
    } catch(e) { console.error(e); }
}

// === IMPORT GRAFIKU ===
function importScheduleData(rawData) {
    if (!rawData || rawData.trim() === "") {
        throw new Error("Brak danych do importu.");
    }
    const lock = LockService.getScriptLock();
    lock.waitLock(30000); 
    try {
        const ss = SpreadsheetApp.openById(MAIN_SHEET_ID);
        let sheet = ss.getSheetByName(grafikSheetName);
        if (!sheet) { sheet = ss.insertSheet(grafikSheetName); } else { sheet.clear(); }

        let separator = '\t';
        if (rawData.indexOf('\t') === -1 && rawData.indexOf(';') !== -1) separator = ';';

        const rows = rawData.split('\n');
        const parsedData = [];
        let maxCols = 0;

        rows.forEach(rowStr => {
            if (rowStr.trim() === "") return;
            const cells = rowStr.split(separator);
            const cleanCells = cells.map(c => c.replace(/^"|"$/g, '').trim());
            if (cleanCells.length > maxCols) maxCols = cleanCells.length;
            parsedData.push(cleanCells);
        });
        const finalData = parsedData.map(row => { while (row.length < maxCols) row.push(""); return row; });
        if (finalData.length === 0) throw new Error("Nie rozpoznano danych.");
        sheet.getRange(1, 1, finalData.length, maxCols).setValues(finalData);
        try { sheet.getRange(1, 3, finalData.length, 1).setNumberFormat("yyyy-MM-dd"); } catch(e) {}
        return { success: true, message: `Zaimportowano ${finalData.length} wierszy grafiku.` };
    } catch (e) { throw new Error("Import nieudany: " + e.message); } finally { lock.releaseLock(); }
}

// === ZADANIA ===
function addTask(taskData) {
  if (!taskData) throw new Error("Brak danych.");
  try {
    const sheet = SpreadsheetApp.openById(tasksSheetId).getSheetByName(tasksSheetName);
    const taskGroupId = Utilities.getUuid();
    const creationTime = new Date();
    const creatorEmail = Session.getActiveUser().getEmail();
    const formattedCreationTime = "'" + Utilities.formatDate(creationTime, Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm");
    const formattedDeadline = taskData.deadline ? "'" + taskData.deadline : '';
    const newRows = [];
    const newTasksData = [];
    taskData.employees.forEach(employeeName => {
      const taskId = Utilities.getUuid();
      newRows.push([ employeeName, taskData.description, formattedCreationTime, formattedDeadline, "Do realizacji", taskId, taskGroupId, taskData.category, taskData.priority, creatorEmail, '[]' ]);
      newTasksData.push({ employee: employeeName, description: taskData.description, creationTime: formattedCreationTime.replace(/'/g, ''), deadline: formattedDeadline.replace(/'/g, ''), status: "Do realizacji", taskId: taskId, taskGroupId: taskGroupId, category: taskData.category, priority: taskData.priority, creator: creatorEmail, replies: [] });
    });
    if (newRows.length > 0) sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
    return newTasksData;
  } catch (e) { throw new Error("Nie udało się dodać zadania: " + e.message); }
}

function addReply(replyData) {
    const { taskId, replyText } = replyData;
    const lock = LockService.getScriptLock(); lock.waitLock(15000);
    try {
        const sheet = SpreadsheetApp.openById(tasksSheetId).getSheetByName(tasksSheetName);
        const rowIndex = _findRowIndexById(sheet, taskId);
        if (rowIndex === -1) throw new Error("Nie znaleziono zadania.");
        const repliesCell = sheet.getRange(rowIndex, 11);
        const currentRepliesJson = repliesCell.getValue();
        let replies = [];
        if (currentRepliesJson) { try { replies = JSON.parse(currentRepliesJson); } catch (e) {} }
        const newReply = { creator: Session.getActiveUser().getEmail(), text: replyText, timestamp: new Date().toISOString() };
        replies.push(newReply);
        repliesCell.setValue(JSON.stringify(replies));
        SpreadsheetApp.flush();
        return newReply;
    } catch(e) { throw new Error("Błąd odpowiedzi."); } finally { lock.releaseLock(); }
}

function updateTask(updateData) {
    if (!updateData || !updateData.taskId) throw new Error("Brak ID.");
    try {
        const sheet = SpreadsheetApp.openById(tasksSheetId).getSheetByName(tasksSheetName);
        const rowIndex = _findRowIndexById(sheet, updateData.taskId);
        if (rowIndex === -1) throw new Error("Nie znaleziono zadania.");
        sheet.getRange(rowIndex, 2).setValue(updateData.description);
        sheet.getRange(rowIndex, 4).setValue(updateData.deadline ? "'" + updateData.deadline : '');
        sheet.getRange(rowIndex, 9).setValue(updateData.priority);
        const updatedRow = sheet.getRange(rowIndex, 1, 1, 11).getValues()[0];
        let replies = []; try { replies = JSON.parse(updatedRow[10]); } catch(e){}
        return { employee: updatedRow[0], description: updatedRow[1], creationTime: updatedRow[2].toString().replace(/'/g, ''), deadline: updatedRow[3].toString().replace(/'/g, ''), status: updatedRow[4], taskId: updatedRow[5], taskGroupId: updatedRow[6], category: updatedRow[7], priority: updatedRow[8], creator: updatedRow[9], replies: replies };
    } catch (e) { throw new Error("Błąd edycji."); }
}
function updateTaskStatuses(taskIds, newStatus) {
    try { const sheet = SpreadsheetApp.openById(tasksSheetId).getSheetByName(tasksSheetName); const ids = sheet.getRange("F2:F" + sheet.getLastRow()).getValues().flat(); taskIds.forEach(taskId => { const index = ids.indexOf(taskId); if (index !== -1) sheet.getRange(index + 2, 5).setValue(newStatus); }); } catch (e) { throw new Error("Błąd statusu."); }
}
function deleteTasksByIds(taskIds) {
    try { const sheet = SpreadsheetApp.openById(tasksSheetId).getSheetByName(tasksSheetName); const idsInSheet = sheet.getRange("F2:F" + sheet.getLastRow()).getValues().flat(); const idsToDelete = new Set(taskIds); let rowsToDelete = []; idsInSheet.forEach((id, index) => { if (idsToDelete.has(id)) rowsToDelete.push(index + 2); }); if (rowsToDelete.length > 0) rowsToDelete.reverse().forEach(rowIndex => sheet.deleteRow(rowIndex)); } catch (e) { throw new Error("Błąd usuwania."); }
}
function deleteReply(data) {
   const { taskId, timestamp } = data; const lock = LockService.getScriptLock(); lock.waitLock(5000); try { const sheet = SpreadsheetApp.openById(tasksSheetId).getSheetByName(tasksSheetName); const rowIndex = _findRowIndexById(sheet, taskId); if(rowIndex === -1) return; const cell = sheet.getRange(rowIndex, 11); const json = cell.getValue(); let replies = []; try { replies = JSON.parse(json); } catch(e){} const newReplies = replies.filter(r => r.timestamp !== timestamp); cell.setValue(JSON.stringify(newReplies)); } finally { lock.releaseLock(); }
}

// === ROTACJA TURBO ===
function processRotationalAssignments(mode, assignments) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.openById(MAIN_SHEET_ID);
    let sheet = ss.getSheetByName(dbRotacjaSheetName);
    
    if (!sheet) {
      sheet = ss.insertSheet(dbRotacjaSheetName);
      sheet.appendRow(["Data", "Typ", "Pracownik", "Param1", "Param2", "Timestamp"]);
      sheet.setFrozenRows(1);
    }
    
    const today = new Date();
    const newRows = assignments.map(a => [today, mode, a.name, a.break, a.floor, new Date()]);
    
    if (newRows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 6).setValues(newRows);
    }
    return { success: true, message: `Zapisano ${newRows.length} osób.` };
  } catch (e) { throw new Error("Błąd zapisu: " + e.message); } finally { lock.releaseLock(); }
}

function deleteRotationalEntry(mode, name) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.openById(MAIN_SHEET_ID);
    const sheet = ss.getSheetByName(dbRotacjaSheetName);
    if (!sheet) throw new Error("Brak bazy danych.");
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, message: "Brak danych." };

    const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    let rowsToDelete = [];
    for (let i = data.length - 1; i >= 0; i--) {
      const rowDate = data[i][0];
      const rowMode = data[i][1];
      const rowName = String(data[i][2]);
      if (rowDate) {
        const rowDateStr = Utilities.formatDate(new Date(rowDate), Session.getScriptTimeZone(), "yyyy-MM-dd");
        if (rowDateStr === todayStr && rowMode === mode && rowName === name) rowsToDelete.push(i + 2);
      }
    }
    if (rowsToDelete.length > 0) {
      rowsToDelete.forEach(rowIdx => { sheet.deleteRow(rowIdx); });
      return { success: true, message: `Usunięto ${name}.` };
    } else { return { success: false, message: "Nie znaleziono wpisu." }; }
  } catch (e) { throw new Error("Błąd usuwania: " + e.message); } finally { lock.releaseLock(); }
}

function deleteBeautyPlanEntries() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.openById(MAIN_SHEET_ID);
    const sheet = ss.getSheetByName(dbRotacjaSheetName);
    if (!sheet) throw new Error("Brak bazy danych.");
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, message: "Brak danych." };

    const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues(); 
    const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    let rowsToDelete = [];
    
    for (let i = data.length - 1; i >= 0; i--) {
      const rowDate = data[i][0];
      const rowMode = data[i][1];
      const param1 = String(data[i][3]); 
      
      if (rowDate) {
        const rowDateStr = Utilities.formatDate(new Date(rowDate), Session.getScriptTimeZone(), "yyyy-MM-dd");
        if (rowDateStr === todayStr && rowMode === 'beauty' && param1 === 'Plan') {
          rowsToDelete.push(i + 2);
        }
      }
    }
    
    if (rowsToDelete.length > 0) {
      rowsToDelete.forEach(rowIdx => { sheet.deleteRow(rowIdx); });
      return { success: true, message: `Anulowano plan (usunięto ${rowsToDelete.length} wpisów).` };
    } else {
      return { success: false, message: "Brak zatwierdzonego planu do anulowania." };
    }
  } catch (e) { 
    throw new Error("Błąd usuwania: " + e.message); 
  } finally { 
    lock.releaseLock(); 
  }
}

function toggleMissionDuration(name) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const ss = SpreadsheetApp.openById(MAIN_SHEET_ID);
    const sheet = ss.getSheetByName(dbRotacjaSheetName);
    if (!sheet) throw new Error("Brak bazy danych.");
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error("Brak danych.");

    const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues(); 
    const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    let newVal = '';
    let found = false;

    for (let i = data.length - 1; i >= 0; i--) {
      const rowDate = data[i][0];
      const rowMode = data[i][1];
      const rowName = String(data[i][2]);
      
      if (rowDate) {
        const rowDateStr = Utilities.formatDate(new Date(rowDate), Session.getScriptTimeZone(), "yyyy-MM-dd");
        if (rowDateStr === todayStr && rowMode === 'missions' && rowName === name) {
           const currentVal = String(data[i][4]); 
           newVal = (currentVal === 'Full') ? '' : 'Full'; 
           sheet.getRange(i + 2, 5).setValue(newVal); 
           found = true;
           break; 
        }
      }
    }
    
    if (found) {
        return { success: true, newVal: newVal, message: `Zmieniono czas misji dla ${name}.` };
    } else {
        throw new Error("Nie znaleziono aktywnej misji dla tego pracownika.");
    }

  } catch (e) { 
    throw new Error("Błąd edycji misji: " + e.message); 
  } finally { 
    lock.releaseLock(); 
  }
}

function toggleExclusion(name, mode) {
  try {
    const key = `EXCLUSIONS_${mode}`;
    let currentList = _readFromGlobalCache(key) || [];
    const index = currentList.indexOf(name);
    let isExcluded = false;
    
    if (index !== -1) {
        currentList.splice(index, 1);
        isExcluded = false;
    } else {
        currentList.push(name);
        isExcluded = true;
    }
    _saveToGlobalCache(key, currentList);
    return { excluded: isExcluded, message: isExcluded ? "Wykluczono pracownika." : "Przywrócono pracownika." };
  } catch (e) { throw new Error("Błąd wykluczania: " + e.message); }
}

function runOneTimeMigration() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); 
  try {
    const ss = SpreadsheetApp.openById(MAIN_SHEET_ID);
    let dbSheet = ss.getSheetByName(dbRotacjaSheetName);
    if (!dbSheet) {
      dbSheet = ss.insertSheet(dbRotacjaSheetName);
      dbSheet.appendRow(["Data", "Typ", "Pracownik", "Param1", "Param2", "Timestamp"]);
      dbSheet.setFrozenRows(1);
    }
    return "Migracja zakończona.";
  } catch (e) { throw new Error("Błąd migracji: " + e.message); } finally { lock.releaseLock(); }
}

// =========================================================
// === MENADŻER URLOPÓW (VACATION MANAGER) - EXTERNAL ===
// =========================================================

function addVacationRequest(data) {
   // Delegacja do DataVacations.gs
   return addVacationRequestExternal(data);
}

function updateVacationStatus(id, newStatus) {
   // Delegacja do DataVacations.gs
   return updateVacationStatusExternal(id, newStatus);
}

function deleteVacationRequest(id) {
   // Delegacja do DataVacations.gs
   return deleteVacationRequestExternal(id);
}

function approveVacationSplit(id, weekStartStr, weekEndStr) {
    // ID format: SHEETNAME::ROW::STARTCOL::ENDCOL
    const parts = id.split('::');
    if (parts.length !== 4) throw new Error("Nieprawidłowe ID urlopu (External).");

    const sheetName = parts[0];
    const rowIdx = parseInt(parts[1]); 
    const seasonConfig = VACATION_SEASONS.find(s => s.name === sheetName);
    if (!seasonConfig) throw new Error("Nie znaleziono konfiguracji sezonu.");

    const ss = SpreadsheetApp.openById(VACATION_SHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    
    // Obliczamy kolumny dla zatwierdzanego zakresu
    const seasonStart = new Date(seasonConfig.startDate);
    const splitStart = new Date(weekStartStr);
    const splitEnd = new Date(weekEndStr);

    // Diff in days
    const diffStart = Math.ceil((splitStart - seasonStart) / (1000 * 60 * 60 * 24));
    const diffEnd = Math.ceil((splitEnd - seasonStart) / (1000 * 60 * 60 * 24));

    const targetStartCol = seasonConfig.startCol + diffStart; 
    const targetEndCol = seasonConfig.startCol + diffEnd;

    // Walidacja
    if (targetStartCol < 0) throw new Error("Data poza zakresem arkusza.");
    
    // Zapisz 'U' (Zatwierdzony) w wybranym zakresie
    const numCols = targetEndCol - targetStartCol + 1;
    if (numCols > 0) {
        // +1 bo Range jest 1-based, +1 bo rowIdx jest 0-based index tablicy (więc wiersz to rowIdx+1)
        // Czekaj, w DataVacations rowIdx to 'r' z pętli. Jeśli pętla była po values, to index 0 = row 1.
        // Jednak DataVacations pobierało: `const values = dataRange.getValues();`.
        // Jeśli `dataRange` to cały arkusz, to index 0 = wiersz 1.
        // Sprawdźmy DataVacations.gs (generowane wcześniej).
        // `let r = startRow; r < data.length`. startRow było np. 4.
        // Tak więc rowIdx to indeks tablicy. Wiersz w arkuszu to rowIdx + 1.
        sheet.getRange(rowIdx + 1, targetStartCol + 1, 1, numCols).setValue('U'); 
    }
    
    return { success: true, message: `Zatwierdzono dni: ${weekStartStr} - ${weekEndStr}.` };
}