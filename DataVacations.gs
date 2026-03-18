/**
 * Plik: DataVacations.gs
 * Moduł: Obsługa Zewnętrznego Planera Urlopów (Excel-style)
 * Sezon: 2026/2027
 * Struktura: Kolumny to dni, Wiersze to pracownicy.
 */

// Konfiguracja Arkuszy Sezonowych (Nazwa, Data Startu, Index Kolumny Startowej - I=8)
const VACATION_SEASONS = [
    { name: 'Luty - Maj26', startDate: '2026-02-02', startCol: 8 },
    { name: 'Czerwiec - Wrzesień26', startDate: '2026-06-01', startCol: 8 },
    { name: 'Październik - Styczeń26/27', startDate: '2026-09-28', startCol: 8 }
];

// Mapowanie Kodów
const CODE_PENDING = 'UNP'; // Wniosek
const CODE_APPROVED = 'U';   // Zatwierdzony

// === FUNKCJE GŁÓWNE ===

/**
 * Pobiera urlopy z zewnętrznego arkusza i formatuje do standardu aplikacji.
 */
function getVacationsExternal(yearFilter) {
    const ss = SpreadsheetApp.openById(VACATION_SHEET_ID);
    let allVacations = [];

    VACATION_SEASONS.forEach(season => {
        const sheet = ss.getSheetByName(season.name);
        if (!sheet) return;

        const data = sheet.getDataRange().getValues();
        if (data.length < 5) return; // Za mało danych

        // Iteracja po pracownikach (od wiersza np. 5 - zakładam nagłówki)
        // Szukamy nagłówków, ale tu przyjmijmy bezpieczny start lub wykrywanie
        let startRow = 0;
        for(let r=0; r<data.length; r++) {
            if (String(data[r][5]) === 'Grupa' || String(data[r][5]) === 'GRUPA') { // Kolumna F to indeks 5
                startRow = r + 1;
                break;
            }
        }
        if (startRow === 0) startRow = 4; // Fallback

        const seasonStart = new Date(season.startDate);

        for (let r = startRow; r < data.length; r++) {
            const row = data[r];
            const group = String(row[5]).trim(); // Kolumna F
            
            // Filtrowanie grupy (UB8)
            if (group !== VACATION_GROUP_FILTER) continue;

            const surname = String(row[1]).trim(); // B
            const name = String(row[2]).trim();    // C
            const fullName = `${name} ${surname}`; // Format w aplikacji: Imię Nazwisko

            // Skanowanie dni (kolumny)
            let currentVacation = null;

            for (let c = season.startCol; c < row.length; c++) {
                const cellVal = String(row[c]).trim().toUpperCase();
                
                // Oblicz datę dla kolumny
                const dayOffset = c - season.startCol;
                const currentDate = new Date(seasonStart);
                currentDate.setDate(currentDate.getDate() + dayOffset);
                const dateStr = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

                const isVac = (cellVal === CODE_PENDING || cellVal === CODE_APPROVED);
                const status = (cellVal === CODE_APPROVED) ? 'Potwierdzony' : 'Oczekujący';

                if (isVac) {
                    if (!currentVacation) {
                        // Start nowego bloku
                        currentVacation = {
                            id_prefix: `${season.name}|${r}|${c}`, // Tymczasowe ID startu
                            name: fullName,
                            start: dateStr,
                            end: dateStr,
                            status: status, // Status pierwszego dnia determinuje blok (uproszczenie)
                            seasonName: season.name,
                            rowIdx: r,
                            startCol: c,
                            endCol: c
                        };
                    } else {
                        // Kontynuacja bloku
                        // Jeśli status się zmienił w trakcie (np. UNP -> U), rozbijamy? 
                        // Dla uproszczenia: jeśli status inny, zamykamy stary i otwieramy nowy.
                        if (currentVacation.status !== status) {
                            allVacations.push(_finalizeVacationObj(currentVacation));
                            currentVacation = {
                                id_prefix: `${season.name}|${r}|${c}`,
                                name: fullName,
                                start: dateStr,
                                end: dateStr,
                                status: status,
                                seasonName: season.name,
                                rowIdx: r,
                                startCol: c,
                                endCol: c
                            };
                        } else {
                            currentVacation.end = dateStr;
                            currentVacation.endCol = c;
                        }
                    }
                } else {
                    // Koniec bloku (puste lub inny kod)
                    if (currentVacation) {
                        allVacations.push(_finalizeVacationObj(currentVacation));
                        currentVacation = null;
                    }
                }
            }
            // Zapisz ostatni blok jeśli istniał na końcu wiersza
            if (currentVacation) {
                allVacations.push(_finalizeVacationObj(currentVacation));
            }
        }
    });

    return allVacations;
}

function _finalizeVacationObj(v) {
    // Generujemy unikalne ID pozwalające namierzyć zakres w Excelu
    // ID: SHEETNAME::ROW::STARTCOL::ENDCOL
    const id = `${v.seasonName}::${v.rowIdx}::${v.startCol}::${v.endCol}`;
    
    // Określenie sezonu (tekstowego dla UI)
    const month = new Date(v.start).getMonth();
    let seasonLabel = "Inny";
    if (month >= 1 && month <= 4) seasonLabel = "Luty-Maj";
    else if (month >= 5 && month <= 8) seasonLabel = "Czerwiec-Wrzesień";
    else if (month >= 9 || month === 0) seasonLabel = "Październik-Styczeń";

    return {
        id: id,
        timestamp: new Date(v.start).getTime(), // Używamy daty startu jako timestamp zgłoszenia (bo w excelu nie ma daty zgłoszenia)
        name: v.name,
        startDate: v.start,
        endDate: v.end,
        season: seasonLabel,
        note: "", // Excel nie przechowuje notatek w komórkach dni
        status: v.status,
        approver: (v.status === 'Potwierdzony' ? 'System' : '')
    };
}

/**
 * Dodaje wniosek (wpisuje UNP w komórki).
 * Obsługuje batch (wiele osób).
 */
function addVacationRequestExternal(data) {
    const ss = SpreadsheetApp.openById(VACATION_SHEET_ID);
    const names = Array.isArray(data.name) ? data.name : [data.name];
    const startDate = new Date(data.startDate);
    const endDate = new Date(data.endDate);
    
    // Ustal status (UNP czy U)
    const codeToWrite = (data.status === 'Potwierdzony') ? CODE_APPROVED : CODE_PENDING;

    let log = [];

    // Dla każdego pracownika
    names.forEach(targetName => {
        // Musimy znaleźć odpowiedni arkusz(e) i wiersz pracownika
        VACATION_SEASONS.forEach(season => {
            const sStart = new Date(season.startDate);
            // Sprawdź czy ten sezon w ogóle pokrywa się z zakresem urlopu
            // To proste sprawdzenie, można ulepszyć (czy zakresy na siebie zachodzą)
            // Zakładamy, że sezony są chronologiczne.
            
            // Oblicz indexy kolumn
            const diffStart = Math.ceil((startDate - sStart) / (1000 * 60 * 60 * 24));
            const diffEnd = Math.ceil((endDate - sStart) / (1000 * 60 * 60 * 24));
            
            // Jeśli urlop kończy się przed startem sezonu lub zaczyna po (z grubsza) - pomiń
            // (Tu uproszczone: sprawdzamy czy indeksy wpadają w sensowny zakres arkusza, np. 0-120 dni)
            if (diffEnd < 0) return; 
            
            const sheet = ss.getSheetByName(season.name);
            if (!sheet) return;
            
            // Znajdź wiersz pracownika
            const dataRange = sheet.getDataRange();
            const values = dataRange.getValues();
            let rowIdx = -1;

            for(let r=0; r<values.length; r++) {
                // Sprawdzamy grupę
                if (String(values[r][5]).trim() !== VACATION_GROUP_FILTER) continue;
                
                // Sprawdzamy nazwisko i imię
                const surname = String(values[r][1]).trim();
                const name = String(values[r][2]).trim();
                const fullName1 = `${name} ${surname}`;
                const fullName2 = `${surname} ${name}`; // Czasem format jest odwrotny
                
                if (fullName1 === targetName || fullName2 === targetName) {
                    rowIdx = r;
                    break;
                }
            }

            if (rowIdx === -1) return; // Nie znaleziono pracownika w tym sezonie

            // Wpisz dane w odpowiednie kolumny
            // StartCol to I (index 8). Data Startu Sezonu to index 8.
            // diffStart = 0 oznacza dzień startu sezonu.
            
            const effectiveStartCol = season.startCol + Math.max(0, diffStart);
            const effectiveEndCol = season.startCol + diffEnd;
            
            // Limit kolumn w arkuszu (żeby nie wyjść poza zakres)
            const maxCol = values[0].length - 1;
            
            if (effectiveStartCol > maxCol) return; // Urlop poza zakresem tego arkusza (później)

            const startC = Math.max(season.startCol, effectiveStartCol);
            const endC = Math.min(maxCol, effectiveEndCol);

            if (startC <= endC) {
                // +1 bo getRange jest 1-based
                sheet.getRange(rowIdx + 1, startC + 1, 1, (endC - startC + 1)).setValue(codeToWrite);
                log.push(`Zapisano ${targetName} w ${season.name}`);
            }
        });
    });

    if (log.length === 0) throw new Error("Nie znaleziono pracowników w arkuszach urlopowych lub data poza zakresem.");
    return { success: true, message: `Zapisano wniosek: ${log.join(', ')}` };
}

/**
 * Aktualizuje status (Zatwierdź / Cofnij / Odrzuć)
 * ID zawiera namiary na zakres.
 */
function updateVacationStatusExternal(id, newStatus) {
    // ID format: SHEETNAME::ROW::STARTCOL::ENDCOL
    const parts = id.split('::');
    if (parts.length !== 4) throw new Error("Nieprawidłowe ID urlopu.");

    const sheetName = parts[0];
    const rowIdx = parseInt(parts[1]);
    const startCol = parseInt(parts[2]);
    const endCol = parseInt(parts[3]);

    const ss = SpreadsheetApp.openById(VACATION_SHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error("Nie znaleziono arkusza sezonu.");

    let val = '';
    if (newStatus === 'Potwierdzony') val = CODE_APPROVED;
    else if (newStatus === 'Oczekujący') val = CODE_PENDING;
    else if (newStatus === 'Odrzucony') val = ''; // Czyścimy komórki

    // +1 bo Range jest 1-based
    sheet.getRange(rowIdx + 1, startCol + 1, 1, (endCol - startCol + 1)).setValue(val);

    return { success: true, message: `Status zmieniony na: ${newStatus}` };
}

/**
 * Usuwanie wniosku (to samo co odrzucenie - czyści komórki)
 */
function deleteVacationRequestExternal(id) {
    return updateVacationStatusExternal(id, 'Odrzucony');
}