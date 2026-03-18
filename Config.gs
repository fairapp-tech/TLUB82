/**
 * Plik: Config.gs
 * Konfiguracja ID arkuszy i nazw rejestrów.
 * Wersja: TURBO ARCHITECTURE (External Vacation Sheet)
 */

// ID głównego arkusza (Baza Danych Aplikacji - Code, Tasks, Rotation)
const MAIN_SHEET_ID = '1jmCVCiAsNXrpp_FQkX1EBValext-k8KRQiF1RCKp5K8';

// ID zewnętrznego arkusza URLOPÓW (Planer 2026/27)
const VACATION_SHEET_ID = '1sQU62KG4zFqzxWUnRsjQh8a2aupTpDs2SjWYikwKeGo';

// Filtrowanie grupy w arkuszu urlopowym (Kolumna F)
// Aplikacja będzie pobierać/zapisywać urlopy TYLKO dla tej grupy.
const VACATION_GROUP_FILTER = 'UB8'; 

// 1. Zadania (Pozostają w osobnym arkuszu dla przejrzystości "Tablicy Zadań")
const tasksSheetId = MAIN_SHEET_ID;
const tasksSheetName = 'Strona';

// 2. Pracownicy (Baza HR - źródło read-only)
const employeesSheetId = MAIN_SHEET_ID;
const employeesSheetName = 'Grupa';

// 3. GŁÓWNA BAZA ROTACJI
const dbRotacjaSheetName = 'DB_Rotacja';

// 4. GŁÓWNY CACHE
const globalCacheSheetName = 'DB_Cache'; 

// 5. BAZA URLOPÓW (Stara lokalna - zostawiona, aby nie zerwać referencji, ale logika przechodzi na External)
const dbUrlopySheetName = 'DB_Urlopy';

// 6. Źródła zewnętrzne (Tylko do odczytu przy imporcie/syncu)
const planningSheetName = 'Planowanie'; 
const grafikSheetName = 'Grafik';

// 7. Stare rejestry (Zachowane jako stałe dla migracji)
const LEGACY_SHEET_NAMES = {
  stow: 'Rejestr STOW',
  carts: 'Rejestr Wózki',
  missions: 'Rejestr Misje',
  support: 'Rejestr Wsparcie',
  beauty: 'Rejestr Beauty'
};