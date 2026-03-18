/**
 * Plik: Config.gs
 * Konfiguracja ID arkuszy i nazw rejestrów.
 * Wersja: TURBO ARCHITECTURE (DB_Rotacja + DB_Cache)
 */

// ID głównego arkusza
const MAIN_SHEET_ID = '1jmCVCiAsNXrpp_FQkX1EBValext-k8KRQiF1RCKp5K8';

// 1. Zadania (Pozostają w osobnym arkuszu dla przejrzystości "Tablicy Zadań")
const tasksSheetId = MAIN_SHEET_ID;
const tasksSheetName = 'Strona';

// 2. Pracownicy (Baza HR - źródło read-only)
const employeesSheetId = MAIN_SHEET_ID;
const employeesSheetName = 'Grupa';

// 3. GŁÓWNA BAZA ROTACJI (NOWOŚĆ)
// Zastępuje: Rejestr STOW, Wózki, Misje, Wsparcie, Beauty.
// Struktura: [Data, Typ, Pracownik, Param1, Param2, Timestamp]
const dbRotacjaSheetName = 'DB_Rotacja';

// 4. GŁÓWNY CACHE (NOWOŚĆ)
// Przechowuje JSONy dla: Obecności, Planowania, Ustawień, Grafiku (wykluczenia).
// Zastępuje: Cache Obecności, Cache Planowania, Ustawienia (jako osobny sheet).
const globalCacheSheetName = 'DB_Cache'; 

// 5. Źródła zewnętrzne (Tylko do odczytu przy imporcie/syncu)
const planningSheetName = 'Planowanie'; // Źródło raportu
const grafikSheetName = 'Grafik';       // Źródło z Excela

// 6. Stare rejestry (Zachowane jako stałe dla migracji)
const LEGACY_SHEET_NAMES = {
  stow: 'Rejestr STOW',
  carts: 'Rejestr Wózki',
  missions: 'Rejestr Misje',
  support: 'Rejestr Wsparcie',
  beauty: 'Rejestr Beauty'
};

// Stare nazwy (dla kompatybilności przy migracji)
const cacheSheetName = 'Cache Obecności'; 
const cachePlanningSheetName = 'Cache Planowania';
const settingsSheetName = 'Ustawienia';