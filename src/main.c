#include <windows.h>
#include <windowsx.h>
#include <commctrl.h>
#include <shellapi.h>
#include <shlobj.h>
#include <dwmapi.h>
#include <uxtheme.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <wchar.h>
#include <math.h>
#include <wctype.h>

#ifdef _MSC_VER
#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "dwmapi.lib")
#pragma comment(lib, "uxtheme.lib")
#endif

// Unicode glyph constants (avoid codepage/mojibake issues)
static const WCHAR CH_MUL = L'\x00D7'; // ×
static const WCHAR CH_DIV = L'\x00F7'; // ÷
static const WCHAR DEG    = L'\x00B0'; // °
static const WCHAR MU     = L'\x03BC'; // µ

static const int PANEL_WIDTH  = 380;  // base logical size; scaled by DPI
static const int PANEL_HEIGHT = 260;
static const int INPUT_H      = 36;
static const int ICONROW_H    = 26;
static const int TABSTRIP_H   = 26;   // when symbol dropdown visible
static const int PADDING      = 10;

// UI Element dimensions (logical pixels)
static const int ICON_WIDTH   = 22;
static const int ICON_GAP     = 8;
static const int TAB_WIDTH    = 70;
static const int ITEM_HEIGHT  = 28;
static const int CARET_WIDTH  = 2;
static const int INPUT_MARGIN = 10;
static const int CARET_MARGIN = 6;

// 8-ball responses - easily modifiable array
static const WCHAR* k8BallResponses[] = {
    L"It is certain",
    L"It is decidedly so",
    L"Without a doubt",
    L"Yes definitely",
    L"You may rely on it",
    L"As I see it, yes",
    L"Most likely",
    L"Perhaps",
    L"Yes",
    L"Signs point to yes",
    L"Better not tell you now",
    L"Don't count on it",
    L"My reply is no",
    L"My sources say no",
    L"Outlook not so good",
    L"There is no way",
    L"Not a chance",
    L"Very doubtful"
};
static const int k8BallResponseCount = sizeof(k8BallResponses) / sizeof(k8BallResponses[0]);

// Array size constants
#define ARRAY_COUNT(arr) (sizeof(arr) / sizeof((arr)[0]))
static const int kUnitCount = 10;  // U_IN through U_LY
static const int kUnitWordCount = 38;  // Will be set after kUnitWords definition
static const int kTempWordCount = 9;
static const int kTimeUnitWordCount = 26;
static const int kHolidayCount = 18;
static const int kGenericUnitCategoryCount = 4;

// Buffer size constants
#define MAX_INPUT_BUFFER 256
#define MAX_SMALL_BUFFER 32
#define MAX_MEDIUM_BUFFER 64
#define MAX_LARGE_BUFFER 128
#define MAX_XLARGE_BUFFER 256
#define MAX_HUGE_BUFFER 512
#define MAX_EXPR_BUFFER 1024

// Random number state
static BOOL g_randomInitialized = FALSE;

// Font cache to avoid repeated allocations
static HFONT g_fontCache[4] = {0};  // [0]=input, [1]=icon, [2]=tab, [3]=item
static BOOL g_fontCacheInitialized = FALSE;

// Font sizes
static const int FONT_INPUT   = 12;
static const int FONT_ICON    = 12;
static const int FONT_TAB     = 10;
static const int FONT_ITEM    = 11;

// Timer and scroll constants
static const int CARET_TIMER_ID = 2;
static const int CARET_BLINK_MS = 530;
static const int SCROLL_STEP    = 40;

// Menu dimensions
static const int MENU_ITEM_HEIGHT = 24;
static const int MENU_ITEM_PADDING = 20;
static const int MENU_TEXT_MARGIN = 10;

static const COLORREF COL_BG        = RGB(30, 32, 36);  // dark background
static const COLORREF COL_INPUT     = RGB(38, 41, 46);
static const COLORREF COL_ICONROW   = RGB(33, 36, 41);
static const COLORREF COL_CONTENT   = RGB(26, 28, 32);
static const COLORREF COL_TEXT      = RGB(220, 224, 228);
static const COLORREF COL_SUBTLE    = RGB(86, 90, 96);
static const COLORREF COL_ACCENT    = RGB(90, 160, 255); // hover/selection
static const COLORREF COL_HOVER     = RGB(48, 52, 58);
static const COLORREF COL_SUGGEST   = RGB(150, 156, 164); // slightly brighter than subtle

// Hotkey
#define HOTKEY_ID  1
#define HOTKEY_MOD MOD_ALT
#define HOTKEY_KEY 0x51 // 'Q'
#define IDI_APP_ICON 101

// Context menu command IDs
#define IDM_CLEAR_INPUT  1001
#define IDM_SETTINGS     1002
#define IDM_INSERT_BASE  1100
#define IDM_INS_PLUS     (IDM_INSERT_BASE + 1)
#define IDM_INS_MINUS    (IDM_INSERT_BASE + 2)
#define IDM_INS_TIMES    (IDM_INSERT_BASE + 3)
#define IDM_INS_DIV      (IDM_INSERT_BASE + 4)
#define IDM_INS_LPAREN   (IDM_INSERT_BASE + 5)
#define IDM_INS_RPAREN   (IDM_INSERT_BASE + 6)
#define IDM_INS_PERCENT  (IDM_INSERT_BASE + 7)
#define IDM_INS_POWER    (IDM_INSERT_BASE + 8)

// Lengths submenu IDs
#define IDM_MEAS_BASE    1200
#define IDM_MEAS_IN      (IDM_MEAS_BASE + 1)
#define IDM_MEAS_FT      (IDM_MEAS_BASE + 2)
#define IDM_MEAS_YD      (IDM_MEAS_BASE + 3)
#define IDM_MEAS_M       (IDM_MEAS_BASE + 4)
#define IDM_MEAS_CM      (IDM_MEAS_BASE + 5)
#define IDM_MEAS_MM      (IDM_MEAS_BASE + 6)
#define IDM_MEAS_KM      (IDM_MEAS_BASE + 7)
#define IDM_MEAS_MI      (IDM_MEAS_BASE + 8)
#define IDM_MEAS_NMI     (IDM_MEAS_BASE + 9)
#define IDM_MEAS_LY      (IDM_MEAS_BASE + 10)

// Temperatures submenu IDs
#define IDM_TEMP_BASE    1300
#define IDM_TEMP_C       (IDM_TEMP_BASE + 1)
#define IDM_TEMP_F       (IDM_TEMP_BASE + 2)
#define IDM_TEMP_K       (IDM_TEMP_BASE + 3)

// Time submenu IDs
#define IDM_TIME_BASE    1400
#define IDM_TIME_NS      (IDM_TIME_BASE + 1)
#define IDM_TIME_US      (IDM_TIME_BASE + 2)
#define IDM_TIME_MS      (IDM_TIME_BASE + 3)
#define IDM_TIME_S       (IDM_TIME_BASE + 4)
#define IDM_TIME_MIN     (IDM_TIME_BASE + 5)
#define IDM_TIME_H       (IDM_TIME_BASE + 6)
#define IDM_TIME_DAY     (IDM_TIME_BASE + 7)
#define IDM_TIME_WEEK    (IDM_TIME_BASE + 8)
#define IDM_TIME_MONTH   (IDM_TIME_BASE + 9)
#define IDM_TIME_YEAR    (IDM_TIME_BASE + 10)

// Holiday submenu IDs
#define IDM_HOLIDAY_BASE          1500
#define IDM_HOLIDAY_NEWYEAR_DAY   (IDM_HOLIDAY_BASE + 1)
#define IDM_HOLIDAY_NEWYEAR_EVE   (IDM_HOLIDAY_BASE + 2)
#define IDM_HOLIDAY_MLK           (IDM_HOLIDAY_BASE + 3)
#define IDM_HOLIDAY_PRESIDENTS    (IDM_HOLIDAY_BASE + 4)
#define IDM_HOLIDAY_MEMORIAL      (IDM_HOLIDAY_BASE + 5)
#define IDM_HOLIDAY_INDEPENDENCE  (IDM_HOLIDAY_BASE + 6)
#define IDM_HOLIDAY_LABOR         (IDM_HOLIDAY_BASE + 7)
#define IDM_HOLIDAY_COLUMBUS      (IDM_HOLIDAY_BASE + 8)
#define IDM_HOLIDAY_HALLOWEEN     (IDM_HOLIDAY_BASE + 9)
#define IDM_HOLIDAY_VETERANS      (IDM_HOLIDAY_BASE + 10)
#define IDM_HOLIDAY_THANKSGIVING  (IDM_HOLIDAY_BASE + 11)
#define IDM_HOLIDAY_CHRISTMAS_EVE (IDM_HOLIDAY_BASE + 12)
#define IDM_HOLIDAY_CHRISTMAS_DAY (IDM_HOLIDAY_BASE + 13)
#define IDM_HOLIDAY_VALENTINE     (IDM_HOLIDAY_BASE + 14)
#define IDM_HOLIDAY_MOTHERS       (IDM_HOLIDAY_BASE + 15)
#define IDM_HOLIDAY_FATHERS       (IDM_HOLIDAY_BASE + 16)

// Area submenu IDs
#define IDM_AREA_BASE         1600
#define IDM_AREA_SQIN         (IDM_AREA_BASE + 1)
#define IDM_AREA_SQFT         (IDM_AREA_BASE + 2)
#define IDM_AREA_SQYD         (IDM_AREA_BASE + 3)
#define IDM_AREA_SQM          (IDM_AREA_BASE + 4)
#define IDM_AREA_SQKM         (IDM_AREA_BASE + 5)
#define IDM_AREA_SQMI         (IDM_AREA_BASE + 6)
#define IDM_AREA_ACRE         (IDM_AREA_BASE + 7)
#define IDM_AREA_HECTARE      (IDM_AREA_BASE + 8)

// Weight submenu IDs
#define IDM_WEIGHT_BASE       1700
#define IDM_WEIGHT_MG         (IDM_WEIGHT_BASE + 1)
#define IDM_WEIGHT_G          (IDM_WEIGHT_BASE + 2)
#define IDM_WEIGHT_KG         (IDM_WEIGHT_BASE + 3)
#define IDM_WEIGHT_OZ         (IDM_WEIGHT_BASE + 4)
#define IDM_WEIGHT_LB         (IDM_WEIGHT_BASE + 5)
#define IDM_WEIGHT_TON        (IDM_WEIGHT_BASE + 6)

// Volume submenu IDs
#define IDM_VOLUME_BASE       1800
#define IDM_VOLUME_ML         (IDM_VOLUME_BASE + 1)
#define IDM_VOLUME_L          (IDM_VOLUME_BASE + 2)
#define IDM_VOLUME_FLOZ       (IDM_VOLUME_BASE + 3)
#define IDM_VOLUME_CUP        (IDM_VOLUME_BASE + 4)
#define IDM_VOLUME_PINT       (IDM_VOLUME_BASE + 5)
#define IDM_VOLUME_QUART      (IDM_VOLUME_BASE + 6)
#define IDM_VOLUME_GAL        (IDM_VOLUME_BASE + 7)
#define IDM_VOLUME_TBSP       (IDM_VOLUME_BASE + 8)
#define IDM_VOLUME_TSP        (IDM_VOLUME_BASE + 9)

// Memory submenu IDs
#define IDM_MEMORY_BASE       1900
#define IDM_MEMORY_BIT        (IDM_MEMORY_BASE + 1)
#define IDM_MEMORY_BYTE       (IDM_MEMORY_BASE + 2)
#define IDM_MEMORY_KB         (IDM_MEMORY_BASE + 3)
#define IDM_MEMORY_MB         (IDM_MEMORY_BASE + 4)
#define IDM_MEMORY_GB         (IDM_MEMORY_BASE + 5)
#define IDM_MEMORY_TB         (IDM_MEMORY_BASE + 6)
#define IDM_MEMORY_KIB        (IDM_MEMORY_BASE + 7)
#define IDM_MEMORY_MIB        (IDM_MEMORY_BASE + 8)
#define IDM_MEMORY_GIB        (IDM_MEMORY_BASE + 9)
#define IDM_MEMORY_TIB        (IDM_MEMORY_BASE + 10)
#define IDM_MEMORY_KBIT       (IDM_MEMORY_BASE + 11)
#define IDM_MEMORY_MBIT       (IDM_MEMORY_BASE + 12)
#define IDM_MEMORY_GBIT       (IDM_MEMORY_BASE + 13)
#define IDM_MEMORY_TBIT       (IDM_MEMORY_BASE + 14)

// Settings dialog control IDs
#define IDC_SETTINGS_HOTKEY       2001
#define IDC_SETTINGS_OPACITY      2002
#define IDC_SETTINGS_SCALE        2003
#define IDC_SETTINGS_TOPMOST      2004
#define IDC_SETTINGS_SYMBOLS      2005
#define IDC_SETTINGS_DATEFORMAT   2006
#define IDC_SETTINGS_SAVE         2007
#define IDC_SETTINGS_CANCEL       2008

// Tray icon integration
#define WMAPP_TRAYICON            (WM_APP + 1)
#define TRAY_ICON_UID             1
#define IDM_TRAY_SHOW             2101
#define IDM_TRAY_SETTINGS         2102
#define IDM_TRAY_EXIT             2103

// DPI helpers (loaded dynamically for compatibility)
typedef BOOL (WINAPI *SetProcessDpiAwarenessContext_t)(HANDLE);
typedef UINT (WINAPI *GetDpiForWindow_t)(HWND);
typedef HRESULT (WINAPI *GetDpiForMonitor_t)(HMONITOR, int, UINT*, UINT*);

static GetDpiForWindow_t pGetDpiForWindow = NULL;
static GetDpiForMonitor_t pGetDpiForMonitor = NULL;

typedef struct Holiday Holiday;

typedef struct UiState {
    int dpi;
    int scale;
    int autoScalePercent;
    BOOL hoveringInput;
    int hoverRightIcon;
    int hoverLeftIcon;
    int hoverHistory;
    BOOL symbolsOpen;
    int activeTab;
    int scrollOffset; // pixels
    SIZE panelSizePx; // in physical pixels, updated on DPI change
    // Minimal input buffer for top bar (to handle !exit)
    WCHAR input[256];
    int inputLen;
    BOOL hasFocus;
    BOOL caretVisible;
    int caretPos;
    int selAnchor;
    int selStart;
    int selEnd;
    BOOL selectingInput;
    BOOL suggestActive;
    WCHAR suggest[64];
    int suggestLen;
    BOOL ghostParensActive;
    WCHAR ghostParens[32];
    int ghostParensLen;
    // Last answer suggestion
    BOOL lastAnsSuggestActive;
    WCHAR lastAns[64];
    int lastAnsLen;
    int historyNavIndex; // -1 when not navigating; otherwise index into g_history
    BOOL suppressLastAnsSuggest; // suppress showing last answer when empty
    BOOL showingMenu;
    BOOL showingSettings;
} UiState;

static UiState g_ui = {0};
static const wchar_t *kTabs[] = { L"Greek", L"Operators", L"Arrows", L"Sets" };
static const int kTabCount = 4;
static const wchar_t *kIcons[] = { L"≡" }; // right-side
static const int kIconCount = 1;
static const wchar_t *kOpIcons[] = { L"+", L"-", L"×", L"÷" }; // left-side operators
static const int kOpIconCount = 4;
// Dynamic history buffer
typedef struct HistoryItem {
    WCHAR *text;    // displayed text
    WCHAR *insert;  // text to insert when clicked; NULL => no action
    WCHAR *answer;  // computed answer string (e.g., right side of =), NULL if none
} HistoryItem;

#define HISTORY_CAP 256
static HistoryItem *g_history[HISTORY_CAP];
static int g_historyCount = 0;

// Live preview (rendered as a transient top row above history)
static BOOL g_liveVisible = FALSE;
static WCHAR g_liveText[512];
static BOOL g_liveComplete = FALSE;
// Live time preview tracking (for periodic refresh)
static BOOL g_timePreviewActive = FALSE;

typedef struct AppSettings {
    UINT hotkeyMods;
    UINT hotkeyVk;
    int opacityPercent;
    int interfaceScalePercent;
    BOOL alwaysOnTop;
    BOOL symbolsOpenByDefault;
    int dateFormatStyle;
} AppSettings;

enum {
    DATEFMT_MDY = 0,
    DATEFMT_DMY = 1
};

static AppSettings g_settings = { HOTKEY_MOD, HOTKEY_KEY, 90, 80, TRUE, FALSE, DATEFMT_MDY };
static WCHAR g_settingsPath[MAX_PATH] = {0};
static HWND g_hwndSettings = NULL;
static UINT g_wmTaskbarCreated = 0;
static BOOL g_trayIconAdded = FALSE;

// forward declarations (used before definitions)
static int S(int v);
static void GetLayoutRects(RECT wnd, RECT *rcInput, RECT *rcIcons, RECT *rcTabStrip, RECT *rcContent);
static HFONT CreateUiFont(int sizePt, BOOL semibold);
static void ReleaseUiFont(HFONT font, int sizePt, BOOL semibold);
static void AdjustWindowHeight(HWND hwnd);
static void SpawnNearCursor(HWND hwnd);
static BOOL AddTrayIcon(HWND hwnd);
static void RemoveTrayIcon(HWND hwnd);
static void ShowTrayMenu(HWND hwnd);
static void ToggleTrayWindow(HWND hwnd);
static void UpdateCommandSuggestion(void);
static void UpdateGhostParens(void);
static BOOL ApplySuggestionToInput(void);
static BOOL AppendToInput(const WCHAR *s);
static BOOL HasInputSelection(void);
static void ClearInputSelection(void);
static void SetInputSelection(int anchor, int caret);
static BOOL DeleteSelectedInput(void);
static BOOL CopyInputSelectionToClipboard(HWND hwnd);
static BOOL PasteClipboardText(HWND hwnd);
static int MeasureTextWidth(HDC hdc, const WCHAR *text, int len);
static void ResetCaretBlink(HWND hwnd);
static void RefreshInputDisplay(HWND hwnd);
static void HandleTextDeletion(HWND hwnd, BOOL forward, BOOL wholeWord);
static void HandleCaretMovement(HWND hwnd, WPARAM wParam, BOOL extendSelection);
static void ProcessInputExpression(HWND hwnd);
static void History_Add(HWND hwnd, const WCHAR *text, const WCHAR *insert, const WCHAR *answer, BOOL autoscroll);
static void ShowContextMenu(HWND hwnd, int x, int y);
static BOOL InsertAtCaret(const WCHAR *s);
static void normalize_of_operator(const WCHAR *src, WCHAR *dst, size_t cap);
static void normalize_percent_word(const WCHAR *src, WCHAR *dst, size_t cap);
// Map input operator glyphs to ASCII for parsing
static void normalize_operators_basic(const WCHAR *src, WCHAR *dst, size_t cap);
static BOOL FindLastAnswer(int *idxOut, WCHAR *buf, size_t cap);
static void UpdateLastAnswerSuggestion(void);
static BOOL InsertLastAnswerIfEmpty(HWND hwnd);
static BOOL IsIncompleteString(const WCHAR *input);
// Time query helpers
typedef enum { TQ_NONE, TQ_UNTIL, TQ_SINCE, TQ_BETWEEN, TQ_NOW } TimeQueryType;
typedef enum { TU_DEFAULT, TU_SECONDS, TU_MINUTES, TU_HOURS, TU_DAYS, TU_WEEKS, TU_MONTHS, TU_YEARS } TimeOutUnit;
typedef struct { TimeQueryType type; SYSTEMTIME a; SYSTEMTIME b; BOOL hasA; BOOL hasB; BOOL hasUnit; TimeOutUnit outUnit; } TimeQuery;
static BOOL ParseDateToken(const WCHAR *s, int *consumed, SYSTEMTIME *out);
// NOTE: ParseTimeToken12 was declared previously but never implemented; remove unused forward declaration.
static BOOL ParseDateTimeAny(const WCHAR *s, int *consumed, SYSTEMTIME *out);
static BOOL HandleHolidayQuery(HWND hwnd, const WCHAR *input);
static BOOL DetectTimeQuery(const WCHAR *s, TimeQuery *out);
static ULONGLONG FileTimeToULL(FILETIME ft);
static void FormatDuration(ULONGLONG secs, WCHAR *buf, size_t cap);
static void FormatDateLong(const SYSTEMTIME *st, WCHAR *buf, size_t cap);
static void FormatLocalDateTime(const SYSTEMTIME *st, WCHAR *buf, size_t cap);
static const WCHAR *TimeOutUnitLabel(TimeOutUnit unit);
static double ConvertDurationToUnit(double seconds, TimeOutUnit unit, const WCHAR **labelOut);
static int CompareCalendarDate(const SYSTEMTIME *a, const SYSTEMTIME *b);
static int DaysFromCivilDate(int year, int month, int day);
static int DaysBetweenCalendarDates(const SYSTEMTIME *fromDate, const SYSTEMTIME *toDate);
static BOOL is_digit(WCHAR c);
static BOOL ParseHolidayDateToken(const WCHAR *s, int *consumed, SYSTEMTIME *out);
static BOOL ResolveHolidayDate(const Holiday *holiday, BOOL hasExplicitYear, int targetYear, const SYSTEMTIME *nowLocal, SYSTEMTIME *holidayDateOut, int *resolvedYearOut, BOOL *isTodayOut);
static BOOL FindHolidayForToday(const SYSTEMTIME *nowLocal, const Holiday **holidayOut, SYSTEMTIME *holidayDateOut);
// Comparison queries ("is A > B")
typedef enum { CMP_NONE, CMP_LT, CMP_LE, CMP_GT, CMP_GE, CMP_EQ, CMP_NE } CmpOp;
static BOOL BuildComparisonAnswer(const WCHAR *input, BOOL forLive, WCHAR *outText, size_t outCap, WCHAR *answerOut, size_t ansCap, BOOL *completeOut);
static int wcsieq(const WCHAR *a, const WCHAR *b);

// Random number and dice rolling functions
static void InitializeRandom(void);
static int GetRandomNumber(int min, int max);
static BOOL Handle8BallCommand(HWND hwnd, const WCHAR *input);
static BOOL HandleRollCommand(HWND hwnd, const WCHAR *input);
static BOOL ParseDiceNotation(const WCHAR *input, int *diceCount, int *diceSides, int *modifier, WCHAR *errorMsg, size_t errorCap);
static void LoadSettings(void);
static BOOL SaveSettings(void);
static void ApplyWindowVisuals(HWND hwnd);
static void OpenSettingsDialog(HWND parent);
static LRESULT CALLBACK SettingsWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);
static void ApplyScaleMetrics(HWND hwnd);
static BOOL RegisterConfiguredHotkey(HWND hwnd, BOOL showFailureFeedback);
static void FormatHotkeyText(UINT mods, UINT vk, WCHAR *buf, size_t cap);
static int ClampDateFormatStyle(int value);

// Units and conversions (length)
typedef enum { U_IN, U_FT, U_YD, U_M, U_CM, U_MM, U_KM, U_MI, U_NMI, U_LY, U_INVALID } UnitId;
static void format_double(double v, WCHAR *buf, size_t cap);
static void format_double_full(double v, WCHAR *buf, size_t cap);
static UnitId match_unit_word(const WCHAR *s, int *outLen);
static int match_temp_unit_word(const WCHAR *s, int *outLen);

typedef struct { const WCHAR *name_singular; const WCHAR *name_plural; const WCHAR *abbr; double to_m; } UnitInfo;
static const UnitInfo kUnits[] = {
    { L"inch", L"inches", L"in", 0.0254 },
    { L"foot", L"feet",   L"ft", 0.3048 },
    { L"yard", L"yards",  L"yd", 0.9144 },
    { L"meter", L"meters", L"m", 1.0 },
    { L"centimeter", L"centimeters", L"cm", 0.01 },
    { L"millimeter", L"millimeters", L"mm", 0.001 },
    { L"kilometer", L"kilometers", L"km", 1000.0 },
    { L"mile", L"miles", L"mi", 1609.344 },
    { L"nautical mile", L"nautical miles", L"nmi", 1852.0 },
    { L"light-year", L"light-years", L"ly", 9460730472580.8 /* km */ * 1000.0 },
};

// Temperature conversion constants
typedef struct { const WCHAR *word; int id; } TempWord;
static const TempWord kTempWords[] = {
    { L"c", 0 }, { L"°c", 0 }, { L"celsius", 0 },
    { L"f", 1 }, { L"°f", 1 }, { L"fahrenheit", 1 },
    { L"k", 2 }, { L"kelvin", 2 }
};

// Time units
typedef enum { T_NS, T_US, T_MS, T_S, T_MIN, T_H, T_DAY, T_WEEK, T_MONTH, T_YEAR, T_INVALID } TimeUnitId;
typedef struct { const WCHAR *name_singular; const WCHAR *name_plural; const WCHAR *abbr; double to_s; } TimeUnitInfo;
static const TimeUnitInfo kTimeUnits[] = {
    { L"nanosecond",  L"nanoseconds",  L"ns",  1e-9 },
    { L"microsecond", L"microseconds", L"µs",  1e-6 },
    { L"millisecond", L"milliseconds", L"ms",  1e-3 },
    { L"second",      L"seconds",      L"s",   1.0 },
    { L"minute",      L"minutes",      L"min", 60.0 },
    { L"hour",        L"hours",        L"h",   3600.0 },
    { L"day",         L"days",         L"day", 86400.0 },
    { L"week",        L"weeks",        L"wk",  604800.0 },
    { L"month",       L"months",       L"mo",  2629746.0 }, // Average month (30.44 days)
    { L"year",        L"years",        L"yr",  31557600.0 }, // Julian year
};
typedef struct { const WCHAR *word; TimeUnitId id; } TimeUnitWord;

typedef struct {
    const WCHAR *singular;
    const WCHAR *plural;
    double to_base;
    const WCHAR *aliases[10];
} GenericUnitDef;

typedef struct {
    const WCHAR *name;
    const GenericUnitDef *units;
    int unitCount;
    BOOL distinguishCaseInSymbols;
} GenericUnitCategory;

typedef struct {
    double inputValue;
    double resultValue;
    const GenericUnitDef *sourceUnit;
    const GenericUnitDef *targetUnit;
} GenericConversionResult;

static BOOL IsDegreeChar(WCHAR c) {
    return c == L'\x00B0';
}

static BOOL IsMicroChar(WCHAR c) {
    return c == L'\x00B5' || c == L'\x03BC';
}

static const GenericUnitDef kAreaUnits[] = {
    { L"square inch", L"square inches", 0.00064516, { L"square inch", L"square inches", L"sq in", L"sqin", L"in^2", L"in2", NULL } },
    { L"square foot", L"square feet", 0.09290304, { L"square foot", L"square feet", L"sq ft", L"sqft", L"ft^2", L"ft2", NULL } },
    { L"square yard", L"square yards", 0.83612736, { L"square yard", L"square yards", L"sq yd", L"sqyd", L"yd^2", L"yd2", NULL } },
    { L"square meter", L"square meters", 1.0, { L"square meter", L"square meters", L"square metre", L"square metres", L"sq m", L"sqm", L"m^2", L"m2", NULL } },
    { L"square kilometer", L"square kilometers", 1000000.0, { L"square kilometer", L"square kilometers", L"square kilometre", L"square kilometres", L"sq km", L"sqkm", L"km^2", L"km2", NULL } },
    { L"square mile", L"square miles", 2589988.110336, { L"square mile", L"square miles", L"sq mi", L"sqmi", L"mi^2", L"mi2", NULL } },
    { L"acre", L"acres", 4046.8564224, { L"acre", L"acres", NULL } },
    { L"hectare", L"hectares", 10000.0, { L"hectare", L"hectares", L"ha", NULL } },
};

static const GenericUnitDef kWeightUnits[] = {
    { L"milligram", L"milligrams", 0.000001, { L"milligram", L"milligrams", L"mg", NULL } },
    { L"gram", L"grams", 0.001, { L"gram", L"grams", L"g", NULL } },
    { L"kilogram", L"kilograms", 1.0, { L"kilogram", L"kilograms", L"kg", NULL } },
    { L"ounce", L"ounces", 0.028349523125, { L"ounce", L"ounces", L"oz", NULL } },
    { L"pound", L"pounds", 0.45359237, { L"pound", L"pounds", L"lb", L"lbs", NULL } },
    { L"ton", L"tons", 907.18474, { L"ton", L"tons", L"short ton", NULL } },
};

static const GenericUnitDef kVolumeUnits[] = {
    { L"milliliter", L"milliliters", 0.001, { L"milliliter", L"milliliters", L"millilitre", L"millilitres", L"ml", L"mL", NULL } },
    { L"liter", L"liters", 1.0, { L"liter", L"liters", L"litre", L"litres", L"l", L"L", NULL } },
    { L"fluid ounce", L"fluid ounces", 0.0295735295625, { L"fluid ounce", L"fluid ounces", L"fl oz", L"floz", NULL } },
    { L"cup", L"cups", 0.2365882365, { L"cup", L"cups", NULL } },
    { L"pint", L"pints", 0.473176473, { L"pint", L"pints", L"pt", NULL } },
    { L"quart", L"quarts", 0.946352946, { L"quart", L"quarts", L"qt", NULL } },
    { L"gallon", L"gallons", 3.785411784, { L"gallon", L"gallons", L"gal", NULL } },
    { L"tablespoon", L"tablespoons", 0.01478676478125, { L"tablespoon", L"tablespoons", L"tbsp", NULL } },
    { L"teaspoon", L"teaspoons", 0.00492892159375, { L"teaspoon", L"teaspoons", L"tsp", NULL } },
};

static const GenericUnitDef kMemoryUnits[] = {
    { L"bit", L"bits", 0.125, { L"bit", L"bits", L"b", NULL } },
    { L"byte", L"bytes", 1.0, { L"byte", L"bytes", L"B", NULL } },
    { L"kilobit", L"kilobits", 125.0, { L"kilobit", L"kilobits", L"Kb", NULL } },
    { L"megabit", L"megabits", 125000.0, { L"megabit", L"megabits", L"Mb", NULL } },
    { L"gigabit", L"gigabits", 125000000.0, { L"gigabit", L"gigabits", L"Gb", NULL } },
    { L"terabit", L"terabits", 125000000000.0, { L"terabit", L"terabits", L"Tb", NULL } },
    { L"kilobyte", L"kilobytes", 1000.0, { L"kilobyte", L"kilobytes", L"KB", NULL } },
    { L"megabyte", L"megabytes", 1000000.0, { L"megabyte", L"megabytes", L"MB", NULL } },
    { L"gigabyte", L"gigabytes", 1000000000.0, { L"gigabyte", L"gigabytes", L"GB", NULL } },
    { L"terabyte", L"terabytes", 1000000000000.0, { L"terabyte", L"terabytes", L"TB", NULL } },
    { L"kibibyte", L"kibibytes", 1024.0, { L"kibibyte", L"kibibytes", L"KiB", NULL } },
    { L"mebibyte", L"mebibytes", 1048576.0, { L"mebibyte", L"mebibytes", L"MiB", NULL } },
    { L"gibibyte", L"gibibytes", 1073741824.0, { L"gibibyte", L"gibibytes", L"GiB", NULL } },
    { L"tebibyte", L"tebibytes", 1099511627776.0, { L"tebibyte", L"tebibytes", L"TiB", NULL } },
};

static const GenericUnitCategory kGenericUnitCategories[] = {
    { L"Area",   kAreaUnits,   (int)(sizeof(kAreaUnits)/sizeof(kAreaUnits[0])), FALSE },
    { L"Weight", kWeightUnits, (int)(sizeof(kWeightUnits)/sizeof(kWeightUnits[0])), FALSE },
    { L"Volume", kVolumeUnits, (int)(sizeof(kVolumeUnits)/sizeof(kVolumeUnits[0])), FALSE },
    { L"Memory", kMemoryUnits, (int)(sizeof(kMemoryUnits)/sizeof(kMemoryUnits[0])), TRUE },
};

static BOOL is_word_boundary(WCHAR c) {
    return c == 0 || iswspace(c) || c == L',' || c == L'.' || c == L'?' || c == L'!' || c == L';';
}

static BOOL generic_alias_uses_case_sensitive_match(const GenericUnitCategory *cat, const WCHAR *alias) {
    if (!cat || !cat->distinguishCaseInSymbols || !alias) return FALSE;
    int len = (int)wcslen(alias);
    if (len < 1 || len > 4) return FALSE;
    BOOL hasUpper = FALSE;
    for (int i = 0; i < len; ++i) {
        if (!iswalpha(alias[i])) return FALSE;
        if (iswupper(alias[i])) hasUpper = TRUE;
    }
    return hasUpper || len == 1;
}

static BOOL match_generic_alias(const WCHAR *s, const WCHAR *alias, const GenericUnitCategory *cat, int *consumed) {
    if (!alias) return FALSE;
    int len = (int)wcslen(alias);
    if (len == 0) return FALSE;
    if (generic_alias_uses_case_sensitive_match(cat, alias)) {
        if (wcsncmp(s, alias, len) != 0) return FALSE;
    } else {
        if (_wcsnicmp(s, alias, len) != 0) return FALSE;
    }
    WCHAR next = s[len];
    WCHAR last = alias[len - 1];
    if ((iswalnum(last) || last == L'%') && iswalnum(next)) return FALSE;
    if (consumed) *consumed = len;
    return TRUE;
}

static const GenericUnitDef* find_generic_unit(const WCHAR *s, const GenericUnitCategory *cat, int *consumed) {
    const GenericUnitDef *best = NULL;
    int bestLen = 0;
    for (int i = 0; i < cat->unitCount; ++i) {
        const GenericUnitDef *unit = &cat->units[i];
        const int aliasCount = ARRAY_COUNT(unit->aliases);
        for (int a = 0; a < aliasCount; ++a) {
            const WCHAR *alias = unit->aliases[a];
            if (!alias) break;
            int len = 0;
            if (match_generic_alias(s, alias, cat, &len)) {
                if (len > bestLen) {
                    best = unit;
                    bestLen = len;
                }
            }
        }
    }
    if (best && consumed) *consumed = bestLen;
    return best;
}

static void skip_spaces(const WCHAR **ps) {
    while (**ps && iswspace(**ps)) (*ps)++;
}

static BOOL consume_word(const WCHAR **ps, const WCHAR *word) {
    const WCHAR *p = *ps;
    int len = (int)wcslen(word);
    if (_wcsnicmp(p, word, len) != 0) return FALSE;
    if (!is_word_boundary(p[len])) return FALSE;
    p += len;
    *ps = p;
    return TRUE;
}

static BOOL IsKnownCompactGenericUnitToken(const WCHAR *token, int len) {
    if (!token || len <= 0) return FALSE;
    for (int i = 0; i < kGenericUnitCategoryCount; ++i) {
        const GenericUnitCategory *cat = &kGenericUnitCategories[i];
        for (int u = 0; u < cat->unitCount; ++u) {
            const GenericUnitDef *unit = &cat->units[u];
            const int aliasCount = ARRAY_COUNT(unit->aliases);
            for (int a = 0; a < aliasCount; ++a) {
                const WCHAR *alias = unit->aliases[a];
                if (!alias) break;
                if ((int)wcslen(alias) != len) continue;
                if (generic_alias_uses_case_sensitive_match(cat, alias)) {
                    if (wcsncmp(token, alias, len) == 0) return TRUE;
                } else {
                    if (_wcsnicmp(token, alias, len) == 0) return TRUE;
                }
            }
        }
    }
    return FALSE;
}

static BOOL ParseGenericUnitConversion(const WCHAR *input, GenericConversionResult *out) {
    if (!input) return FALSE;
    const WCHAR *p = input;
    skip_spaces(&p);

    static const WCHAR *prefixes[] = { L"convert", L"calculate", L"what is", L"what's", L"how many" };
    const int prefixCount = ARRAY_COUNT(prefixes);
    for (int i = 0; i < prefixCount; ++i) {
        const WCHAR *pref = prefixes[i];
        const WCHAR *temp = p;
        if (consume_word(&temp, pref)) {
            p = temp;
            skip_spaces(&p);
            break;
        }
    }

    if (*p == L'+') p++;
    WCHAR *end = NULL;
    double value = wcstod(p, &end);
    if (!end || end == p) return FALSE;
    p = end;
    skip_spaces(&p);

    const GenericUnitCategory *category = NULL;
    const GenericUnitDef *srcUnit = NULL;
    int consumed = 0;
    for (int i = 0; i < kGenericUnitCategoryCount; ++i) {
        const GenericUnitCategory *cat = &kGenericUnitCategories[i];
        const GenericUnitDef *unit = find_generic_unit(p, cat, &consumed);
        if (unit) {
            category = cat;
            srcUnit = unit;
            p += consumed;
            break;
        }
    }
    if (!category || !srcUnit) return FALSE;
    skip_spaces(&p);

    if (_wcsnicmp(p, L"per", 3) == 0 && is_word_boundary(p[3])) {
        return FALSE;
    }

    if (_wcsnicmp(p, L"to", 2) == 0 && is_word_boundary(p[2])) {
        p += 2;
    } else if (_wcsnicmp(p, L"into", 4) == 0 && is_word_boundary(p[4])) {
        p += 4;
    } else if (_wcsnicmp(p, L"in", 2) == 0 && is_word_boundary(p[2])) {
        p += 2;
    } else if (p[0] == L'-' && p[1] == L'>') {
        p += 2;
    } else if (_wcsnicmp(p, L"as", 2) == 0 && is_word_boundary(p[2])) {
        p += 2;
    } else {
        return FALSE;
    }
    skip_spaces(&p);

    if (_wcsnicmp(p, L"the", 3) == 0 && is_word_boundary(p[3])) {
        p += 3;
        skip_spaces(&p);
    }
    if (_wcsnicmp(p, L"a", 1) == 0 && is_word_boundary(p[1])) {
        p += 1;
        skip_spaces(&p);
    }

    const GenericUnitDef *dstUnit = find_generic_unit(p, category, &consumed);
    if (!dstUnit) return FALSE;
    p += consumed;
    skip_spaces(&p);

    while (*p) {
        if (*p == L'.' || *p == L'?' || *p == L'!') {
            p++;
            skip_spaces(&p);
            continue;
        }
        return FALSE;
    }

    double baseValue = value * srcUnit->to_base;
    double resultValue = baseValue / dstUnit->to_base;
    if (out) {
        out->inputValue = value;
        out->resultValue = resultValue;
        out->sourceUnit = srcUnit;
        out->targetUnit = dstUnit;
    }
    return TRUE;
}

static const WCHAR* pick_unit_name(const GenericUnitDef *unit, double value) {
    if (!unit) return L"";
    return (fabs(value - 1.0) < 1e-9) ? unit->singular : unit->plural;
}

static BOOL HandleGenericUnitConversion(HWND hwnd, const WCHAR *input) {
    GenericConversionResult conv;
    if (!ParseGenericUnitConversion(input, &conv)) return FALSE;

    WCHAR valueStr[64]; format_double(conv.inputValue, valueStr, 64);
    WCHAR resultStr[64]; format_double(conv.resultValue, resultStr, 64);
    WCHAR resultFull[64]; format_double_full(conv.resultValue, resultFull, 64);

    const WCHAR *srcName = pick_unit_name(conv.sourceUnit, conv.inputValue);
    const WCHAR *dstName = pick_unit_name(conv.targetUnit, conv.resultValue);

    WCHAR line[512];
    swprintf(line, sizeof(line)/sizeof(line[0]), L"%s %s = %s %s", valueStr, srcName, resultStr, dstName);

    History_Add(hwnd, line, resultFull, resultFull, TRUE);
    return TRUE;
}

static const TimeUnitWord kTimeUnitWords[] = {
    { L"ns", T_NS }, { L"nanosecond", T_NS }, { L"nanoseconds", T_NS },
    { L"ms", T_MS }, { L"millisecond", T_MS }, { L"milliseconds", T_MS },
    { L"s", T_S }, { L"sec", T_S }, { L"secs", T_S }, { L"second", T_S }, { L"seconds", T_S },
    { L"min", T_MIN }, { L"mins", T_MIN }, { L"minute", T_MIN }, { L"minutes", T_MIN },
    { L"h", T_H }, { L"hr", T_H }, { L"hour", T_H }, { L"hours", T_H },
    { L"d", T_DAY }, { L"day", T_DAY }, { L"days", T_DAY },
    { L"wk", T_WEEK }, { L"week", T_WEEK }, { L"weeks", T_WEEK },
    { L"mo", T_MONTH }, { L"month", T_MONTH }, { L"months", T_MONTH },
    { L"yr", T_YEAR }, { L"year", T_YEAR }, { L"years", T_YEAR }
};

// Unit word mappings
typedef struct { const WCHAR *word; UnitId id; } UnitWord;
static const UnitWord kUnitWords[] = {
    { L"in", U_IN }, { L"inch", U_IN }, { L"inches", U_IN },
    { L"ft", U_FT }, { L"f", U_FT }, { L"foot", U_FT }, { L"feet", U_FT },
    { L"yd", U_YD }, { L"yard", U_YD }, { L"yards", U_YD },
    { L"m", U_M }, { L"meter", U_M }, { L"meters", U_M }, { L"metre", U_M }, { L"metres", U_M },
    { L"cm", U_CM }, { L"centimeter", U_CM }, { L"centimeters", U_CM }, { L"centimetre", U_CM }, { L"centimetres", U_CM },
    { L"mm", U_MM }, { L"millimeter", U_MM }, { L"millimeters", U_MM },
    { L"km", U_KM }, { L"kilometer", U_KM }, { L"kilometers", U_KM }, { L"kilometre", U_KM }, { L"kilometres", U_KM },
    { L"mi", U_MI }, { L"mile", U_MI }, { L"miles", U_MI },
    { L"nmi", U_NMI }, { L"nauticalmile", U_NMI }, { L"nauticalmiles", U_NMI }, { L"nm", U_NMI },
    { L"ly", U_LY }, { L"lightyear", U_LY }, { L"lightyears", U_LY }
};

// Helpers to detect last-mentioned unit in a snippet
static UnitId find_last_length_unit_in_text(const WCHAR *s) {
    if (!s) return U_INVALID;
    UnitId found = U_INVALID;
    for (const WCHAR *p = s; *p; ) {
        if (iswalpha(*p) || IsMicroChar(*p)) {
            const WCHAR *start = p;
            while (*p && (iswalpha(*p) || IsMicroChar(*p))) p++;
            int len = (int)(p - start);
            if (len > 0) {
                int outLen = 0;
                UnitId uid = match_unit_word(start, &outLen);
                if (uid != U_INVALID && outLen == len) {
                    found = uid;
                }
            }
        } else {
            p++;
        }
    }
    return found;
}

static int find_last_temp_unit_in_text(const WCHAR *s) {
    if (!s) return -1;
    int found = -1;
    for (const WCHAR *p = s; *p; ) {
        if (iswalpha(*p) || IsDegreeChar(*p)) {
            const WCHAR *start = p;
            while (*p && (iswalpha(*p) || IsDegreeChar(*p))) p++;
            int len = (int)(p - start);
            if (len > 0) {
                int outLen = 0;
                int tid = match_temp_unit_word(start, &outLen);
                if (tid >= 0 && outLen == len) {
                    found = tid;
                }
            }
        } else {
            p++;
        }
    }
    return found;
}

// Helper function to extract and match time unit word (reduces code duplication)
static TimeUnitId match_time_unit_from_word(const WCHAR *word, int len) {
    if (len <= 0 || len >= MAX_SMALL_BUFFER) return T_INVALID;
    
    WCHAR buf[MAX_SMALL_BUFFER];
    for (int i = 0; i < len && i < MAX_SMALL_BUFFER - 1; i++) {
        buf[i] = (WCHAR)towlower(word[i]);
    }
    buf[len] = 0;
    
    const int timeUnitWordCount = ARRAY_COUNT(kTimeUnitWords);
    for (int k = 0; k < timeUnitWordCount; ++k) {
        if (wcsieq(buf, kTimeUnitWords[k].word)) {
            return kTimeUnitWords[k].id;
        }
    }
    return T_INVALID;
}

static TimeUnitId find_last_time_unit_in_text(const WCHAR *s) {
    if (!s) return T_INVALID;
    TimeUnitId found = T_INVALID;
    for (const WCHAR *p = s; *p; ) {
        if (iswalpha(*p) || IsMicroChar(*p)) {
            const WCHAR *start = p;
            while (*p && (iswalpha(*p) || IsMicroChar(*p))) p++;
            int len = (int)(p - start);
            if (len > 0) {
                TimeUnitId id = match_time_unit_from_word(start, len);
                if (id != T_INVALID) found = id;
            }
        } else {
            p++;
        }
    }
    return found;
}

// Mixed unit arithmetic support
typedef enum {
    UNIT_SYSTEM_IMPERIAL,  // ft, in, yd
    UNIT_SYSTEM_METRIC,    // m, cm, mm
    UNIT_SYSTEM_MIXED      // when both systems are present
} UnitSystem;

typedef struct {
    double value;           // result in base meters
    UnitSystem inputSystem; // what system was used in input
    UnitId primaryUnit;     // best unit for the result
    BOOL hasExplicitTarget; // user specified "in cm" etc
    UnitId targetUnit;      // user's requested unit
    BOOL useCompound;       // show as "ft + in" format
} MixedUnitResult;

// Forward declarations for functions used by mixed unit formatting
static double unit_to_m(UnitId id);
static void format_double(double v, WCHAR *buf, size_t cap);
static void unit_name_full(UnitId id, double value, WCHAR *buf, size_t cap);

static int wcsieq(const WCHAR *a, const WCHAR *b) { return _wcsicmp(a, b) == 0; }

// Smart time formatting - shows time in the most appropriate units
static void format_time_smart(double seconds, WCHAR *buf, size_t cap) {
    if (cap < 2) return;
    
    // Handle negative values
    BOOL negative = (seconds < 0);
    if (negative) seconds = -seconds;
    
    if (seconds == 0) {
        wcsncpy(buf, L"0 seconds", cap - 1);
        buf[cap - 1] = 0;
        return;
    }
    
    // Define thresholds for each unit (when to prefer that unit)
    const double YEAR_THRESHOLD = 365.25 * 24 * 3600;   // 1 year
    const double MONTH_THRESHOLD = 30.44 * 24 * 3600;   // 1 month  
    const double WEEK_THRESHOLD = 7 * 24 * 3600;        // 1 week
    const double DAY_THRESHOLD = 24 * 3600;             // 1 day
    const double HOUR_THRESHOLD = 3600;                 // 1 hour
    const double MINUTE_THRESHOLD = 60;                 // 1 minute
    
    WCHAR result[256] = {0};
    int pos = 0;
    
    if (negative) {
        result[pos++] = L'-';
    }
    
    // For very large values, show in multiple units (e.g., "3 years, 2 months")
    if (seconds >= YEAR_THRESHOLD) {
        double years = floor(seconds / YEAR_THRESHOLD);
        seconds -= years * YEAR_THRESHOLD;
        
        pos += swprintf(result + pos, 256 - pos, L"%.0f year%s", years, (years == 1) ? L"" : L"s");
        
        if (seconds >= MONTH_THRESHOLD) {
            double months = floor(seconds / MONTH_THRESHOLD);
            seconds -= months * MONTH_THRESHOLD;
            pos += swprintf(result + pos, 256 - pos, L", %.0f month%s", months, (months == 1) ? L"" : L"s");
        } else if (seconds >= WEEK_THRESHOLD) {
            double weeks = floor(seconds / WEEK_THRESHOLD);
            pos += swprintf(result + pos, 256 - pos, L", %.0f week%s", weeks, (weeks == 1) ? L"" : L"s");
        }
    } else if (seconds >= MONTH_THRESHOLD) {
        double months = floor(seconds / MONTH_THRESHOLD);
        seconds -= months * MONTH_THRESHOLD;
        
        pos += swprintf(result + pos, 256 - pos, L"%.0f month%s", months, (months == 1) ? L"" : L"s");
        
        if (seconds >= WEEK_THRESHOLD) {
            double weeks = floor(seconds / WEEK_THRESHOLD);
            seconds -= weeks * WEEK_THRESHOLD;
            pos += swprintf(result + pos, 256 - pos, L", %.0f week%s", weeks, (weeks == 1) ? L"" : L"s");
        } else if (seconds >= DAY_THRESHOLD) {
            double days = floor(seconds / DAY_THRESHOLD);
            pos += swprintf(result + pos, 256 - pos, L", %.0f day%s", days, (days == 1) ? L"" : L"s");
        }
    } else if (seconds >= WEEK_THRESHOLD) {
        double weeks = floor(seconds / WEEK_THRESHOLD);
        seconds -= weeks * WEEK_THRESHOLD;
        
        pos += swprintf(result + pos, 256 - pos, L"%.0f week%s", weeks, (weeks == 1) ? L"" : L"s");
        
        if (seconds >= DAY_THRESHOLD) {
            double days = floor(seconds / DAY_THRESHOLD);
            pos += swprintf(result + pos, 256 - pos, L", %.0f day%s", days, (days == 1) ? L"" : L"s");
        }
    } else if (seconds >= DAY_THRESHOLD) {
        double days = seconds / DAY_THRESHOLD;
        if (days == floor(days)) {
            pos += swprintf(result + pos, 256 - pos, L"%.0f day%s", days, (days == 1) ? L"" : L"s");
        } else {
            pos += swprintf(result + pos, 256 - pos, L"%.1f day%s", days, L"s");
        }
    } else if (seconds >= HOUR_THRESHOLD) {
        double hours = seconds / HOUR_THRESHOLD;
        if (hours == floor(hours)) {
            pos += swprintf(result + pos, 256 - pos, L"%.0f hour%s", hours, (hours == 1) ? L"" : L"s");
        } else {
            pos += swprintf(result + pos, 256 - pos, L"%.1f hour%s", hours, L"s");
        }
    } else if (seconds >= MINUTE_THRESHOLD) {
        double minutes = seconds / MINUTE_THRESHOLD;
        if (minutes == floor(minutes)) {
            pos += swprintf(result + pos, 256 - pos, L"%.0f minute%s", minutes, (minutes == 1) ? L"" : L"s");
        } else {
            pos += swprintf(result + pos, 256 - pos, L"%.1f minute%s", minutes, L"s");
        }
    } else {
        // Less than a minute - show in seconds
        if (seconds == floor(seconds)) {
            pos += swprintf(result + pos, 256 - pos, L"%.0f second%s", seconds, (seconds == 1) ? L"" : L"s");
        } else {
            pos += swprintf(result + pos, 256 - pos, L"%.2f second%s", seconds, L"s");
        }
    }
    
    wcsncpy(buf, result, cap - 1);
    buf[cap - 1] = 0;
}

// Holiday support
struct Holiday {
    const WCHAR *name;
    int month;    // 1-12
    int day;      // 1-31, or special codes for relative dates
    int weekday;  // 0-6 (Sunday=0), -1 for fixed dates
    int week;     // 1-5 for "nth weekday", negative for "last weekday"
};

typedef struct {
    WCHAR lower[64];
    const Holiday *holiday;
} HolidayLookup;

static HolidayLookup g_holidayLookup[32];  // Extra space for aliases
static int g_holidayLookupCount = 0;

static const Holiday kHolidays[] = {
    // Fixed date holidays
    { L"New Year's Day", 1, 1, -1, 0 },
    { L"Valentine's Day", 2, 14, -1, 0 },
    { L"Independence Day", 7, 4, -1, 0 },
    { L"Halloween", 10, 31, -1, 0 },
    { L"Veterans Day", 11, 11, -1, 0 },
    { L"Christmas Eve", 12, 24, -1, 0 },
    { L"Christmas Day", 12, 25, -1, 0 },
    { L"New Year's Eve", 12, 31, -1, 0 },
    
    // Relative date holidays (US)
    { L"Martin Luther King Jr. Day", 1, 0, 1, 3 },     // 3rd Monday in January
    { L"Presidents' Day", 2, 0, 1, 3 },                // 3rd Monday in February  
    { L"Memorial Day", 5, 0, 1, -1 },                  // Last Monday in May
    { L"Labor Day", 9, 0, 1, 1 },                      // 1st Monday in September
    { L"Columbus Day", 10, 0, 1, 2 },                  // 2nd Monday in October
    { L"Thanksgiving", 11, 0, 4, 4 },                  // 4th Thursday in November
    
    // Mother's Day, Father's Day
    { L"Mother's Day", 5, 0, 0, 2 },                   // 2nd Sunday in May
    { L"Father's Day", 6, 0, 0, 3 },                   // 3rd Sunday in June
};

static int calculate_holiday_date(const Holiday *h, int year, SYSTEMTIME *out) {
    if (!h || !out) return 0;
    
    memset(out, 0, sizeof(*out));
    out->wYear = (WORD)year;
    out->wMonth = (WORD)h->month;
    
    if (h->weekday == -1) {
        // Fixed date
        out->wDay = (WORD)h->day;
        return 1;
    }
    
    // Relative date - find the nth occurrence of weekday in month
    SYSTEMTIME temp = {0};
    temp.wYear = (WORD)year;
    temp.wMonth = (WORD)h->month;
    temp.wDay = 1;
    
    FILETIME ft;
    if (!SystemTimeToFileTime(&temp, &ft)) return 0;
    if (!FileTimeToSystemTime(&ft, &temp)) return 0;
    
    int firstWeekday = temp.wDayOfWeek;  // 0=Sunday
    int targetWeekday = h->weekday;
    
    if (h->week > 0) {
        // Find nth occurrence
        int dayOffset = (targetWeekday - firstWeekday + 7) % 7;
        int targetDay = 1 + dayOffset + (h->week - 1) * 7;
        
        // Check if date is valid for this month
        int daysInMonth = 31;
        if (h->month == 2) {
            daysInMonth = ((year % 4 == 0 && year % 100 != 0) || (year % 400 == 0)) ? 29 : 28;
        } else if (h->month == 4 || h->month == 6 || h->month == 9 || h->month == 11) {
            daysInMonth = 30;
        }
        
        if (targetDay > daysInMonth) return 0;
        out->wDay = (WORD)targetDay;
    } else {
        // Find last occurrence (h->week is negative)
        int daysInMonth = 31;
        if (h->month == 2) {
            daysInMonth = ((year % 4 == 0 && year % 100 != 0) || (year % 400 == 0)) ? 29 : 28;
        } else if (h->month == 4 || h->month == 6 || h->month == 9 || h->month == 11) {
            daysInMonth = 30;
        }
        
        // Find the last occurrence by working backwards
        for (int day = daysInMonth; day >= 1; day--) {
            temp.wDay = (WORD)day;
            if (SystemTimeToFileTime(&temp, &ft) && FileTimeToSystemTime(&ft, &temp)) {
                if (temp.wDayOfWeek == targetWeekday) {
                    out->wDay = (WORD)day;
                    return 1;
                }
            }
        }
        return 0;
    }
    
    return 1;
}

// Helper to add a holiday name to the lookup table
static void AddHolidayLookup(const WCHAR *name, const Holiday *holiday) {
    if (g_holidayLookupCount >= 32) return;
    
    int i = 0;
    while (name[i] && i < 63) {
        g_holidayLookup[g_holidayLookupCount].lower[i] = (WCHAR)towlower(name[i]);
        i++;
    }
    g_holidayLookup[g_holidayLookupCount].lower[i] = 0;
    g_holidayLookup[g_holidayLookupCount].holiday = holiday;
    g_holidayLookupCount++;
}

// Initialize holiday lookup table with pre-computed lowercase names and aliases
static void InitializeHolidayLookup(void) {
    if (g_holidayLookupCount > 0) return;  // Already initialized
    
    const int holidayCount = ARRAY_COUNT(kHolidays);
    for (int i = 0; i < holidayCount; i++) {
        const Holiday *h = &kHolidays[i];
        // Add primary name
        AddHolidayLookup(h->name, h);
        
        // Add common aliases
        if (wcscmp(h->name, L"Christmas Day") == 0) {
            AddHolidayLookup(L"christmas", h);
            AddHolidayLookup(L"xmas", h);
        } else if (wcscmp(h->name, L"Christmas Eve") == 0) {
            AddHolidayLookup(L"xmas eve", h);
        } else if (wcscmp(h->name, L"New Year's Day") == 0) {
            AddHolidayLookup(L"new year", h);
            AddHolidayLookup(L"newyear", h);
            AddHolidayLookup(L"new years day", h);
        } else if (wcscmp(h->name, L"New Year's Eve") == 0) {
            AddHolidayLookup(L"new years eve", h);
        } else if (wcscmp(h->name, L"Martin Luther King Jr. Day") == 0) {
            AddHolidayLookup(L"martin luther king day", h);
            AddHolidayLookup(L"mlk day", h);
        } else if (wcscmp(h->name, L"Presidents' Day") == 0) {
            AddHolidayLookup(L"presidents day", h);
            AddHolidayLookup(L"president's day", h);
        } else if (wcscmp(h->name, L"Thanksgiving") == 0) {
            AddHolidayLookup(L"thanksgiving", h);
        } else if (wcscmp(h->name, L"Halloween") == 0) {
            AddHolidayLookup(L"halloween", h);
        } else if (wcscmp(h->name, L"Valentine's Day") == 0) {
            AddHolidayLookup(L"valentines day", h);
        } else if (wcscmp(h->name, L"Mother's Day") == 0) {
            AddHolidayLookup(L"mothers day", h);
        } else if (wcscmp(h->name, L"Father's Day") == 0) {
            AddHolidayLookup(L"fathers day", h);
        } else if (wcscmp(h->name, L"Veterans Day") == 0) {
            AddHolidayLookup(L"veterans day", h);
        }
    }
}

static BOOL find_holiday_by_name(const WCHAR *name, const Holiday **outHoliday) {
    if (!name || !outHoliday) return FALSE;
    
    // Ensure lookup table is initialized
    InitializeHolidayLookup();
    
    // Convert input to lowercase for comparison
    WCHAR lowerName[64];
    int i = 0;
    while (name[i] && i < 63) {
        lowerName[i] = (WCHAR)towlower(name[i]);
        i++;
    }
    lowerName[i] = 0;
    
    // Single pass through pre-computed lookup table
    for (int j = 0; j < g_holidayLookupCount; j++) {
        if (wcscmp(lowerName, g_holidayLookup[j].lower) == 0) {
            *outHoliday = g_holidayLookup[j].holiday;
            return TRUE;
        }
    }
    
    return FALSE;
}

static int CompareCalendarDate(const SYSTEMTIME *a, const SYSTEMTIME *b) {
    if (a->wYear != b->wYear) return (a->wYear < b->wYear) ? -1 : 1;
    if (a->wMonth != b->wMonth) return (a->wMonth < b->wMonth) ? -1 : 1;
    if (a->wDay != b->wDay) return (a->wDay < b->wDay) ? -1 : 1;
    return 0;
}

static int DaysFromCivilDate(int year, int month, int day) {
    year -= month <= 2;
    {
        int era = (year >= 0 ? year : year - 399) / 400;
        unsigned yoe = (unsigned)(year - era * 400);
        unsigned mp = (unsigned)(month + (month > 2 ? -3 : 9));
        unsigned doy = (153U * mp + 2U) / 5U + (unsigned)day - 1U;
        unsigned doe = yoe * 365U + yoe / 4U - yoe / 100U + doy;
        return era * 146097 + (int)doe;
    }
}

static int DaysBetweenCalendarDates(const SYSTEMTIME *fromDate, const SYSTEMTIME *toDate) {
    int fromDays = DaysFromCivilDate(fromDate->wYear, fromDate->wMonth, fromDate->wDay);
    int toDays = DaysFromCivilDate(toDate->wYear, toDate->wMonth, toDate->wDay);
    return toDays - fromDays;
}

static BOOL ResolveHolidayDate(const Holiday *holiday, BOOL hasExplicitYear, int targetYear, const SYSTEMTIME *nowLocal, SYSTEMTIME *holidayDateOut, int *resolvedYearOut, BOOL *isTodayOut) {
    SYSTEMTIME holidayDate;
    int resolvedYear = targetYear;
    BOOL isToday = FALSE;

    if (!holiday || !nowLocal || !holidayDateOut) return FALSE;

    if (hasExplicitYear) {
        if (!calculate_holiday_date(holiday, targetYear, &holidayDate)) return FALSE;
    } else {
        resolvedYear = nowLocal->wYear;
        if (!calculate_holiday_date(holiday, resolvedYear, &holidayDate)) {
            resolvedYear++;
            if (!calculate_holiday_date(holiday, resolvedYear, &holidayDate)) return FALSE;
        } else {
            int cmp = CompareCalendarDate(&holidayDate, nowLocal);
            if (cmp < 0) {
                resolvedYear++;
                if (!calculate_holiday_date(holiday, resolvedYear, &holidayDate)) return FALSE;
            } else if (cmp == 0) {
                isToday = TRUE;
            }
        }
    }

    if (!isToday) {
        isToday = (CompareCalendarDate(&holidayDate, nowLocal) == 0);
    }

    *holidayDateOut = holidayDate;
    if (resolvedYearOut) *resolvedYearOut = resolvedYear;
    if (isTodayOut) *isTodayOut = isToday;
    return TRUE;
}

static BOOL FindHolidayForToday(const SYSTEMTIME *nowLocal, const Holiday **holidayOut, SYSTEMTIME *holidayDateOut) {
    int holidayCount = ARRAY_COUNT(kHolidays);
    if (!nowLocal) return FALSE;
    for (int i = 0; i < holidayCount; ++i) {
        SYSTEMTIME holidayDate;
        if (calculate_holiday_date(&kHolidays[i], nowLocal->wYear, &holidayDate) &&
            CompareCalendarDate(&holidayDate, nowLocal) == 0) {
            if (holidayOut) *holidayOut = &kHolidays[i];
            if (holidayDateOut) *holidayDateOut = holidayDate;
            return TRUE;
        }
    }
    return FALSE;
}

static BOOL ParseHolidayDateToken(const WCHAR *s, int *consumed, SYSTEMTIME *out) {
    int i = 0;
    const Holiday *holiday = NULL;
    WCHAR nameBuf[64];
    WCHAR work[96];
    int workLen = 0;
    BOOL hasExplicitYear = FALSE;
    int targetYear = 0;

    if (!s || !*s) return FALSE;

    while (s[i] == L' ') i++;
    while (s[i + workLen] && s[i + workLen] != L',' && s[i + workLen] != L';') {
        if (workLen >= (int)ARRAY_COUNT(work) - 1) break;
        work[workLen] = s[i + workLen];
        workLen++;
    }
    work[workLen] = 0;
    while (workLen > 0 && iswspace(work[workLen - 1])) work[--workLen] = 0;
    if (workLen == 0) return FALSE;

    for (int len = workLen; len > 0; ) {
        int candidateConsumed = len;
        while (len > 0 && iswspace(work[len - 1])) work[--len] = 0;
        if (len <= 0) break;
        candidateConsumed = len;

        {
            int digitEnd = len;
            int digitStart = digitEnd;
            while (digitStart > 0 && is_digit(work[digitStart - 1])) digitStart--;
            if (digitStart < digitEnd && (digitEnd - digitStart == 2 || digitEnd - digitStart == 4) &&
                (digitStart == 0 || iswspace(work[digitStart - 1]))) {
                int parsedYear = 0;
                int digitCount = digitEnd - digitStart;
                for (int idx = digitStart; idx < digitEnd; ++idx) parsedYear = parsedYear * 10 + (work[idx] - L'0');
                if (digitCount == 2) parsedYear += 2000;
                if (parsedYear >= 1601) {
                    hasExplicitYear = TRUE;
                    targetYear = parsedYear;
                    candidateConsumed = digitEnd;
                    work[digitStart] = 0;
                    len = digitStart;
                    while (len > 0 && iswspace(work[len - 1])) work[--len] = 0;
                }
            }
        }

        if (len <= 0) break;
        if (len >= (int)ARRAY_COUNT(nameBuf)) len = (int)ARRAY_COUNT(nameBuf) - 1;
        wcsncpy(nameBuf, work, len);
        nameBuf[len] = 0;
        while (len > 0 && iswspace(nameBuf[len - 1])) nameBuf[--len] = 0;
        if (len <= 0) break;

        if (find_holiday_by_name(nameBuf, &holiday)) {
            SYSTEMTIME nowLocal;
            SYSTEMTIME holidayDate;
            GetLocalTime(&nowLocal);
            nowLocal.wHour = nowLocal.wMinute = nowLocal.wSecond = nowLocal.wMilliseconds = 0;
            if (!ResolveHolidayDate(holiday, hasExplicitYear, hasExplicitYear ? targetYear : nowLocal.wYear, &nowLocal, &holidayDate, NULL, NULL)) {
                return FALSE;
            }
            if (out) *out = holidayDate;
            if (consumed) *consumed = i + candidateConsumed;
            return TRUE;
        }

        while (len > 0 && !iswspace(work[len - 1])) len--;
        work[len] = 0;
    }

    return FALSE;
}

// Mixed unit system detection and formatting
static UnitSystem detect_unit_system(const WCHAR *expr) {
    BOOL hasImperial = FALSE, hasMetric = FALSE;
    
    for (const WCHAR *p = expr; *p; p++) {
        if (iswalpha(*p)) {
            // Extract word
            const WCHAR *start = p;
            while (*p && iswalpha(*p)) p++;
            int len = (int)(p - start);
            WCHAR word[32];
            if (len < 32) {
                wcsncpy(word, start, len);
                word[len] = 0;
                
                // Convert to lowercase for comparison
                for (int i = 0; word[i]; i++) {
                    word[i] = (WCHAR)towlower(word[i]);
                }
                
                // Check against known units
                const int unitWordCount = ARRAY_COUNT(kUnitWords);
                for (int i = 0; i < unitWordCount; i++) {
                    if (wcscmp(word, kUnitWords[i].word) == 0) {
                        UnitId id = kUnitWords[i].id;
                        if (id == U_FT || id == U_IN || id == U_YD || id == U_MI || id == U_NMI) {
                            hasImperial = TRUE;
                        } else if (id == U_M || id == U_CM || id == U_MM || id == U_KM) {
                            hasMetric = TRUE;
                        }
                        break;
                    }
                }
            }
            p--; // Back up since loop will increment
        }
    }
    
    if (hasImperial && hasMetric) return UNIT_SYSTEM_MIXED;
    if (hasImperial) return UNIT_SYSTEM_IMPERIAL;
    if (hasMetric) return UNIT_SYSTEM_METRIC;
    return UNIT_SYSTEM_IMPERIAL; // default
}

static BOOL is_clean_inches(double inches) {
    // Check if inches value is "clean" for compound display
    double rounded = floor(inches + 0.5);
    return fabs(inches - rounded) < 0.01; // Within 0.01 inch
}

static void format_imperial_result(MixedUnitResult *result, WCHAR *output, size_t cap) {
    double totalInches = result->value / unit_to_m(U_IN);
    
    // Try compound feet + inches first
    if (totalInches >= 12.0) {
        int feet = (int)(totalInches / 12.0);
        double remainingInches = totalInches - (feet * 12.0);
        
        // Use compound if remaining inches are "clean"
        if (is_clean_inches(remainingInches)) {
            if (remainingInches < 0.5) {
                swprintf(output, cap, L"%d feet", feet);
            } else {
                swprintf(output, cap, L"%d ft %.0f in", feet, remainingInches);
            }
            return;
        }
    }
    
    // Fall back to decimal feet for larger values, inches for smaller
    if (totalInches >= 12.0) {
        double feet = totalInches / 12.0;
        WCHAR num[64]; format_double(feet, num, 64);
        swprintf(output, cap, L"%s feet", num);
    } else {
        WCHAR num[64]; format_double(totalInches, num, 64);
        swprintf(output, cap, L"%s inches", num);
    }
}

static void format_metric_result(MixedUnitResult *result, WCHAR *output, size_t cap) {
    double meters = result->value;
    
    if (meters >= 1.0) {
        WCHAR num[64]; format_double(meters, num, 64);
        swprintf(output, cap, L"%s meters", num);
    } else if (meters >= 0.01) {
        double cm = meters / unit_to_m(U_CM);
        WCHAR num[64]; format_double(cm, num, 64);
        swprintf(output, cap, L"%s centimeters", num);
    } else {
        double mm = meters / unit_to_m(U_MM);
        WCHAR num[64]; format_double(mm, num, 64);
        swprintf(output, cap, L"%s millimeters", num);
    }
}

static void format_mixed_system_result(MixedUnitResult *result, WCHAR *output, size_t cap) {
    // When both systems are present, show result in both
    WCHAR primary[128];
    
    // Use imperial as primary for mixed (common in US)
    format_imperial_result(result, primary, 128);
    
    // Add metric equivalent
    double cm = result->value / unit_to_m(U_CM);
    swprintf(output, cap, L"%s (%.1f cm)", primary, cm);
}

static void format_mixed_result(MixedUnitResult *result, WCHAR *output, size_t cap) {
    if (result->hasExplicitTarget) {
        // User specified target - always honor it
        double targetValue = result->value / unit_to_m(result->targetUnit);
        WCHAR num[64]; format_double(targetValue, num, 64);
        WCHAR unitName[32]; unit_name_full(result->targetUnit, targetValue, unitName, 32);
        swprintf(output, cap, L"%s %s", num, unitName);
        return;
    }
    
    switch (result->inputSystem) {
        case UNIT_SYSTEM_IMPERIAL:
            format_imperial_result(result, output, cap);
            break;
        case UNIT_SYSTEM_METRIC:
            format_metric_result(result, output, cap);
            break;
        case UNIT_SYSTEM_MIXED:
            format_mixed_system_result(result, output, cap);
            break;
    }
}

static UnitId match_unit_word(const WCHAR *s, int *outLen) {
    // Recognize quotes first: ' and " and '' (two single quotes) as feet
    if (*s == L'"') {
        if (outLen) *outLen = 1;
        return U_IN;
    }
    if (*s == L'\'') {
        if (s[1] == L'\'') {
            if (outLen) *outLen = 2;
            return U_FT;
        }
        if (outLen) *outLen = 1;
        return U_FT;
    }
    
    // Match alpha unit words and plurals/aliases - require exact word boundaries
    WCHAR buf[32]; 
    int i = 0; 
    while (s[i] && iswalpha(s[i]) && i < (int)(sizeof(buf)/sizeof(buf[0])) - 1) { 
        buf[i] = (WCHAR)towlower(s[i]); 
        i++; 
    }
    buf[i] = 0;
    
    if (i == 0) return U_INVALID;
    
    // Check for exact word match in our known units
    const int unitWordCount = ARRAY_COUNT(kUnitWords);
    for (int k = 0; k < unitWordCount; ++k) {
        if (wcsieq(buf, kUnitWords[k].word)) { 
            if (outLen) *outLen = i; 
            return kUnitWords[k].id; 
        }
    }
    
    // If we found letters but no exact match, it's invalid (like "ftt")
    return U_INVALID;
}

// Temperature units: returns 0=C,1=F,2=K, or -1 if not matched
static int match_temp_unit_word(const WCHAR *s, int *outLen) {
    WCHAR buf[32]; 
    int i = 0; 
    while (s[i] && (iswalpha(s[i]) || IsDegreeChar(s[i])) && i < (int)(sizeof(buf)/sizeof(buf[0])) - 1) { 
        buf[i] = (WCHAR)towlower(s[i]); 
        i++; 
    }
    buf[i] = 0;
    
    if (i == 0) return -1;
    
    const int tempWordCount = ARRAY_COUNT(kTempWords);
    for (int k = 0; k < tempWordCount; ++k) {
        if (wcsieq(buf, kTempWords[k].word)) { 
            if (outLen) *outLen = i; 
            return kTempWords[k].id; 
        }
    }
    return -1;
}

static double temp_to_kelvin(double v, int unitId) {
    switch (unitId) { // 0=C,1=F,2=K
        case 0: return v + 273.15;
        case 1: return (v - 32.0) * (5.0/9.0) + 273.15;
        case 2: return v;
        default: return v;
    }
}

static double kelvin_to_temp(double k, int unitId) {
    switch (unitId) { // 0=C,1=F,2=K
        case 0: return k - 273.15;
        case 1: return (k - 273.15) * (9.0/5.0) + 32.0;
        case 2: return k;
        default: return k;
    }
}

static const WCHAR* temp_symbol(int unitId) {
    switch (unitId) {
        case 0: return L"°C";
        case 1: return L"°F";
        case 2: return L"K";
        default: return L"";
    }
}

// Small internal usage removed; temp helpers may be referenced where needed.

static double unit_to_m(UnitId id) {
    if (id >= 0 && id < U_INVALID) {
        return kUnits[id].to_m;
    }
    return 1.0;
}

static const UnitInfo* unit_info(UnitId id) { return (id >= 0 && id < U_INVALID) ? &kUnits[id] : NULL; }

static void unit_name_full(UnitId id, double value, WCHAR *buf, size_t cap) {
    const UnitInfo *ui = unit_info(id);
    if (!ui || !buf || cap == 0) return;
    const WCHAR *name = (fabs(value - 1.0) < 1e-12) ? ui->name_singular : ui->name_plural;
    wcsncpy(buf, name, cap - 1); buf[cap-1] = 0;
}

static void format_unit_value(UnitId id, double value, WCHAR *buf, size_t cap) {
    if (!buf || cap == 0) return;
    // Unit-aware rounding: inches/cm/mm -> 2 decimals; meters/feet/yards/km/mi/nmi -> 3; ly -> 6
    int decimals = 2;
    switch (id) {
        case U_IN: case U_CM: case U_MM: decimals = 2; break;
        case U_FT: case U_YD: case U_M: case U_KM: case U_MI: case U_NMI:  decimals = 3; break;
        case U_LY: decimals = 6; break;
        default: decimals = 3; break;
    }
    // Build a printf format like %.2f, then trim trailing zeros and dot
    WCHAR fmt[16]; swprintf(fmt, 16, L"%%.%df", decimals);
    swprintf(buf, cap, fmt, value);
    // Trim trailing zeros
    size_t n = wcslen(buf);
    while (n > 0 && buf[n-1] == L'0') { buf[--n] = 0; }
    if (n > 0 && buf[n-1] == L'.') { buf[--n] = 0; }
    if (n == 0) { wcsncpy(buf, L"0", cap - 1); buf[cap-1] = 0; }
}

typedef struct UnitsPre {
    BOOL hasUnits;
    BOOL hasTarget;
    UnitId target;
    WCHAR eval[1024];
    WCHAR prettyLeft[1024];
    // Temperature support
    BOOL hasTemp;
    BOOL hasTempTarget;
    int  tempTarget; // 0=C,1=F,2=K
    WCHAR prettyLeftFull[1024];
    // Mixed unit support
    BOOL hasMixedUnits;
    UnitSystem unitSystem;
    double mixedResult; // Final result in meters for mixed calculations
    // Time support
    BOOL hasTime;
    BOOL hasTimeTarget;
    TimeUnitId timeTarget;
} UnitsPre;

typedef enum {
    TARGET_KIND_NONE,
    TARGET_KIND_LENGTH,
    TARGET_KIND_TEMP,
    TARGET_KIND_TIME
} TargetKind;

static TargetKind DetectPreferredTargetToken(const WCHAR *s, const UnitsPre *ctx, int *lenOut, UnitId *unitOut, int *tempOut, TimeUnitId *timeOut) {
    int ulen = 0, tlen = 0, ttlen = 0;
    UnitId uid = match_unit_word(s, &ulen);
    int tid = match_temp_unit_word(s, &tlen);
    TimeUnitId ttid = T_INVALID;
    WCHAR timeBuf[MAX_SMALL_BUFFER];

    if (uid != U_INVALID && ulen > 0 && iswalpha(s[ulen])) uid = U_INVALID;
    if (tid >= 0 && tlen > 0 && (iswalpha(s[tlen]) || IsDegreeChar(s[tlen]))) tid = -1;

    {
        int i = 0;
        while (s[i] && (iswalpha(s[i]) || IsMicroChar(s[i])) && i < MAX_SMALL_BUFFER - 1) {
            timeBuf[i] = (WCHAR)towlower(s[i]);
            i++;
        }
        timeBuf[i] = 0;
        ttlen = i;
        ttid = match_time_unit_from_word(timeBuf, ttlen);
    }

    if (lenOut) *lenOut = 0;
    if (unitOut) *unitOut = U_INVALID;
    if (tempOut) *tempOut = -1;
    if (timeOut) *timeOut = T_INVALID;

    if (ctx) {
        if (ctx->hasTime && ttid != T_INVALID) {
            if (lenOut) *lenOut = ttlen;
            if (timeOut) *timeOut = ttid;
            return TARGET_KIND_TIME;
        }
        if (ctx->hasUnits && uid != U_INVALID) {
            if (lenOut) *lenOut = ulen;
            if (unitOut) *unitOut = uid;
            return TARGET_KIND_LENGTH;
        }
        if (ctx->hasTemp && tid >= 0) {
            if (lenOut) *lenOut = tlen;
            if (tempOut) *tempOut = tid;
            return TARGET_KIND_TEMP;
        }
    }

    if (uid != U_INVALID && tid < 0 && ttid == T_INVALID) {
        if (lenOut) *lenOut = ulen;
        if (unitOut) *unitOut = uid;
        return TARGET_KIND_LENGTH;
    }
    if (tid >= 0 && uid == U_INVALID && ttid == T_INVALID) {
        if (lenOut) *lenOut = tlen;
        if (tempOut) *tempOut = tid;
        return TARGET_KIND_TEMP;
    }
    if (ttid != T_INVALID && uid == U_INVALID && tid < 0) {
        if (lenOut) *lenOut = ttlen;
        if (timeOut) *timeOut = ttid;
        return TARGET_KIND_TIME;
    }

    if (uid != U_INVALID) {
        if (lenOut) *lenOut = ulen;
        if (unitOut) *unitOut = uid;
        return TARGET_KIND_LENGTH;
    }
    if (tid >= 0) {
        if (lenOut) *lenOut = tlen;
        if (tempOut) *tempOut = tid;
        return TARGET_KIND_TEMP;
    }
    if (ttid != T_INVALID) {
        if (lenOut) *lenOut = ttlen;
        if (timeOut) *timeOut = ttid;
        return TARGET_KIND_TIME;
    }

    return TARGET_KIND_NONE;
}

static BOOL HasIncompatibleConversionDomain(const UnitsPre *up) {
    BOOL srcLen, srcTemp, srcTime, dstLen, dstTemp, dstTime;
    if (!up) return FALSE;
    srcLen = up->hasUnits;
    srcTemp = up->hasTemp;
    srcTime = up->hasTime;
    dstLen = up->hasTarget;
    dstTemp = up->hasTempTarget;
    dstTime = up->hasTimeTarget;

    if ((srcLen && (dstTemp || dstTime)) ||
        (srcTemp && (dstLen || dstTime)) ||
        (srcTime && (dstLen || dstTemp))) {
        return TRUE;
    }
    return FALSE;
}

static void preprocess_units(const WCHAR *src, UnitsPre *out) {
    if (!out) return;
    out->hasUnits = FALSE;
    out->hasTarget = FALSE;
    out->target = U_INVALID;
    out->hasTemp = FALSE;
    out->hasTempTarget = FALSE;
    out->tempTarget = -1;
    out->hasTime = FALSE;
    out->hasTimeTarget = FALSE;
    out->timeTarget = T_INVALID;
    out->eval[0] = 0;
    out->prettyLeft[0] = 0;
    // Initialize mixed unit fields
    out->hasMixedUnits = FALSE;
    out->unitSystem = detect_unit_system(src);
    out->mixedResult = 0.0;
    
    size_t ei = 0, pi = 0, pif = 0; size_t ecap = sizeof(out->eval)/sizeof(out->eval[0]); size_t pcap = sizeof(out->prettyLeft)/sizeof(out->prettyLeft[0]); size_t pfcap = sizeof(out->prettyLeftFull)/sizeof(out->prettyLeftFull[0]);
    const WCHAR *s = src;
    while (*s && ei < ecap - 1 && pi < pcap - 1) {
        // Skip spaces copying to prettyLeft
        if (iswspace(*s)) { out->prettyLeft[pi++] = *s; if (pif < pfcap - 1) out->prettyLeftFull[pif++] = *s; s++; continue; }
                // Also apply same validation to target detection in = and 'to' patterns
                if (*s == L'=') {
                    s++; while (iswspace(*s)) s++;
                    {
                        int targetLen = 0, tempId = -1;
                        UnitId uid = U_INVALID;
                        TimeUnitId ttid = T_INVALID;
                        TargetKind kind = DetectPreferredTargetToken(s, out, &targetLen, &uid, &tempId, &ttid);
                        if (kind == TARGET_KIND_LENGTH) { out->hasTarget = TRUE; out->target = uid; s += targetLen; continue; }
                        if (kind == TARGET_KIND_TEMP) { out->hasTempTarget = TRUE; out->tempTarget = tempId; s += targetLen; continue; }
                        if (kind == TARGET_KIND_TIME) { out->hasTimeTarget = TRUE; out->timeTarget = ttid; s += targetLen; continue; }
                    }
                    
                    // If not a unit target, treat '=' as a delimiter in eval too
                    if (ei < ecap - 1) out->eval[ei++] = L'=';
                    out->prettyLeft[pi++] = L'='; if (pif < pfcap - 1) out->prettyLeftFull[pif++] = L'='; continue;
                }
        if ((s[0] == L't' || s[0] == L'T') && (s[1] == L'o' || s[1] == L'O') && !iswalpha(s[2])) {
            // likely 'to'
            const WCHAR *p = s + 2; while (iswspace(*p)) p++;
            {
                int targetLen = 0, tempId = -1;
                UnitId uid = U_INVALID;
                TimeUnitId ttid = T_INVALID;
                TargetKind kind = DetectPreferredTargetToken(p, out, &targetLen, &uid, &tempId, &ttid);
                if (kind == TARGET_KIND_LENGTH) { out->hasTarget = TRUE; out->target = uid; s = p + targetLen; continue; }
                if (kind == TARGET_KIND_TEMP) { out->hasTempTarget = TRUE; out->tempTarget = tempId; s = p + targetLen; continue; }
                if (kind == TARGET_KIND_TIME) { out->hasTimeTarget = TRUE; out->timeTarget = ttid; s = p + targetLen; continue; }
            }
            // else copy 'to'
        }
        // Treat standalone word 'in' as a conversion operator to target (e.g., "1 in in cm")
        if ((s[0] == L'i' || s[0] == L'I') && (s[1] == L'n' || s[1] == L'N') && !iswalpha(s[2])) {
            const WCHAR *p = s + 2; while (iswspace(*p)) p++;
            
            {
                int targetLen = 0, tempId = -1;
                UnitId uid = U_INVALID;
                TimeUnitId ttid = T_INVALID;
                TargetKind kind = DetectPreferredTargetToken(p, out, &targetLen, &uid, &tempId, &ttid);
                if (kind == TARGET_KIND_LENGTH) { out->hasTarget = TRUE; out->target = uid; s = p + targetLen; continue; }
                if (kind == TARGET_KIND_TEMP) { out->hasTempTarget = TRUE; out->tempTarget = tempId; s = p + targetLen; continue; }
                if (kind == TARGET_KIND_TIME) { out->hasTimeTarget = TRUE; out->timeTarget = ttid; s = p + targetLen; continue; }
                if (out->hasTime) {
                    // Keep legacy shorthand for time-only targets after "in".
                    WCHAR buf[MAX_SMALL_BUFFER];
                    int i = 0;
                    while (p[i] && (iswalpha(p[i]) || IsMicroChar(p[i])) && i < MAX_SMALL_BUFFER - 1) {
                        buf[i] = (WCHAR)towlower(p[i]);
                        i++;
                    }
                    buf[i] = 0;
                    if (wcslen(buf) == 1) {
                        if (buf[0] == L'm') { out->hasTimeTarget = TRUE; out->timeTarget = T_MONTH; s = p + 1; continue; }
                        if (buf[0] == L'h') { out->hasTimeTarget = TRUE; out->timeTarget = T_H; s = p + 1; continue; }
                        if (buf[0] == L's') { out->hasTimeTarget = TRUE; out->timeTarget = T_S; s = p + 1; continue; }
                    } else if (wcslen(buf) == 2 && wcscmp(buf, L"mo") == 0) {
                        out->hasTimeTarget = TRUE; out->timeTarget = T_MONTH; s = p + 2; continue;
                    } else if (wcslen(buf) == 3 && wcscmp(buf, L"wee") == 0) {
                        out->hasTimeTarget = TRUE; out->timeTarget = T_WEEK; s = p + 3; continue;
                    }
                }
            }
            // else copy 'in' as letters
        }
        // Try quote pattern N' M"
        const WCHAR *save = s; WCHAR *endptr = NULL; double feet = wcstod(s, &endptr);
        if (endptr != s) {
            const WCHAR *p = endptr; const WCHAR *origNumEnd = endptr;
            while (iswspace(*p)) p++;
            if (*p == L'\'') {
                // Feet marker can be ' or ''
                if (p[0] == L'\'' && p[1] == L'\'') p += 2; else p++;
                while (iswspace(*p)) p++;
                // Optional inches
                double inches = 0.0; WCHAR *end2 = NULL; const WCHAR *q = p;
                double maybeIn = wcstod(q, &end2);
                if (end2 != q) { inches = maybeIn; q = end2; while (iswspace(*q)) q++; }
                if (*q == L'"' || inches == 0.0) {
                    if (*q == L'"') q++;
                    double totalIn = feet * 12.0 + inches;
                    double meters = totalIn * unit_to_m(U_IN);
                    // write eval number
                    WCHAR numbuf[64]; swprintf(numbuf, 64, L"%.15g", meters);
                    size_t nlen = wcslen(numbuf); if (ei + nlen >= ecap) nlen = ecap - ei - 1;
                    wmemcpy(out->eval + ei, numbuf, nlen); ei += nlen;
                    // pretty copy original slice from start to q
                    size_t slen = (size_t)(q - save);
                    if (pi + slen >= pcap) slen = pcap - pi - 1;
                    wmemcpy(out->prettyLeft + pi, save, slen); pi += slen;
                    // prettyFull: keep original quote form
                    if (pif + slen >= pfcap) slen = pfcap - pif - 1;
                    wmemcpy(out->prettyLeftFull + pif, save, slen); pif += slen;
                    out->hasUnits = TRUE; s = q; continue;
                }
            }
            // If number followed by unit word (temperature first, then length, then time)
            s = save; feet = wcstod(s, &endptr); if (endptr != s) {
                const WCHAR *p2 = endptr; while (iswspace(*p2)) p2++;
                int tlen3 = 0; int tid3 = match_temp_unit_word(p2, &tlen3);
                if (tid3 >= 0) {
                    // Check if there are remaining alphabetic characters after the temp unit
                    if (tlen3 > 0 && (iswalpha(p2[tlen3]) || IsDegreeChar(p2[tlen3]))) {
                        // Invalid temp unit - contains extra letters
                        tid3 = -1;
                    }
                }
                if (tid3 >= 0) {
                    double kel = temp_to_kelvin(feet, tid3);
                    WCHAR numbuf[64]; swprintf(numbuf, 64, L"%.15g", kel);
                    size_t nlen = wcslen(numbuf); if (ei + nlen >= ecap) nlen = ecap - ei - 1;
                    wmemcpy(out->eval + ei, numbuf, nlen); ei += nlen;
                    size_t plen = (size_t)((p2 + tlen3) - s);
                    if (pi + plen >= pcap) plen = pcap - pi - 1;
                    wmemcpy(out->prettyLeft + pi, s, plen); pi += plen;
                    // prettyFull: number + space + symbol
                    size_t numlen = (size_t)(endptr - s);
                    if (pif + numlen + 3 >= pfcap) numlen = (pfcap - pif - 4);
                    wmemcpy(out->prettyLeftFull + pif, s, numlen); pif += numlen;
                    if (pif < pfcap - 1) out->prettyLeftFull[pif++] = L' ';
                    const WCHAR *sym = temp_symbol(tid3);
                    size_t symlen = wcslen(sym);
                    if (pif + symlen >= pfcap) symlen = pfcap - pif - 1;
                    wmemcpy(out->prettyLeftFull + pif, sym, symlen); pif += symlen;
                    out->hasTemp = TRUE; s = p2 + tlen3; continue;
                }
                int ulen2 = 0; UnitId uid2 = match_unit_word(p2, &ulen2);
                if (uid2 != U_INVALID) {
                    // Check if there are remaining alphabetic characters after the unit
                    // This catches cases like "ftt" where "ft" matches but "t" remains
                    if (ulen2 > 0 && iswalpha(p2[ulen2])) {
                        // Invalid unit - contains extra letters
                        uid2 = U_INVALID;
                    }
                }
                if (uid2 != U_INVALID) {
                    double meters = feet * unit_to_m(uid2);
                    WCHAR numbuf[64]; swprintf(numbuf, 64, L"%.15g", meters);
                    size_t nlen = wcslen(numbuf); if (ei + nlen >= ecap) nlen = ecap - ei - 1;
                    wmemcpy(out->eval + ei, numbuf, nlen); ei += nlen;
                    // pretty copy original number+unit
                    size_t plen = (size_t)((p2 + ulen2) - s);
                    if (pi + plen >= pcap) plen = pcap - pi - 1;
                    wmemcpy(out->prettyLeft + pi, s, plen); pi += plen;
                    // prettyFull: number + space + full unit name
                    size_t numlen = (size_t)(endptr - s);
                    if (pif + numlen + 1 >= pfcap) numlen = (pfcap - pif - 2);
                    wmemcpy(out->prettyLeftFull + pif, s, numlen); pif += numlen;
                    if (pif < pfcap - 1) out->prettyLeftFull[pif++] = L' ';
                    // decide plural
                    double qty = wcstod(s, NULL);
                    WCHAR uname[32]; unit_name_full(uid2, qty, uname, 32);
                    size_t unlen = wcslen(uname);
                    if (pif + unlen >= pfcap) unlen = pfcap - pif - 1;
                    wmemcpy(out->prettyLeftFull + pif, uname, unlen); pif += unlen;
                    out->hasUnits = TRUE; s = p2 + ulen2; continue;
                }
                // time unit
                {
                    // Read possible time word
                    const WCHAR *q = p2; int i = 0; WCHAR buf[32];
                    while (q[i] && (iswalpha(q[i]) || IsMicroChar(q[i])) && i < 31) { buf[i] = (WCHAR)towlower(q[i]); i++; }
                    buf[i] = 0; int ttlen = i; TimeUnitId ttid = T_INVALID;
                    for (int k = 0; k < (int)(sizeof(kTimeUnitWords)/sizeof(kTimeUnitWords[0])); ++k) if (wcsieq(buf, kTimeUnitWords[k].word)) { ttid = kTimeUnitWords[k].id; break; }
                    if (ttid != T_INVALID) {
                        // seconds value
                        double seconds = feet * kTimeUnits[ttid].to_s;
                        WCHAR numbuf[64]; swprintf(numbuf, 64, L"%.15g", seconds);
                        size_t nlen = wcslen(numbuf); if (ei + nlen >= ecap) nlen = ecap - ei - 1;
                        wmemcpy(out->eval + ei, numbuf, nlen); ei += nlen;
                        // pretty copy original text
                        size_t plen = (size_t)((p2 + ttlen) - s);
                        if (pi + plen >= pcap) plen = pcap - pi - 1;
                        wmemcpy(out->prettyLeft + pi, s, plen); pi += plen;
                        // prettyFull: number + space + full unit name
                        size_t numlen = (size_t)(endptr - s);
                        if (pif + numlen + 1 >= pfcap) numlen = (pfcap - pif - 2);
                        wmemcpy(out->prettyLeftFull + pif, s, numlen); pif += numlen;
                        if (pif < pfcap - 1) out->prettyLeftFull[pif++] = L' ';
                        // decide plural
                        double qty = wcstod(s, NULL);
                        const TimeUnitInfo *ti = &kTimeUnits[ttid];
                        const WCHAR *name = (fabs(qty - 1.0) < 1e-12) ? ti->name_singular : ti->name_plural;
                        size_t unlen = wcslen(name);
                        if (pif + unlen >= pfcap) unlen = pfcap - pif - 1;
                        wmemcpy(out->prettyLeftFull + pif, name, unlen); pif += unlen;
                        out->hasTime = TRUE; s = p2 + ttlen; continue;
                    }
                }
            }
            // plain number: copy to eval and pretty
            int consumed = (int)(origNumEnd - s);
            if (consumed > 0) {
                if (ei + consumed >= ecap) consumed = (int)(ecap - ei - 1);
                wmemcpy(out->eval + ei, s, consumed); ei += consumed;
                if (pi + consumed >= pcap) consumed = (int)(pcap - pi - 1);
                wmemcpy(out->prettyLeft + pi, s, consumed); pi += consumed;
                if (pif + consumed >= pfcap) consumed = (int)(pfcap - pif - 1);
                wmemcpy(out->prettyLeftFull + pif, s, consumed); pif += consumed;
                s = origNumEnd; continue;
            }
        }
        // Operators and others: copy to both
        out->eval[ei++] = *s;
        out->prettyLeft[pi++] = *s;
        if (pif < pfcap - 1) out->prettyLeftFull[pif++] = *s;
        s++;
    }
    out->eval[ei] = 0; out->prettyLeft[pi] = 0;
    out->prettyLeftFull[pif] = 0;
}

// Expression parsing/evaluation
typedef struct Parser {
    const WCHAR *s;
    int len;
    int pos;
} Parser;

static void skip_ws(Parser *p){ while(p->pos < p->len && (p->s[p->pos] == L' ' || p->s[p->pos] == L'\t')) p->pos++; }

static BOOL parse_number(Parser *p, double *out){
    skip_ws(p);
    int start = p->pos;
    BOOL seen = FALSE;
    while (p->pos < p->len && iswdigit(p->s[p->pos])) { p->pos++; seen = TRUE; }
    if (p->pos < p->len && p->s[p->pos] == L'.') {
        p->pos++;
        while (p->pos < p->len && iswdigit(p->s[p->pos])) { p->pos++; seen = TRUE; }
    }
    if (!seen) { p->pos = start; return FALSE; }
    WCHAR buf[128]; int n = p->pos - start; if (n >= (int)(sizeof(buf)/sizeof(buf[0]))) n = (int)(sizeof(buf)/sizeof(buf[0])) - 1;
    wcsncpy(buf, p->s + start, n); buf[n] = 0;
    *out = wcstod(buf, NULL);
    return TRUE;
}

static BOOL parse_factor(Parser *p, double *out, BOOL *outIsPercent);
static BOOL parse_power(Parser *p, double *out, BOOL *outIsPercent);
static BOOL parse_expr(Parser *p, double *out);

static BOOL parse_unary(Parser *p, double *out, BOOL *outIsPercent){
    skip_ws(p);
    if (p->pos < p->len && (p->s[p->pos] == L'+' || p->s[p->pos] == L'-')) {
        WCHAR op = p->s[p->pos++];
        double val;
        BOOL isPct = FALSE;
        if (!parse_unary(p, &val, &isPct)) return FALSE;
        *out = (op == L'-') ? -val : val;
        if (outIsPercent) *outIsPercent = isPct;
        return TRUE;
    }
    return parse_power(p, out, outIsPercent);
}

static BOOL parse_factor(Parser *p, double *out, BOOL *outIsPercent){
    skip_ws(p);
    if (p->pos < p->len && p->s[p->pos] == L'(') {
        p->pos++;
        // Parse a full expression inside parentheses to respect precedence
        if (!parse_expr(p, out)) return FALSE;
        skip_ws(p);
        if (p->pos < p->len && p->s[p->pos] == L')') { p->pos++; return TRUE; }
        return FALSE;
    }
    if (!parse_number(p, out)) return FALSE;
    skip_ws(p);
    if (p->pos < p->len && p->s[p->pos] == L'%') {
        p->pos++;
        *out = (*out) / 100.0;
        if (outIsPercent) *outIsPercent = TRUE;
    } else if (outIsPercent) {
        *outIsPercent = FALSE;
    }
    return TRUE;
}

static BOOL parse_power(Parser *p, double *out, BOOL *outIsPercent){
    BOOL isPct = FALSE;
    if (!parse_factor(p, out, &isPct)) return FALSE;
    // Right-associative: 2^3^2 = 2^(3^2) = 2^9 = 512
    skip_ws(p);
    if (p->pos < p->len && p->s[p->pos] == L'^') {
        p->pos++;
        double exponent;
        BOOL expPct = FALSE;
        if (!parse_power(p, &exponent, &expPct)) return FALSE;
        *out = pow(*out, exponent);
        isPct = FALSE; // Power operation results in non-percent
    }
    if (outIsPercent) *outIsPercent = isPct;
    return TRUE;
}

static BOOL parse_term(Parser *p, double *out, BOOL *outTermIsPercent){
    BOOL isPct = FALSE;
    if (!parse_unary(p, out, &isPct)) return FALSE;
    BOOL termIsPct = isPct;
    while (1) {
        skip_ws(p);
        if (p->pos >= p->len) break;
        WCHAR op = p->s[p->pos];
        if (op != L'*' && op != L'/' && op != L'×' && op != L'÷') break;
        p->pos++;
        double rhs;
        BOOL rhsPct = FALSE;
        if (!parse_unary(p, &rhs, &rhsPct)) return FALSE;
        // Any multiplication/division means composite term
        termIsPct = FALSE;
        if (op == L'*' || op == L'×') *out *= rhs; else *out /= rhs;
    }
    if (outTermIsPercent) *outTermIsPercent = termIsPct;
    return TRUE;
}

static BOOL parse_expr(Parser *p, double *out){
    BOOL lhsPct = FALSE;
    if (!parse_term(p, out, &lhsPct)) return FALSE;
    while (1) {
        skip_ws(p);
        if (p->pos >= p->len) break;
        WCHAR op = p->s[p->pos];
        if (op != L'+' && op != L'-') break;
        p->pos++;
        double rhs;
        BOOL rhsPct = FALSE;
        if (!parse_term(p, &rhs, &rhsPct)) return FALSE;
        if (rhsPct) {
            double delta = (*out) * rhs; // rhs already divided by 100
            if (op == L'+') *out += delta; else *out -= delta;
        } else {
            if (op == L'+') *out += rhs; else *out -= rhs;
        }
    }
    return TRUE;
}

static BOOL eval_expression(const WCHAR *s, double *out, BOOL *complete, int *consumed){
    Parser p = { s, (int)wcslen(s), 0 };
    skip_ws(&p);
    int start = p.pos;
    double val;
    if (!parse_expr(&p, &val)) { if (consumed) *consumed = 0; if (complete) *complete = FALSE; return FALSE; }
    // success for part of the string
    int pos_after = p.pos;
    // Allow trailing whitespace
    skip_ws(&p);
    if (complete) *complete = (p.pos >= p.len);
    if (consumed) *consumed = pos_after - start;
    if (out) *out = val;
    return TRUE;
}

static void trim_trailing_zeros(WCHAR *buf) {
    if (!buf) return;
    size_t n = wcslen(buf);
    while (n > 0 && buf[n-1] == L'0') { buf[--n] = 0; }
    if (n > 0 && buf[n-1] == L'.') { buf[--n] = 0; }
    if (n > 1 && buf[0] == L'-' && buf[1] == L'0' && buf[2] == 0) { buf[0] = L'0'; buf[1] = 0; }
}

// Display formatting: cap to 6 decimals, trim trailing zeros
static void format_double(double v, WCHAR *buf, size_t cap){
    if (!buf || cap == 0) return;
    double r = round(v);
    if (fabs(v - r) < 1e-12) {
        swprintf(buf, cap, L"%.0f", r);
        return;
    }
    swprintf(buf, cap, L"%.6f", v);
    trim_trailing_zeros(buf);
    // If formatting collapsed to 0 but value is non-zero and very small, show with 6 significant digits
    if ((wcscmp(buf, L"0") == 0 || wcscmp(buf, L"-0") == 0) && fabs(v) > 0.0 && fabs(v) < 1e-6) {
        swprintf(buf, cap, L"%.6g", v);
    }
}

// Display formatting: cap to 3 decimals, trim trailing zeros
static void format_double_3(double v, WCHAR *buf, size_t cap){
    if (!buf || cap == 0) return;
    double r = round(v);
    if (fabs(v - r) < 1e-12) {
        swprintf(buf, cap, L"%.0f", r);
        return;
    }
    swprintf(buf, cap, L"%.3f", v);
    trim_trailing_zeros(buf);
    if ((wcscmp(buf, L"0") == 0 || wcscmp(buf, L"-0") == 0) && fabs(v) > 0.0 && fabs(v) < 1e-3) {
        swprintf(buf, cap, L"%.3g", v);
    }
}

// Full-precision insertion formatting: avoid scientific notation, keep up to 12 decimals
static void format_double_full(double v, WCHAR *buf, size_t cap) {
    if (!buf || cap == 0) return;
    swprintf(buf, cap, L"%.12f", v);
    trim_trailing_zeros(buf);
}

static void format_expr_spaced(const WCHAR *src, WCHAR *dst, size_t cap){
    if (!dst || cap == 0) return;
    size_t di = 0;
    int expectOperand = 1; // at start we expect an operand; + or - here are unary
    for (size_t i = 0; src && src[i] && di < cap - 1; ) {
        WCHAR c = src[i];
        if (iswspace(c)) { i++; continue; }
        if (c == L'(') { dst[di++] = c; i++; expectOperand = 1; continue; }
        if (c == L')') { dst[di++] = c; i++; expectOperand = 0; continue; }
        if (c == L'+' || c == L'-') {
            if (expectOperand) {
                // Unary sign: attach to next number without spaces
                dst[di++] = c; i++; // keep expectOperand = 1 to allow sequences like + - 1
                if (di >= cap - 1) break;
                continue;
            } else {
                if (di > 0 && dst[di-1] != L' ') dst[di++] = L' ';
                dst[di++] = c;
                if (di < cap - 1) dst[di++] = L' ';
                i++; expectOperand = 1; continue;
            }
        }
        if (c == L'*' || c == L'/' || c == L'×' || c == L'÷') {
            if (di > 0 && dst[di-1] != L' ') dst[di++] = L' ';
            dst[di++] = c;
            if (di < cap - 1) dst[di++] = L' ';
            i++; expectOperand = 1; continue;
        }
        if (c == L'^') {
            if (di > 0 && dst[di-1] != L' ') dst[di++] = L' ';
            dst[di++] = c;
            if (di < cap - 1) dst[di++] = L' ';
            i++; expectOperand = 1; continue;
        }
        // Normal character (digits, dot, etc.)
        dst[di++] = c;
        i++; expectOperand = 0;
    }
    // Trim trailing spaces
    while (di > 0 && dst[di-1] == L' ') di--;
    // Remove spaces after '(' and before ')'
    dst[di] = 0;
    size_t r = 0, w = 0; 
    while (dst[r]) {
        if (dst[r] == L'(' && dst[r+1] == L' ') { r++; continue; }
        if (dst[r] == L' ' && dst[r+1] == L')') { r++; continue; }
        dst[w++] = dst[r++];
    }
    dst[w] = 0;
}

// Comparison query parsing and answering
static void trim_ws(const WCHAR *s, const WCHAR **outStart, const WCHAR **outEnd) {
    const WCHAR *a = s; while (*a && iswspace(*a)) a++;
    const WCHAR *b = s + wcslen(s);
    while (b > a && iswspace(*(b-1))) b--;
    *outStart = a; *outEnd = b;
}

static BOOL starts_with_is(const WCHAR *s, int *consumed) {
    const WCHAR *p = s; while (*p && iswspace(*p)) p++;
    if ((p[0] == L'i' || p[0] == L'I') && (p[1] == L's' || p[1] == L'S') && !iswalpha(p[2])) {
        p += 2; while (*p && iswspace(*p)) p++;
        if (consumed) *consumed = (int)(p - s);
        return TRUE;
    }
    if (consumed) *consumed = 0;
    return FALSE;
}

static BOOL find_comparator(const WCHAR *s, int *opIndex, CmpOp *op, int *opLen) {
    int depth = 0; int idx = -1; CmpOp o = CMP_NONE; int olen = 1;
    int len = (int)wcslen(s);
    for (int i = 0; i < len; ++i) {
        WCHAR c = s[i];
        if (c == L'(') { depth++; continue; }
        if (c == L')') { if (depth > 0) depth--; continue; }
        if (depth > 0) continue;
        if (c == L'>' || c == L'<' || c == L'=' || c == L'!' || c == L'≥' || c == L'≤' || c == L'≠') {
            // Multi-character operators first
            if (c == L'>') { if (i+1 < len && s[i+1] == L'=') { o = CMP_GE; olen = 2; } else { o = CMP_GT; olen = 1; } }
            else if (c == L'<') { if (i+1 < len && s[i+1] == L'=') { o = CMP_LE; olen = 2; } else { o = CMP_LT; olen = 1; } }
            else if (c == L'=') { if (i+1 < len && s[i+1] == L'=') { o = CMP_EQ; olen = 2; } else { o = CMP_EQ; olen = 1; } }
            else if (c == L'!') { if (i+1 < len && s[i+1] == L'=') { o = CMP_NE; olen = 2; } else { continue; } }
            else if (c == L'≥') { o = CMP_GE; olen = 1; }
            else if (c == L'≤') { o = CMP_LE; olen = 1; }
            else if (c == L'≠') { o = CMP_NE; olen = 1; }
            idx = i; break;
        }
    }
    if (idx < 0 || o == CMP_NONE) return FALSE;
    if (opIndex) *opIndex = idx; if (op) *op = o; if (opLen) *opLen = olen; return TRUE;
}

static BOOL split_comparison(const WCHAR *input, WCHAR *left, size_t lcap, WCHAR *right, size_t rcap, CmpOp *opOut, BOOL *completeOut) {
    if (left && lcap) left[0] = 0; if (right && rcap) right[0] = 0; if (completeOut) *completeOut = FALSE;
    int off = 0; starts_with_is(input, &off);
    const WCHAR *core = input + off;
    // Trim trailing '?'
    const WCHAR *a, *b; trim_ws(core, &a, &b);
    if (b > a && *(b-1) == L'?') { b--; while (b > a && iswspace(*(b-1))) b--; }
    if (a >= b) return FALSE;
    WCHAR buf[1024]; int n = (int)(b - a); if (n >= (int)(sizeof(buf)/sizeof(buf[0]))) n = (int)sizeof(buf)/sizeof(buf[0]) - 1; wcsncpy(buf, a, n); buf[n] = 0;
    int idx = -1; CmpOp op = CMP_NONE; int olen = 0;
    if (!find_comparator(buf, &idx, &op, &olen)) return FALSE;
    // Split
    WCHAR ltmp[1024], rtmp[1024];
    int ln = idx; if (ln < 0) ln = 0; if (ln >= (int)sizeof(ltmp)/2) ln = (int)(sizeof(ltmp)/sizeof(ltmp[0])) - 1;
    wcsncpy(ltmp, buf, ln); ltmp[ln] = 0;
    const WCHAR *rstart = buf + idx + olen;
    while (*rstart && iswspace(*rstart)) rstart++;
    while (ln > 0 && iswspace(ltmp[ln-1])) { ltmp[--ln] = 0; }
    wcsncpy(rtmp, rstart, (sizeof(rtmp)/sizeof(rtmp[0])) - 1); rtmp[(sizeof(rtmp)/sizeof(rtmp[0])) - 1] = 0;
    // Copy out
    if (left && lcap) { wcsncpy(left, ltmp, lcap - 1); left[lcap - 1] = 0; }
    if (right && rcap) { wcsncpy(right, rtmp, rcap - 1); right[rcap - 1] = 0; }
    if (opOut) *opOut = op;
    if (completeOut) *completeOut = TRUE;
    return TRUE;
}

static int sign_from_cmp(CmpOp op, double L, double R) {
    switch (op) {
        case CMP_GT: return (L > R) ? 1 : 0;
        case CMP_GE: return (L >= R) ? 1 : 0;
        case CMP_LT: return (L < R) ? 1 : 0;
        case CMP_LE: return (L <= R) ? 1 : 0;
        case CMP_EQ: return (fabs(L - R) < 1e-12) ? 1 : 0;
        case CMP_NE: return (fabs(L - R) >= 1e-12) ? 1 : 0;
        default: return 0;
    }
}

static const WCHAR* cmp_readable(CmpOp op, BOOL truth, BOOL invert) {
    // When building explanations like "which is less than", choose phrase aligned with actual relation
    switch (op) {
        case CMP_GT: return truth ? L"greater than" : (invert ? L"greater than or equal to" : L"less than or equal to");
        case CMP_GE: return truth ? L"greater than or equal to" : L"less than";
        case CMP_LT: return truth ? L"less than" : (invert ? L"less than or equal to" : L"greater than or equal to");
        case CMP_LE: return truth ? L"less than or equal to" : L"greater than";
        case CMP_EQ: return truth ? L"equal to" : L"not equal to";
        case CMP_NE: return truth ? L"not equal to" : L"equal to";
        default: return L"";
    }
}

static BOOL BuildComparisonAnswer(const WCHAR *input, BOOL forLive, WCHAR *outText, size_t outCap, WCHAR *answerOut, size_t ansCap, BOOL *completeOut) {
    if (outText && outCap) outText[0] = 0; if (answerOut && ansCap) answerOut[0] = 0; if (completeOut) *completeOut = FALSE;
    // Split into left/right and operator
    WCHAR left[512], right[512]; CmpOp op = CMP_NONE; BOOL splitOk = split_comparison(input, left, 512, right, 512, &op, completeOut);
    if (!splitOk || op == CMP_NONE) return FALSE;

    // Normalize text on each side similar to normal evaluation
    WCHAR l1[512], l2[512], l3[512]; normalize_percent_word(left, l1, 512); normalize_of_operator(l1, l2, 512); normalize_operators_basic(l2, l3, 512);
    WCHAR r1[512], r2[512], r3[512]; normalize_percent_word(right, r1, 512); normalize_of_operator(r1, r2, 512); normalize_operators_basic(r2, r3, 512);

    UnitsPre ul = {0}; UnitsPre ur = {0}; preprocess_units(l3, &ul); preprocess_units(r3, &ur);
    const WCHAR *le = ul.eval[0] ? ul.eval : l3; const WCHAR *re = ur.eval[0] ? ur.eval : r3;

    double Lv = 0.0, Rv = 0.0; BOOL lc = FALSE, rc = FALSE; int lcons = 0, rcons = 0;
    if (!eval_expression(le, &Lv, &lc, &lcons) || lcons <= 0) return FALSE;
    if (!eval_expression(re, &Rv, &rc, &rcons) || rcons <= 0) return FALSE;

    // Determine domain for formatting (length/time/temp/none)
    BOOL hasLen = (ul.hasUnits || ul.hasTarget || ur.hasUnits || ur.hasTarget);
    BOOL hasTmp = (ul.hasTemp || ul.hasTempTarget || ur.hasTemp || ur.hasTempTarget);
    BOOL hasTim = (ul.hasTime || ul.hasTimeTarget || ur.hasTime || ur.hasTimeTarget);

    // Check for mixed unit/unitless comparisons - treat as incomplete
    if (hasLen && ((ul.hasUnits || ul.hasTarget) != (ur.hasUnits || ur.hasTarget))) {
        // One side has units, other doesn't - incomplete expression
        return FALSE;
    }
    if (hasTmp && ((ul.hasTemp || ul.hasTempTarget) != (ur.hasTemp || ur.hasTempTarget))) {
        // One side has temperature, other doesn't - incomplete expression
        return FALSE;
    }
    if (hasTim && ((ul.hasTime || ul.hasTimeTarget) != (ur.hasTime || ur.hasTimeTarget))) {
        // One side has time, other doesn't - incomplete expression
        return FALSE;
    }

    // Choose display unit for explanation
    WCHAR expl[512] = {0};
    int truth = sign_from_cmp(op, Lv, Rv);
    if (answerOut && ansCap) {
        swprintf(answerOut, ansCap, L"%d", truth ? 1 : 0);
    }

    if (hasTmp && !(hasLen || hasTim)) {
        // Kelvin base; pick symbol from right then left then °C
        int disp = find_last_temp_unit_in_text(right); if (disp < 0) disp = find_last_temp_unit_in_text(left); if (disp < 0) disp = 0; // °C default
        double Ld = kelvin_to_temp(Lv, disp), Rd = kelvin_to_temp(Rv, disp);
        WCHAR ln[64], rn[64]; format_double(Ld, ln, 64); format_double(Rd, rn, 64);
        const WCHAR *sym = temp_symbol(disp);
        const WCHAR *rel = cmp_readable(op, truth, FALSE);
        if (forLive) {
            _snwprintf(expl, sizeof(expl)/sizeof(expl[0]), L"%s, %s %s %s %s", truth?L"Yes":L"No", left, rel, rn, sym);
        } else {
            _snwprintf(expl, sizeof(expl)/sizeof(expl[0]), L"%s, %s ≈ %s %s, which is %s %s %s", truth?L"Yes":L"No", left, ln, sym, rel, rn, sym);
        }
    } else if (hasTim && !(hasLen || hasTmp)) {
        // Seconds base
        TimeUnitId disp = find_last_time_unit_in_text(right); if (disp == T_INVALID) disp = find_last_time_unit_in_text(left); if (disp == T_INVALID) disp = T_S;
        double Ld = Lv / kTimeUnits[disp].to_s, Rd = Rv / kTimeUnits[disp].to_s;
        WCHAR ln[64], rn[64]; format_double(Ld, ln, 64); format_double(Rd, rn, 64);
        const WCHAR *unit = (fabs(Ld - 1.0) < 1e-12) ? kTimeUnits[disp].name_singular : kTimeUnits[disp].name_plural; // not exact for both but fine for short text
        const WCHAR *rel = cmp_readable(op, truth, FALSE);
        if (forLive) {
            _snwprintf(expl, sizeof(expl)/sizeof(expl[0]), L"%s, %s %s %s %s", truth?L"Yes":L"No", left, rel, rn, unit);
        } else {
            _snwprintf(expl, sizeof(expl)/sizeof(expl[0]), L"%s, %s ≈ %s %s, which is %s %s %s", truth?L"Yes":L"No", left, ln, unit, rel, rn, unit);
        }
    } else if (hasLen && !(hasTmp || hasTim)) {
        // Meters base; choose display unit from right then left, else meters
        UnitId disp = find_last_length_unit_in_text(right); if (disp == U_INVALID) disp = find_last_length_unit_in_text(left); if (disp == U_INVALID) disp = U_M;
        double Ld = Lv / unit_to_m(disp), Rd = Rv / unit_to_m(disp);
        WCHAR ln[64], rn[64]; format_unit_value(disp, Ld, ln, 64); format_unit_value(disp, Rd, rn, 64);
        WCHAR unameL[32]; unit_name_full(disp, Ld, unameL, 32);
        const WCHAR *rel = cmp_readable(op, truth, FALSE);
        // Build concise explanation
        if (forLive) {
            _snwprintf(expl, sizeof(expl)/sizeof(expl[0]), L"%s, %s %s %s %s", truth?L"Yes":L"No", left, rel, rn, unameL);
        } else {
            _snwprintf(expl, sizeof(expl)/sizeof(expl[0]), L"%s, %s ≈ %s %s, which is %s %s %s", truth?L"Yes":L"No", left, ln, unameL, rel, rn, unameL);
        }
    } else {
        // Pure numeric
        WCHAR ln[64], rn[64]; format_double(Lv, ln, 64); format_double(Rv, rn, 64);
        const WCHAR *rel = cmp_readable(op, truth, FALSE);
        if (forLive) {
            _snwprintf(expl, sizeof(expl)/sizeof(expl[0]), L"%s, %s %s %s", truth?L"Yes":L"No", left, rel, rn);
        } else {
            _snwprintf(expl, sizeof(expl)/sizeof(expl[0]), L"%s, %s ≈ %s, which is %s %s", truth?L"Yes":L"No", left, ln, rel, rn);
        }
    }

    if (outText && outCap) { wcsncpy(outText, expl, outCap - 1); outText[outCap - 1] = 0; }
    if (completeOut) *completeOut = TRUE;
    return TRUE;
}

static void normalize_expr_for_history(const WCHAR *src, WCHAR *dst, size_t cap) {
    // src should be the spaced expression; we will:
    // - replace '*' with '×', '/' with '÷'
    // - replace tokens like [-+]?[0-9]+(.[0-9]+)?% with their decimal (value/100)
    if (!dst || cap == 0) return;
    size_t di = 0;
    for (size_t i = 0; src && src[i] && di < cap - 1; ) {
        WCHAR c = src[i];
        if (c == L'*') { dst[di++] = L'×'; i++; continue; }
        if (c == L'/') { dst[di++] = L'÷'; i++; continue; }
        // Attempt to parse a signed number followed by '%'
    size_t j = i;
    // token start if at beginning or previous is space or '('
    BOOL tokenStart = (j == 0) || src[j-1] == L' ' || src[j-1] == L'(';
    if (tokenStart && (src[j] == L'+' || src[j] == L'-')) { j++; }
        BOOL sawDigits = FALSE;
        size_t k = j;
        while (iswdigit(src[k])) { k++; sawDigits = TRUE; }
        if (src[k] == L'.') {
            size_t k2 = k + 1; BOOL any = FALSE;
            while (iswdigit(src[k2])) { k2++; any = TRUE; }
            if (any) { k = k2; sawDigits = TRUE; }
        }
        if (sawDigits && src[k] == L'%') {
            // Convert the number in [i..k) to double, divide by 100, write formatted
            int nlen = (int)(k - i);
            if (nlen > 0) {
                WCHAR buf[128];
                if (nlen >= (int)(sizeof(buf)/sizeof(buf[0]))) nlen = (int)(sizeof(buf)/sizeof(buf[0])) - 1;
                wcsncpy(buf, src + i, nlen); buf[nlen] = 0;
                double v = wcstod(buf, NULL);
                v = (v / 100.0);
                // format_double already drops trailing .0
                WCHAR outNum[64]; format_double(v, outNum, 64);
                size_t on = wcslen(outNum);
                for (size_t t = 0; t < on && di < cap - 1; ++t) dst[di++] = outNum[t];
                i = k + 1; // skip the '%'
                continue;
            }
        }
        dst[di++] = c; i++;
    }
    dst[di] = 0;
}

// Special-case pretty formatter for expressions like:  A + B%  or  A - B%
// Displays as: "A ± (A × (B/100))"
static BOOL PrettyPercentOfLeft(const WCHAR *src, WCHAR *dst, size_t cap) {
    if (!src || !dst || cap == 0) return FALSE;
    const WCHAR *p = src;
    // Skip leading spaces
    while (*p && iswspace(*p)) p++;
    // Parse A
    wchar_t *end = NULL;
    double A = wcstod(p, &end);
    if (end == p) return FALSE; // no number
    p = end;
    // Skip spaces
    while (*p && iswspace(*p)) p++;
    // Op must be + or -
    if (*p != L'+' && *p != L'-') return FALSE;
    WCHAR op = *p;
    p++;
    // Skip spaces
    while (*p && iswspace(*p)) p++;
    // Parse B
    double B = wcstod(p, &end);
    if (end == p) return FALSE;
    p = end;
    // Spaces then percent
    while (*p && iswspace(*p)) p++;
    if (*p != L'%') return FALSE;
    p++;
    // Skip trailing spaces
    while (*p && iswspace(*p)) p++;
    // Ensure end
    if (*p != 0) return FALSE;

    // Format A op (A × B/100)
    WCHAR aStr[64]; WCHAR decStr[64];
    format_double(A, aStr, 64);
    format_double(B / 100.0, decStr, 64);
    _snwprintf(dst, cap, L"%s %c (%s × %s)", aStr, op, aStr, decStr);
    dst[cap - 1] = 0;
    return TRUE;
}

static BOOL has_binary_operator(const WCHAR *s) {
    if (!s) return FALSE;
    int len = (int)wcslen(s);
    int i = 0;
    while (i < len && iswspace(s[i])) i++;
    int expectOperand = 1; // at start, unary +/− allowed
    for (; i < len; ++i) {
        WCHAR c = s[i];
        if (iswspace(c)) continue;
        if (c == L'(') { expectOperand = 1; continue; }
        if (c == L')') { expectOperand = 0; continue; }
        if (c == L'+' || c == L'-' || c == L'*' || c == L'/' || c == L'×' || c == L'÷') {
            if (c == L'-') {
                if (!expectOperand) return TRUE; // binary minus
            } else if (c == L'+') {
                if (!expectOperand) return TRUE; // binary plus
                // else unary plus; continue
            } else {
                // *, /, ×, ÷ are binary operators
                return TRUE;
            }
            expectOperand = 1;
            continue;
        }
        // Any other non-space char counts as part of an operand (digits, dot, etc.)
        expectOperand = 0;
    }
    return FALSE;
}

// Random number and dice rolling implementation
static void InitializeRandom(void) {
    if (!g_randomInitialized) {
        srand((unsigned int)GetTickCount());
        g_randomInitialized = TRUE;
    }
}

static int GetRandomNumber(int min, int max) {
    InitializeRandom();
    if (min > max) {
        int temp = min;
        min = max;
        max = temp;
    }
    return min + (rand() % (max - min + 1));
}

static BOOL Handle8BallCommand(HWND hwnd, const WCHAR *input) {
    // Extract question if present (we'll ignore it as specified)
    WCHAR response[256];
    int responseIndex = GetRandomNumber(0, k8BallResponseCount - 1);
    const WCHAR *selectedResponse = k8BallResponses[responseIndex];
    
    swprintf(response, sizeof(response)/sizeof(response[0]), L"🎱 %s", selectedResponse);
    History_Add(hwnd, response, NULL, NULL, TRUE);
    return TRUE;
}

static BOOL ParseDiceNotation(const WCHAR *input, int *diceCount, int *diceSides, int *modifier, WCHAR *errorMsg, size_t errorCap) {
    if (!input || !diceCount || !diceSides || !modifier) return FALSE;
    
    *diceCount = 1;
    *diceSides = 6;
    *modifier = 0;
    if (errorMsg && errorCap > 0) errorMsg[0] = 0;
    
    const WCHAR *p = input;
    // Skip whitespace
    while (*p && iswspace(*p)) p++;
    
    // Check if input starts with 'd' or 'D' - single die notation
    if (*p == L'd' || *p == L'D') {
        p++;
        if (iswdigit(*p)) {
            WCHAR *endptr;
            long sides = wcstol(p, &endptr, 10);
            if (sides < 2 || sides > 1000) {
                if (errorMsg && errorCap > 0) 
                    swprintf(errorMsg, errorCap, L"Die sides must be between 2 and 1000");
                return FALSE;
            }
            *diceCount = 1;
            *diceSides = (int)sides;
            p = endptr;
        }
    }
    // Check if input is just a number (like "12") - treat as d12
    else if (iswdigit(*p)) {
        WCHAR *endptr;
        long firstNum = wcstol(p, &endptr, 10);
        
        // Check what comes after the number
        const WCHAR *after = endptr;
        while (*after && iswspace(*after)) after++;
        
        if (*after == L'd' || *after == L'D') {
            // Format: "2d6" - count then 'd' then sides
            if (firstNum < 1 || firstNum > 100) {
                if (errorMsg && errorCap > 0) 
                    swprintf(errorMsg, errorCap, L"Dice count must be between 1 and 100");
                return FALSE;
            }
            *diceCount = (int)firstNum;
            p = after + 1;
            
            if (iswdigit(*p)) {
                WCHAR *endptr2;
                long sides = wcstol(p, &endptr2, 10);
                if (sides < 2 || sides > 1000) {
                    if (errorMsg && errorCap > 0) 
                        swprintf(errorMsg, errorCap, L"Die sides must be between 2 and 1000");
                    return FALSE;
                }
                *diceSides = (int)sides;
                p = endptr2;
            }
        } else if (*after == L'+' || *after == L'-' || *after == 0 || iswspace(*after)) {
            // Just a number (like "12" or "6 + 2") - treat as single die with that many sides
            if (firstNum < 2 || firstNum > 1000) {
                if (errorMsg && errorCap > 0) 
                    swprintf(errorMsg, errorCap, L"Die sides must be between 2 and 1000");
                return FALSE;
            }
            *diceCount = 1;
            *diceSides = (int)firstNum;
            p = endptr;
        } else {
            // Invalid format
            if (errorMsg && errorCap > 0) 
                swprintf(errorMsg, errorCap, L"Invalid dice notation");
            return FALSE;
        }
    }
    
    // Skip whitespace
    while (*p && iswspace(*p)) p++;
    
    // Parse modifier (optional)
    if (*p == L'+' || *p == L'-') {
        WCHAR sign = *p;
        p++;
        while (*p && iswspace(*p)) p++;
        
        if (iswdigit(*p)) {
            WCHAR *endptr;
            long mod = wcstol(p, &endptr, 10);
            if (mod > 1000) {
                if (errorMsg && errorCap > 0) 
                    swprintf(errorMsg, errorCap, L"Modifier too large");
                return FALSE;
            }
            *modifier = (sign == L'-') ? -(int)mod : (int)mod;
            p = endptr;
        } else {
            if (errorMsg && errorCap > 0) {
                swprintf(errorMsg, errorCap, L"Invalid modifier");
            }
            return FALSE;
        }
        
        // Parse additional modifiers
        while (*p) {
            while (*p && iswspace(*p)) p++;
            if (*p == L'+' || *p == L'-') {
                WCHAR sign2 = *p;
                p++;
                while (*p && iswspace(*p)) p++;
                
                if (iswdigit(*p)) {
                    WCHAR *endptr;
                    long mod = wcstol(p, &endptr, 10);
                    if (abs(*modifier) + mod > 1000) {
                        if (errorMsg && errorCap > 0) 
                            swprintf(errorMsg, errorCap, L"Total modifier too large");
                        return FALSE;
                    }
                    *modifier += (sign2 == L'-') ? -(int)mod : (int)mod;
                    p = endptr;
                } else {
                    if (errorMsg && errorCap > 0) {
                        swprintf(errorMsg, errorCap, L"Invalid modifier");
                    }
                    return FALSE;
                }
            } else {
                break;
            }
        }
    } else if (*p != 0) {
        if (errorMsg && errorCap > 0) {
            swprintf(errorMsg, errorCap, L"Invalid dice notation");
        }
        return FALSE;
    }

    while (*p && iswspace(*p)) p++;
    if (*p != 0) {
        if (errorMsg && errorCap > 0) {
            swprintf(errorMsg, errorCap, L"Unexpected trailing text");
        }
        return FALSE;
    }

    return TRUE;
}

static BOOL HandleRollCommand(HWND hwnd, const WCHAR *input) {
    // Skip "!roll" or "roll" and whitespace
    const WCHAR *p = input;
    if (*p == L'!' || *p == L'/') p++;
    while (*p && !iswspace(*p)) p++; // skip "roll"
    while (*p && iswspace(*p)) p++;
    
    int diceCount, diceSides, modifier;
    WCHAR errorMsg[256];
    
    if (!ParseDiceNotation(p, &diceCount, &diceSides, &modifier, errorMsg, sizeof(errorMsg))) {
        WCHAR response[512];
        swprintf(response, sizeof(response)/sizeof(response[0]), L"Error: %s", errorMsg);
        History_Add(hwnd, response, NULL, NULL, TRUE);
        return TRUE;
    }
    
    // Roll the dice
    WCHAR rollDetails[256] = {0};
    WCHAR response[512];
    int total = 0;
    
    if (diceCount == 1) {
        int roll = GetRandomNumber(1, diceSides);
        total = roll + modifier;
        
        if (modifier == 0) {
            swprintf(response, sizeof(response)/sizeof(response[0]), 
                L"You rolled a %d on a d%d", roll, diceSides);
        } else if (modifier > 0) {
            swprintf(response, sizeof(response)/sizeof(response[0]), 
                L"You rolled a %d on a d%d +%d = %d", roll, diceSides, modifier, total);
        } else {
            swprintf(response, sizeof(response)/sizeof(response[0]), 
                L"You rolled a %d on a d%d %d = %d", roll, diceSides, modifier, total);
        }
    } else {
        WCHAR rolls[512] = {0};
        int rollsLen = 0;
        
        for (int i = 0; i < diceCount; i++) {
            int roll = GetRandomNumber(1, diceSides);
            total += roll;
            
            if (i == 0) {
                rollsLen += swprintf(rolls + rollsLen, sizeof(rolls)/sizeof(rolls[0]) - rollsLen, L"%d", roll);
            } else if (i == diceCount - 1) {
                rollsLen += swprintf(rolls + rollsLen, sizeof(rolls)/sizeof(rolls[0]) - rollsLen, L", and %d", roll);
            } else {
                rollsLen += swprintf(rolls + rollsLen, sizeof(rolls)/sizeof(rolls[0]) - rollsLen, L", %d", roll);
            }
        }
        
        total += modifier;
        
        if (modifier == 0) {
            swprintf(response, sizeof(response)/sizeof(response[0]), 
                L"%d - Rolled %s", total, rolls);
        } else {
            swprintf(response, sizeof(response)/sizeof(response[0]), 
                L"%d - Rolled %s with %s%d modifier", 
                total, rolls, (modifier > 0) ? L"+" : L"", modifier);
        }
    }
    
    WCHAR totalStr[32];
    swprintf(totalStr, sizeof(totalStr)/sizeof(totalStr[0]), L"%d", total);
    History_Add(hwnd, response, totalStr, totalStr, TRUE);
    return TRUE;
}

static const WCHAR *kCmds[] = { L"help", L"settings", L"exit", L"clear", L"8ball", L"roll" };
static const int kCmdCount = 6;

static void UpdateCommandSuggestion(void) {
    g_ui.suggestActive = FALSE;
    g_ui.suggest[0] = 0;
    g_ui.suggestLen = 0;
    if (g_ui.inputLen <= 0) return;
    WCHAR pfx = g_ui.input[0];
    if (pfx != L'!' && pfx != L'/') return;
    // Accept only [a-z] letters after prefix for suggestion
    int i = 1;
    while (i < g_ui.inputLen && iswalpha(g_ui.input[i])) i++;
    // If there are extra non-alpha chars (like space) then stop suggesting
    if (i < g_ui.inputLen) return;
    // Extract typed part after prefix
    WCHAR typed[64];
    int tlen = g_ui.inputLen - 1;
    if (tlen < 0) tlen = 0;
    if (tlen >= (int)(sizeof(typed)/sizeof(typed[0])) - 1) tlen = (int)(sizeof(typed)/sizeof(typed[0])) - 2;
    if (tlen > 0) { wcsncpy(typed, g_ui.input + 1, tlen); typed[tlen] = 0; }
    else { typed[0] = 0; }
    // Default to help when no letters typed
    int chosen = -1;
    if (tlen == 0) {
        chosen = 0; // help
    } else {
        for (int c = 0; c < kCmdCount; ++c) {
            const WCHAR *cmd = kCmds[c];
            if (_wcsnicmp(cmd, typed, tlen) == 0) { chosen = c; break; }
        }
    }
    if (chosen < 0) return;
    const WCHAR *full = kCmds[chosen];
    int flen = (int)wcslen(full);
    if (tlen >= flen) return; // already complete
    // Suggest the remaining suffix only
    const WCHAR *suffix = full + tlen;
    wcsncpy(g_ui.suggest, suffix, (sizeof(g_ui.suggest)/sizeof(g_ui.suggest[0])) - 1);
    g_ui.suggest[(sizeof(g_ui.suggest)/sizeof(g_ui.suggest[0])) - 1] = 0;
    g_ui.suggestLen = (int)wcslen(g_ui.suggest);
    g_ui.suggestActive = (g_ui.suggestLen > 0);
}

static BOOL ApplySuggestionToInput(void) {
    if (!g_ui.suggestActive || g_ui.suggestLen <= 0) return FALSE;
    if (g_ui.inputLen <= 0) return FALSE;
    WCHAR pfx = g_ui.input[0];
    if (pfx != L'!' && pfx != L'/') return FALSE;
    size_t cur = (size_t)g_ui.inputLen;
    size_t suf = (size_t)g_ui.suggestLen;
    size_t cap = (sizeof(g_ui.input)/sizeof(g_ui.input[0])) - 1;
    if (cur + suf > cap) suf = (cap > cur) ? (cap - cur) : 0;
    if (suf == 0) return FALSE;
    wmemcpy(g_ui.input + cur, g_ui.suggest, suf);
    g_ui.input[cur + suf] = 0;
    g_ui.inputLen = (int)(cur + suf);
    g_ui.caretPos = g_ui.inputLen;
    ClearInputSelection();
    g_ui.suggestActive = FALSE;
    g_ui.suggest[0] = 0;
    g_ui.suggestLen = 0;
    return TRUE;
}

static void UpdateGhostParens(void) {
    g_ui.ghostParensActive = FALSE;
    g_ui.ghostParens[0] = 0;
    g_ui.ghostParensLen = 0;
    if (g_ui.inputLen <= 0) return;
    // Don't show ghost parens for commands
    WCHAR pfx = g_ui.input[0];
    if (pfx == L'!' || pfx == L'/') return;
    int depth = 0;
    for (int i = 0; i < g_ui.inputLen; ++i) {
        WCHAR c = g_ui.input[i];
        if (c == L'(') depth++;
        else if (c == L')') { if (depth > 0) depth--; }
    }
    if (depth <= 0) return;
    if (depth > (int)((sizeof(g_ui.ghostParens)/sizeof(g_ui.ghostParens[0])) - 1)) {
        depth = (int)((sizeof(g_ui.ghostParens)/sizeof(g_ui.ghostParens[0])) - 1);
    }
    for (int i = 0; i < depth; ++i) g_ui.ghostParens[i] = L')';
    g_ui.ghostParens[depth] = 0;
    g_ui.ghostParensLen = depth;
    g_ui.ghostParensActive = TRUE;
}

static BOOL is_word_char(WCHAR c) {
    return iswalnum(c) || c == L'_' || c == L'.';
}

static int PrevWordPos(const WCHAR *s, int len, int pos) {
    int i = pos;
    if (i <= 0) return 0;
    i--; // start left of caret
    (void)len; // silence unused parameter
    while (i > 0 && iswspace(s[i])) i--;
    if (i < 0) return 0;
    if (is_word_char(s[i])) {
        while (i > 0 && is_word_char(s[i-1])) i--;
        return i;
    } else {
        while (i > 0 && !iswspace(s[i-1]) && !is_word_char(s[i-1])) i--;
        return i;
    }
}

static int NextWordPos(const WCHAR *s, int len, int pos) {
    int i = pos;
    if (i >= len) return len;
    while (i < len && iswspace(s[i])) i++;
    if (i >= len) return len;
    if (is_word_char(s[i])) {
        while (i < len && is_word_char(s[i])) i++;
        return i;
    } else {
        while (i < len && !iswspace(s[i]) && !is_word_char(s[i])) i++;
        return i;
    }
}

static int ClampInputPos(int pos) {
    if (pos < 0) return 0;
    if (pos > g_ui.inputLen) return g_ui.inputLen;
    return pos;
}

static BOOL HasInputSelection(void) {
    return g_ui.selStart != g_ui.selEnd;
}

static void ClearInputSelection(void) {
    int pos = ClampInputPos(g_ui.caretPos);
    g_ui.caretPos = pos;
    g_ui.selAnchor = pos;
    g_ui.selStart = pos;
    g_ui.selEnd = pos;
}

static void SetInputSelection(int anchor, int caret) {
    anchor = ClampInputPos(anchor);
    caret = ClampInputPos(caret);
    g_ui.selAnchor = anchor;
    g_ui.caretPos = caret;
    if (anchor <= caret) {
        g_ui.selStart = anchor;
        g_ui.selEnd = caret;
    } else {
        g_ui.selStart = caret;
        g_ui.selEnd = anchor;
    }
}

static BOOL DeleteSelectedInput(void) {
    if (!HasInputSelection()) return FALSE;
    int start = g_ui.selStart;
    int end = g_ui.selEnd;
    memmove(g_ui.input + start,
            g_ui.input + end,
            (g_ui.inputLen - end + 1) * sizeof(WCHAR));
    g_ui.inputLen -= (end - start);
    g_ui.caretPos = start;
    ClearInputSelection();
    return TRUE;
}

static int MeasureTextWidth(HDC hdc, const WCHAR *text, int len) {
    SIZE sz = {0};
    if (!hdc || !text || len <= 0) return 0;
    GetTextExtentPoint32W(hdc, text, len, &sz);
    return sz.cx;
}

static int CaretIndexFromPoint(HWND hwnd, RECT inText, int x) {
    int rel = x - inText.left;
    if (rel <= 0) return 0;
    if (g_ui.inputLen <= 0) return 0;
    HDC hdc = GetDC(hwnd);
    HFONT f = CreateUiFont(FONT_INPUT, FALSE);
    HFONT old = (HFONT)SelectObject(hdc, f);
    int fit = 0;
    int lo = 0, hi = g_ui.inputLen;
    while (lo < hi) {
        int mid = (lo + hi + 1) / 2;
        int width = MeasureTextWidth(hdc, g_ui.input, mid);
        if (width <= rel) lo = mid; else hi = mid - 1;
    }
    fit = lo;
    SelectObject(hdc, old);
    ReleaseUiFont(f, FONT_INPUT, FALSE);
    ReleaseDC(hwnd, hdc);
    if (fit < 0) fit = 0;
    if (fit > g_ui.inputLen) fit = g_ui.inputLen;
    return fit;
}

static BOOL IsIncompleteString(const WCHAR *input) {
    if (!input || !*input) return FALSE;

    // Let fully-formed generic conversions through before the heuristic checks below.
    if (ParseGenericUnitConversion(input, NULL)) return FALSE;
    
    int len = (int)wcslen(input);
    
    // Check for trailing operators
    WCHAR lastChar = input[len - 1];
    if (lastChar == L'+' || lastChar == L'-' || lastChar == L'*' || lastChar == L'/' || 
        lastChar == L'×' || lastChar == L'÷' || lastChar == L'=' || lastChar == L'(') {
        return TRUE;
    }
    
    // Check for expressions with = that should be treated as incomplete algebraic equations
    // Examples: "33=3", "x=5", but NOT "500+3*2=" which is just a trailing =
    const WCHAR *eqPos = wcsstr(input, L"=");
    if (eqPos) {
        // If = is at the end, it's just a trailing operator (handled above)
        // Only treat as algebraic if there's content after the =
        const WCHAR *afterEq = eqPos + 1;
        while (*afterEq && iswspace(*afterEq)) afterEq++;
        
        if (*afterEq) {
            // There's content after =, check if it looks like an algebraic equation
            // Look for patterns like "number=number" or "variable=something"
            const WCHAR *beforeEq = eqPos - 1;
            while (beforeEq >= input && iswspace(*beforeEq)) beforeEq--;
            
            // Check if left side is just a simple number/variable and right side exists
            // This catches "33=3", "x=5" but allows unit conversions to work
            BOOL isSimpleLeft = TRUE;
            for (const WCHAR *p = input; p <= beforeEq; p++) {
                if (!iswdigit(*p) && !iswspace(*p) && !iswalpha(*p) && *p != L'.') {
                    isSimpleLeft = FALSE;
                    break;
                }
            }
            
            if (isSimpleLeft && beforeEq >= input) {
                return TRUE; // Treat as incomplete algebraic equation
            }
        }
    }
    
    // Check for numbers followed by invalid unit sequences
    for (const WCHAR *p = input; *p; p++) {
        if (iswdigit(*p)) {
            // Skip the number
            while (*p && (iswdigit(*p) || *p == L'.')) p++;
            if (!*p) break;
            
            // Skip spaces
            while (*p && iswspace(*p)) p++;
            if (!*p) break;
            
            // Check if followed by alphabetic characters
            if (iswalpha(*p) || IsMicroChar(*p)) {
                const WCHAR *unitStart = p;
                while (*p && (iswalpha(*p) || IsMicroChar(*p))) p++;
                
                // Check if this forms a valid unit
                int unitLen = (int)(p - unitStart);
                WCHAR unitBuf[32];
                if (unitLen < 32) {
                    wcsncpy(unitBuf, unitStart, unitLen);
                    unitBuf[unitLen] = 0;
                    
                    // Check if it's a known unit
                    int tempUnitLen = 0;
                    int tempId = match_temp_unit_word(unitBuf, &tempUnitLen);
                    if (tempId < 0 || tempUnitLen != unitLen) {
                        int regularUnitLen = 0;
                        UnitId unitId = match_unit_word(unitBuf, &regularUnitLen);
                        if (unitId == U_INVALID || regularUnitLen != unitLen) {
                            // Check time units as valid too
                            int i = 0; WCHAR buf[32];
                            while (unitBuf[i] && i < 31) { buf[i] = (WCHAR)towlower(unitBuf[i]); i++; }
                            buf[i] = 0;
                            BOOL ok = FALSE;
                            for (int k = 0; k < (int)(sizeof(kTimeUnitWords)/sizeof(kTimeUnitWords[0])); ++k) {
                                if (wcsieq(buf, kTimeUnitWords[k].word)) { ok = TRUE; break; }
                            }
                            if (!ok) {
                                ok = IsKnownCompactGenericUnitToken(unitBuf, unitLen);
                            }
                            if (!ok) return TRUE; // Invalid unit
                        }
                    }
                }
                p--; // Back up since the for loop will increment
            }
        }
    }
    
    return FALSE;
}

static void UpdateLivePreview(HWND hwnd){
    if (g_ui.inputLen <= 0) { g_liveVisible = FALSE; g_liveText[0] = 0; g_liveComplete = FALSE; g_timePreviewActive = FALSE; AdjustWindowHeight(hwnd); InvalidateRect(hwnd, NULL, FALSE); return; }
    
    // Time queries (live, dynamic)
    TimeQuery tq = {0};
    if (DetectTimeQuery(g_ui.input, &tq) && tq.type != TQ_NONE) {
        SYSTEMTIME nowLocal; GetLocalTime(&nowLocal);
        SYSTEMTIME nowUtc; TzSpecificLocalTimeToSystemTime(NULL, &nowLocal, &nowUtc);
        FILETIME ftNow; SystemTimeToFileTime(&nowUtc, &ftNow);
        WCHAR leftBuf[256];
        switch (tq.type) {
            case TQ_NOW: {
                WCHAR cur[128]; FormatLocalDateTime(&nowLocal, cur, 128);
                swprintf(g_liveText, sizeof(g_liveText)/sizeof(g_liveText[0]), L"%s", cur);
                break;
            }
            case TQ_UNTIL: {
                SYSTEMTIME tUtc; TzSpecificLocalTimeToSystemTime(NULL, &tq.a, &tUtc);
                FILETIME ftA; SystemTimeToFileTime(&tUtc, &ftA);
                ULONGLONG now = FileTimeToULL(ftNow), a = FileTimeToULL(ftA);
                ULONGLONG diffSec = (now < a) ? (a - now) / 10000000ULL : 0ULL;
                FormatLocalDateTime(&tq.a, leftBuf, 256);
                if (tq.hasUnit && tq.outUnit != TU_DEFAULT) {
                    const WCHAR *unitLabel = NULL;
                    double v = ConvertDurationToUnit((double)diffSec, tq.outUnit, &unitLabel);
                    WCHAR num[64]; format_double_3(v, num, 64);
                    swprintf(g_liveText, sizeof(g_liveText)/sizeof(g_liveText[0]), L"%s until %s: %s %s",
                             TimeOutUnitLabel(tq.outUnit),
                             leftBuf, num, unitLabel);
                } else {
                    WCHAR dur[128]; FormatDuration(diffSec, dur, 128);
                    swprintf(g_liveText, sizeof(g_liveText)/sizeof(g_liveText[0]), L"Time until %s: %s", leftBuf, dur);
                }
                break;
            }
            case TQ_SINCE: {
                SYSTEMTIME tUtc; TzSpecificLocalTimeToSystemTime(NULL, &tq.a, &tUtc);
                FILETIME ftA; SystemTimeToFileTime(&tUtc, &ftA);
                ULONGLONG now = FileTimeToULL(ftNow), a = FileTimeToULL(ftA);
                ULONGLONG diffSec = (now > a) ? (now - a) / 10000000ULL : 0ULL;
                FormatLocalDateTime(&tq.a, leftBuf, 256);
                if (tq.hasUnit && tq.outUnit != TU_DEFAULT) {
                    const WCHAR *unitLabel = NULL;
                    double v = ConvertDurationToUnit((double)diffSec, tq.outUnit, &unitLabel);
                    WCHAR num[64]; format_double_3(v, num, 64);
                    swprintf(g_liveText, sizeof(g_liveText)/sizeof(g_liveText[0]), L"%s since %s: %s %s",
                             TimeOutUnitLabel(tq.outUnit),
                             leftBuf, num, unitLabel);
                } else {
                    WCHAR dur[128]; FormatDuration(diffSec, dur, 128);
                    swprintf(g_liveText, sizeof(g_liveText)/sizeof(g_liveText[0]), L"Time since %s: %s", leftBuf, dur);
                }
                break;
            }
            case TQ_BETWEEN: {
                SYSTEMTIME aUtc, bUtc; TzSpecificLocalTimeToSystemTime(NULL, &tq.a, &aUtc); TzSpecificLocalTimeToSystemTime(NULL, &tq.b, &bUtc);
                FILETIME ftA, ftB; SystemTimeToFileTime(&aUtc, &ftA); SystemTimeToFileTime(&bUtc, &ftB);
                ULONGLONG A = FileTimeToULL(ftA), B = FileTimeToULL(ftB);
                ULONGLONG diffSec = (A > B ? A - B : B - A) / 10000000ULL;
                WCHAR aStr[128], bStr[128]; FormatLocalDateTime(&tq.a, aStr, 128); FormatLocalDateTime(&tq.b, bStr, 128);
                if (tq.hasUnit && tq.outUnit != TU_DEFAULT) {
                    const WCHAR *unitLabel = NULL;
                    double v = ConvertDurationToUnit((double)diffSec, tq.outUnit, &unitLabel);
                    WCHAR num[64]; format_double(v, num, 64);
                    swprintf(g_liveText, sizeof(g_liveText)/sizeof(g_liveText[0]), L"%s between %s and %s: %s %s",
                             TimeOutUnitLabel(tq.outUnit),
                             aStr, bStr, num, unitLabel);
                } else {
                    WCHAR dur[128]; FormatDuration(diffSec, dur, 128);
                    swprintf(g_liveText, sizeof(g_liveText)/sizeof(g_liveText[0]), L"Between %s and %s: %s", aStr, bStr, dur);
                }
                break;
            }
            default: break;
        }
        g_liveComplete = TRUE;
        g_liveVisible = TRUE;
        g_timePreviewActive = TRUE;
        AdjustWindowHeight(hwnd);
        InvalidateRect(hwnd, NULL, FALSE);
        return;
    } else {
        g_timePreviewActive = FALSE;
    }
    
    // Comparison queries like: "is 11 in > 100 cm"
    {
        WCHAR cmpText[512]; BOOL cmpComplete = FALSE; WCHAR ansBuf[8];
        if (BuildComparisonAnswer(g_ui.input, TRUE, cmpText, sizeof(cmpText)/sizeof(cmpText[0]), ansBuf, sizeof(ansBuf)/sizeof(ansBuf[0]), &cmpComplete)) {
            wcsncpy(g_liveText, cmpText, sizeof(g_liveText)/sizeof(g_liveText[0]) - 1);
            g_liveText[sizeof(g_liveText)/sizeof(g_liveText[0]) - 1] = 0;
            g_liveVisible = TRUE;
            g_liveComplete = cmpComplete;
            AdjustWindowHeight(hwnd);
            InvalidateRect(hwnd, NULL, FALSE);
            return;
        }
    }
    
    // Check for incomplete/broken strings first
    if (IsIncompleteString(g_ui.input)) {
        wcscpy(g_liveText, L"...");
        g_liveVisible = TRUE;
        g_liveComplete = FALSE;
        AdjustWindowHeight(hwnd);
        InvalidateRect(hwnd, NULL, FALSE);
        return;
    }
    

    GenericConversionResult convPrev;
    if (ParseGenericUnitConversion(g_ui.input, &convPrev)) {
        WCHAR valueStr[64]; format_double(convPrev.inputValue, valueStr, 64);
        WCHAR resultStr[64]; format_double(convPrev.resultValue, resultStr, 64);
        const WCHAR *srcName = pick_unit_name(convPrev.sourceUnit, convPrev.inputValue);
        const WCHAR *dstName = pick_unit_name(convPrev.targetUnit, convPrev.resultValue);
        swprintf(g_liveText, sizeof(g_liveText)/sizeof(g_liveText[0]), L"%s %s = %s %s", valueStr, srcName, resultStr, dstName);
        g_liveVisible = TRUE;
        g_liveComplete = TRUE;
        AdjustWindowHeight(hwnd);
        InvalidateRect(hwnd, NULL, FALSE);
        return;
    }

    UpdateCommandSuggestion();
    UpdateGhostParens();
    // Normalize NL tokens, then preprocess units
    WCHAR tmpIn[512]; normalize_percent_word(g_ui.input, tmpIn, 512);
    WCHAR ofIn[512]; normalize_of_operator(tmpIn, ofIn, 512);
    WCHAR opIn[512]; normalize_operators_basic(ofIn, opIn, 512);
    UnitsPre up = {0}; preprocess_units(opIn, &up);
    if (HasIncompatibleConversionDomain(&up)) {
        wcscpy(g_liveText, L"...");
        g_liveVisible = TRUE;
        g_liveComplete = FALSE;
        AdjustWindowHeight(hwnd);
        InvalidateRect(hwnd, NULL, FALSE);
        return;
    }
    const WCHAR *evalIn = up.eval[0] ? up.eval : opIn;
    double val; BOOL complete; int consumed;
    if (!eval_expression(evalIn, &val, &complete, &consumed) || consumed <= 0) {
        // Show "..." for expressions that don't evaluate
        wcscpy(g_liveText, L"...");
        g_liveVisible = TRUE;
        g_liveComplete = FALSE;
        AdjustWindowHeight(hwnd);
        InvalidateRect(hwnd, NULL, FALSE);
        return;
    }    if ((up.hasUnits && !up.hasTarget) || (up.hasTemp && !up.hasTempTarget)) {
        g_liveVisible = FALSE; g_liveText[0] = 0; g_liveComplete = FALSE;
        AdjustWindowHeight(hwnd); InvalidateRect(hwnd, NULL, FALSE); return;
    }
    if (up.hasTemp) {
        // Temperature conversion path
        int tgtT = up.hasTempTarget ? up.tempTarget : 0; // default °C for commit; live preview guarded above
        double outT = kelvin_to_temp(val, tgtT);
        WCHAR num[64]; format_unit_value(U_M, outT, num, 64); // reuse rounding
        WCHAR symBuf[4]; symBuf[0] = (tgtT==2? L'K' : DEG); symBuf[1] = (tgtT==2? 0 : (tgtT==0? L'C' : L'F')); symBuf[2] = 0; const WCHAR *sym = symBuf;
        if (complete) {
            const WCHAR *left = up.prettyLeftFull[0] ? up.prettyLeftFull : (up.prettyLeft[0] ? up.prettyLeft : evalIn);
            {
                WCHAR spaced[512]; format_expr_spaced(left, spaced, 512);
                swprintf(g_liveText, sizeof(g_liveText)/sizeof(g_liveText[0]), L"%s = %s %s", spaced, num, sym);
            }
        } else {
            wcsncpy(g_liveText, num, sizeof(g_liveText)/sizeof(g_liveText[0]) - 1);
            g_liveText[sizeof(g_liveText)/sizeof(g_liveText[0]) - 1] = 0;
        }
    } else if (up.hasUnits || up.hasTarget) {
        // Length conversion path - check for mixed units first
        if (up.unitSystem == UNIT_SYSTEM_MIXED || up.unitSystem == UNIT_SYSTEM_IMPERIAL || up.unitSystem == UNIT_SYSTEM_METRIC) {
            // Use mixed unit formatting
            MixedUnitResult mixedResult = {0};
            mixedResult.value = val; // val should already be in meters from eval_expression
            mixedResult.inputSystem = up.unitSystem;
            mixedResult.hasExplicitTarget = up.hasTarget;
            mixedResult.targetUnit = up.hasTarget ? up.target : U_INVALID;
            
            WCHAR formattedResult[256];
            format_mixed_result(&mixedResult, formattedResult, sizeof(formattedResult)/sizeof(formattedResult[0]));
            
            if (complete) {
                const WCHAR *left = up.prettyLeftFull[0] ? up.prettyLeftFull : (up.prettyLeft[0] ? up.prettyLeft : evalIn);
                {
                    WCHAR spaced[512]; format_expr_spaced(left, spaced, 512);
                    swprintf(g_liveText, sizeof(g_liveText)/sizeof(g_liveText[0]), L"%s = %s", spaced, formattedResult);
                }
            } else {
                wcsncpy(g_liveText, formattedResult, sizeof(g_liveText)/sizeof(g_liveText[0]) - 1);
                g_liveText[sizeof(g_liveText)/sizeof(g_liveText[0]) - 1] = 0;
            }
        } else {
            // Original single unit conversion path
            UnitId tgt = up.hasTarget ? up.target : U_IN;
            double outVal = val / unit_to_m(tgt);
            WCHAR num[64]; format_unit_value(tgt, outVal, num, 64);
            WCHAR unitName[32]; unit_name_full(tgt, outVal, unitName, 32);
            if (complete) {
                const WCHAR *left = up.prettyLeftFull[0] ? up.prettyLeftFull : (up.prettyLeft[0] ? up.prettyLeft : evalIn);
                {
                    WCHAR spaced[512]; format_expr_spaced(left, spaced, 512);
                    swprintf(g_liveText, sizeof(g_liveText)/sizeof(g_liveText[0]), L"%s = %s %s", spaced, num, unitName);
                }
            } else {
                wcsncpy(g_liveText, num, sizeof(g_liveText)/sizeof(g_liveText[0]) - 1);
                g_liveText[sizeof(g_liveText)/sizeof(g_liveText[0]) - 1] = 0;
            }
        }
    } else if (up.hasTime || up.hasTimeTarget) {
        // Time conversion path
        TimeUnitId tgt = up.hasTimeTarget ? up.timeTarget : T_S;
        double secondsVal = val; // already normalized to seconds in eval
        double outVal = secondsVal / kTimeUnits[tgt].to_s;
        WCHAR num[64]; format_double(outVal, num, 64);
        const WCHAR *unitName = (fabs(outVal - 1.0) < 1e-12) ? kTimeUnits[tgt].name_singular : kTimeUnits[tgt].name_plural;
        if (complete) {
            const WCHAR *left = up.prettyLeftFull[0] ? up.prettyLeftFull : (up.prettyLeft[0] ? up.prettyLeft : evalIn);
            {
                WCHAR spaced[512]; format_expr_spaced(left, spaced, 512);
                swprintf(g_liveText, sizeof(g_liveText)/sizeof(g_liveText[0]), L"%s = %s %s", spaced, num, unitName);
            }
        } else {
            wcsncpy(g_liveText, num, sizeof(g_liveText)/sizeof(g_liveText[0]) - 1);
            g_liveText[sizeof(g_liveText)/sizeof(g_liveText[0]) - 1] = 0;
        }
    } else {
        // No units: behave as before
        WCHAR num[64]; format_double(val, num, 64);
        if (complete) {
            if (has_binary_operator(evalIn)) {
                WCHAR spaced[512]; format_expr_spaced(evalIn, spaced, 512);
                WCHAR pretty[512];
                if (!PrettyPercentOfLeft(spaced, pretty, 512)) {
                    // Build pretty by copying spaced and swapping ASCII ops to Unicode
                    size_t cap = 512; size_t di = 0; for (const WCHAR *p = spaced; *p && di < cap - 1; ++p) { WCHAR c = *p; if (c == L'*') c = CH_MUL; else if (c == L'/') c = CH_DIV; pretty[di++] = c; } pretty[di] = 0;
                }
                swprintf(g_liveText, sizeof(g_liveText)/sizeof(g_liveText[0]), L"%s = %s", pretty, num);
            } else {
                wcsncpy(g_liveText, num, sizeof(g_liveText)/sizeof(g_liveText[0]) - 1);
                g_liveText[sizeof(g_liveText)/sizeof(g_liveText[0]) - 1] = 0;
            }
        } else {
            wcsncpy(g_liveText, num, sizeof(g_liveText)/sizeof(g_liveText[0]) - 1);
            g_liveText[sizeof(g_liveText)/sizeof(g_liveText[0]) - 1] = 0;
        }
    }
    g_liveComplete = complete;
    g_liveVisible = TRUE;
    AdjustWindowHeight(hwnd);
    InvalidateRect(hwnd, NULL, FALSE);
}

static BOOL AppendToInput(const WCHAR *s) {
    if (!s || !*s) return FALSE;
    size_t cur = (size_t)g_ui.inputLen;
    size_t add = wcslen(s);
    size_t cap = (sizeof(g_ui.input)/sizeof(g_ui.input[0])) - 1;
    if (cur >= cap) return FALSE;
    if (cur + add > cap) add = cap - cur;
    wmemcpy(g_ui.input + cur, s, add);
    g_ui.input[cur + add] = 0;
    g_ui.inputLen = (int)(cur + add);
    g_ui.suppressLastAnsSuggest = FALSE;
    g_ui.caretPos = g_ui.inputLen;
    ClearInputSelection();
    return add > 0;
}

static void ProcessInputExpression(HWND hwnd) {
    // Check for incomplete/broken strings first
    if (IsIncompleteString(g_ui.input)) {
        History_Add(hwnd, L"Invalid Input - Broken/Incomplete String", g_ui.input, NULL, TRUE);
        return;
    }
    if (HandleGenericUnitConversion(hwnd, g_ui.input)) { return; }
    // Time queries: snapshot numeric value at commit, save to history and suggestion
    TimeQuery tq = {0};
    if (DetectTimeQuery(g_ui.input, &tq) && tq.type != TQ_NONE) {
        SYSTEMTIME nowLocal; GetLocalTime(&nowLocal);
        SYSTEMTIME nowUtc; TzSpecificLocalTimeToSystemTime(NULL, &nowLocal, &nowUtc);
        FILETIME ftNow; SystemTimeToFileTime(&nowUtc, &ftNow);
        WCHAR line[720];
        if (tq.type == TQ_NOW) {
            WCHAR stamp[128]; FormatLocalDateTime(&nowLocal, stamp, 128);
            swprintf(line, sizeof(line)/sizeof(line[0]), L"now = %s", stamp);
            History_Add(hwnd, line, L"now", stamp, TRUE);
            return;
        }
        // compute seconds diff
        ULONGLONG diffSec = 0ULL;
        if (tq.type == TQ_UNTIL || tq.type == TQ_SINCE) {
            SYSTEMTIME tUtc; TzSpecificLocalTimeToSystemTime(NULL, &tq.a, &tUtc);
            FILETIME ftA; SystemTimeToFileTime(&tUtc, &ftA);
            ULONGLONG now = FileTimeToULL(ftNow), a = FileTimeToULL(ftA);
            diffSec = (tq.type == TQ_UNTIL) ? ((now < a) ? (a - now) / 10000000ULL : 0ULL)
                                           : ((now > a) ? (now - a) / 10000000ULL : 0ULL);
            WCHAR left[256]; FormatLocalDateTime(&tq.a, left, 256);
            // unit conversion
            TimeOutUnit outU = tq.hasUnit ? tq.outUnit : TU_DEFAULT;
            if (outU == TU_DEFAULT) outU = TU_SECONDS; // default store seconds
            const WCHAR *unitLabel = NULL;
            double v = ConvertDurationToUnit((double)diffSec, outU, &unitLabel);
            WCHAR numDisp[64]; format_double_3(v, numDisp, 64);
            WCHAR numFull[64]; format_double_full(v, numFull, 64);
            const WCHAR *verb = (tq.type == TQ_UNTIL) ? L"until" : L"since";
            swprintf(line, sizeof(line)/sizeof(line[0]), L"%s %s %s = %s %s",
                     TimeOutUnitLabel(outU),
                     verb, left, numDisp, unitLabel);
            History_Add(hwnd, line, g_ui.input, numFull, TRUE);
            return;
        }
        if (tq.type == TQ_BETWEEN) {
            SYSTEMTIME aUtc, bUtc; TzSpecificLocalTimeToSystemTime(NULL, &tq.a, &aUtc); TzSpecificLocalTimeToSystemTime(NULL, &tq.b, &bUtc);
            FILETIME ftA, ftB; SystemTimeToFileTime(&aUtc, &ftA); SystemTimeToFileTime(&bUtc, &ftB);
            ULONGLONG A = FileTimeToULL(ftA), B = FileTimeToULL(ftB);
            diffSec = (A > B ? A - B : B - A) / 10000000ULL;
            WCHAR aStr[128], bStr[128]; FormatLocalDateTime(&tq.a, aStr, 128); FormatLocalDateTime(&tq.b, bStr, 128);
            TimeOutUnit outU = tq.hasUnit ? tq.outUnit : TU_SECONDS;
            const WCHAR *unitLabel = NULL;
            double v = ConvertDurationToUnit((double)diffSec, outU, &unitLabel);
            WCHAR numDisp[64]; format_double(v, numDisp, 64);
            WCHAR numFull[64]; format_double_full(v, numFull, 64);
            swprintf(line, sizeof(line)/sizeof(line[0]), L"%s between %s and %s = %s %s",
                     TimeOutUnitLabel(outU),
                     aStr, bStr, numDisp, unitLabel);
            History_Add(hwnd, line, g_ui.input, numFull, TRUE);
            return;
        }
        // Fallback
        History_Add(hwnd, g_ui.input, g_ui.input, NULL, TRUE);
        return;
    }
    
    if (HandleHolidayQuery(hwnd, g_ui.input)) {
        return;
    }

    // Commit expression or echo input
    // First, handle comparison queries (e.g., "is 11 in > 100 cm")
    {
        WCHAR line[720]; WCHAR ans[8]; BOOL cmpComplete = FALSE;
        if (BuildComparisonAnswer(g_ui.input, FALSE, line, sizeof(line)/sizeof(line[0]), ans, sizeof(ans)/sizeof(ans[0]), &cmpComplete)) {
            WCHAR spaced[512]; format_expr_spaced(g_ui.input, spaced, 512);
            History_Add(hwnd, line, spaced, ans, TRUE);
            return;
        }
    }

    double val; BOOL complete; int consumed;
    WCHAR tmpIn[512]; normalize_percent_word(g_ui.input, tmpIn, 512);
    WCHAR ofIn[512]; normalize_of_operator(tmpIn, ofIn, 512);
    WCHAR opIn[512]; normalize_operators_basic(ofIn, opIn, 512);
    UnitsPre up = {0}; preprocess_units(opIn, &up);
    if (HasIncompatibleConversionDomain(&up)) {
        History_Add(hwnd, L"Invalid Input - Incompatible unit conversion", g_ui.input, NULL, TRUE);
        return;
    }
    const WCHAR *evalIn = up.eval[0] ? up.eval : opIn;
    
    if (eval_expression(evalIn, &val, &complete, &consumed) && consumed > 0 && complete) {
        if (up.hasTempTarget) {
            // Temperature conversion only when explicit target provided
            int tgtT = up.tempTarget;
            double outT = kelvin_to_temp(val, tgtT);
            WCHAR num[64]; format_unit_value(U_M, outT, num, 64); // display
            WCHAR numFull[64]; format_double_full(outT, numFull, 64); // precise insert
            WCHAR symBuf[4]; symBuf[0] = (tgtT==2? L'K' : DEG); symBuf[1] = (tgtT==2? 0 : (tgtT==0? L'C' : L'F')); symBuf[2] = 0; const WCHAR *sym = symBuf;
            const WCHAR *left = up.prettyLeftFull[0] ? up.prettyLeftFull : (up.prettyLeft[0] ? up.prettyLeft : evalIn);
            WCHAR spaced[512]; format_expr_spaced(left, spaced, 512);
            WCHAR line[720]; swprintf(line, sizeof(line)/sizeof(line[0]), L"%s = %s %s", spaced, num, sym);
            History_Add(hwnd, line, spaced, numFull, TRUE);
        } else if (up.hasTemp && !up.hasTempTarget) {
            // Do not assume default temp target; just echo input
            History_Add(hwnd, g_ui.input, g_ui.input, NULL, TRUE);
        } else if (up.hasTarget) {
            // Length conversion with explicit target: compute target numeric + display string
            UnitId tgt = up.target;
            double outVal = val / unit_to_m(tgt);
            WCHAR num[64]; format_unit_value(tgt, outVal, num, 64); // display
            WCHAR numFull[64]; format_double_full(outVal, numFull, 64); // precise insert
            WCHAR unitName[32]; unit_name_full(tgt, outVal, unitName, 32);
            const WCHAR *left = up.prettyLeftFull[0] ? up.prettyLeftFull : (up.prettyLeft[0] ? up.prettyLeft : evalIn);
            WCHAR spaced[512]; format_expr_spaced(left, spaced, 512);
            WCHAR line[700]; swprintf(line, sizeof(line)/sizeof(line[0]), L"%s = %s %s", spaced, num, unitName);
            History_Add(hwnd, line, spaced, numFull, TRUE);
        } else if (up.hasUnits && !up.hasTarget) {
            // Mixed unit arithmetic without explicit target: show formatted result, but don't override last-answer insertion
            MixedUnitResult mixedResult = {0};
            mixedResult.value = val;
            mixedResult.inputSystem = up.unitSystem;
            mixedResult.hasExplicitTarget = FALSE;
            WCHAR formattedResult[256];
            format_mixed_result(&mixedResult, formattedResult, sizeof(formattedResult)/sizeof(formattedResult[0]));
            const WCHAR *left = up.prettyLeftFull[0] ? up.prettyLeftFull : (up.prettyLeft[0] ? up.prettyLeft : evalIn);
            WCHAR spaced[512]; format_expr_spaced(left, spaced, 512);
            WCHAR line[700]; swprintf(line, sizeof(line)/sizeof(line[0]), L"%s = %s", spaced, formattedResult);
            History_Add(hwnd, line, spaced, NULL, TRUE);
        } else if (up.hasTimeTarget) {
            // Time conversion only when explicit target provided
            TimeUnitId tgt = up.timeTarget;
            double outVal = val / kTimeUnits[tgt].to_s;
            WCHAR num[64]; format_double(outVal, num, 64);
            WCHAR numFull[64]; format_double_full(outVal, numFull, 64);
            const WCHAR *unitName = (fabs(outVal - 1.0) < 1e-12) ? kTimeUnits[tgt].name_singular : kTimeUnits[tgt].name_plural;
            const WCHAR *left = up.prettyLeftFull[0] ? up.prettyLeftFull : (up.prettyLeft[0] ? up.prettyLeft : evalIn);
            WCHAR spaced[512]; format_expr_spaced(left, spaced, 512);
            WCHAR line[720]; swprintf(line, sizeof(line)/sizeof(line[0]), L"%s = %s %s", spaced, num, unitName);
            History_Add(hwnd, line, spaced, numFull, TRUE);
        } else if (up.hasTime && !up.hasTimeTarget) {
            // Smart time formatting without explicit target
            WCHAR smartTime[256];
            format_time_smart(val, smartTime, sizeof(smartTime)/sizeof(smartTime[0]));
            const WCHAR *left = up.prettyLeftFull[0] ? up.prettyLeftFull : (up.prettyLeft[0] ? up.prettyLeft : evalIn);
            WCHAR spaced[512]; format_expr_spaced(left, spaced, 512);
            WCHAR line[720]; swprintf(line, sizeof(line)/sizeof(line[0]), L"%s = %s", spaced, smartTime);
            History_Add(hwnd, line, spaced, NULL, TRUE);
        } else if (has_binary_operator(evalIn)) {
            WCHAR spaced[512]; format_expr_spaced(evalIn, spaced, 512);
            WCHAR pretty[512];
            if (!PrettyPercentOfLeft(spaced, pretty, 512)) {
                size_t cap = 512; size_t di = 0; for (const WCHAR *p = spaced; *p && di < cap - 1; ++p) { WCHAR c = *p; if (c == L'*') c = CH_MUL; else if (c == L'/') c = CH_DIV; pretty[di++] = c; } pretty[di] = 0;
            }
            WCHAR numDisp[64]; format_double(val, numDisp, 64);
            WCHAR numFull[64]; format_double_full(val, numFull, 64);
            WCHAR line[600]; swprintf(line, sizeof(line)/sizeof(line[0]), L"%s = %s", pretty, numDisp);
            History_Add(hwnd, line, pretty, numFull, TRUE);
        } else {
            History_Add(hwnd, g_ui.input, g_ui.input, g_ui.input, TRUE);
        }
    } else {
        // Expression failed to evaluate - check if it's incomplete
        if (IsIncompleteString(g_ui.input)) {
            History_Add(hwnd, L"Invalid Input - Broken/Incomplete String", g_ui.input, NULL, TRUE);
        } else {
            History_Add(hwnd, g_ui.input, g_ui.input, NULL, TRUE);
        }
    }
}

static void HandleTextDeletion(HWND hwnd, BOOL forward, BOOL wholeWord) {
    if (DeleteSelectedInput()) {
        RefreshInputDisplay(hwnd);
        return;
    }
    if (wholeWord) {
        if (forward) {
            // Delete to next word boundary
            int next = NextWordPos(g_ui.input, g_ui.inputLen, g_ui.caretPos);
            if (next > g_ui.caretPos) {
                memmove(g_ui.input + g_ui.caretPos,
                        g_ui.input + next,
                        (g_ui.inputLen - next + 1) * sizeof(WCHAR));
                g_ui.inputLen -= (next - g_ui.caretPos);
            }
        } else {
            // Delete to previous word boundary
            int newPos = PrevWordPos(g_ui.input, g_ui.inputLen, g_ui.caretPos);
            if (newPos < g_ui.caretPos) {
                memmove(g_ui.input + newPos,
                        g_ui.input + g_ui.caretPos,
                        (g_ui.inputLen - g_ui.caretPos + 1) * sizeof(WCHAR));
                g_ui.inputLen -= (g_ui.caretPos - newPos);
                g_ui.caretPos = newPos;
            }
        }
    } else {
        if (forward) {
            // Delete character at caret
            if (g_ui.caretPos < g_ui.inputLen) {
                memmove(g_ui.input + g_ui.caretPos,
                        g_ui.input + g_ui.caretPos + 1,
                        (g_ui.inputLen - g_ui.caretPos) * sizeof(WCHAR));
                g_ui.inputLen--;
            }
        } else {
            // Delete character before caret
            if (g_ui.caretPos > 0) {
                memmove(g_ui.input + g_ui.caretPos - 1,
                        g_ui.input + g_ui.caretPos,
                        (g_ui.inputLen - g_ui.caretPos + 1) * sizeof(WCHAR));
                g_ui.caretPos--;
                g_ui.inputLen--;
            }
        }
    }
    RefreshInputDisplay(hwnd);
}

static void HandleCaretMovement(HWND hwnd, WPARAM wParam, BOOL extendSelection) {
    if (g_ui.inputLen == 0) InsertLastAnswerIfEmpty(hwnd);
    int newPos = g_ui.caretPos;

    switch (wParam) {
        case VK_LEFT:
            if (!extendSelection && HasInputSelection() && !(GetKeyState(VK_CONTROL) & 0x8000)) {
                newPos = g_ui.selStart;
                break;
            }
            if (GetKeyState(VK_CONTROL) & 0x8000) {
                newPos = PrevWordPos(g_ui.input, g_ui.inputLen, g_ui.caretPos);
            } else {
                if (newPos > 0) newPos--;
            }
            break;
        case VK_RIGHT:
            if (!extendSelection && HasInputSelection() && !(GetKeyState(VK_CONTROL) & 0x8000)) {
                newPos = g_ui.selEnd;
                break;
            }
            if (GetKeyState(VK_CONTROL) & 0x8000) {
                newPos = NextWordPos(g_ui.input, g_ui.inputLen, g_ui.caretPos);
            } else {
                if (newPos < g_ui.inputLen) newPos++;
            }
            break;
        case VK_HOME:
            newPos = 0;
            break;
        case VK_END:
            newPos = g_ui.inputLen;
            break;
    }
    if (extendSelection) SetInputSelection(g_ui.selAnchor, newPos);
    else {
        g_ui.caretPos = ClampInputPos(newPos);
        ClearInputSelection();
    }
    ResetCaretBlink(hwnd);
    InvalidateRect(hwnd, NULL, FALSE);
}

static void RefreshInputDisplay(HWND hwnd) {
    ResetCaretBlink(hwnd);
    UpdateCommandSuggestion();
    UpdateGhostParens();
    UpdateLivePreview(hwnd);
    InvalidateRect(hwnd, NULL, FALSE);
}

static void ResetCaretBlink(HWND hwnd) {
    g_ui.caretVisible = TRUE;
    KillTimer(hwnd, CARET_TIMER_ID);
    SetTimer(hwnd, CARET_TIMER_ID, CARET_BLINK_MS, NULL);
}

static void normalize_of_operator(const WCHAR *src, WCHAR *dst, size_t cap) {
    if (!dst || cap == 0) return;
    size_t di = 0;
    for (size_t i = 0; src && src[i] && di < cap - 1; ) {
        // Case-insensitive match for "of" with non-alpha boundaries
        if ((src[i] == L'o' || src[i] == L'O') && (src[i+1] == L'f' || src[i+1] == L'F')) {
            WCHAR prev = (i == 0) ? L' ' : src[i-1];
            WCHAR next = src[i+2];
            if (!iswalpha(prev) && !iswalpha(next)) {
                // replace with '*'
                dst[di++] = L'*';
                i += 2;
                continue;
            }
        }
        dst[di++] = src[i++];
    }
    dst[di] = 0;
}

static void normalize_percent_word(const WCHAR *src, WCHAR *dst, size_t cap) {
    if (!dst || cap == 0) return;
    size_t di = 0;
    const WCHAR *w = src;
    while (w && *w && di < cap - 1) {
        // Look for case-insensitive "percent" with non-alpha boundaries
        if ((w[0] == L'p' || w[0] == L'P') &&
            (w[1] == L'e' || w[1] == L'E') &&
            (w[2] == L'r' || w[2] == L'R') &&
            (w[3] == L'c' || w[3] == L'C') &&
            (w[4] == L'e' || w[4] == L'E') &&
            (w[5] == L'n' || w[5] == L'N') &&
            (w[6] == L't' || w[6] == L'T')) {
            WCHAR prev = (w == src) ? L' ' : *(w - 1);
            WCHAR next = w[7];
            if (!iswalpha(prev) && !iswalpha(next)) {
                dst[di++] = L'%';
                w += 7;
                continue;
            }
        }
        dst[di++] = *w++;
    }
    dst[di] = 0;
}

// Replace '×', '÷', and 'x'/'X' with ASCII '*' and '/' for evaluation
static void normalize_operators_basic(const WCHAR *src, WCHAR *dst, size_t cap) {
    if (!dst || cap == 0) return;
    size_t di = 0;
    for (const WCHAR *p = src; p && *p && di < cap - 1; ++p) {
        WCHAR c = *p;
        if (c == CH_MUL || c == L'x' || c == L'X') {
            dst[di++] = L'*';
        } else if (c == CH_DIV) {
            dst[di++] = L'/';
        } else {
            dst[di++] = c;
        }
    }
    dst[di] = 0;
}

static BOOL FindLastAnswer(int *idxOut, WCHAR *buf, size_t cap) {
    for (int i = g_historyCount - 1; i >= 0; --i) {
        HistoryItem *it = g_history[i];
        if (!it) continue;
        if (it->answer && *it->answer) {
            if (buf && cap > 0) {
                wcsncpy(buf, it->answer, cap - 1);
                buf[cap - 1] = 0;
            }
            if (idxOut) *idxOut = i;
            return TRUE;
        }
    }
    if (idxOut) *idxOut = -1;
    if (buf && cap > 0) buf[0] = 0;
    return FALSE;
}

static void UpdateLastAnswerSuggestion(void) {
    WCHAR buf[64]; int idx;
    if (FindLastAnswer(&idx, buf, 64)) {
        wcsncpy(g_ui.lastAns, buf, sizeof(g_ui.lastAns)/sizeof(g_ui.lastAns[0]) - 1);
        g_ui.lastAns[sizeof(g_ui.lastAns)/sizeof(g_ui.lastAns[0]) - 1] = 0;
        g_ui.lastAnsLen = (int)wcslen(g_ui.lastAns);
        g_ui.lastAnsSuggestActive = TRUE;
    } else {
        g_ui.lastAnsSuggestActive = FALSE;
        g_ui.lastAns[0] = 0; g_ui.lastAnsLen = 0;
    }
}

static BOOL InsertLastAnswerIfEmpty(HWND hwnd) {
    if (g_ui.inputLen == 0) {
        if (g_ui.suppressLastAnsSuggest) return FALSE; // honor suppression after backspace
        if (!g_ui.lastAnsSuggestActive) UpdateLastAnswerSuggestion();
        if (g_ui.lastAnsSuggestActive && g_ui.lastAnsLen > 0) {
            if (InsertAtCaret(g_ui.lastAns)) {
                ResetCaretBlink(hwnd);
                UpdateLivePreview(hwnd);
                InvalidateRect(hwnd, NULL, FALSE);
                return TRUE;
            }
        }
    }
    return FALSE;
}

static BOOL CopyInputSelectionToClipboard(HWND hwnd) {
    if (!HasInputSelection()) return FALSE;
    int len = g_ui.selEnd - g_ui.selStart;
    HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, (SIZE_T)(len + 1) * sizeof(WCHAR));
    if (!hMem) return FALSE;

    WCHAR *dst = (WCHAR*)GlobalLock(hMem);
    if (!dst) {
        GlobalFree(hMem);
        return FALSE;
    }
    wmemcpy(dst, g_ui.input + g_ui.selStart, len);
    dst[len] = 0;
    GlobalUnlock(hMem);

    if (!OpenClipboard(hwnd)) {
        GlobalFree(hMem);
        return FALSE;
    }
    EmptyClipboard();
    if (!SetClipboardData(CF_UNICODETEXT, hMem)) {
        CloseClipboard();
        GlobalFree(hMem);
        return FALSE;
    }
    CloseClipboard();
    return TRUE;
}

static BOOL PasteClipboardText(HWND hwnd) {
    BOOL inserted = FALSE;
    if (!OpenClipboard(hwnd)) return FALSE;
    HANDLE hData = GetClipboardData(CF_UNICODETEXT);
    if (hData) {
        const WCHAR *src = (const WCHAR*)GlobalLock(hData);
        if (src) {
            size_t srcLen = wcslen(src);
            WCHAR *clean = (WCHAR*)malloc((srcLen + 1) * sizeof(WCHAR));
            if (clean) {
                size_t di = 0;
                for (size_t i = 0; i < srcLen; ++i) {
                    WCHAR ch = src[i];
                    if (ch == L'\r' || ch == L'\n' || ch == L'\t') {
                        if (di > 0 && clean[di - 1] != L' ') clean[di++] = L' ';
                    } else if (ch >= 0x20 || ch == CH_MUL || ch == CH_DIV || ch == DEG || ch == MU) {
                        clean[di++] = ch;
                    }
                }
                clean[di] = 0;
                inserted = InsertAtCaret(clean);
                free(clean);
            }
            GlobalUnlock(hData);
        }
    }
    CloseClipboard();
    return inserted;
}

static BOOL InsertAtCaret(const WCHAR *s) {
    if (!s || !*s) return FALSE;
    if (HasInputSelection()) DeleteSelectedInput();
    size_t add = wcslen(s);
    size_t cap = (sizeof(g_ui.input)/sizeof(g_ui.input[0])) - 1;
    size_t curLen = (size_t)g_ui.inputLen;
    size_t pos = (size_t)g_ui.caretPos;
    if (pos > curLen) pos = curLen;
    if (curLen >= cap) return FALSE;
    if (curLen + add > cap) add = cap - curLen;
    // Move tail right
    memmove(g_ui.input + pos + add, g_ui.input + pos, (curLen - pos + 1) * sizeof(WCHAR));
    // Insert new
    wmemcpy(g_ui.input + pos, s, add);
    g_ui.inputLen = (int)(curLen + add);
    g_ui.suppressLastAnsSuggest = FALSE;
    g_ui.caretPos = (int)(pos + add);
    ClearInputSelection();
    return add > 0;
}

// Time parsing helpers
static BOOL is_digit(WCHAR c) { return c >= L'0' && c <= L'9'; }

static BOOL ParseDateTokenText(const WCHAR *s, int *consumed, SYSTEMTIME *out) {
    if (!s || !*s) return FALSE;
    if (!iswalpha(s[0])) return FALSE;

    int i = 0;
    int nameStart = 0;
    while (s[i] && (iswalpha(s[i]) || s[i] == L'.')) i++;
    if (i == nameStart) return FALSE;
    int nameEnd = i;
    while (nameEnd > nameStart && s[nameEnd - 1] == L'.') nameEnd--;
    if (nameEnd <= nameStart) return FALSE;
    int nameLen = nameEnd - nameStart;
    if (nameLen >= 16) nameLen = 15;

    WCHAR monthBuf[16];
    for (int k = 0; k < nameLen; ++k) {
        monthBuf[k] = (WCHAR)towlower(s[nameStart + k]);
    }
    monthBuf[nameLen] = 0;

    static const struct { const WCHAR *name; int month; } kMonthNames[] = {
        { L"january", 1 }, { L"jan", 1 },
        { L"february", 2 }, { L"feb", 2 },
        { L"march", 3 }, { L"mar", 3 },
        { L"april", 4 }, { L"apr", 4 },
        { L"may", 5 },
        { L"june", 6 }, { L"jun", 6 },
        { L"july", 7 }, { L"jul", 7 },
        { L"august", 8 }, { L"aug", 8 },
        { L"september", 9 }, { L"sept", 9 }, { L"sep", 9 },
        { L"october", 10 }, { L"oct", 10 },
        { L"november", 11 }, { L"nov", 11 },
        { L"december", 12 }, { L"dec", 12 },
    };

    int month = 0;
    const int monthNameCount = ARRAY_COUNT(kMonthNames);
    for (int idx = 0; idx < monthNameCount; ++idx) {
        if (wcscmp(monthBuf, kMonthNames[idx].name) == 0) {
            month = kMonthNames[idx].month;
            break;
        }
    }
    if (month == 0) return FALSE;

    while (s[i] == L' ' || s[i] == L'	') i++;
    if (s[i] == L',') {
        i++;
        while (s[i] == L' ' || s[i] == L'	') i++;
    }

    int day = 0;
    int dayDigits = 0;
    int numberStart = i;
    while (is_digit(s[i]) && dayDigits < 4) {
        day = day * 10 + (int)(s[i] - L'0');
        i++;
        dayDigits++;
    }
    if (is_digit(s[i])) return FALSE;

    if (dayDigits > 0 && dayDigits <= 2) {
        if (s[i]) {
            WCHAR c1 = (WCHAR)towlower(s[i]);
            WCHAR c2 = s[i+1] ? (WCHAR)towlower(s[i+1]) : 0;
            if ((c1 == L's' && c2 == L't') || (c1 == L'n' && c2 == L'd') ||
                (c1 == L'r' && c2 == L'd') || (c1 == L't' && c2 == L'h')) {
                i += c2 ? 2 : 1;
            }
        }
        if (s[i] == L'.') i++;
    } else if (dayDigits != 0) {
        i = numberStart;
        day = 0;
        dayDigits = 0;
    }

    while (s[i] == L' ' || s[i] == L'	') i++;
    if (s[i] == L',') {
        i++;
        while (s[i] == L' ' || s[i] == L'	') i++;
    }

    int year = 0;
    int yearDigits = 0;
    while (is_digit(s[i]) && yearDigits < 4) {
        year = year * 10 + (int)(s[i] - L'0');
        i++;
        yearDigits++;
    }
    if (is_digit(s[i])) return FALSE;

    if (dayDigits == 0) {
        day = 1;
    }

    BOOL hasYear = (yearDigits > 0);
    if (yearDigits == 2) {
        year += 2000;
    } else if (yearDigits != 0 && yearDigits != 4) {
        return FALSE;
    }
    if (!hasYear) {
        SYSTEMTIME now;
        GetLocalTime(&now);
        year = now.wYear;
        if (dayDigits == 0) {
            if (month < now.wMonth || (month == now.wMonth && now.wDay > 1)) {
                year++;
            }
        }
    }

    if (day < 1 || day > 31 || year < 1601) return FALSE;

    SYSTEMTIME st = {0};
    st.wYear = (WORD)year;
    st.wMonth = (WORD)month;
    st.wDay = (WORD)day;
    st.wHour = 0;
    st.wMinute = 0;
    st.wSecond = 0;
    if (out) *out = st;
    if (consumed) *consumed = i;
    return TRUE;
}

static BOOL ParseDateTokenNumeric(const WCHAR *s, int *consumed, SYSTEMTIME *out) {
    if (!s || !*s) return FALSE;
    // Accept numeric formats using the configured date order. If year is omitted, assume current year.
    int i = 0; int first = 0, second = 0, m = 0, d = 0, y = 0; int n = 0; BOOL hasYear = FALSE;
    while (is_digit(s[i]) && n < 2) { first = first*10 + (s[i]-L'0'); i++; n++; }
    if (n == 0) return FALSE;
    if (s[i] != L'/') return FALSE;
    i++;
    n = 0; while (is_digit(s[i]) && n < 2) { second = second*10 + (s[i]-L'0'); i++; n++; }
    if (n == 0) return FALSE;
    if (s[i] == L'/') {
        i++;
        n = 0; while (is_digit(s[i]) && n < 4) { y = y*10 + (s[i]-L'0'); i++; n++; }
        if (n < 2) return FALSE;
        if (n == 2) y += 2000;
        hasYear = TRUE;
    }

    if (ClampDateFormatStyle(g_settings.dateFormatStyle) == DATEFMT_DMY) {
        d = first;
        m = second;
    } else {
        m = first;
        d = second;
    }

    if (!hasYear) {
        SYSTEMTIME now; GetLocalTime(&now); y = now.wYear;
    }
    if (m < 1 || m > 12 || d < 1 || d > 31 || y < 1601) return FALSE;
    SYSTEMTIME st = {0};
    st.wYear = (WORD)y; st.wMonth = (WORD)m; st.wDay = (WORD)d; st.wHour = 0; st.wMinute = 0; st.wSecond = 0;
    if (out) *out = st;
    if (consumed) *consumed = i;
    return TRUE;
}

static BOOL ParseDateToken(const WCHAR *s, int *consumed, SYSTEMTIME *out) {
    if (ParseDateTokenText(s, consumed, out)) return TRUE;
    return ParseDateTokenNumeric(s, consumed, out);
}

static BOOL ParseTimeTokenAny(const WCHAR *s, int *consumed, SYSTEMTIME *out) {
    // Accept HH:MM with optional AM/PM. Default is 24-hour if AM/PM not present.
    int i = 0; int h = 0, min = 0; int n = 0;
    while (is_digit(s[i]) && n < 2) { h = h*10 + (s[i]-L'0'); i++; n++; }
    if (n == 0) return FALSE;
    if (s[i] != L':') return FALSE;
    i++;
    if (!is_digit(s[i]) || !is_digit(s[i+1])) return FALSE;
    min = (s[i]-L'0')*10 + (s[i+1]-L'0');
    i += 2;
    // Check for optional AM/PM
    int j = i; while (s[j] == L' ') j++;
    WCHAR a = (WCHAR)towlower(s[j]); WCHAR b = (WCHAR)towlower(s[j+1]);
    BOOL hasAmPm = ((a == L'a' || a == L'p') && b == L'm');
    int hour24;
    if (hasAmPm) {
        j += 2;
        if (h < 1 || h > 12 || min < 0 || min > 59) return FALSE;
        hour24 = (a == L'p') ? ((h % 12) + 12) : (h % 12);
        i = j; // consume AM/PM
    } else {
        if (h < 0 || h > 23 || min < 0 || min > 59) return FALSE;
        hour24 = h;
    }
    SYSTEMTIME st = {0};
    st.wHour = (WORD)hour24; st.wMinute = (WORD)min; st.wSecond = 0;
    if (out) *out = st;
    if (consumed) *consumed = i;
    return TRUE;
}

static BOOL ParseDateTimeAny(const WCHAR *s, int *consumed, SYSTEMTIME *out) {
    int i = 0;
    SYSTEMTIME date = {0}, time = {0}; BOOL hasDate = FALSE, hasTime = FALSE;
    // skip spaces
    while (s[i] == L' ') i++;

    {
        int holidayConsumed = 0;
        SYSTEMTIME holidayDate;
        if (ParseHolidayDateToken(s + i, &holidayConsumed, &holidayDate)) {
            if (out) *out = holidayDate;
            if (consumed) *consumed = i + holidayConsumed;
            return TRUE;
        }
    }
    
    // Fallback to original date/time parsing
    int c = 0; SYSTEMTIME tmp;
    if (ParseDateToken(s + i, &c, &tmp)) { date = tmp; hasDate = TRUE; i += c; while (s[i] == L' ') i++; }
    if (ParseTimeTokenAny(s + i, &c, &tmp)) { time = tmp; hasTime = TRUE; i += c; while (s[i] == L' ') i++; }
    if (!hasDate && !hasTime) return FALSE;
    SYSTEMTIME st = {0};
    if (hasDate) {
        st = date;
    } else {
        GetLocalTime(&st); // today
        st.wMilliseconds = 0;
    }
    if (hasTime) {
        st.wHour = time.wHour; st.wMinute = time.wMinute; st.wSecond = 0;
    }
    if (out) *out = st;
    if (consumed) *consumed = i;
    return TRUE;
}

static BOOL HandleHolidayQuery(HWND hwnd, const WCHAR *input) {
    const Holiday *holiday = NULL;
    SYSTEMTIME now;
    SYSTEMTIME holidayDate;
    WCHAR query[128];
    WCHAR lower[128];
    const WCHAR *p = input;
    int targetYear = 0;
    BOOL hasExplicitYear = FALSE;
    BOOL isToday = FALSE;
    BOOL isTodayQuery = FALSE;
    int len = 0;

    if (!hwnd || !input) return FALSE;

    GetLocalTime(&now);
    now.wHour = now.wMinute = now.wSecond = now.wMilliseconds = 0;

    while (*p && iswspace(*p)) p++;
    if (_wcsicmp(p, L"is today a holiday") == 0 || _wcsicmp(p, L"is today holiday") == 0) {
        isTodayQuery = TRUE;
    } else if (_wcsnicmp(p, L"when is", 7) == 0) {
        p += 7;
        while (*p && iswspace(*p)) p++;
        if (_wcsnicmp(p, L"the ", 4) == 0) {
            p += 4;
            while (*p && iswspace(*p)) p++;
        }
    }

    if (isTodayQuery) {
        SYSTEMTIME todayHolidayDate;
        if (FindHolidayForToday(&now, &holiday, &todayHolidayDate)) {
            WCHAR longDate[128];
            WCHAR shortDate[32];
            WCHAR message[256];
            FormatDateLong(&todayHolidayDate, longDate, ARRAY_COUNT(longDate));
            FormatDateLong(&todayHolidayDate, shortDate, ARRAY_COUNT(shortDate));
            swprintf(message, ARRAY_COUNT(message), L"Today is %s (%s).", holiday->name, longDate);
            History_Add(hwnd, message, shortDate, shortDate, TRUE);
        } else {
            History_Add(hwnd, L"Today is not one of the supported holidays.", NULL, NULL, TRUE);
        }
        return TRUE;
    }

    if (!*p) return FALSE;

    wcsncpy(query, p, ARRAY_COUNT(query) - 1);
    query[ARRAY_COUNT(query) - 1] = 0;
    len = (int)wcslen(query);
    while (len > 0 && iswspace(query[len - 1])) query[--len] = 0;
    while (len > 0 && (query[len - 1] == L'?' || query[len - 1] == L'.' || query[len - 1] == L'!')) query[--len] = 0;
    while (len > 0 && iswspace(query[len - 1])) query[--len] = 0;
    if (len == 0) return FALSE;

    for (int idx = 0; idx <= len; ++idx) lower[idx] = (WCHAR)towlower(query[idx]);

    if (len >= 10 && wcscmp(lower + len - 10, L" next year") == 0) {
        hasExplicitYear = TRUE;
        targetYear = now.wYear + 1;
        query[len - 10] = 0;
        len -= 10;
    } else if (len >= 10 && wcscmp(lower + len - 10, L" this year") == 0) {
        hasExplicitYear = TRUE;
        targetYear = now.wYear;
        query[len - 10] = 0;
        len -= 10;
    } else {
        int pos = len;
        while (pos > 0 && iswspace(query[pos - 1])) pos--;
        {
            int digitStart = pos;
            while (digitStart > 0 && is_digit(query[digitStart - 1])) digitStart--;
            if (digitStart < pos && (pos - digitStart == 2 || pos - digitStart == 4) &&
                (digitStart == 0 || iswspace(query[digitStart - 1]))) {
                int parsedYear = 0;
                for (int idx = digitStart; idx < pos; ++idx) parsedYear = parsedYear * 10 + (query[idx] - L'0');
                if (pos - digitStart == 2) parsedYear += 2000;
                if (parsedYear >= 1601) {
                    hasExplicitYear = TRUE;
                    targetYear = parsedYear;
                    query[digitStart] = 0;
                    len = digitStart;
                }
            }
        }
    }

    while (len > 0 && iswspace(query[len - 1])) query[--len] = 0;
    if (len == 0) return FALSE;

    for (int idx = 0; idx <= len; ++idx) lower[idx] = (WCHAR)towlower(query[idx]);

    if (!find_holiday_by_name(query, &holiday)) {
        if (wcscmp(lower, L"4th of july") == 0 || wcscmp(lower, L"fourth of july") == 0 ||
            wcscmp(lower, L"july 4") == 0 || wcscmp(lower, L"july 4th") == 0) {
            find_holiday_by_name(L"Independence Day", &holiday);
        } else if (wcscmp(lower, L"mlk day") == 0 || wcscmp(lower, L"martin luther king day") == 0) {
            find_holiday_by_name(L"Martin Luther King Jr. Day", &holiday);
        }
    }
    if (!holiday) return FALSE;

    if (!ResolveHolidayDate(holiday, hasExplicitYear, hasExplicitYear ? targetYear : now.wYear, &now, &holidayDate, &targetYear, &isToday)) {
        WCHAR message[256];
        if (hasExplicitYear) {
            swprintf(message, ARRAY_COUNT(message), L"Sorry, I couldn't determine %s in %d.", holiday->name, targetYear);
        } else {
            swprintf(message, ARRAY_COUNT(message), L"Sorry, I couldn't determine %s.", holiday->name);
        }
        History_Add(hwnd, message, NULL, NULL, TRUE);
        return TRUE;
    }

    {
        WCHAR longDate[128];
        WCHAR shortDate[32];
        WCHAR message[256];
        FormatDateLong(&holidayDate, longDate, ARRAY_COUNT(longDate));
        FormatDateLong(&holidayDate, shortDate, ARRAY_COUNT(shortDate));

        if (hasExplicitYear) {
            swprintf(message, ARRAY_COUNT(message), L"%s %d falls on %s.", holiday->name, targetYear, longDate);
        } else if (isToday) {
            swprintf(message, ARRAY_COUNT(message), L"Today is %s (%s).", holiday->name, longDate);
        } else {
            int daysAway = DaysBetweenCalendarDates(&now, &holidayDate);
            if (targetYear == now.wYear) {
                swprintf(message, ARRAY_COUNT(message), L"This year's %s is on %s (%d days from now).", holiday->name, longDate, daysAway);
            } else {
                swprintf(message, ARRAY_COUNT(message), L"The next %s is on %s (%d days from now).", holiday->name, longDate, daysAway);
            }
        }

        History_Add(hwnd, message, shortDate, shortDate, TRUE);
    }
    return TRUE;
}

static const WCHAR* wcsistr(const WCHAR* hay, const WCHAR* needle) {
    if (!hay || !needle || !*needle) return NULL;
    size_t nlen = wcslen(needle);
    for (const WCHAR *p = hay; *p; ++p) {
        if (_wcsnicmp(p, needle, nlen) == 0) return p;
    }
    return NULL;
}

static const WCHAR *TimeOutUnitLabel(TimeOutUnit unit) {
    switch (unit) {
        case TU_SECONDS: return L"seconds";
        case TU_MINUTES: return L"minutes";
        case TU_HOURS: return L"hours";
        case TU_DAYS: return L"days";
        case TU_WEEKS: return L"weeks";
        case TU_MONTHS: return L"months";
        case TU_YEARS: return L"years";
        default: return L"seconds";
    }
}

static double ConvertDurationToUnit(double seconds, TimeOutUnit unit, const WCHAR **labelOut) {
    const WCHAR *label = TimeOutUnitLabel(unit);
    double value = seconds;

    switch (unit) {
        case TU_MINUTES: value = seconds / 60.0; break;
        case TU_HOURS:   value = seconds / 3600.0; break;
        case TU_DAYS:    value = seconds / 86400.0; break;
        case TU_WEEKS:   value = seconds / 604800.0; break;
        case TU_MONTHS:  value = seconds / (31557600.0 / 12.0); break;
        case TU_YEARS:   value = seconds / 31557600.0; break;
        case TU_SECONDS:
        case TU_DEFAULT:
        default:
            value = seconds;
            label = L"seconds";
            break;
    }

    if (labelOut) *labelOut = label;
    return value;
}

static BOOL DetectTimeQuery(const WCHAR *s, TimeQuery *out) {
    if (!s || !*s) return FALSE;
    if (out) { ZeroMemory(out, sizeof(*out)); out->type = TQ_NONE; }
    // now
    const WCHAR *pnow = s; while (*pnow && iswspace(*pnow)) pnow++;
    if (_wcsicmp(pnow, L"now") == 0 || _wcsicmp(pnow, L"current time") == 0) {
        if (out) { out->type = TQ_NOW; }
        return TRUE;
    }
    // Optional unit prefix: seconds/minutes/hours/days/weeks/months/years
    TimeOutUnit outUnit = TU_DEFAULT; const WCHAR *p = s; while (*p && iswspace(*p)) p++;
    const struct { const WCHAR *w1; const WCHAR *w2; TimeOutUnit u; } unitWords[] = {
        { L"second", L"seconds", TU_SECONDS },
        { L"minute", L"minutes", TU_MINUTES },
        { L"hour",   L"hours",   TU_HOURS   },
        { L"day",    L"days",    TU_DAYS    },
        { L"week",   L"weeks",   TU_WEEKS   },
        { L"month",  L"months",  TU_MONTHS  },
        { L"year",   L"years",   TU_YEARS   },
    };
    const WCHAR *afterUnit = p; BOOL sawUnit = FALSE;
    for (int i = 0; i < (int)(sizeof(unitWords)/sizeof(unitWords[0])); ++i) {
        size_t l1 = wcslen(unitWords[i].w1), l2 = wcslen(unitWords[i].w2);
        if (_wcsnicmp(p, unitWords[i].w1, l1) == 0 && !iswalpha(p[l1])) { outUnit = unitWords[i].u; afterUnit = p + l1; sawUnit = TRUE; break; }
        if (_wcsnicmp(p, unitWords[i].w2, l2) == 0 && !iswalpha(p[l2])) { outUnit = unitWords[i].u; afterUnit = p + l2; sawUnit = TRUE; break; }
    }
    if (sawUnit) { p = afterUnit; while (*p && iswspace(*p)) p++; }
    // between X and Y
    const WCHAR *pbetween = wcsistr(p, L"between");
    if (pbetween) {
        pbetween += 7; // skip 'between'
        int c1 = 0; SYSTEMTIME a;
        if (!ParseDateTimeAny(pbetween, &c1, &a)) return FALSE;
        const WCHAR *pand = wcsistr(pbetween + c1, L"and"); if (!pand) return FALSE; pand += 3;
        int c2 = 0; SYSTEMTIME b; if (!ParseDateTimeAny(pand, &c2, &b)) return FALSE;
        if (out) { out->type = TQ_BETWEEN; out->a = a; out->b = b; out->hasA = out->hasB = TRUE; out->hasUnit = sawUnit; out->outUnit = outUnit; }
        return TRUE;
    }
    // until X
    const WCHAR *puntil = wcsistr(p, L"until");
    if (puntil) {
        puntil += 5; int c = 0; SYSTEMTIME a;
        if (!ParseDateTimeAny(puntil, &c, &a)) return FALSE;
        if (out) { out->type = TQ_UNTIL; out->a = a; out->hasA = TRUE; out->hasUnit = sawUnit; out->outUnit = outUnit; }
        return TRUE;
    }
    // since X
    const WCHAR *psince = wcsistr(p, L"since");
    if (psince) {
        psince += 5; int c = 0; SYSTEMTIME a;
        if (!ParseDateTimeAny(psince, &c, &a)) return FALSE;
        if (out) { out->type = TQ_SINCE; out->a = a; out->hasA = TRUE; out->hasUnit = sawUnit; out->outUnit = outUnit; }
        return TRUE;
    }
    return FALSE;
}

static ULONGLONG FileTimeToULL(FILETIME ft) {
    ULARGE_INTEGER u; u.LowPart = ft.dwLowDateTime; u.HighPart = ft.dwHighDateTime; return u.QuadPart;
}

static void FormatDuration(ULONGLONG secs, WCHAR *buf, size_t cap) {
    // Use the smart time formatter instead of the old day/hour format
    format_time_smart((double)secs, buf, cap);
}

static void FormatDateLong(const SYSTEMTIME *st, WCHAR *buf, size_t cap) {
    if (!st || !buf || cap == 0) return;
    switch (ClampDateFormatStyle(g_settings.dateFormatStyle)) {
        case DATEFMT_MDY:
            swprintf(buf, cap, L"%02d/%02d/%04d", st->wMonth, st->wDay, st->wYear);
            return;
        case DATEFMT_DMY:
            swprintf(buf, cap, L"%02d/%02d/%04d", st->wDay, st->wMonth, st->wYear);
            return;
        default:
            swprintf(buf, cap, L"%02d/%02d/%04d", st->wMonth, st->wDay, st->wYear);
            return;
    }
}

static void FormatLocalDateTime(const SYSTEMTIME *st, WCHAR *buf, size_t cap) {
    if (!st || !buf || cap == 0) return;
    int h = st->wHour % 12; if (h == 0) h = 12;
    const WCHAR *ampm = (st->wHour >= 12) ? L"PM" : L"AM";
    switch (ClampDateFormatStyle(g_settings.dateFormatStyle)) {
        case DATEFMT_MDY:
            swprintf(buf, cap, L"%02d/%02d/%04d %d:%02d %s", st->wMonth, st->wDay, st->wYear, h, st->wMinute, ampm);
            break;
        case DATEFMT_DMY:
            swprintf(buf, cap, L"%02d/%02d/%04d %d:%02d %s", st->wDay, st->wMonth, st->wYear, h, st->wMinute, ampm);
            break;
        default:
            swprintf(buf, cap, L"%02d/%02d/%04d %d:%02d %s", st->wMonth, st->wDay, st->wYear, h, st->wMinute, ampm);
            break;
    }
}

static BOOL AddTrayIcon(HWND hwnd) {
    NOTIFYICONDATAW nid;
    ZeroMemory(&nid, sizeof(nid));
    nid.cbSize = sizeof(nid);
    nid.hWnd = hwnd;
    nid.uID = TRAY_ICON_UID;
    nid.uFlags = NIF_MESSAGE | NIF_ICON | NIF_TIP;
    nid.uCallbackMessage = WMAPP_TRAYICON;
    wcscpy_s(nid.szTip, ARRAY_COUNT(nid.szTip), L"Calico");
    nid.hIcon = (HICON)LoadImageW(GetModuleHandleW(NULL), MAKEINTRESOURCEW(IDI_APP_ICON), IMAGE_ICON,
                                  GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), 0);
    if (!nid.hIcon) nid.hIcon = LoadIconW(NULL, IDI_APPLICATION);

    BOOL ok = Shell_NotifyIconW(g_trayIconAdded ? NIM_MODIFY : NIM_ADD, &nid);
    if (nid.hIcon) DestroyIcon(nid.hIcon);
    if (ok) g_trayIconAdded = TRUE;
    return ok;
}

static void RemoveTrayIcon(HWND hwnd) {
    if (!g_trayIconAdded) return;
    NOTIFYICONDATAW nid;
    ZeroMemory(&nid, sizeof(nid));
    nid.cbSize = sizeof(nid);
    nid.hWnd = hwnd;
    nid.uID = TRAY_ICON_UID;
    Shell_NotifyIconW(NIM_DELETE, &nid);
    g_trayIconAdded = FALSE;
}

static void ToggleTrayWindow(HWND hwnd) {
    if (g_hwndSettings && IsWindow(g_hwndSettings)) {
        SetForegroundWindow(g_hwndSettings);
        return;
    }
    if (IsWindowVisible(hwnd) && GetForegroundWindow() == hwnd) {
        ShowWindow(hwnd, SW_HIDE);
        return;
    }
    SpawnNearCursor(hwnd);
}

static void ShowTrayMenu(HWND hwnd) {
    HMENU menu = CreatePopupMenu();
    if (!menu) return;

    AppendMenuW(menu, MF_STRING, IDM_TRAY_SHOW, IsWindowVisible(hwnd) ? L"Hide Calico" : L"Show Calico");
    AppendMenuW(menu, MF_STRING, IDM_TRAY_SETTINGS, L"Settings...");
    AppendMenuW(menu, MF_SEPARATOR, 0, NULL);
    AppendMenuW(menu, MF_STRING, IDM_TRAY_EXIT, L"Exit");

    POINT pt;
    GetCursorPos(&pt);
    g_ui.showingMenu = TRUE;
    SetForegroundWindow(hwnd);
    TrackPopupMenu(menu, TPM_RIGHTBUTTON | TPM_LEFTALIGN | TPM_BOTTOMALIGN, pt.x, pt.y, 0, hwnd, NULL);
    PostMessageW(hwnd, WM_NULL, 0, 0);
    g_ui.showingMenu = FALSE;
    DestroyMenu(menu);
}

static void ShowContextMenu(HWND hwnd, int x, int y) {
    HMENU hMenu = CreatePopupMenu();
    HMENU hSub = CreatePopupMenu();
    HMENU hMeas = CreatePopupMenu();
    HMENU hArea = CreatePopupMenu();
    HMENU hWeight = CreatePopupMenu();
    HMENU hVolume = CreatePopupMenu();
    HMENU hMemory = CreatePopupMenu();
    HMENU hTemp = CreatePopupMenu();
    HMENU hTime = CreatePopupMenu();
    HMENU hSym  = CreatePopupMenu();

    // Apply dark background to menus to better match app styling
    MENUINFO mi = {0};
    mi.cbSize = sizeof(mi);
    mi.fMask = MIM_BACKGROUND | MIM_APPLYTOSUBMENUS;
    HBRUSH darkBrush = CreateSolidBrush(COL_CONTENT);
    mi.hbrBack = darkBrush;
    SetMenuInfo(hMenu, &mi);
    SetMenuInfo(hSub, &mi);
    SetMenuInfo(hMeas, &mi);
    SetMenuInfo(hArea, &mi);
    SetMenuInfo(hWeight, &mi);
    SetMenuInfo(hVolume, &mi);
    SetMenuInfo(hMemory, &mi);
    SetMenuInfo(hTemp, &mi);
    SetMenuInfo(hTime, &mi);
    SetMenuInfo(hSym, &mi);

    // Build Insert > Symbols (owner-draw)
    MENUITEMINFOW mii;
    ZeroMemory(&mii, sizeof(mii));
    mii.cbSize = sizeof(mii);
    mii.fMask = MIIM_ID | MIIM_FTYPE | MIIM_DATA | MIIM_STRING;
    mii.fType = MFT_OWNERDRAW;

    const struct { UINT id; const WCHAR *text; } items[] = {
        { IDM_INS_PLUS,   L"+" },
        { IDM_INS_MINUS,  L"-" },
        { IDM_INS_TIMES,  L"×" },
        { IDM_INS_DIV,    L"÷" },
        { IDM_INS_LPAREN, L"(" },
        { IDM_INS_RPAREN, L")" },
        { IDM_INS_PERCENT,L"%" },
        { IDM_INS_POWER,  L"^" },
    };
    for (int i = 0; i < (int)(sizeof(items)/sizeof(items[0])); ++i) {
        mii.wID = items[i].id;
        mii.dwTypeData = (LPWSTR)items[i].text;
        mii.cch = (UINT)wcslen(items[i].text);
        mii.dwItemData = (ULONG_PTR)items[i].text;
        InsertMenuItemW(hSym, i, TRUE, &mii);
    }

    // Override displayed glyphs for multiply/divide to reliable Unicode (owner-draw uses itemData)
    {
        static const WCHAR kMulSym[] = L"\x00D7";
        static const WCHAR kDivSym[] = L"\x00F7";
        MENUITEMINFOW mm = {0}; mm.cbSize = sizeof(mm); mm.fMask = MIIM_DATA;
        mm.dwItemData = (ULONG_PTR)kMulSym; SetMenuItemInfoW(hSym, IDM_INS_TIMES, FALSE, &mm);
        mm.dwItemData = (ULONG_PTR)kDivSym; SetMenuItemInfoW(hSym, IDM_INS_DIV, FALSE, &mm);
    }

    // Build Insert > Lengths submenu
    const struct { UINT id; const WCHAR *text; } meas[] = {
        { IDM_MEAS_IN,  L"in" },
        { IDM_MEAS_FT,  L"ft" },
        { IDM_MEAS_YD,  L"yd" },
        { IDM_MEAS_M,   L"m" },
        { IDM_MEAS_CM,  L"cm" },
        { IDM_MEAS_MM,  L"mm" },
        { IDM_MEAS_KM,  L"km" },
        { IDM_MEAS_MI,  L"mi" },
        { IDM_MEAS_NMI, L"nmi" },
        { IDM_MEAS_LY,  L"ly" },
    };
    for (int i = 0; i < (int)(sizeof(meas)/sizeof(meas[0])); ++i) {
        mii.wID = meas[i].id;
        mii.dwTypeData = (LPWSTR)meas[i].text;
        mii.cch = (UINT)wcslen(meas[i].text);
        mii.dwItemData = (ULONG_PTR)meas[i].text;
        InsertMenuItemW(hMeas, i, TRUE, &mii);
    }

    // Build Insert > Areas submenu
    const struct { UINT id; const WCHAR *text; } areas[] = {
        { IDM_AREA_SQIN, L"sq in" },
        { IDM_AREA_SQFT, L"sq ft" },
        { IDM_AREA_SQYD, L"sq yd" },
        { IDM_AREA_SQM,  L"sq m" },
        { IDM_AREA_SQKM, L"sq km" },
        { IDM_AREA_SQMI, L"sq mi" },
        { IDM_AREA_ACRE, L"acre" },
        { IDM_AREA_HECTARE, L"hectare" },
    };
    for (int i = 0; i < (int)(sizeof(areas)/sizeof(areas[0])); ++i) {
        mii.wID = areas[i].id;
        mii.dwTypeData = (LPWSTR)areas[i].text;
        mii.cch = (UINT)wcslen(areas[i].text);
        mii.dwItemData = (ULONG_PTR)areas[i].text;
        InsertMenuItemW(hArea, i, TRUE, &mii);
    }

    // Build Insert > Weight submenu
    const struct { UINT id; const WCHAR *text; } weights[] = {
        { IDM_WEIGHT_MG,  L"mg" },
        { IDM_WEIGHT_G,   L"g" },
        { IDM_WEIGHT_KG,  L"kg" },
        { IDM_WEIGHT_OZ,  L"oz" },
        { IDM_WEIGHT_LB,  L"lb" },
        { IDM_WEIGHT_TON, L"ton" },
    };
    for (int i = 0; i < (int)(sizeof(weights)/sizeof(weights[0])); ++i) {
        mii.wID = weights[i].id;
        mii.dwTypeData = (LPWSTR)weights[i].text;
        mii.cch = (UINT)wcslen(weights[i].text);
        mii.dwItemData = (ULONG_PTR)weights[i].text;
        InsertMenuItemW(hWeight, i, TRUE, &mii);
    }

    // Build Insert > Volume submenu
    const struct { UINT id; const WCHAR *text; } volumes[] = {
        { IDM_VOLUME_ML,   L"mL" },
        { IDM_VOLUME_L,    L"L" },
        { IDM_VOLUME_FLOZ, L"fl oz" },
        { IDM_VOLUME_CUP,  L"cup" },
        { IDM_VOLUME_PINT, L"pt" },
        { IDM_VOLUME_QUART,L"qt" },
        { IDM_VOLUME_GAL,  L"gal" },
        { IDM_VOLUME_TBSP, L"tbsp" },
        { IDM_VOLUME_TSP,  L"tsp" },
    };
    for (int i = 0; i < (int)(sizeof(volumes)/sizeof(volumes[0])); ++i) {
        mii.wID = volumes[i].id;
        mii.dwTypeData = (LPWSTR)volumes[i].text;
        mii.cch = (UINT)wcslen(volumes[i].text);
        mii.dwItemData = (ULONG_PTR)volumes[i].text;
        InsertMenuItemW(hVolume, i, TRUE, &mii);
    }

    // Build Insert > Memory submenu
    const struct { UINT id; const WCHAR *text; } memories[] = {
        { IDM_MEMORY_BIT,  L"bit" },
        { IDM_MEMORY_BYTE, L"byte" },
        { IDM_MEMORY_KBIT, L"Kb" },
        { IDM_MEMORY_MBIT, L"Mb" },
        { IDM_MEMORY_GBIT, L"Gb" },
        { IDM_MEMORY_TBIT, L"Tb" },
        { IDM_MEMORY_KB,   L"KB" },
        { IDM_MEMORY_MB,   L"MB" },
        { IDM_MEMORY_GB,   L"GB" },
        { IDM_MEMORY_TB,   L"TB" },
        { IDM_MEMORY_KIB,  L"KiB" },
        { IDM_MEMORY_MIB,  L"MiB" },
        { IDM_MEMORY_GIB,  L"GiB" },
        { IDM_MEMORY_TIB,  L"TiB" },
    };
    for (int i = 0; i < (int)(sizeof(memories)/sizeof(memories[0])); ++i) {
        mii.wID = memories[i].id;
        mii.dwTypeData = (LPWSTR)memories[i].text;
        mii.cch = (UINT)wcslen(memories[i].text);
        mii.dwItemData = (ULONG_PTR)memories[i].text;
        InsertMenuItemW(hMemory, i, TRUE, &mii);
    }

    // Build Insert > Temperatures submenu
    const struct { UINT id; const WCHAR *text; } temps[] = {
        { IDM_TEMP_C, L"\u00B0C" },
        { IDM_TEMP_F, L"\u00B0F" },
        { IDM_TEMP_K, L"K" },
    };
    for (int i = 0; i < (int)(sizeof(temps)/sizeof(temps[0])); ++i) {
        mii.wID = temps[i].id;
        mii.dwTypeData = (LPWSTR)temps[i].text;
        mii.cch = (UINT)wcslen(temps[i].text);
        mii.dwItemData = (ULONG_PTR)temps[i].text;
        InsertMenuItemW(hTemp, i, TRUE, &mii);
    }

    // Override temperature labels to use degree symbols (owner-draw uses itemData)
    {
        static const WCHAR kDegC[] = L"\u00B0C";
        static const WCHAR kDegF[] = L"\u00B0F";
        MENUITEMINFOW mm2 = {0}; mm2.cbSize = sizeof(mm2); mm2.fMask = MIIM_DATA;
        mm2.dwItemData = (ULONG_PTR)kDegC; SetMenuItemInfoW(hTemp, IDM_TEMP_C, FALSE, &mm2);
        mm2.dwItemData = (ULONG_PTR)kDegF; SetMenuItemInfoW(hTemp, IDM_TEMP_F, FALSE, &mm2);
    }

    // Build Insert > Time submenu
    const struct { UINT id; const WCHAR *text; } times[] = {
        { IDM_TIME_NS,  L"ns" },
        { IDM_TIME_US,  L"\x00B5s" },
        { IDM_TIME_MS,  L"ms" },
        { IDM_TIME_S,   L"s"  },
        { IDM_TIME_MIN, L"min" },
        { IDM_TIME_H,   L"h" },
        { IDM_TIME_DAY, L"day" },
        { IDM_TIME_WEEK,L"week" },
        { IDM_TIME_MONTH,L"month" },
        { IDM_TIME_YEAR,L"year" },
    };
    for (int i = 0; i < (int)(sizeof(times)/sizeof(times[0])); ++i) {
        mii.wID = times[i].id;
        mii.dwTypeData = (LPWSTR)times[i].text;
        mii.cch = (UINT)wcslen(times[i].text);
        mii.dwItemData = (ULONG_PTR)times[i].text;
        InsertMenuItemW(hTime, i, TRUE, &mii);
    }

    // Append holiday shortcuts beneath the time units
    InsertMenuW(hTime, GetMenuItemCount(hTime), MF_BYPOSITION | MF_SEPARATOR, 0, NULL);
    const struct { UINT id; const WCHAR *text; } holidays[] = {
        { IDM_HOLIDAY_NEWYEAR_DAY,   L"New Year's Day" },
        { IDM_HOLIDAY_NEWYEAR_EVE,   L"New Year's Eve" },
        { IDM_HOLIDAY_MLK,           L"Martin Luther King Jr. Day" },
        { IDM_HOLIDAY_PRESIDENTS,    L"Presidents' Day" },
        { IDM_HOLIDAY_MEMORIAL,      L"Memorial Day" },
        { IDM_HOLIDAY_INDEPENDENCE,  L"Independence Day" },
        { IDM_HOLIDAY_LABOR,         L"Labor Day" },
        { IDM_HOLIDAY_COLUMBUS,      L"Columbus Day" },
        { IDM_HOLIDAY_HALLOWEEN,     L"Halloween" },
        { IDM_HOLIDAY_VETERANS,      L"Veterans Day" },
        { IDM_HOLIDAY_THANKSGIVING,  L"Thanksgiving" },
        { IDM_HOLIDAY_CHRISTMAS_EVE, L"Christmas Eve" },
        { IDM_HOLIDAY_CHRISTMAS_DAY, L"Christmas Day" },
        { IDM_HOLIDAY_VALENTINE,     L"Valentine's Day" },
        { IDM_HOLIDAY_MOTHERS,       L"Mother's Day" },
        { IDM_HOLIDAY_FATHERS,       L"Father's Day" },
    };
    for (int i = 0; i < (int)(sizeof(holidays)/sizeof(holidays[0])); ++i) {
        mii.wID = holidays[i].id;
        mii.dwTypeData = (LPWSTR)holidays[i].text;
        mii.cch = (UINT)wcslen(holidays[i].text);
        mii.dwItemData = (ULONG_PTR)holidays[i].text;
        InsertMenuItemW(hTime, GetMenuItemCount(hTime), TRUE, &mii);
    }

    // Top-level: Clear Input Bar, Settings, separator, Insert -> submenu
    MENUITEMINFOW mTop;
    ZeroMemory(&mTop, sizeof(mTop));
    mTop.cbSize = sizeof(mTop);
    mTop.fMask = MIIM_ID | MIIM_FTYPE | MIIM_DATA | MIIM_STRING;
    mTop.fType = MFT_OWNERDRAW;
    mTop.wID = IDM_CLEAR_INPUT;
    mTop.dwTypeData = L"Clear Input Bar";
    mTop.cch = (UINT)wcslen(L"Clear Input Bar");
    mTop.dwItemData = (ULONG_PTR)L"Clear Input Bar";
    InsertMenuItemW(hMenu, 0, TRUE, &mTop);

    mTop.wID = IDM_SETTINGS;
    mTop.dwTypeData = L"Settings...";
    mTop.cch = (UINT)wcslen(L"Settings...");
    mTop.dwItemData = (ULONG_PTR)L"Settings...";
    InsertMenuItemW(hMenu, 1, TRUE, &mTop);

    InsertMenuW(hMenu, 2, MF_BYPOSITION | MF_SEPARATOR, 0, NULL);

    MENUITEMINFOW mIns;
    ZeroMemory(&mIns, sizeof(mIns));
    mIns.cbSize = sizeof(mIns);
    mIns.fMask = MIIM_SUBMENU | MIIM_FTYPE | MIIM_DATA | MIIM_STRING;
    mIns.fType = MFT_OWNERDRAW;
    mIns.hSubMenu = hSub;
    mIns.dwTypeData = L"Insert";
    mIns.cch = (UINT)wcslen(L"Insert");
    mIns.dwItemData = (ULONG_PTR)L"Insert";
    InsertMenuItemW(hMenu, 3, TRUE, &mIns);

    // Add Insert > Operators
    MENUITEMINFOW mCat;
    ZeroMemory(&mCat, sizeof(mCat));
    mCat.cbSize = sizeof(mCat);
    mCat.fMask = MIIM_SUBMENU | MIIM_FTYPE | MIIM_DATA | MIIM_STRING;
    mCat.fType = MFT_OWNERDRAW;
    mCat.hSubMenu = hSym;
    mCat.dwTypeData = L"Operators";
    mCat.cch = (UINT)wcslen(L"Operators");
    mCat.dwItemData = (ULONG_PTR)L"Operators";
    int pos = 0;
    InsertMenuItemW(hSub, pos, TRUE, &mCat);

    // Add Insert > Lengths
    ZeroMemory(&mCat, sizeof(mCat));
    mCat.cbSize = sizeof(mCat);
    mCat.fMask = MIIM_SUBMENU | MIIM_FTYPE | MIIM_DATA | MIIM_STRING;
    mCat.fType = MFT_OWNERDRAW;
    mCat.hSubMenu = hMeas;
    mCat.dwTypeData = L"Lengths";
    mCat.cch = (UINT)wcslen(L"Lengths");
    mCat.dwItemData = (ULONG_PTR)L"Lengths";
    pos = GetMenuItemCount(hSub);
    InsertMenuItemW(hSub, pos, TRUE, &mCat);

    // Add Insert > Areas
    ZeroMemory(&mCat, sizeof(mCat));
    mCat.cbSize = sizeof(mCat);
    mCat.fMask = MIIM_SUBMENU | MIIM_FTYPE | MIIM_DATA | MIIM_STRING;
    mCat.fType = MFT_OWNERDRAW;
    mCat.hSubMenu = hArea;
    mCat.dwTypeData = L"Areas";
    mCat.cch = (UINT)wcslen(L"Areas");
    mCat.dwItemData = (ULONG_PTR)L"Areas";
    pos = GetMenuItemCount(hSub);
    InsertMenuItemW(hSub, pos, TRUE, &mCat);

    // Add Insert > Weight
    ZeroMemory(&mCat, sizeof(mCat));
    mCat.cbSize = sizeof(mCat);
    mCat.fMask = MIIM_SUBMENU | MIIM_FTYPE | MIIM_DATA | MIIM_STRING;
    mCat.fType = MFT_OWNERDRAW;
    mCat.hSubMenu = hWeight;
    mCat.dwTypeData = L"Weight";
    mCat.cch = (UINT)wcslen(L"Weight");
    mCat.dwItemData = (ULONG_PTR)L"Weight";
    pos = GetMenuItemCount(hSub);
    InsertMenuItemW(hSub, pos, TRUE, &mCat);

    // Add Insert > Volume
    ZeroMemory(&mCat, sizeof(mCat));
    mCat.cbSize = sizeof(mCat);
    mCat.fMask = MIIM_SUBMENU | MIIM_FTYPE | MIIM_DATA | MIIM_STRING;
    mCat.fType = MFT_OWNERDRAW;
    mCat.hSubMenu = hVolume;
    mCat.dwTypeData = L"Volume";
    mCat.cch = (UINT)wcslen(L"Volume");
    mCat.dwItemData = (ULONG_PTR)L"Volume";
    pos = GetMenuItemCount(hSub);
    InsertMenuItemW(hSub, pos, TRUE, &mCat);

    // Add Insert > Memory
    ZeroMemory(&mCat, sizeof(mCat));
    mCat.cbSize = sizeof(mCat);
    mCat.fMask = MIIM_SUBMENU | MIIM_FTYPE | MIIM_DATA | MIIM_STRING;
    mCat.fType = MFT_OWNERDRAW;
    mCat.hSubMenu = hMemory;
    mCat.dwTypeData = L"Memory";
    mCat.cch = (UINT)wcslen(L"Memory");
    mCat.dwItemData = (ULONG_PTR)L"Memory";
    pos = GetMenuItemCount(hSub);
    InsertMenuItemW(hSub, pos, TRUE, &mCat);

    // Add Insert > Temperatures
    ZeroMemory(&mCat, sizeof(mCat));
    mCat.cbSize = sizeof(mCat);
    mCat.fMask = MIIM_SUBMENU | MIIM_FTYPE | MIIM_DATA | MIIM_STRING;
    mCat.fType = MFT_OWNERDRAW;
    mCat.hSubMenu = hTemp;
    mCat.dwTypeData = L"Temperatures";
    mCat.cch = (UINT)wcslen(L"Temperatures");
    mCat.dwItemData = (ULONG_PTR)L"Temperatures";
    pos = GetMenuItemCount(hSub);
    InsertMenuItemW(hSub, pos, TRUE, &mCat);

    // Add Insert > Time
    ZeroMemory(&mCat, sizeof(mCat));
    mCat.cbSize = sizeof(mCat);
    mCat.fMask = MIIM_SUBMENU | MIIM_FTYPE | MIIM_DATA | MIIM_STRING;
    mCat.fType = MFT_OWNERDRAW;
    mCat.hSubMenu = hTime;
    mCat.dwTypeData = L"Time";
    mCat.cch = (UINT)wcslen(L"Time");
    mCat.dwItemData = (ULONG_PTR)L"Time";
    pos = GetMenuItemCount(hSub);
    InsertMenuItemW(hSub, pos, TRUE, &mCat);

        g_ui.showingMenu = TRUE;
    SetForegroundWindow(hwnd);
    TrackPopupMenuEx(hMenu, TPM_RIGHTBUTTON | TPM_LEFTALIGN, x, y, hwnd, NULL);
    DestroyMenu(hMenu); // destroys submenu too
    DeleteObject(darkBrush);
    g_ui.showingMenu = FALSE;
}

static WCHAR* wcsdup_heap(const WCHAR *s) {
    if (!s) return NULL;
    size_t len = wcslen(s);
    WCHAR *p = (WCHAR*)malloc((len + 1) * sizeof(WCHAR));
    if (!p) return NULL;
    memcpy(p, s, (len + 1) * sizeof(WCHAR));
    return p;
}

static void History_FreeAll(void) {
    for (int i = 0; i < g_historyCount; ++i) {
        HistoryItem *it = g_history[i];
        if (it) {
            if (it->text) free(it->text);
            if (it->insert) free(it->insert);
            if (it->answer) free(it->answer);
            free(it);
            g_history[i] = NULL;
        }
    }
    g_historyCount = 0;
}

static void History_Add(HWND hwnd, const WCHAR *text, const WCHAR *insert, const WCHAR *answer, BOOL autoscroll) {
    if (!text || !*text) return;
    HistoryItem *it = (HistoryItem*)calloc(1, sizeof(HistoryItem));
    if (!it) return;
    it->text = wcsdup_heap(text);
    if (!it->text) { free(it); return; }
    if (insert && *insert) it->insert = wcsdup_heap(insert);
    else it->insert = NULL;
    if (answer && *answer) it->answer = wcsdup_heap(answer);
    else it->answer = NULL;

    if (g_historyCount == HISTORY_CAP) {
        // drop oldest
        HistoryItem *old = g_history[0];
        if (old) {
            if (old->text) free(old->text);
            if (old->insert) free(old->insert);
            if (old->answer) free(old->answer);
            free(old);
        }
        memmove(&g_history[0], &g_history[1], sizeof(HistoryItem*) * (HISTORY_CAP - 1));
        g_history[HISTORY_CAP - 1] = it;
    } else {
        g_history[g_historyCount++] = it;
    }
    if (autoscroll && hwnd) {
        // First resize the window to the new desired height
        AdjustWindowHeight(hwnd);
        // Then compute autoscroll against the updated visible content area
        RECT wnd; GetClientRect(hwnd, &wnd);
        RECT rcInput, rcIcons, rcTabs, rcContent;
        GetLayoutRects(wnd, &rcInput, &rcIcons, &rcTabs, &rcContent);
        int itemH = S(ITEM_HEIGHT);
        int total = g_historyCount * itemH;
        int vis = rcContent.bottom - rcContent.top;
        g_ui.scrollOffset = (total > vis) ? (total - vis) : 0;
    }
    UpdateLastAnswerSuggestion();
}

static void History_Clear(HWND hwnd) {
    History_FreeAll();
    g_ui.scrollOffset = 0;
    if (hwnd) AdjustWindowHeight(hwnd);
    g_ui.lastAnsSuggestActive = FALSE; g_ui.lastAns[0] = 0; g_ui.lastAnsLen = 0;
}

static int ClampOpacityPercent(int value) {
    if (value < 40) return 40;
    if (value > 100) return 100;
    return value;
}

static int ClampDateFormatStyle(int value) {
    if (value < DATEFMT_MDY) return DATEFMT_MDY;
    if (value > DATEFMT_DMY) return DATEFMT_MDY;
    return value;
}

static int ClampInterfaceScalePercent(int value) {
    if (value < 75) return 75;
    if (value > 200) return 200;
    return value;
}

static UINT ClampHotkeyMods(UINT mods) {
    UINT allowed = MOD_ALT | MOD_CONTROL | MOD_SHIFT | MOD_WIN;
    return mods & allowed;
}

static UINT ClampHotkeyVk(UINT vk) {
    if (vk == 0) return 0;
    if ((vk >= 'A' && vk <= 'Z') || (vk >= '0' && vk <= '9')) return vk;
    if (vk >= VK_F1 && vk <= VK_F24) return vk;
    switch (vk) {
    case VK_OEM_3:
    case VK_OEM_MINUS:
    case VK_OEM_PLUS:
    case VK_OEM_4:
    case VK_OEM_6:
    case VK_OEM_5:
    case VK_OEM_1:
    case VK_OEM_7:
    case VK_OEM_COMMA:
    case VK_OEM_PERIOD:
    case VK_OEM_2:
        return vk;
    default:
        return HOTKEY_KEY;
    }
}

static BYTE ModMaskToHotkeyFlags(UINT mods) {
    BYTE flags = 0;
    if (mods & MOD_SHIFT) flags |= HOTKEYF_SHIFT;
    if (mods & MOD_CONTROL) flags |= HOTKEYF_CONTROL;
    if (mods & MOD_ALT) flags |= HOTKEYF_ALT;
    return flags;
}

static UINT HotkeyFlagsToModMask(BYTE flags) {
    UINT mods = 0;
    if (flags & HOTKEYF_SHIFT) mods |= MOD_SHIFT;
    if (flags & HOTKEYF_CONTROL) mods |= MOD_CONTROL;
    if (flags & HOTKEYF_ALT) mods |= MOD_ALT;
    return mods;
}

static void AppendHotkeyPart(WCHAR *buf, size_t cap, const WCHAR *part) {
    if (!buf || cap == 0 || !part || !*part) return;
    if (buf[0]) wcscat_s(buf, cap, L"+");
    wcscat_s(buf, cap, part);
}

static void FormatHotkeyText(UINT mods, UINT vk, WCHAR *buf, size_t cap) {
    if (!buf || cap == 0) return;
    buf[0] = 0;
    if (mods & MOD_CONTROL) AppendHotkeyPart(buf, cap, L"Ctrl");
    if (mods & MOD_ALT) AppendHotkeyPart(buf, cap, L"Alt");
    if (mods & MOD_SHIFT) AppendHotkeyPart(buf, cap, L"Shift");
    if (mods & MOD_WIN) AppendHotkeyPart(buf, cap, L"Win");

    WCHAR keyName[64] = {0};
    UINT scanCode = MapVirtualKeyW(vk, MAPVK_VK_TO_VSC);
    LONG lParam = (LONG)(scanCode << 16);
    if (!GetKeyNameTextW(lParam, keyName, (int)ARRAY_COUNT(keyName))) {
        if (vk >= 'A' && vk <= 'Z') {
            keyName[0] = (WCHAR)vk;
            keyName[1] = 0;
        } else {
            wcscpy_s(keyName, ARRAY_COUNT(keyName), L"Unknown");
        }
    }
    AppendHotkeyPart(buf, cap, keyName);
}

static void InitSettingsPath(void) {
    WCHAR appDataPath[MAX_PATH];
    WCHAR settingsDir[MAX_PATH];

    if (g_settingsPath[0]) return;

    if (SUCCEEDED(SHGetFolderPathW(NULL, CSIDL_APPDATA | CSIDL_FLAG_CREATE, NULL, SHGFP_TYPE_CURRENT, appDataPath))) {
        if (wcscpy_s(settingsDir, ARRAY_COUNT(settingsDir), appDataPath) == 0 &&
            wcscat_s(settingsDir, ARRAY_COUNT(settingsDir), L"\\Calico") == 0) {
            if (CreateDirectoryW(settingsDir, NULL) || GetLastError() == ERROR_ALREADY_EXISTS) {
                if (wcscpy_s(g_settingsPath, ARRAY_COUNT(g_settingsPath), settingsDir) == 0 &&
                    wcscat_s(g_settingsPath, ARRAY_COUNT(g_settingsPath), L"\\Calico.ini") == 0) {
                    return;
                }
            }
        }
    }

    DWORD len = GetModuleFileNameW(NULL, g_settingsPath, ARRAY_COUNT(g_settingsPath));
    if (len == 0 || len >= ARRAY_COUNT(g_settingsPath)) {
        wcscpy_s(g_settingsPath, ARRAY_COUNT(g_settingsPath), L".\\Calico.ini");
        return;
    }

    WCHAR *slash = wcsrchr(g_settingsPath, L'\\');
    if (slash) {
        slash[1] = 0;
        wcscat_s(g_settingsPath, ARRAY_COUNT(g_settingsPath), L"Calico.ini");
        return;
    }

    wcscpy_s(g_settingsPath, ARRAY_COUNT(g_settingsPath), L".\\Calico.ini");
}

static void LoadSettings(void) {
    const int DATE_FORMAT_STYLE_SCHEMA = 2;
    InitSettingsPath();
    g_settings.hotkeyMods = ClampHotkeyMods(
        GetPrivateProfileIntW(L"Window", L"HotkeyMods", g_settings.hotkeyMods, g_settingsPath));
    g_settings.hotkeyVk = ClampHotkeyVk(
        GetPrivateProfileIntW(L"Window", L"HotkeyVk", g_settings.hotkeyVk, g_settingsPath));
    g_settings.opacityPercent = ClampOpacityPercent(
        GetPrivateProfileIntW(L"Window", L"OpacityPercent", g_settings.opacityPercent, g_settingsPath));
    g_settings.interfaceScalePercent = ClampInterfaceScalePercent(
        GetPrivateProfileIntW(L"Window", L"InterfaceScalePercent", g_settings.interfaceScalePercent, g_settingsPath));
    g_settings.alwaysOnTop =
        GetPrivateProfileIntW(L"Window", L"AlwaysOnTop", g_settings.alwaysOnTop ? 1 : 0, g_settingsPath) != 0;
    g_settings.symbolsOpenByDefault =
        GetPrivateProfileIntW(L"Window", L"SymbolsOpenByDefault", g_settings.symbolsOpenByDefault ? 1 : 0, g_settingsPath) != 0;
    {
        int rawDateStyle = GetPrivateProfileIntW(L"Window", L"DateFormatStyle", g_settings.dateFormatStyle, g_settingsPath);
        int styleSchema = GetPrivateProfileIntW(L"Window", L"DateFormatStyleSchema", 0, g_settingsPath);
        if (styleSchema < DATE_FORMAT_STYLE_SCHEMA) {
            // Migrate from the earlier 3-option scheme: 0=Long, 1=MDY, 2=DMY.
            g_settings.dateFormatStyle = (rawDateStyle == 2) ? DATEFMT_DMY : DATEFMT_MDY;
        } else {
            g_settings.dateFormatStyle = ClampDateFormatStyle(rawDateStyle);
        }
    }
}

static BOOL SaveSettings(void) {
    WCHAR buf[16];
    InitSettingsPath();
    _snwprintf_s(buf, ARRAY_COUNT(buf), _TRUNCATE, L"%u", ClampHotkeyMods(g_settings.hotkeyMods));
    if (!WritePrivateProfileStringW(L"Window", L"HotkeyMods", buf, g_settingsPath)) return FALSE;
    _snwprintf_s(buf, ARRAY_COUNT(buf), _TRUNCATE, L"%u", ClampHotkeyVk(g_settings.hotkeyVk));
    if (!WritePrivateProfileStringW(L"Window", L"HotkeyVk", buf, g_settingsPath)) return FALSE;
    _snwprintf_s(buf, ARRAY_COUNT(buf), _TRUNCATE, L"%d", ClampOpacityPercent(g_settings.opacityPercent));
    if (!WritePrivateProfileStringW(L"Window", L"OpacityPercent", buf, g_settingsPath)) return FALSE;
    _snwprintf_s(buf, ARRAY_COUNT(buf), _TRUNCATE, L"%d", ClampInterfaceScalePercent(g_settings.interfaceScalePercent));
    if (!WritePrivateProfileStringW(L"Window", L"InterfaceScalePercent", buf, g_settingsPath)) return FALSE;
    if (!WritePrivateProfileStringW(L"Window", L"AlwaysOnTop", g_settings.alwaysOnTop ? L"1" : L"0", g_settingsPath)) return FALSE;
    if (!WritePrivateProfileStringW(L"Window", L"SymbolsOpenByDefault", g_settings.symbolsOpenByDefault ? L"1" : L"0", g_settingsPath)) return FALSE;
    _snwprintf_s(buf, ARRAY_COUNT(buf), _TRUNCATE, L"%d", ClampDateFormatStyle(g_settings.dateFormatStyle));
    if (!WritePrivateProfileStringW(L"Window", L"DateFormatStyle", buf, g_settingsPath)) return FALSE;
    if (!WritePrivateProfileStringW(L"Window", L"DateFormatStyleSchema", L"2", g_settingsPath)) return FALSE;
    return TRUE;
}

static int GetEffectiveUiScalePercent(void) {
    int autoScale = g_ui.autoScalePercent > 0 ? g_ui.autoScalePercent : 100;
    return MulDiv(autoScale, ClampInterfaceScalePercent(g_settings.interfaceScalePercent), 100);
}

static int GetFontCacheIndex(int sizePt, BOOL semibold) {
    if (semibold) return -1;
    if (sizePt == FONT_INPUT) return 0;
    if (sizePt == FONT_ICON) return 1;
    if (sizePt == FONT_TAB) return 2;
    if (sizePt == FONT_ITEM) return 3;
    return -1;
}

static int GetMonitorDpi(HMONITOR monitor, int fallbackDpi) {
    enum { MDT_EFFECTIVE_DPI = 0 };
    static BOOL triedLoad = FALSE;

    if (!pGetDpiForMonitor && !triedLoad) {
        triedLoad = TRUE;
        HMODULE shcore = LoadLibraryW(L"shcore.dll");
        if (shcore) {
            FARPROC fp = GetProcAddress(shcore, "GetDpiForMonitor");
            if (fp) {
                union { FARPROC p; GetDpiForMonitor_t f; } u;
                u.p = fp;
                pGetDpiForMonitor = u.f;
            }
        }
    }

    if (pGetDpiForMonitor && monitor) {
        UINT dpiX = 0, dpiY = 0;
        if (SUCCEEDED(pGetDpiForMonitor(monitor, MDT_EFFECTIVE_DPI, &dpiX, &dpiY)) && dpiX > 0) {
            return (int)dpiX;
        }
    }

    return fallbackDpi > 0 ? fallbackDpi : 96;
}

static int GetAutoScalePercentForMonitor(HMONITOR monitor) {
    MONITORINFO mi;
    ZeroMemory(&mi, sizeof(mi));
    mi.cbSize = sizeof(mi);
    if (!monitor || !GetMonitorInfoW(monitor, &mi)) return 100;

    int workWidth = mi.rcWork.right - mi.rcWork.left;
    int workHeight = mi.rcWork.bottom - mi.rcWork.top;
    int widthPercent = MulDiv(workWidth, 100, 1920);
    int heightPercent = MulDiv(workHeight, 100, 1080);
    int autoScale = min(widthPercent, heightPercent);

    if (autoScale < 100) autoScale = 100;
    if (autoScale > 200) autoScale = 200;
    return autoScale;
}

static void ApplyScaleMetricsForMonitor(HMONITOR monitor, int fallbackDpi) {
    g_ui.dpi = GetMonitorDpi(monitor, fallbackDpi);
    g_ui.autoScalePercent = GetAutoScalePercentForMonitor(monitor);
    g_ui.scale = GetEffectiveUiScalePercent();
    g_ui.panelSizePx.cx = S(PANEL_WIDTH);
    g_ui.panelSizePx.cy = S(PANEL_HEIGHT);
}

static void ApplyScaleMetrics(HWND hwnd) {
    HMONITOR monitor = MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST);
    int fallbackDpi = g_ui.dpi > 0 ? g_ui.dpi : 96;
    if (pGetDpiForWindow) fallbackDpi = (int)pGetDpiForWindow(hwnd);
    ApplyScaleMetricsForMonitor(monitor, fallbackDpi);
}

static BOOL RegisterConfiguredHotkey(HWND hwnd, BOOL showFailureFeedback) {
    WCHAR hotkeyText[64];
    UINT mods = ClampHotkeyMods(g_settings.hotkeyMods);
    UINT vk = ClampHotkeyVk(g_settings.hotkeyVk);

    UnregisterHotKey(hwnd, HOTKEY_ID);
    if (mods == 0 || vk == 0) return FALSE;
    if (RegisterHotKey(hwnd, HOTKEY_ID, mods, vk)) return TRUE;

    if (showFailureFeedback) {
        FormatHotkeyText(mods, vk, hotkeyText, ARRAY_COUNT(hotkeyText));
        WCHAR line[128];
        _snwprintf_s(line, ARRAY_COUNT(line), _TRUNCATE, L"Global hotkey %ls could not be registered.", hotkeyText);
        History_Add(hwnd, line, NULL, NULL, TRUE);
        History_Add(hwnd, L"Tip: Close conflicting apps or change the hotkey in /settings.", NULL, NULL, TRUE);
    }
    return FALSE;
}

static void InitDpi(HWND hwnd) {
    HMODULE user = GetModuleHandleW(L"user32.dll");
    if (user) {
        FARPROC fp = GetProcAddress(user, "GetDpiForWindow");
        if (fp) {
            union { FARPROC p; GetDpiForWindow_t f; } u;
            u.p = fp;
            pGetDpiForWindow = u.f;
        }
    }
    if (pGetDpiForWindow) {
        g_ui.dpi = pGetDpiForWindow(hwnd);
    } else {
        HDC hdc = GetDC(hwnd);
        g_ui.dpi = GetDeviceCaps(hdc, LOGPIXELSX);
        ReleaseDC(hwnd, hdc);
    }
    ApplyScaleMetrics(hwnd);
}

static int S(int v) { // scale logical -> px
    return MulDiv(MulDiv(v, g_ui.dpi, 96), GetEffectiveUiScalePercent(), 100);
}

static int HeaderHeightPx(void) {
    int h = 0;
    h += S(PADDING);            // top padding
    h += S(INPUT_H);            // input
    h += S(6);                  // spacing
    h += S(ICONROW_H);          // icon row
    if (g_ui.symbolsOpen) {
        h += S(4) + S(TABSTRIP_H) + S(6);
    } else {
        h += S(10);
    }
    return h;
}

static int DesiredWindowHeightPx(void) {
    int itemH = S(ITEM_HEIGHT);
    int rows = g_historyCount + (g_liveVisible ? 1 : 0);
    // Hide history area when there are no rows to show
    int content = rows > 0 ? rows * itemH : 0;
    int total = HeaderHeightPx() + content + S(PADDING); // bottom padding
    int maxH = S(PANEL_HEIGHT);
    if (total > maxH) total = maxH;
    return total;
}

static void AdjustWindowHeight(HWND hwnd) {
    if (!hwnd) return;
    RECT wr; GetWindowRect(hwnd, &wr);
    int w = wr.right - wr.left;
    int h = DesiredWindowHeightPx();
    g_ui.panelSizePx.cy = h;
    SetWindowPos(hwnd, NULL, 0, 0, w, h, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
}

static void ApplyWindowVisuals(HWND hwnd) {
    // Layered alpha for whole window
    LONG_PTR exStyle = GetWindowLongPtrW(hwnd, GWL_EXSTYLE);
    exStyle |= WS_EX_LAYERED | WS_EX_TOOLWINDOW;
    if (g_settings.alwaysOnTop) exStyle |= WS_EX_TOPMOST;
    else exStyle &= ~WS_EX_TOPMOST;
    SetWindowLongPtrW(hwnd, GWL_EXSTYLE, exStyle);
    SetLayeredWindowAttributes(hwnd, 0, (BYTE)MulDiv(ClampOpacityPercent(g_settings.opacityPercent), 255, 100), LWA_ALPHA);
    SetWindowPos(hwnd, g_settings.alwaysOnTop ? HWND_TOPMOST : HWND_NOTOPMOST,
                 0, 0, 0, 0,
                 SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_FRAMECHANGED);

    // Windows 11 rounded corners if available
    const DWORD DWMWA_WINDOW_CORNER_PREFERENCE = 33;
    const DWORD DWMWCP_ROUND = 2;
    DwmSetWindowAttribute(hwnd, DWMWA_WINDOW_CORNER_PREFERENCE, (void*)&DWMWCP_ROUND, sizeof(DWMWCP_ROUND));

    // Allow dark border/shadow
    BOOL val = TRUE;
    const DWORD DWMWA_USE_IMMERSIVE_DARK_MODE = 20;
    DwmSetWindowAttribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE, &val, sizeof(val));
}

static HFONT CreateUiFont(int sizePt, BOOL semibold) {
    int cacheIdx = GetFontCacheIndex(sizePt, semibold);

    if (cacheIdx >= 0 && g_fontCache[cacheIdx]) {
        return g_fontCache[cacheIdx];
    }

    LOGFONTW lf = {0};
    lf.lfHeight = -MulDiv(MulDiv(sizePt, GetEffectiveUiScalePercent(), 100), g_ui.dpi, 72);
    lf.lfWeight = semibold ? FW_SEMIBOLD : FW_NORMAL;
    lf.lfQuality = CLEARTYPE_QUALITY;
    wcscpy_s(lf.lfFaceName, LF_FACESIZE, L"Segoe UI");
    HFONT font = CreateFontIndirectW(&lf);

    if (cacheIdx >= 0 && !g_fontCache[cacheIdx]) {
        g_fontCache[cacheIdx] = font;
        g_fontCacheInitialized = TRUE;
    }

    return font;
}

static void ReleaseUiFont(HFONT font, int sizePt, BOOL semibold) {
    if (!font) return;
    if (GetFontCacheIndex(sizePt, semibold) >= 0) return;
    DeleteObject(font);
}

static void CleanupFontCache(void) {
    for (int i = 0; i < 4; i++) {
        if (g_fontCache[i]) {
            DeleteObject(g_fontCache[i]);
            g_fontCache[i] = NULL;
        }
    }
    g_fontCacheInitialized = FALSE;
}

static void FillRectColor(HDC hdc, RECT *rc, COLORREF c) {
    HBRUSH b = CreateSolidBrush(c);
    FillRect(hdc, rc, b);
    DeleteObject(b);
}

static void DrawTextClipped(HDC hdc, const wchar_t *txt, RECT rc, COLORREF col, UINT fmt) {
    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, col);
    DrawTextW(hdc, txt, -1, &rc, fmt);
}

static int SettingsScale(HWND hwnd, int v) {
    int dpi = 96;
    if (pGetDpiForWindow) dpi = (int)pGetDpiForWindow(hwnd);
    return MulDiv(MulDiv(v, dpi, 96), GetEffectiveUiScalePercent(), 100);
}

static void ApplyDialogFont(HWND hwnd) {
    HFONT font = (HFONT)GetStockObject(DEFAULT_GUI_FONT);
    SendMessageW(hwnd, WM_SETFONT, (WPARAM)font, TRUE);
}

static void CreateSettingsControls(HWND hwnd) {
    int margin = SettingsScale(hwnd, 14);
    int labelW = SettingsScale(hwnd, 180);
    int editW = SettingsScale(hwnd, 90);
    int comboW = SettingsScale(hwnd, 170);
    int lineH = SettingsScale(hwnd, 22);
    int rowGap = SettingsScale(hwnd, 12);
    int btnW = SettingsScale(hwnd, 86);
    int btnH = SettingsScale(hwnd, 28);
    int noteH = SettingsScale(hwnd, 32);
    int y = margin;

    HWND hLabel = CreateWindowExW(0, L"STATIC", L"Hotkey", WS_CHILD | WS_VISIBLE,
                                  margin, y + SettingsScale(hwnd, 3), labelW, lineH,
                                  hwnd, NULL, NULL, NULL);
    ApplyDialogFont(hLabel);
    HWND hHotkey = CreateWindowExW(WS_EX_CLIENTEDGE, HOTKEY_CLASSW, L"", WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                                   margin + labelW, y, editW, lineH + SettingsScale(hwnd, 2),
                                   hwnd, (HMENU)(INT_PTR)IDC_SETTINGS_HOTKEY, NULL, NULL);
    ApplyDialogFont(hHotkey);
    y += lineH + rowGap;

    hLabel = CreateWindowExW(0, L"STATIC", L"Opacity (40-100%)", WS_CHILD | WS_VISIBLE,
                             margin, y + SettingsScale(hwnd, 3), labelW, lineH,
                             hwnd, NULL, NULL, NULL);
    ApplyDialogFont(hLabel);
    HWND hEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"", WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL | ES_NUMBER,
                                 margin + labelW, y, editW, lineH + SettingsScale(hwnd, 2),
                                 hwnd, (HMENU)(INT_PTR)IDC_SETTINGS_OPACITY, NULL, NULL);
    ApplyDialogFont(hEdit);
    SendMessageW(hEdit, EM_LIMITTEXT, 3, 0);
    y += lineH + rowGap;

    hLabel = CreateWindowExW(0, L"STATIC", L"Scale Multiplier (75-200%)", WS_CHILD | WS_VISIBLE,
                             margin, y + SettingsScale(hwnd, 3), labelW, lineH,
                             hwnd, NULL, NULL, NULL);
    ApplyDialogFont(hLabel);
    hEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"", WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL | ES_NUMBER,
                            margin + labelW, y, editW, lineH + SettingsScale(hwnd, 2),
                            hwnd, (HMENU)(INT_PTR)IDC_SETTINGS_SCALE, NULL, NULL);
    ApplyDialogFont(hEdit);
    SendMessageW(hEdit, EM_LIMITTEXT, 3, 0);
    y += lineH + rowGap;

    hLabel = CreateWindowExW(0, L"STATIC", L"Date Order", WS_CHILD | WS_VISIBLE,
                             margin, y + SettingsScale(hwnd, 3), labelW, lineH,
                             hwnd, NULL, NULL, NULL);
    ApplyDialogFont(hLabel);
    {
        HWND hCombo = CreateWindowExW(0, WC_COMBOBOXW, L"", WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWNLIST | WS_VSCROLL,
                                      margin + labelW, y, comboW, SettingsScale(hwnd, 220),
                                      hwnd, (HMENU)(INT_PTR)IDC_SETTINGS_DATEFORMAT, NULL, NULL);
        ApplyDialogFont(hCombo);
        SendMessageW(hCombo, CB_ADDSTRING, 0, (LPARAM)L"MM/DD/YYYY");
        SendMessageW(hCombo, CB_ADDSTRING, 0, (LPARAM)L"DD/MM/YYYY");
    }
    y += lineH + rowGap;

    hLabel = CreateWindowExW(0, L"STATIC", L"100 uses automatic monitor sizing. 125 makes the whole UI 25% larger than automatic.", WS_CHILD | WS_VISIBLE,
                             margin, y, SettingsScale(hwnd, 280), noteH,
                             hwnd, NULL, NULL, NULL);
    ApplyDialogFont(hLabel);
    y += noteH + SettingsScale(hwnd, 6);

    HWND hTopmost = CreateWindowExW(0, L"BUTTON", L"Always on top", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
                                    margin, y, SettingsScale(hwnd, 180), lineH,
                                    hwnd, (HMENU)(INT_PTR)IDC_SETTINGS_TOPMOST, NULL, NULL);
    ApplyDialogFont(hTopmost);
    y += lineH + SettingsScale(hwnd, 6);

    HWND hSymbols = CreateWindowExW(0, L"BUTTON", L"Show symbols row by default", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
                                    margin, y, SettingsScale(hwnd, 220), lineH,
                                    hwnd, (HMENU)(INT_PTR)IDC_SETTINGS_SYMBOLS, NULL, NULL);
    ApplyDialogFont(hSymbols);
    y += lineH + SettingsScale(hwnd, 20);

    HWND hSave = CreateWindowExW(0, L"BUTTON", L"Save", WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON,
                                 margin, y, btnW, btnH,
                                 hwnd, (HMENU)(INT_PTR)IDC_SETTINGS_SAVE, NULL, NULL);
    ApplyDialogFont(hSave);
    HWND hCancel = CreateWindowExW(0, L"BUTTON", L"Cancel", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                                   margin + btnW + SettingsScale(hwnd, 10), y, btnW, btnH,
                                   hwnd, (HMENU)(INT_PTR)IDC_SETTINGS_CANCEL, NULL, NULL);
    ApplyDialogFont(hCancel);
}

static void SyncSettingsControls(HWND hwnd) {
    WCHAR buf[8];
    SendDlgItemMessageW(hwnd, IDC_SETTINGS_HOTKEY, HKM_SETHOTKEY,
                        MAKEWORD(ClampHotkeyVk(g_settings.hotkeyVk), ModMaskToHotkeyFlags(g_settings.hotkeyMods)), 0);
    _snwprintf_s(buf, ARRAY_COUNT(buf), _TRUNCATE, L"%d", ClampOpacityPercent(g_settings.opacityPercent));
    SetDlgItemTextW(hwnd, IDC_SETTINGS_OPACITY, buf);
    _snwprintf_s(buf, ARRAY_COUNT(buf), _TRUNCATE, L"%d", ClampInterfaceScalePercent(g_settings.interfaceScalePercent));
    SetDlgItemTextW(hwnd, IDC_SETTINGS_SCALE, buf);
    CheckDlgButton(hwnd, IDC_SETTINGS_TOPMOST, g_settings.alwaysOnTop ? BST_CHECKED : BST_UNCHECKED);
    CheckDlgButton(hwnd, IDC_SETTINGS_SYMBOLS, g_settings.symbolsOpenByDefault ? BST_CHECKED : BST_UNCHECKED);
    SendDlgItemMessageW(hwnd, IDC_SETTINGS_DATEFORMAT, CB_SETCURSEL, (WPARAM)ClampDateFormatStyle(g_settings.dateFormatStyle), 0);
}

static void CloseSettingsDialog(HWND hwnd) {
    HWND owner = (HWND)GetWindowLongPtrW(hwnd, GWLP_USERDATA);
    g_hwndSettings = NULL;
    g_ui.showingSettings = FALSE;
    if (owner && IsWindow(owner)) {
        EnableWindow(owner, TRUE);
        SetForegroundWindow(owner);
    }
    DestroyWindow(hwnd);
}

static void OpenSettingsDialog(HWND parent) {
    if (g_hwndSettings && IsWindow(g_hwndSettings)) {
        ShowWindow(g_hwndSettings, SW_SHOWNORMAL);
        SetForegroundWindow(g_hwndSettings);
        return;
    }

    static BOOL registered = FALSE;
    if (!registered) {
        WNDCLASSEXW wc;
        ZeroMemory(&wc, sizeof(wc));
        wc.cbSize = sizeof(wc);
        wc.lpfnWndProc = SettingsWndProc;
        wc.hInstance = GetModuleHandleW(NULL);
        wc.hCursor = LoadCursor(NULL, IDC_ARROW);
        wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
        wc.lpszClassName = L"CalicoSettings";
        wc.style = CS_HREDRAW | CS_VREDRAW;
        RegisterClassExW(&wc);
        registered = TRUE;
    }

    g_ui.showingSettings = TRUE;
    EnableWindow(parent, FALSE);
    g_hwndSettings = CreateWindowExW(
        WS_EX_DLGMODALFRAME,
        L"CalicoSettings",
        L"Calico Settings",
        WS_CAPTION | WS_SYSMENU | WS_POPUP | WS_VISIBLE,
        CW_USEDEFAULT, CW_USEDEFAULT,
        SettingsScale(parent, 390), SettingsScale(parent, 370),
        parent, NULL, GetModuleHandleW(NULL), parent);
    if (!g_hwndSettings) {
        g_ui.showingSettings = FALSE;
        EnableWindow(parent, TRUE);
        return;
    }

    RECT ownerRc = {0}, dlgRc = {0};
    GetWindowRect(parent, &ownerRc);
    GetWindowRect(g_hwndSettings, &dlgRc);
    int x = ownerRc.left + ((ownerRc.right - ownerRc.left) - (dlgRc.right - dlgRc.left)) / 2;
    int y = ownerRc.top + ((ownerRc.bottom - ownerRc.top) - (dlgRc.bottom - dlgRc.top)) / 2;
    SetWindowPos(g_hwndSettings, HWND_TOP, x, y, 0, 0, SWP_NOSIZE);
    SetForegroundWindow(g_hwndSettings);
}

static void GetLayoutRects(RECT wnd, RECT *rcInput, RECT *rcIcons, RECT *rcTabStrip, RECT *rcContent) {
    int top = wnd.top;
    int left = wnd.left;
    (void)top; (void)left; // silence unused warnings if any
    int y = wnd.top;

    rcInput->left = left + S(PADDING);
    rcInput->top = y + S(PADDING);
    rcInput->right = wnd.right - S(PADDING);
    rcInput->bottom = rcInput->top + S(INPUT_H);
    y = rcInput->bottom + S(6);

    rcIcons->left = rcInput->left;
    rcIcons->top = y;
    rcIcons->right = rcInput->right;
    rcIcons->bottom = rcIcons->top + S(ICONROW_H);
    y = rcIcons->bottom + (g_ui.symbolsOpen ? S(6) : S(10));

    if (g_ui.symbolsOpen) {
        rcTabStrip->left = rcIcons->left;
        rcTabStrip->top = rcIcons->bottom + S(4);
        rcTabStrip->right = rcIcons->right;
        rcTabStrip->bottom = rcTabStrip->top + S(TABSTRIP_H);
        y = rcTabStrip->bottom + S(6);
    } else {
        SetRect(rcTabStrip, 0, 0, 0, 0);
    }

    rcContent->left = rcInput->left;
    rcContent->top = y;
    rcContent->right = rcInput->right;
    rcContent->bottom = wnd.bottom - S(PADDING);
}

static int HitTestIconsRight(RECT rcIcons, POINT pt) {
    int gap = S(ICON_GAP);
    int iconW = S(ICON_WIDTH);
    int totalW = kIconCount * iconW + (kIconCount - 1) * gap;
    int x0 = rcIcons.right - totalW; // right-aligned
    for (int i = 0; i < kIconCount; ++i) {
        RECT cell = { x0 + i*(iconW + gap), rcIcons.top, x0 + i*(iconW + gap) + iconW, rcIcons.bottom };
        if (PtInRect(&cell, pt)) return i;
    }
    return -1;
}

static int HitTestIconsLeft(RECT rcIcons, POINT pt) {
    int gap = S(ICON_GAP);
    int iconW = S(ICON_WIDTH);
    int x0 = rcIcons.left + gap; // left-aligned
    for (int i = 0; i < kOpIconCount; ++i) {
        RECT cell = { x0 + i*(iconW + gap), rcIcons.top, x0 + i*(iconW + gap) + iconW, rcIcons.bottom };
        if (PtInRect(&cell, pt)) return i;
    }
    return -1;
}

static int HitTestTabs(RECT rcTab, POINT pt) {
    if (!g_ui.symbolsOpen) return -1;
    int gap = S(ICON_GAP);
    int x = rcTab.left;
    for (int i = 0; i < kTabCount; ++i) {
        int w = S(TAB_WIDTH);
        RECT r = { x, rcTab.top, x + w, rcTab.bottom };
        if (PtInRect(&r, pt)) return i;
        x += w + gap;
    }
    return -1;
}

static int HitTestHistory(RECT rcContent, POINT pt, int *outIndex, RECT *outRect) {
    // Each item height
    int itemH = S(ITEM_HEIGHT);
    int y0 = rcContent.top - g_ui.scrollOffset;
    // Skip live preview area for hit testing (non-clickable)
    if (g_liveVisible) {
        RECT rprev = { rcContent.left, y0, rcContent.right, y0 + itemH };
        if (PtInRect(&rprev, pt)) return FALSE;
        y0 += itemH;
    }
    for (int i = 0; i < g_historyCount; ++i) {
        RECT r = { rcContent.left, y0 + i*itemH, rcContent.right, y0 + (i+1)*itemH };
        if (r.bottom < rcContent.top) continue;
        if (r.top > rcContent.bottom) break;
        if (PtInRect(&r, pt)) {
            if (outIndex) *outIndex = i;
            if (outRect) *outRect = r;
            return TRUE;
        }
    }
    return FALSE;
}

// Double buffering to reduce flicker
static HDC g_memDC = NULL;
static HBITMAP g_memBmp = NULL;
static HBITMAP g_memOld = NULL;
static SIZE g_memSize = {0};

static void DestroyBackBuffer(void) {
    if (g_memDC) {
        if (g_memOld) { SelectObject(g_memDC, g_memOld); g_memOld = NULL; }
        if (g_memBmp) { DeleteObject(g_memBmp); g_memBmp = NULL; }
        DeleteDC(g_memDC); g_memDC = NULL;
    }
    g_memSize.cx = g_memSize.cy = 0;
}

static void EnsureBackBuffer(HWND hwnd) {
    RECT rc; GetClientRect(hwnd, &rc);
    int w = max(1, rc.right - rc.left);
    int h = max(1, rc.bottom - rc.top);
    if (g_memDC && g_memSize.cx == w && g_memSize.cy == h) return;
    DestroyBackBuffer();
    HDC hdc = GetDC(hwnd);
    g_memDC = CreateCompatibleDC(hdc);
    g_memBmp = CreateCompatibleBitmap(hdc, w, h);
    g_memOld = (HBITMAP)SelectObject(g_memDC, g_memBmp);
    g_memSize.cx = w; g_memSize.cy = h;
    ReleaseDC(hwnd, hdc);
}

static void DrawUI(HDC hdc, RECT wnd) {
    RECT rcInput, rcIcons, rcTabs, rcContent;
    GetLayoutRects(wnd, &rcInput, &rcIcons, &rcTabs, &rcContent);

    // Background
    FillRectColor(hdc, &wnd, COL_BG);

    // Input area
    RECT in = rcInput;
    FillRectColor(hdc, &in, g_ui.hoveringInput ? COL_HOVER : COL_INPUT);
    RECT inText = { in.left + S(INPUT_MARGIN), in.top, in.right - S(INPUT_MARGIN), in.bottom };

    HFONT fInput = CreateUiFont(FONT_INPUT, FALSE);
    HFONT fPrev = (HFONT)SelectObject(hdc, fInput);
    if (g_ui.inputLen > 0) {
        int xOff = inText.left;
        if (HasInputSelection()) {
            WCHAR prefix[256];
            WCHAR selected[256];
            WCHAR suffix[256];
            int prefixLen = g_ui.selStart;
            int selectedLen = g_ui.selEnd - g_ui.selStart;
            int suffixLen = g_ui.inputLen - g_ui.selEnd;

            if (prefixLen > 0) {
                wcsncpy(prefix, g_ui.input, prefixLen);
                prefix[prefixLen] = 0;
                DrawTextClipped(hdc, prefix, inText, COL_TEXT, DT_SINGLELINE | DT_VCENTER | DT_LEFT | DT_NOPREFIX);
            }

            int prefixW = MeasureTextWidth(hdc, g_ui.input, prefixLen);
            int selectedW = MeasureTextWidth(hdc, g_ui.input + g_ui.selStart, selectedLen);
            RECT selRect = {
                inText.left + prefixW,
                in.top + S(4),
                inText.left + prefixW + selectedW,
                in.bottom - S(4)
            };
            if (selRect.left > inText.right) selRect.left = inText.right;
            if (selRect.right > inText.right) selRect.right = inText.right;
            FillRectColor(hdc, &selRect, COL_ACCENT);

            if (selectedLen > 0) {
                wcsncpy(selected, g_ui.input + g_ui.selStart, selectedLen);
                selected[selectedLen] = 0;
                RECT selTextRect = { inText.left + prefixW, inText.top, inText.right, inText.bottom };
                DrawTextClipped(hdc, selected, selTextRect, COL_BG, DT_SINGLELINE | DT_VCENTER | DT_LEFT | DT_NOPREFIX);
            }

            if (suffixLen > 0) {
                wcsncpy(suffix, g_ui.input + g_ui.selEnd, suffixLen);
                suffix[suffixLen] = 0;
                RECT suffixRect = { inText.left + prefixW + selectedW, inText.top, inText.right, inText.bottom };
                DrawTextClipped(hdc, suffix, suffixRect, COL_TEXT, DT_SINGLELINE | DT_VCENTER | DT_LEFT | DT_NOPREFIX);
            }

            xOff = inText.left + MeasureTextWidth(hdc, g_ui.input, g_ui.inputLen);
        } else {
            DrawTextClipped(hdc, g_ui.input, inText, COL_TEXT, DT_SINGLELINE | DT_VCENTER | DT_LEFT | DT_NOPREFIX);
            xOff = inText.left + MeasureTextWidth(hdc, g_ui.input, g_ui.inputLen);
        }
        // Draw command suggestion suffix inline
        if (g_ui.suggestActive && g_ui.suggestLen > 0) {
            RECT sugRc = { xOff, inText.top, inText.right, inText.bottom };
            DrawTextClipped(hdc, g_ui.suggest, sugRc, COL_SUGGEST, DT_SINGLELINE | DT_VCENTER | DT_LEFT | DT_NOPREFIX);
            // advance x offset by suggestion width
            xOff += MeasureTextWidth(hdc, g_ui.suggest, g_ui.suggestLen);
        }
        // Draw ghost closing parentheses if any
        if (g_ui.ghostParensActive && g_ui.ghostParensLen > 0) {
            RECT ghostRc = { xOff, inText.top, inText.right, inText.bottom };
            DrawTextClipped(hdc, g_ui.ghostParens, ghostRc, COL_SUBTLE, DT_SINGLELINE | DT_VCENTER | DT_LEFT | DT_NOPREFIX);
        }
    } else {
        // When empty, show last answer as a grey inline suggestion if available and not suppressed
        if (!g_ui.suppressLastAnsSuggest) {
            if (!g_ui.lastAnsSuggestActive) UpdateLastAnswerSuggestion();
            if (g_ui.lastAnsSuggestActive && g_ui.lastAnsLen > 0) {
                DrawTextClipped(hdc, g_ui.lastAns, inText, COL_SUGGEST, DT_SINGLELINE | DT_VCENTER | DT_LEFT | DT_NOPREFIX);
            } else {
                DrawTextClipped(hdc, L"Type expression...", inText, COL_SUBTLE, DT_SINGLELINE | DT_VCENTER | DT_LEFT | DT_NOPREFIX);
            }
        } else {
            DrawTextClipped(hdc, L"Type expression...", inText, COL_SUBTLE, DT_SINGLELINE | DT_VCENTER | DT_LEFT | DT_NOPREFIX);
        }
    }
    // Draw blinking caret when focused
    if (g_ui.hasFocus && g_ui.caretVisible) {
        int caretX = inText.left + MeasureTextWidth(hdc, g_ui.input, g_ui.caretPos);
        int caretW = max(1, S(CARET_WIDTH));
        RECT caret = {
            caretX,
            in.top + S(CARET_MARGIN),
            caretX + caretW,
            in.bottom - S(CARET_MARGIN)
        };
        FillRectColor(hdc, &caret, COL_TEXT);
    }
    SelectObject(hdc, fPrev);
    ReleaseUiFont(fInput, FONT_INPUT, FALSE);

    // Icon row
    FillRectColor(hdc, &rcIcons, COL_ICONROW);
    int gap = S(ICON_GAP), iconW = S(ICON_WIDTH);
    HFONT fIcon = CreateUiFont(FONT_ICON, TRUE); fPrev = (HFONT)SelectObject(hdc, fIcon);
    // Left operators (draw +, -, ×, ÷)
    int lx0 = rcIcons.left + gap;
    for (int i = 0; i < kOpIconCount; ++i) {
        RECT cell = { lx0 + i*(iconW + gap), rcIcons.top, lx0 + i*(iconW + gap) + iconW, rcIcons.bottom };
        if (g_ui.hoverLeftIcon == i) FillRectColor(hdc, &cell, COL_HOVER);
        WCHAR sym = L' ';
        if (i == 0) sym = L'+';
        else if (i == 1) sym = L'-';
        else if (i == 2) sym = CH_MUL;
        else if (i == 3) sym = CH_DIV;
        WCHAR safe[2] = { sym, 0 };
        DrawTextClipped(hdc, safe, cell, COL_TEXT, DT_SINGLELINE | DT_CENTER | DT_VCENTER);
    }
    // Right side icons
    int rx0 = rcIcons.right - (kIconCount * iconW + (kIconCount - 1) * gap);
    for (int i = 0; i < kIconCount; ++i) {
        RECT cell = { rx0 + i*(iconW + gap), rcIcons.top, rx0 + i*(iconW + gap) + iconW, rcIcons.bottom };
        if (g_ui.hoverRightIcon == i) FillRectColor(hdc, &cell, COL_HOVER);
        WCHAR safeMenu[2] = { L'\u2261', 0 }; // Hamburger menu symbol ≡
        DrawTextClipped(hdc, safeMenu, cell, COL_TEXT, DT_SINGLELINE | DT_CENTER | DT_VCENTER);
    }
    SelectObject(hdc, fPrev);
    ReleaseUiFont(fIcon, FONT_ICON, TRUE);

    // Symbols tab strip (inline)
    if (g_ui.symbolsOpen) {
        FillRectColor(hdc, &rcTabs, COL_ICONROW);
        int x = rcTabs.left;
        HFONT fTab = CreateUiFont(FONT_TAB, FALSE); fPrev = (HFONT)SelectObject(hdc, fTab);
        for (int i = 0; i < kTabCount; ++i) {
            int w = S(TAB_WIDTH);
            RECT r = { x, rcTabs.top, x + w, rcTabs.bottom };
            if (g_ui.activeTab == i) {
                RECT rr = r; rr.bottom -= S(2);
                FillRectColor(hdc, &rr, COL_HOVER);
                // underline accent
                RECT ul = { r.left + S(6), r.bottom - S(2), r.left + w - S(6), r.bottom };
                FillRectColor(hdc, &ul, COL_ACCENT);
            }
            DrawTextClipped(hdc, kTabs[i], r, COL_TEXT, DT_SINGLELINE | DT_VCENTER | DT_CENTER);
            x += w + gap;
        }
        SelectObject(hdc, fPrev);
        ReleaseUiFont(fTab, FONT_TAB, FALSE);
    }

    // Content area (history list for now) — hide when no rows to display
    int rows = g_historyCount + (g_liveVisible ? 1 : 0);
    if (rows > 0 && rcContent.bottom > rcContent.top) {
        FillRectColor(hdc, &rcContent, COL_CONTENT);
        int itemH = S(ITEM_HEIGHT);
        HFONT fItem = CreateUiFont(FONT_ITEM, FALSE); fPrev = (HFONT)SelectObject(hdc, fItem);
        int y0 = rcContent.top - g_ui.scrollOffset;
        int saved = SaveDC(hdc);
        IntersectClipRect(hdc, rcContent.left, rcContent.top, rcContent.right, rcContent.bottom);
        // Live preview row at top
        if (g_liveVisible) {
            RECT rprev = { rcContent.left + S(8), y0, rcContent.right - S(8), y0 + itemH };
            DrawTextClipped(hdc, g_liveText, rprev, COL_TEXT, DT_SINGLELINE | DT_VCENTER | DT_LEFT);
            y0 += itemH;
        }
        for (int i = 0; i < g_historyCount; ++i) {
            RECT r = { rcContent.left + S(8), y0 + i*itemH, rcContent.right - S(8), y0 + (i+1)*itemH };
            if (r.bottom < rcContent.top) continue;
            if (r.top > rcContent.bottom) break;
            if (g_ui.hoverHistory == i) {
                RECT hr = r; hr.left -= S(8); hr.right += S(8);
                FillRectColor(hdc, &hr, COL_HOVER);
            }
            if (g_history[i] && g_history[i]->text)
                DrawTextClipped(hdc, g_history[i]->text, r, COL_TEXT, DT_SINGLELINE | DT_VCENTER | DT_LEFT);
        }
        RestoreDC(hdc, saved);
        SelectObject(hdc, fPrev);
        ReleaseUiFont(fItem, FONT_ITEM, FALSE);
    }
}

static void SpawnNearCursor(HWND hwnd) {
    POINT pt; GetCursorPos(&pt);
    HMONITOR mon = MonitorFromPoint(pt, MONITOR_DEFAULTTONEAREST);
    int oldDpi = g_ui.dpi;
    int oldAutoScale = g_ui.autoScalePercent;
    ApplyScaleMetrics(hwnd);
    if (mon) {
        ApplyScaleMetricsForMonitor(mon, g_ui.dpi);
        if (g_ui.autoScalePercent != oldAutoScale || g_ui.dpi != oldDpi) {
            CleanupFontCache();
            EnsureBackBuffer(hwnd);
        }
    }
    int w = g_ui.panelSizePx.cx;
    int h = DesiredWindowHeightPx();
    int x = pt.x + S(12);
    int y = pt.y - S(20);
    MONITORINFO mi = { .cbSize = sizeof(mi) };
    if (GetMonitorInfoW(mon, &mi)) {
        // Keep inside work area
        if (x + w > mi.rcWork.right) x = mi.rcWork.right - w - S(8);
        if (y + h > mi.rcWork.bottom) y = mi.rcWork.bottom - h - S(8);
        if (y < mi.rcWork.top) y = mi.rcWork.top + S(8);
    }
    SetWindowPos(hwnd, g_settings.alwaysOnTop ? HWND_TOPMOST : HWND_NOTOPMOST, x, y, w, h, SWP_SHOWWINDOW);
    ShowWindow(hwnd, SW_SHOW);
    SetForegroundWindow(hwnd);
    SetFocus(hwnd);
}

static LRESULT CALLBACK SettingsWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_CREATE: {
        CREATESTRUCTW *cs = (CREATESTRUCTW*)lParam;
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, (LONG_PTR)cs->lpCreateParams);
        CreateSettingsControls(hwnd);
        SyncSettingsControls(hwnd);
        return 0;
    }
    case WM_COMMAND: {
        switch (LOWORD(wParam)) {
        case IDC_SETTINGS_SAVE: {
            HWND owner = (HWND)GetWindowLongPtrW(hwnd, GWLP_USERDATA);
            AppSettings oldSettings = g_settings;
            WCHAR buf[16];
            DWORD hotkeyValue = (DWORD)SendDlgItemMessageW(hwnd, IDC_SETTINGS_HOTKEY, HKM_GETHOTKEY, 0, 0);
            UINT hotkeyVk = ClampHotkeyVk(LOBYTE(hotkeyValue));
            UINT hotkeyMods = HotkeyFlagsToModMask(HIBYTE(hotkeyValue));
            LRESULT dateFormatSel = SendDlgItemMessageW(hwnd, IDC_SETTINGS_DATEFORMAT, CB_GETCURSEL, 0, 0);
            GetDlgItemTextW(hwnd, IDC_SETTINGS_OPACITY, buf, ARRAY_COUNT(buf));
            int opacity = _wtoi(buf);
            GetDlgItemTextW(hwnd, IDC_SETTINGS_SCALE, buf, ARRAY_COUNT(buf));
            int interfaceScale = _wtoi(buf);
            if (hotkeyVk == 0 || hotkeyMods == 0) {
                MessageBoxW(hwnd, L"Choose a hotkey with at least one modifier key.", L"Calico Settings", MB_OK | MB_ICONWARNING);
                SetFocus(GetDlgItem(hwnd, IDC_SETTINGS_HOTKEY));
                return 0;
            }
            if (opacity < 40 || opacity > 100) {
                MessageBoxW(hwnd, L"Opacity must be between 40 and 100.", L"Calico Settings", MB_OK | MB_ICONWARNING);
                SetFocus(GetDlgItem(hwnd, IDC_SETTINGS_OPACITY));
                return 0;
            }
            if (interfaceScale < 75 || interfaceScale > 200) {
                MessageBoxW(hwnd, L"Scale must be between 75 and 200.", L"Calico Settings", MB_OK | MB_ICONWARNING);
                SetFocus(GetDlgItem(hwnd, IDC_SETTINGS_SCALE));
                return 0;
            }
            g_settings.hotkeyVk = hotkeyVk;
            g_settings.hotkeyMods = hotkeyMods;
            g_settings.opacityPercent = opacity;
            g_settings.interfaceScalePercent = interfaceScale;
            g_settings.alwaysOnTop = (IsDlgButtonChecked(hwnd, IDC_SETTINGS_TOPMOST) == BST_CHECKED);
            g_settings.symbolsOpenByDefault = (IsDlgButtonChecked(hwnd, IDC_SETTINGS_SYMBOLS) == BST_CHECKED);
            g_settings.dateFormatStyle = ClampDateFormatStyle((dateFormatSel == CB_ERR) ? DATEFMT_MDY : (int)dateFormatSel);
            if (owner && IsWindow(owner) && !RegisterConfiguredHotkey(owner, FALSE)) {
                WCHAR hotkeyText[64];
                g_settings = oldSettings;
                RegisterConfiguredHotkey(owner, FALSE);
                FormatHotkeyText(hotkeyMods, hotkeyVk, hotkeyText, ARRAY_COUNT(hotkeyText));
                WCHAR msg[192];
                _snwprintf_s(msg, ARRAY_COUNT(msg), _TRUNCATE, L"The hotkey %ls is unavailable. Choose another shortcut.", hotkeyText);
                MessageBoxW(hwnd, msg, L"Calico Settings", MB_OK | MB_ICONWARNING);
                SyncSettingsControls(hwnd);
                SetFocus(GetDlgItem(hwnd, IDC_SETTINGS_HOTKEY));
                return 0;
            }
            if (!SaveSettings()) {
                g_settings = oldSettings;
                if (owner && IsWindow(owner)) RegisterConfiguredHotkey(owner, FALSE);
                MessageBoxW(hwnd, L"Could not save settings.", L"Calico Settings", MB_OK | MB_ICONERROR);
                return 0;
            }
            if (owner && IsWindow(owner)) {
                ApplyScaleMetrics(owner);
                CleanupFontCache();
                ApplyWindowVisuals(owner);
                g_ui.symbolsOpen = g_settings.symbolsOpenByDefault;
                SetWindowPos(owner, NULL, 0, 0, g_ui.panelSizePx.cx, DesiredWindowHeightPx(), SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
                EnsureBackBuffer(owner);
                AdjustWindowHeight(owner);
                InvalidateRect(owner, NULL, TRUE);
            }
            CloseSettingsDialog(hwnd);
            return 0;
        }
        case IDC_SETTINGS_CANCEL:
        case IDCANCEL:
            CloseSettingsDialog(hwnd);
            return 0;
        }
        break;
    }
    case WM_CLOSE:
        CloseSettingsDialog(hwnd);
        return 0;
    }
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

static LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    if (g_wmTaskbarCreated != 0 && msg == g_wmTaskbarCreated) {
        AddTrayIcon(hwnd);
        return 0;
    }

    switch (msg) {
    case WM_CREATE: {
        LoadSettings();
        InitDpi(hwnd);
        g_ui.symbolsOpen = g_settings.symbolsOpenByDefault;
        ApplyWindowVisuals(hwnd);
        
        // Initialize lookup tables
        InitializeHolidayLookup();
        
        // Size window now (dynamic height based on history)
        SetWindowPos(hwnd, HWND_TOPMOST, 100, 100, g_ui.panelSizePx.cx, DesiredWindowHeightPx(), SWP_NOZORDER);
        EnsureBackBuffer(hwnd);
        AddTrayIcon(hwnd);
        if (!RegisterConfiguredHotkey(hwnd, TRUE)) {
            SpawnNearCursor(hwnd);
        }
        // Start with empty history
        g_ui.historyNavIndex = -1;
        return 0;
    }
    case WM_DESTROY:
        RemoveTrayIcon(hwnd);
        if (g_hwndSettings && IsWindow(g_hwndSettings)) DestroyWindow(g_hwndSettings);
        UnregisterHotKey(hwnd, HOTKEY_ID);
        History_FreeAll();
        DestroyBackBuffer();
        CleanupFontCache();
        PostQuitMessage(0);
        return 0;
    case WMAPP_TRAYICON:
        switch ((UINT)lParam) {
        case WM_LBUTTONUP:
        case WM_LBUTTONDBLCLK:
            ToggleTrayWindow(hwnd);
            return 0;
        case WM_RBUTTONUP:
        case WM_CONTEXTMENU:
            ShowTrayMenu(hwnd);
            return 0;
        }
        return 0;
    case WM_HOTKEY:
        if (wParam == HOTKEY_ID) {
            SpawnNearCursor(hwnd);
        }
        return 0;
    case WM_DPICHANGED: {
        g_ui.dpi = HIWORD(wParam);
        ApplyScaleMetrics(hwnd);
        g_ui.panelSizePx.cy = DesiredWindowHeightPx();
        
        // Invalidate font cache since DPI changed
        CleanupFontCache();
        
        RECT *suggested = (RECT*)lParam;
        SetWindowPos(hwnd, NULL, suggested->left, suggested->top,
                     g_ui.panelSizePx.cx, g_ui.panelSizePx.cy,
                     SWP_NOZORDER | SWP_NOACTIVATE);
        EnsureBackBuffer(hwnd);
        InvalidateRect(hwnd, NULL, TRUE);
        return 0;
    }
    case WM_ERASEBKGND:
        return 1; // reduce flicker, we draw the whole background
    case WM_SIZE:
        EnsureBackBuffer(hwnd);
        return 0;
    case WM_PAINT: {
        EnsureBackBuffer(hwnd);
        PAINTSTRUCT ps; HDC hdc = BeginPaint(hwnd, &ps);
        RECT wnd; GetClientRect(hwnd, &wnd);
        DrawUI(g_memDC, wnd);
        BitBlt(hdc, 0, 0, g_memSize.cx, g_memSize.cy, g_memDC, 0, 0, SRCCOPY);
        EndPaint(hwnd, &ps);
        return 0;
    }
    case WM_COMMAND: {
        // Quick remap for certain inserts before the main switch
        UINT id = LOWORD(wParam);
        if (id == IDM_INS_TIMES) { InsertAtCaret(L"*"); RefreshInputDisplay(hwnd); return 0; }
        if (id == IDM_INS_DIV)   { InsertAtCaret(L"/"); RefreshInputDisplay(hwnd); return 0; }
        if (id == IDM_TEMP_C)    { InsertAtCaret(L" \x00B0C"); RefreshInputDisplay(hwnd); return 0; }
        if (id == IDM_TEMP_F)    { InsertAtCaret(L" \x00B0F"); RefreshInputDisplay(hwnd); return 0; }
        // Context menu commands
        switch (LOWORD(wParam)) {
        case IDM_TRAY_SHOW:
            ToggleTrayWindow(hwnd);
            return 0;
        case IDM_TRAY_SETTINGS:
            OpenSettingsDialog(hwnd);
            return 0;
        case IDM_TRAY_EXIT:
            DestroyWindow(hwnd);
            return 0;
        case IDM_CLEAR_INPUT:
            g_ui.inputLen = 0; g_ui.input[0] = 0; g_ui.caretPos = 0;
            ClearInputSelection();
            RefreshInputDisplay(hwnd);
            return 0;
        case IDM_SETTINGS:
            OpenSettingsDialog(hwnd);
            return 0;
        case IDM_INS_PLUS:    InsertAtCaret(L"+"); break;
        case IDM_INS_MINUS:   InsertAtCaret(L"-"); break;
        case IDM_INS_TIMES:   InsertAtCaret(L"×"); break;
        case IDM_INS_DIV:     InsertAtCaret(L"÷"); break;
        case IDM_INS_LPAREN:  InsertAtCaret(L"("); break;
        case IDM_INS_RPAREN:  InsertAtCaret(L")"); break;
        case IDM_INS_PERCENT: InsertAtCaret(L"%"); break;
        case IDM_INS_POWER:   InsertAtCaret(L"^"); break;
        // Lengths
        case IDM_MEAS_IN:     InsertAtCaret(L" in"); break;
        case IDM_MEAS_FT:     InsertAtCaret(L" ft"); break;
        case IDM_MEAS_YD:     InsertAtCaret(L" yd"); break;
        case IDM_MEAS_M:      InsertAtCaret(L" m"); break;
        case IDM_MEAS_CM:     InsertAtCaret(L" cm"); break;
        case IDM_MEAS_MM:     InsertAtCaret(L" mm"); break;
        case IDM_MEAS_KM:     InsertAtCaret(L" km"); break;
        case IDM_MEAS_MI:     InsertAtCaret(L" mi"); break;
        case IDM_MEAS_NMI:    InsertAtCaret(L" nmi"); break;
        case IDM_MEAS_LY:     InsertAtCaret(L" ly"); break;
        // Areas
        case IDM_AREA_SQIN:     InsertAtCaret(L" sq in"); break;
        case IDM_AREA_SQFT:     InsertAtCaret(L" sq ft"); break;
        case IDM_AREA_SQYD:     InsertAtCaret(L" sq yd"); break;
        case IDM_AREA_SQM:      InsertAtCaret(L" sq m"); break;
        case IDM_AREA_SQKM:     InsertAtCaret(L" sq km"); break;
        case IDM_AREA_SQMI:     InsertAtCaret(L" sq mi"); break;
        case IDM_AREA_ACRE:     InsertAtCaret(L" acre"); break;
        case IDM_AREA_HECTARE:  InsertAtCaret(L" hectare"); break;
        // Weight
        case IDM_WEIGHT_MG:     InsertAtCaret(L" mg"); break;
        case IDM_WEIGHT_G:      InsertAtCaret(L" g"); break;
        case IDM_WEIGHT_KG:     InsertAtCaret(L" kg"); break;
        case IDM_WEIGHT_OZ:     InsertAtCaret(L" oz"); break;
        case IDM_WEIGHT_LB:     InsertAtCaret(L" lb"); break;
        case IDM_WEIGHT_TON:    InsertAtCaret(L" ton"); break;
        // Volume
        case IDM_VOLUME_ML:     InsertAtCaret(L" mL"); break;
        case IDM_VOLUME_L:      InsertAtCaret(L" L"); break;
        case IDM_VOLUME_FLOZ:   InsertAtCaret(L" fl oz"); break;
        case IDM_VOLUME_CUP:    InsertAtCaret(L" cup"); break;
        case IDM_VOLUME_PINT:   InsertAtCaret(L" pt"); break;
        case IDM_VOLUME_QUART:  InsertAtCaret(L" qt"); break;
        case IDM_VOLUME_GAL:    InsertAtCaret(L" gal"); break;
        case IDM_VOLUME_TBSP:   InsertAtCaret(L" tbsp"); break;
        case IDM_VOLUME_TSP:    InsertAtCaret(L" tsp"); break;
        // Memory
        case IDM_MEMORY_BIT:    InsertAtCaret(L" bit"); break;
        case IDM_MEMORY_BYTE:   InsertAtCaret(L" byte"); break;
        case IDM_MEMORY_KBIT:   InsertAtCaret(L" Kb"); break;
        case IDM_MEMORY_MBIT:   InsertAtCaret(L" Mb"); break;
        case IDM_MEMORY_GBIT:   InsertAtCaret(L" Gb"); break;
        case IDM_MEMORY_TBIT:   InsertAtCaret(L" Tb"); break;
        case IDM_MEMORY_KB:     InsertAtCaret(L" KB"); break;
        case IDM_MEMORY_MB:     InsertAtCaret(L" MB"); break;
        case IDM_MEMORY_GB:     InsertAtCaret(L" GB"); break;
        case IDM_MEMORY_TB:     InsertAtCaret(L" TB"); break;
        case IDM_MEMORY_KIB:    InsertAtCaret(L" KiB"); break;
        case IDM_MEMORY_MIB:    InsertAtCaret(L" MiB"); break;
        case IDM_MEMORY_GIB:    InsertAtCaret(L" GiB"); break;
        case IDM_MEMORY_TIB:    InsertAtCaret(L" TiB"); break;

        // Temperatures
        case IDM_TEMP_C:      InsertAtCaret(L" °C"); break;
        case IDM_TEMP_F:      InsertAtCaret(L" °F"); break;
        case IDM_TEMP_K:      InsertAtCaret(L" K"); break;
        // Time
        case IDM_TIME_NS:     InsertAtCaret(L" ns"); break;
        case IDM_TIME_US:     InsertAtCaret(L" \x00B5s"); break;
        case IDM_TIME_MS:     InsertAtCaret(L" ms"); break;
        case IDM_TIME_S:      InsertAtCaret(L" s"); break;
        case IDM_TIME_MIN:    InsertAtCaret(L" min"); break;
        case IDM_TIME_H:      InsertAtCaret(L" h"); break;
        case IDM_TIME_DAY:    InsertAtCaret(L" day"); break;
        case IDM_TIME_WEEK:   InsertAtCaret(L" week"); break;
        case IDM_TIME_MONTH:  InsertAtCaret(L" month"); break;
        case IDM_TIME_YEAR:   InsertAtCaret(L" year"); break;
        // Holidays
        case IDM_HOLIDAY_NEWYEAR_DAY:   InsertAtCaret(L"New Year's Day"); break;
        case IDM_HOLIDAY_NEWYEAR_EVE:   InsertAtCaret(L"New Year's Eve"); break;
        case IDM_HOLIDAY_MLK:           InsertAtCaret(L"Martin Luther King Jr. Day"); break;
        case IDM_HOLIDAY_PRESIDENTS:    InsertAtCaret(L"Presidents' Day"); break;
        case IDM_HOLIDAY_MEMORIAL:      InsertAtCaret(L"Memorial Day"); break;
        case IDM_HOLIDAY_INDEPENDENCE:  InsertAtCaret(L"Independence Day"); break;
        case IDM_HOLIDAY_LABOR:         InsertAtCaret(L"Labor Day"); break;
        case IDM_HOLIDAY_COLUMBUS:      InsertAtCaret(L"Columbus Day"); break;
        case IDM_HOLIDAY_HALLOWEEN:     InsertAtCaret(L"Halloween"); break;
        case IDM_HOLIDAY_VETERANS:      InsertAtCaret(L"Veterans Day"); break;
        case IDM_HOLIDAY_THANKSGIVING:  InsertAtCaret(L"Thanksgiving"); break;
        case IDM_HOLIDAY_CHRISTMAS_EVE: InsertAtCaret(L"Christmas Eve"); break;
        case IDM_HOLIDAY_CHRISTMAS_DAY: InsertAtCaret(L"Christmas Day"); break;
        case IDM_HOLIDAY_VALENTINE:     InsertAtCaret(L"Valentine's Day"); break;
        case IDM_HOLIDAY_MOTHERS:       InsertAtCaret(L"Mother's Day"); break;
        case IDM_HOLIDAY_FATHERS:       InsertAtCaret(L"Father's Day"); break;
        default: return 0;
        }
        RefreshInputDisplay(hwnd);
        return 0;
    }
    case WM_MEASUREITEM: {
        LPMEASUREITEMSTRUCT mis = (LPMEASUREITEMSTRUCT)lParam;
        if (mis->CtlType == ODT_MENU) {
            HFONT f = CreateUiFont(FONT_INPUT, FALSE);
            HDC hdc = GetDC(hwnd);
            HFONT old = (HFONT)SelectObject(hdc, f);
            const WCHAR *txt = (const WCHAR*)mis->itemData;
            SIZE sz = { 0 };
            if (txt) GetTextExtentPoint32W(hdc, txt, (int)wcslen(txt), &sz);
            SelectObject(hdc, old);
            ReleaseUiFont(f, FONT_INPUT, FALSE);
            ReleaseDC(hwnd, hdc);
            mis->itemHeight = S(MENU_ITEM_HEIGHT);
            mis->itemWidth = sz.cx + S(MENU_ITEM_PADDING);
            return TRUE;
        }
        break;
    }
    case WM_DRAWITEM: {
        LPDRAWITEMSTRUCT dis = (LPDRAWITEMSTRUCT)lParam;
        if (dis->CtlType == ODT_MENU) {
            RECT rc = dis->rcItem;
            COLORREF bg = (dis->itemState & ODS_SELECTED) ? COL_HOVER : COL_ICONROW;
            HBRUSH b = CreateSolidBrush(bg);
            FillRect(dis->hDC, &rc, b);
            DeleteObject(b);
            // Draw subtle inner border for visual separation from system frame
            HPEN pen = CreatePen(PS_SOLID, 1, COL_BG);
            HPEN oldPen = (HPEN)SelectObject(dis->hDC, pen);
            HBRUSH oldBrush = (HBRUSH)SelectObject(dis->hDC, GetStockObject(HOLLOW_BRUSH));
            Rectangle(dis->hDC, rc.left, rc.top, rc.right, rc.bottom);
            SelectObject(dis->hDC, oldBrush);
            SelectObject(dis->hDC, oldPen);
            DeleteObject(pen);
            const WCHAR *txt = (const WCHAR*)dis->itemData;
            if (txt) {
                HFONT f = CreateUiFont(FONT_INPUT, FALSE);
                HFONT old = (HFONT)SelectObject(dis->hDC, f);
                SetBkMode(dis->hDC, TRANSPARENT);
                SetTextColor(dis->hDC, COL_TEXT);
                RECT tr = { rc.left + S(MENU_TEXT_MARGIN), rc.top, rc.right - S(MENU_TEXT_MARGIN), rc.bottom };
                DrawTextW(dis->hDC, txt, -1, &tr, DT_SINGLELINE | DT_VCENTER | DT_LEFT | DT_NOPREFIX);
                SelectObject(dis->hDC, old);
                ReleaseUiFont(f, FONT_INPUT, FALSE);
            }
            return TRUE;
        }
        break;
    }
    case WM_MOUSEMOVE: {
        TRACKMOUSEEVENT tme = { sizeof(tme), TME_LEAVE, hwnd, 0 };
        TrackMouseEvent(&tme);
        POINT pt = { GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
        RECT wnd; GetClientRect(hwnd, &wnd);
        RECT rcInput, rcIcons, rcTabs, rcContent;
        GetLayoutRects(wnd, &rcInput, &rcIcons, &rcTabs, &rcContent);
        if (g_ui.selectingInput) {
            RECT inText = {
                rcInput.left + S(INPUT_MARGIN),
                rcInput.top,
                rcInput.right - S(INPUT_MARGIN),
                rcInput.bottom
            };
            int pos = CaretIndexFromPoint(hwnd, inText, pt.x);
            SetInputSelection(g_ui.selAnchor, pos);
            ResetCaretBlink(hwnd);
            InvalidateRect(hwnd, &rcInput, FALSE);
        }
        BOOL hoverInput = PtInRect(&rcInput, pt);
        int iconR = HitTestIconsRight(rcIcons, pt);
        int iconL = HitTestIconsLeft(rcIcons, pt);
        int histIdx = -1; RECT tmp;
        if (PtInRect(&rcContent, pt)) {
            HitTestHistory(rcContent, pt, &histIdx, &tmp);
        }
        BOOL changed = (hoverInput != g_ui.hoveringInput) || (iconR != g_ui.hoverRightIcon) || (iconL != g_ui.hoverLeftIcon) || (histIdx != g_ui.hoverHistory);
        g_ui.hoveringInput = hoverInput;
        g_ui.hoverRightIcon = iconR;
        g_ui.hoverLeftIcon = iconL;
        g_ui.hoverHistory = histIdx;
        if (changed) InvalidateRect(hwnd, NULL, FALSE);
        return 0;
    }
    case WM_MOUSELEAVE:
        g_ui.hoveringInput = FALSE; g_ui.hoverRightIcon = -1; g_ui.hoverLeftIcon = -1; g_ui.hoverHistory = -1; InvalidateRect(hwnd, NULL, FALSE);
        return 0;
    case WM_RBUTTONUP: {
        POINT pt; GetCursorPos(&pt);
        ShowContextMenu(hwnd, pt.x, pt.y);
        return 0;
    }
    case WM_LBUTTONDOWN: {
        POINT pt = { GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
        RECT wnd; GetClientRect(hwnd, &wnd);
        RECT rcInput, rcIcons, rcTabs, rcContent;
        GetLayoutRects(wnd, &rcInput, &rcIcons, &rcTabs, &rcContent);
        int iconR = HitTestIconsRight(rcIcons, pt);
        int iconL = HitTestIconsLeft(rcIcons, pt);
        if (iconR >= 0) {
            // Open the same context menu when clicking the menu (≡) icon
            POINT scr = pt; ClientToScreen(hwnd, &scr);
            ShowContextMenu(hwnd, scr.x, scr.y);
            return 0;
        }
        if (iconL >= 0) {
            // Insert operator symbol at caret (map ×/÷ to * and / for parsing)
            const WCHAR *op = kOpIcons[iconL];
            if (iconL == 2) op = L"*";       // multiply
            else if (iconL == 3) op = L"/";  // divide
            if (g_ui.inputLen == 0) { InsertLastAnswerIfEmpty(hwnd); }
            if (InsertAtCaret(op)) {
                RefreshInputDisplay(hwnd);
            }
            SetFocus(hwnd);
            return 0;
        }
        if (PtInRect(&rcInput, pt)) {
            SetFocus(hwnd);
            g_ui.hasFocus = TRUE;
            // Compute caret index from click X
            RECT inText = {
                rcInput.left + S(INPUT_MARGIN),
                rcInput.top,
                rcInput.right - S(INPUT_MARGIN),
                rcInput.bottom
            };
            int pos = CaretIndexFromPoint(hwnd, inText, pt.x);
            if (GetKeyState(VK_SHIFT) & 0x8000) SetInputSelection(g_ui.selAnchor, pos);
            else {
                g_ui.caretPos = pos;
                ClearInputSelection();
            }
            g_ui.selectingInput = TRUE;
            SetCapture(hwnd);
            // Make caret visible immediately and reset blink timer
            g_ui.caretVisible = TRUE;
            KillTimer(hwnd, CARET_TIMER_ID);
            SetTimer(hwnd, CARET_TIMER_ID, CARET_BLINK_MS, NULL);
            InvalidateRect(hwnd, &rcInput, FALSE);
            return 0;
        }
        if (g_ui.selectingInput) {
            g_ui.selectingInput = FALSE;
            ReleaseCapture();
        }
        int tab = HitTestTabs(rcTabs, pt);
        if (tab >= 0) { g_ui.activeTab = tab; InvalidateRect(hwnd, NULL, FALSE); return 0; }
        // Click on history -> load into input
        if (PtInRect(&rcContent, pt)) {
            int idx = -1; RECT r;
            if (HitTestHistory(rcContent, pt, &idx, &r) && idx >= 0 && idx < g_historyCount && g_history[idx]) {
                HistoryItem *it = g_history[idx];
                if (it->answer && *it->answer) {
                    // Insert the answer at caret
                    InsertAtCaret(it->answer);
                } else if (it->insert && *it->insert) {
                    // Fallback to command insertion for non-answer rows
                    InsertAtCaret(it->insert);
                }
                g_ui.historyNavIndex = -1;
                UpdateCommandSuggestion();
                UpdateLivePreview(hwnd);
                InvalidateRect(hwnd, NULL, FALSE);
                return 0;
            }
        }
        return 0;
    }
    case WM_LBUTTONUP:
        if (g_ui.selectingInput) {
            g_ui.selectingInput = FALSE;
            ReleaseCapture();
        }
        return 0;
    case WM_CAPTURECHANGED:
        g_ui.selectingInput = FALSE;
        return 0;
    case WM_SETFOCUS:
        g_ui.hasFocus = TRUE; g_ui.caretVisible = TRUE; SetTimer(hwnd, CARET_TIMER_ID, CARET_BLINK_MS, NULL); InvalidateRect(hwnd, NULL, FALSE);
        return 0;
    case WM_KILLFOCUS:
        g_ui.hasFocus = FALSE; g_ui.caretVisible = FALSE; KillTimer(hwnd, CARET_TIMER_ID); InvalidateRect(hwnd, NULL, FALSE);
        // Hide the panel when it loses focus unless a menu or modal settings window is open.
        if (!g_ui.showingMenu && !g_ui.showingSettings) ShowWindow(hwnd, SW_HIDE);
        return 0;
    case WM_ACTIVATE:
        if (LOWORD(wParam) == WA_INACTIVE) {
            if (!g_ui.showingMenu && !g_ui.showingSettings) ShowWindow(hwnd, SW_HIDE);
            return 0;
        }
        break;
    case WM_CONTEXTMENU: {
        int x = GET_X_LPARAM(lParam);
        int y = GET_Y_LPARAM(lParam);
        if (x == -1 && y == -1) {
            POINT pt; GetCursorPos(&pt); x = pt.x; y = pt.y;
        }
        ShowContextMenu(hwnd, x, y);
        return 0;
    }
    case WM_TIMER:
        if (wParam == (WPARAM)CARET_TIMER_ID) {
            if (g_ui.hasFocus) {
                g_ui.caretVisible = !g_ui.caretVisible;
                // Invalidate only input area for efficiency
                RECT wnd; GetClientRect(hwnd, &wnd);
                RECT rcInput, rcIcons, rcTabs, rcContent;
                GetLayoutRects(wnd, &rcInput, &rcIcons, &rcTabs, &rcContent);
                InvalidateRect(hwnd, &rcInput, FALSE);
                if (g_timePreviewActive) {
                    // Refresh live preview for time queries
                    UpdateLivePreview(hwnd);
                }
            }
            return 0;
        }
        break;
    case WM_MOUSEWHEEL: {
        int delta = GET_WHEEL_DELTA_WPARAM(wParam);
        g_ui.scrollOffset -= (delta / WHEEL_DELTA) * S(SCROLL_STEP);
        if (g_ui.scrollOffset < 0) g_ui.scrollOffset = 0;
        // Clamp to content height
        int itemH = S(ITEM_HEIGHT);
        RECT wnd; GetClientRect(hwnd, &wnd);
        RECT rcInput, rcIcons, rcTabs, rcContent;
        GetLayoutRects(wnd, &rcInput, &rcIcons, &rcTabs, &rcContent);
        
        // Add overflow check
        int count = g_historyCount + (g_liveVisible ? 1 : 0);
        if (count > 1000) count = 1000;  // Safety limit
        int total = count * itemH;
        
        int vis = rcContent.bottom - rcContent.top;
        int maxOff = total > vis ? total - vis : 0;
        if (g_ui.scrollOffset > maxOff) g_ui.scrollOffset = maxOff;
        InvalidateRect(hwnd, NULL, FALSE);
        return 0;
    }
    case WM_KEYDOWN:
        if (GetKeyState(VK_CONTROL) & 0x8000) {
            switch (wParam) {
            case 'A':
                SetInputSelection(0, g_ui.inputLen);
                ResetCaretBlink(hwnd);
                InvalidateRect(hwnd, NULL, FALSE);
                return 0;
            case 'C':
                CopyInputSelectionToClipboard(hwnd);
                return 0;
            case 'X':
                if (CopyInputSelectionToClipboard(hwnd) && DeleteSelectedInput()) {
                    RefreshInputDisplay(hwnd);
                }
                return 0;
            case 'V':
                if (PasteClipboardText(hwnd)) {
                    RefreshInputDisplay(hwnd);
                    g_ui.historyNavIndex = -1;
                }
                return 0;
            }
        }
        if (wParam == VK_ESCAPE) { ShowWindow(hwnd, SW_HIDE); return 0; }
        if (wParam == VK_TAB) {
            BOOL changed = FALSE;
            if (g_ui.suggestActive) {
                if (ApplySuggestionToInput()) changed = TRUE;
            }
            if (g_ui.ghostParensActive && g_ui.ghostParensLen > 0) {
                if (AppendToInput(g_ui.ghostParens)) changed = TRUE;
                g_ui.ghostParensActive = FALSE; g_ui.ghostParens[0] = 0; g_ui.ghostParensLen = 0;
            }
            // Accept last answer suggestion when input is empty
            if (g_ui.inputLen == 0) {
                if (InsertLastAnswerIfEmpty(hwnd)) changed = TRUE;
            }
            if (changed) {
                RefreshInputDisplay(hwnd);
            }
            return 0; // consume Tab
        }
        if (wParam == VK_UP || wParam == VK_DOWN) {
            // Cycle through history answers
            int start = g_ui.historyNavIndex;
            if (start < 0 || start >= g_historyCount) {
                start = (wParam == VK_UP) ? g_historyCount : -1;
            }
            int step = (wParam == VK_UP) ? -1 : 1;
            int idx = start + step;
            while (idx >= 0 && idx < g_historyCount) {
                HistoryItem *it = g_history[idx];
                if (it && it->answer && *it->answer) {
                    wcsncpy(g_ui.input, it->answer, sizeof(g_ui.input)/sizeof(g_ui.input[0]) - 1);
                    g_ui.input[sizeof(g_ui.input)/sizeof(g_ui.input[0]) - 1] = 0;
                    g_ui.inputLen = (int)wcslen(g_ui.input);
                    g_ui.caretPos = g_ui.inputLen;
                    ClearInputSelection();
                    g_ui.historyNavIndex = idx;
                    RefreshInputDisplay(hwnd);
                    break;
                }
                idx += step;
            }
            return 0;
        }
        if (wParam == VK_LEFT || wParam == VK_RIGHT || wParam == VK_HOME || wParam == VK_END) {
            HandleCaretMovement(hwnd, wParam, (GetKeyState(VK_SHIFT) & 0x8000) != 0);
            return 0;
        }
        if (wParam == VK_DELETE) {
            BOOL isWordOp = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
            HandleTextDeletion(hwnd, TRUE, isWordOp);
            return 0;
        }
        // Ctrl-modified deletion
        if (GetKeyState(VK_CONTROL) & 0x8000) {
            if (wParam == VK_BACK) {
                HandleTextDeletion(hwnd, FALSE, TRUE);
                return 0; // consume Ctrl+Backspace
            }
        }
        return 0;
    case WM_CHAR: {
        // Minimal text handling for commands + backspace, and echo to history
        wchar_t ch = (wchar_t)wParam;
        if ((GetKeyState(VK_CONTROL) & 0x8000) && ch != L'\r' && ch != L'\b') return 0;
        if (ch == L'\r') { // Enter
            if (g_ui.inputLen > 0) {
                // If there is an active suggestion for a command, complete it before handling
                if (g_ui.suggestActive) {
                    ApplySuggestionToInput();
                    UpdateCommandSuggestion();
                }
                // If there are ghost parens, close them before handling
                if (g_ui.ghostParensActive && g_ui.ghostParensLen > 0) {
                    AppendToInput(g_ui.ghostParens);
                    g_ui.ghostParensActive = FALSE; g_ui.ghostParens[0] = 0; g_ui.ghostParensLen = 0;
                    UpdateGhostParens();
                }
                ResetCaretBlink(hwnd);
                if (wcscmp(g_ui.input, L"!exit") == 0 || wcscmp(g_ui.input, L"/exit") == 0) {
                    PostMessage(hwnd, WM_CLOSE, 0, 0);
                } else if (wcscmp(g_ui.input, L"!help") == 0 || wcscmp(g_ui.input, L"/help") == 0) {
                    // Only commands; no shortcuts
                    History_Add(hwnd, L"!help - Shows a list of commands.", L"!help", NULL, TRUE);
                    History_Add(hwnd, L"!settings or /settings - Opens the settings dialog.", L"!settings", NULL, TRUE);
                    History_Add(hwnd, L"!exit - Force closes the calculator.", L"!exit", NULL, TRUE);
                    History_Add(hwnd, L"!clear - Clears the history.", L"!clear", NULL, TRUE);
                    History_Add(hwnd, L"!8ball [Optional question] - Responds to a question with a random answer.", L"!8ball", NULL, TRUE);
                    History_Add(hwnd, L"!roll [number/die: ex. 6, d12, 3d20] [+/- x (modifier)] - Rolls dice.", L"!roll", NULL, TRUE);
                } else if (wcscmp(g_ui.input, L"/settings") == 0 || wcscmp(g_ui.input, L"!settings") == 0) {
                    OpenSettingsDialog(hwnd);
                } else if (wcscmp(g_ui.input, L"!clear") == 0 || wcscmp(g_ui.input, L"/clear") == 0) {
                    History_Clear(hwnd);
                } else if (_wcsnicmp(g_ui.input, L"!8ball", 6) == 0 || _wcsnicmp(g_ui.input, L"/8ball", 6) == 0) {
                    Handle8BallCommand(hwnd, g_ui.input);
                } else if (_wcsnicmp(g_ui.input, L"!roll", 5) == 0 || _wcsnicmp(g_ui.input, L"/roll", 5) == 0) {
                    HandleRollCommand(hwnd, g_ui.input);
                } else {
                    ProcessInputExpression(hwnd);
                }
                // Clear input after handling Enter
                g_ui.inputLen = 0; g_ui.input[0] = 0; g_ui.caretPos = 0; g_ui.suppressLastAnsSuggest = FALSE;
                ClearInputSelection();
                UpdateCommandSuggestion();
                UpdateLivePreview(hwnd);
                InvalidateRect(hwnd, NULL, FALSE);
            }
            return 0;
        }
        if (ch == L'\b') { // Backspace
            // If Ctrl is held, WM_KEYDOWN already handled word-delete
            if (GetKeyState(VK_CONTROL) & 0x8000) return 0;
            if (g_ui.inputLen > 0) {
                HandleTextDeletion(hwnd, FALSE, FALSE);
            } else {
                // No text; hide last-answer suggestion if visible
                if (g_ui.lastAnsSuggestActive || !g_ui.suppressLastAnsSuggest) {
                    g_ui.suppressLastAnsSuggest = TRUE;
                    g_ui.lastAnsSuggestActive = FALSE;
                    g_ui.lastAns[0] = 0; g_ui.lastAnsLen = 0;
                    InvalidateRect(hwnd, NULL, FALSE);
                }
            }
            return 0;
        }
        if (ch >= 0x20 && ch != 0x7F) {
            if (g_ui.inputLen < (int)(sizeof(g_ui.input)/sizeof(g_ui.input[0])) - 1) {
                // If first char is an operator and we have a last-answer suggestion, prefill it
                if (g_ui.inputLen == 0) {
                    if (ch == L'+' || ch == L'-' || ch == L'*' || ch == L'/' || ch == L'×' || ch == L'÷') {
                        // Special-case: after backspace suppression, do NOT prefill for '/'
                        if (!(ch == L'/' && g_ui.suppressLastAnsSuggest)) {
                            InsertLastAnswerIfEmpty(hwnd);
                        }
                    }
                }
                WCHAR tmp[2] = { ch, 0 };
                InsertAtCaret(tmp);
                RefreshInputDisplay(hwnd);
                // Any typing cancels history navigation
                g_ui.historyNavIndex = -1;
            }
            return 0;
        }
        return 0;
    }
    }
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

int APIENTRY wWinMain(HINSTANCE hInst, HINSTANCE hPrev, LPWSTR lpCmdLine, int nCmdShow) {
    (void)hPrev; (void)lpCmdLine; (void)nCmdShow; // mark unused
    g_wmTaskbarCreated = RegisterWindowMessageW(L"TaskbarCreated");

    // Make per-monitor DPI aware if possible
    HMODULE user = GetModuleHandleW(L"user32.dll");
    if (user) {
        FARPROC fp = GetProcAddress(user, "SetProcessDpiAwarenessContext");
        if (fp) {
            union { FARPROC p; SetProcessDpiAwarenessContext_t f; } u;
            u.p = fp;
            SetProcessDpiAwarenessContext_t pSet = u.f;
            if (pSet) pSet((HANDLE)-4 /*DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2*/);
        }
    }

    INITCOMMONCONTROLSEX icc;
    ZeroMemory(&icc, sizeof(icc));
    icc.dwSize = sizeof(icc);
    icc.dwICC = ICC_WIN95_CLASSES;
    InitCommonControlsEx(&icc);

    const wchar_t *clsName = L"CalicoPanel";
    WNDCLASSEXW wc;
    ZeroMemory(&wc, sizeof(wc));
    wc.cbSize = sizeof(wc);
    wc.style = CS_DROPSHADOW | CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInst;
    wc.hIcon = (HICON)LoadImageW(hInst, MAKEINTRESOURCEW(IDI_APP_ICON), IMAGE_ICON,
                                 GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CYICON), 0);
    wc.hIconSm = (HICON)LoadImageW(hInst, MAKEINTRESOURCEW(IDI_APP_ICON), IMAGE_ICON,
                                   GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), 0);
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW+1);
    wc.lpszClassName = clsName;
    RegisterClassExW(&wc);

    // Reference some helpers so compilers won't warn if certain feature paths are unused
    (void)match_temp_unit_word;
    (void)temp_to_kelvin;

    HWND hwnd = CreateWindowExW(
        WS_EX_TOOLWINDOW | WS_EX_TOPMOST | WS_EX_LAYERED,
        clsName, L"Calico",
        WS_POPUP,
        CW_USEDEFAULT, CW_USEDEFAULT, PANEL_WIDTH, PANEL_HEIGHT,
        NULL, NULL, hInst, NULL);

    if (!hwnd) return 0;

    // Initially hidden; the configured hotkey shows it near the cursor.
    ShowWindow(hwnd, SW_HIDE);
    UpdateWindow(hwnd);

    MSG msg;
    while (GetMessageW(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }
    return (int)msg.wParam;
}


