// HomeOfficeHelper.cpp
// Build (MSVC Developer Prompt):
// rc /nologo /fo HomeOfficeHelper.res HomeOfficeHelper.rc
// cl HomeOfficeHelper.cpp HomeOfficeHelper.res user32.lib shell32.lib advapi32.lib /DUNICODE /D_UNICODE /W4 /link /SUBSYSTEM:WINDOWS

#include <windows.h>
#include <shellapi.h>

#define APP_NAME        L"HomeOfficeHelper"
#define RUN_KEY         L"Software\\Microsoft\\Windows\\CurrentVersion\\Run"

#define WM_TRAYICON     (WM_APP + 1)
#define ID_TRAY_ENABLED         1001
#define ID_TRAY_AUTOSTART       1002
#define ID_TRAY_EXIT            1003
#define ID_TIMER_NUMJIGGLE      2001
#define JIGGLE_INTERVAL_MS      (4 * 60 * 1000) // 4 Minuten

// resource id (aus resource.h), hier zusätzlich, falls nicht inkludiert
#ifndef IDI_APP
#define IDI_APP 101
#endif

HINSTANCE g_hInst = nullptr;
HWND      g_hWnd  = nullptr;
NOTIFYICONDATAW g_nid = {};
BOOL g_enabled   = FALSE;
BOOL g_autostart = FALSE;

static BOOL GetSelfPath(wchar_t* buf, DWORD cch) {
    DWORD len = GetModuleFileNameW(nullptr, buf, cch);
    return (len > 0 && len < cch); // Pfad der laufenden EXE [6]
}

static BOOL IsAutostartEnabled() {
    HKEY hKey;
    if (RegOpenKeyExW(HKEY_CURRENT_USER, RUN_KEY, 0, KEY_READ, &hKey) != ERROR_SUCCESS) return FALSE;
    DWORD type = 0, cb = 0;
    LONG rc = RegQueryValueExW(hKey, APP_NAME, 0, &type, nullptr, &cb);
    RegCloseKey(hKey);
    return (rc == ERROR_SUCCESS && type == REG_SZ && cb > 0); // HKCU\Run [13]
}

static BOOL SetAutostart(BOOL enable) {
    HKEY hKey;
    if (RegCreateKeyExW(HKEY_CURRENT_USER, RUN_KEY, 0, nullptr, 0, KEY_SET_VALUE | KEY_QUERY_VALUE, nullptr, &hKey, nullptr) != ERROR_SUCCESS) {
        return FALSE; // [13]
    }
    BOOL ok = FALSE;
    if (enable) {
        wchar_t path[MAX_PATH];
        if (GetSelfPath(path, MAX_PATH)) {
            // Befehl in Anführungszeichen (Pfad mit Leerzeichen)
            wchar_t cmd[MAX_PATH + 2];
            int n = wsprintfW(cmd, L"\"%s\"", path);
            if (n > 0) {
                ok = (RegSetValueExW(hKey, APP_NAME, 0, REG_SZ, (const BYTE*)cmd,
                                     (DWORD)((lstrlenW(cmd) + 1) * sizeof(wchar_t))) == ERROR_SUCCESS);
            }
        }
    } else {
        ok = (RegDeleteValueW(hKey, APP_NAME) == ERROR_SUCCESS || GetLastError() == ERROR_FILE_NOT_FOUND);
    }
    RegCloseKey(hKey);
    return ok; // [13]
}

static void EnsureTrayIcon() {
    g_nid = {};
    g_nid.cbSize = sizeof(g_nid);
    g_nid.hWnd = g_hWnd;
    g_nid.uID = 1;
    g_nid.uFlags = NIF_MESSAGE | NIF_ICON; // Tooltip weggelassen
    g_nid.uCallbackMessage = WM_TRAYICON;

    // Icon direkt aus eingebetteter Ressource laden (Fix A: ohne LoadIconMetric)
    int cx = GetSystemMetrics(SM_CXSMICON);
    int cy = GetSystemMetrics(SM_CYSMICON);
    HICON hIco = (HICON)LoadImageW(
        g_hInst,
        MAKEINTRESOURCE(IDI_APP),
        IMAGE_ICON,
        cx, cy,
        LR_DEFAULTCOLOR
    ); // [1][9]
    if (!hIco) {
        hIco = LoadIconW(nullptr, IDI_APPLICATION);
    }
    g_nid.hIcon = hIco;

    Shell_NotifyIconW(NIM_ADD, &g_nid); // Tray-Icon hinzufügen [6]
}

static void RemoveTrayIcon() {
    if (g_nid.cbSize) {
        Shell_NotifyIconW(NIM_DELETE, &g_nid); // [6]
        if (g_nid.hIcon) {
            DestroyIcon(g_nid.hIcon);
            g_nid.hIcon = nullptr;
        }
        g_nid = {};
    }
}

static void JiggleNumLock() {
    INPUT inp[15] = {};
    for (int i = 0; i < 4; ++i) {
        inp[i].type = INPUT_KEYBOARD;
        inp[i].ki.wVk = VK_NUMLOCK;
        inp[i].ki.dwFlags = (i % 2 == 1) ? KEYEVENTF_KEYUP : 0;
    }
    SendInput(4, inp, sizeof(INPUT)); // Tastendruck injizieren [9]
}

static void SetEnabled(BOOL on) {
    g_enabled = on;
    if (on) {
        SetTimer(g_hWnd, ID_TIMER_NUMJIGGLE, JIGGLE_INTERVAL_MS, nullptr);
    } else {
        KillTimer(g_hWnd, ID_TIMER_NUMJIGGLE);
    }
}

static void ShowContextMenu(POINT pt) {
    HMENU hMenu = CreatePopupMenu();
    if (!hMenu) return;

    AppendMenuW(hMenu, MF_STRING | (g_enabled ? MF_CHECKED : 0), ID_TRAY_ENABLED,   L"Aktiviert"); // [16]
    AppendMenuW(hMenu, MF_STRING | (g_autostart ? MF_CHECKED : 0), ID_TRAY_AUTOSTART, L"Mit Windows starten");
    AppendMenuW(hMenu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(hMenu, MF_STRING, ID_TRAY_EXIT, L"Beenden");

    SetForegroundWindow(g_hWnd);
    TrackPopupMenu(hMenu, TPM_RIGHTBUTTON | TPM_BOTTOMALIGN | TPM_LEFTALIGN, pt.x, pt.y, 0, g_hWnd, nullptr); // [16]
    DestroyMenu(hMenu);
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_TRAYICON:
        if (lParam == WM_RBUTTONUP || lParam == WM_CONTEXTMENU) {
            POINT pt; GetCursorPos(&pt);
            ShowContextMenu(pt);
        } else if (lParam == WM_LBUTTONDBLCLK) {
            SetEnabled(!g_enabled);
        }
        break;
    case WM_COMMAND:
        switch (LOWORD(wParam)) {
        case ID_TRAY_ENABLED:
            SetEnabled(!g_enabled);
            break;
        case ID_TRAY_AUTOSTART:
            if (SetAutostart(!g_autostart)) {
                g_autostart = IsAutostartEnabled();
            }
            break;
        case ID_TRAY_EXIT:
            PostQuitMessage(0);
            break;
        default:
            break;
        }
        break;
    case WM_TIMER:
        if (wParam == ID_TIMER_NUMJIGGLE && g_enabled) {
            JiggleNumLock();
        }
        break;
    case WM_DESTROY:
        KillTimer(hWnd, ID_TIMER_NUMJIGGLE);
        RemoveTrayIcon();
        PostQuitMessage(0);
        break;
    default:
        return DefWindowProcW(hWnd, msg, wParam, lParam);
    }
    return 0;
}

int APIENTRY wWinMain(HINSTANCE hInstance, HINSTANCE, LPWSTR, int) {
    g_hInst = hInstance;

    WNDCLASSW wc = {};
    wc.lpfnWndProc   = WndProc;
    wc.hInstance     = hInstance;
    wc.lpszClassName = L"HomeOfficeHelperWnd";
    RegisterClassW(&wc); // [6]

    g_hWnd = CreateWindowExW(0, wc.lpszClassName, APP_NAME,
                             WS_OVERLAPPEDWINDOW,
                             CW_USEDEFAULT, CW_USEDEFAULT, 300, 200,
                             nullptr, nullptr, hInstance, nullptr);

    ShowWindow(g_hWnd, SW_HIDE);

    g_autostart = IsAutostartEnabled(); // [13]
    SetEnabled(FALSE);
    EnsureTrayIcon(); // [6]

    MSG msg;
    while (GetMessageW(&msg, nullptr, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }
    return 0;
}