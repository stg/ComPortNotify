#include <windows.h>
#include <shellapi.h>
#include <dbt.h>
#include <stdio.h>

#define ICO1 101
#define ID_TRAY_APP_ICON    1001
#define ID_TRAY_EXIT        1002
#define ID_TRAY_VOID        1003
#define WM_SYSICON          (WM_USER + 1)
