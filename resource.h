#include <windows.h>
#include <shellapi.h>
#include <dbt.h>
#include <stdio.h>

#define ICO1 101
#define ID_TRAY_APP_ICON    1001
#define ID_TRAY_EXIT        1002
#define ID_TRAY_VOID        1003
#define ID_TRAY_STARTUP     1004
#define ID_TRAY_NOTIF_OFF   1005
#define ID_TRAY_NOTIF_BALLOON 1006
#define ID_TRAY_NOTIF_TOAST 1007
#define ID_TRAY_NOTIF_TEST  1008
#define ID_TRAY_DISC_SHOW   1010
#define ID_TRAY_DISC_HIDE   1011
#define ID_TRAY_DISC_AFTER_10   1012
#define ID_TRAY_DISC_AFTER_60   1013
#define ID_TRAY_DISC_AFTER_900  1014
#define ID_TRAY_DISC_AFTER_1800 1015
#define ID_TRAY_DISC_AFTER_3600 1016
#define WM_SYSICON          (WM_USER + 1)
