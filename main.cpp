#include <stdarg.h>
#include <stdio.h>
#include <time.h>
#include "resource.h"
#include "serial.h"
#include <shlobj.h>
#include <shobjidl.h>
#include <propkey.h>
#include <propvarutil.h>

// Memory allocating printf
char * mpprintf(const char *fmt, ...) {
    va_list args;
    va_start(args, fmt);
    va_list args2;
    va_copy(args2, args);
	size_t needed = vsnprintf(NULL, 0, fmt, args2) + 1;
    va_end(args2);
    char * buffer = (char *)malloc(needed);
    if(buffer) vsnprintf(buffer, needed, fmt, args);
    va_end(args);
	return buffer;
}

// Variables
HMENU Hmenu;
UINT WM_TASKBAR = 0;
HWND Hwnd;
NOTIFYICONDATA notifyIconData;
TCHAR szTIP[64] = TEXT("No serial port changes detected");
char szClassName[ ] = "Serial port notifier";
bool die = false;

// This GUID is for all USB serial host PnP drivers, but you can replace it with any valid device class guid
GUID WceusbshGUID = {0x25dbce51, 0x6c8f, 0x4a72, 0x8a, 0x6d, 0xb5, 0x4c, 0x2b, 0x4f, 0xc8, 0x35};
					  
// Procedures
LRESULT CALLBACK WindowProcedure (HWND, UINT, WPARAM, LPARAM);
void minimize();
void restore();
void InitNotifyIconData();
bool has_toast_shortcut();
bool create_toast_shortcut();
bool has_startup_shortcut();
bool create_startup_shortcut();
bool delete_startup_shortcut();
extern "C" bool toast_show_winrt(const wchar_t *title, const wchar_t *body);
extern "C" long toast_last_error();
int get_notification_mode();
bool set_notification_mode(int mode);
bool show_notification(const wchar_t *title, const wchar_t *body);
void show_balloon(const char *title, const char *body);
int get_disconnected_mode();
int get_disconnected_timeout();
bool set_disconnected_mode(int mode);
bool set_disconnected_timeout(int seconds);

// Port linked list
typedef struct lport {
	char * device;
	char * name;
	char * hwid;
	struct lport *next;
} lport_t;

// Port history list
typedef struct hport {
	char * device;
	char * name;
	char * hwid;
	time_t connected_at;
	time_t disconnected_at;
	bool connected;
	bool seen;
	struct hport *next;
} hport_t;

static hport_t *find_hport(const char *device);
static void remove_hport(hport_t *prev, hport_t *cur);
static void move_hport_to_head(hport_t *prev, hport_t *cur);

// Toast settings
static const char *SETTINGS_KEY = "Software\\ComPortNotify";
static const char *TOAST_AUMID = "DSp.Tools.CPNotify.1";
static const char *SETTINGS_STARTUP_LINK = "StartupLinkName";
static const char *DEFAULT_STARTUP_LINK = "ComPortNotify.lnk";
static const char *SETTINGS_NOTIF_MODE = "NotificationMode";

enum {
	NOTIF_MODE_OFF = 0,
	NOTIF_MODE_BALLOON = 1,
	NOTIF_MODE_TOAST = 2
};

static bool get_programs_dir(char *buffer, DWORD size) {
	PWSTR wide = NULL;
	if(FAILED(SHGetKnownFolderPath(FOLDERID_Programs, 0, NULL, &wide))) return false;
	int len = WideCharToMultiByte(CP_ACP, 0, wide, -1, buffer, size, NULL, NULL);
	CoTaskMemFree(wide);
	return len > 0;
}

static bool get_startup_dir(char *buffer, DWORD size) {
	PWSTR wide = NULL;
	if(FAILED(SHGetKnownFolderPath(FOLDERID_Startup, 0, NULL, &wide))) return false;
	int len = WideCharToMultiByte(CP_ACP, 0, wide, -1, buffer, size, NULL, NULL);
	CoTaskMemFree(wide);
	return len > 0;
}

static bool get_startup_link_name(char *buffer, DWORD size) {
	HKEY hKey;
	if(RegOpenKeyExA(HKEY_CURRENT_USER, SETTINGS_KEY, 0, KEY_QUERY_VALUE, &hKey) != ERROR_SUCCESS) return false;
	DWORD type = 0;
	DWORD bufSize = size;
	LONG result = RegQueryValueExA(hKey, SETTINGS_STARTUP_LINK, NULL, &type, (LPBYTE)buffer, &bufSize);
	RegCloseKey(hKey);
	return (result == ERROR_SUCCESS && type == REG_SZ && buffer[0]);
}

static void set_startup_link_name(const char *name) {
	HKEY hKey;
	if(RegCreateKeyExA(HKEY_CURRENT_USER, SETTINGS_KEY, 0, NULL, 0, KEY_SET_VALUE, NULL, &hKey, NULL) != ERROR_SUCCESS) {
		return;
	}
	if(name && name[0]) {
		RegSetValueExA(hKey, SETTINGS_STARTUP_LINK, 0, REG_SZ, (const BYTE *)name, (DWORD)(strlen(name) + 1));
	} else {
		RegDeleteValueA(hKey, SETTINGS_STARTUP_LINK);
	}
	RegCloseKey(hKey);
}

int get_notification_mode() {
	HKEY hKey;
	if(RegOpenKeyExA(HKEY_CURRENT_USER, SETTINGS_KEY, 0, KEY_QUERY_VALUE, &hKey) != ERROR_SUCCESS) return NOTIF_MODE_BALLOON;
	DWORD value = 0;
	DWORD size = sizeof(value);
	DWORD type = 0;
	LONG result = RegQueryValueExA(hKey, SETTINGS_NOTIF_MODE, NULL, &type, (LPBYTE)&value, &size);
	RegCloseKey(hKey);
	if(result != ERROR_SUCCESS || type != REG_DWORD) return NOTIF_MODE_BALLOON;
	if(value > NOTIF_MODE_TOAST) return NOTIF_MODE_BALLOON;
	return (int)value;
}

bool set_notification_mode(int mode) {
	HKEY hKey;
	if(RegCreateKeyExA(HKEY_CURRENT_USER, SETTINGS_KEY, 0, NULL, 0, KEY_SET_VALUE, NULL, &hKey, NULL) != ERROR_SUCCESS) {
		return false;
	}
	DWORD value = (DWORD)mode;
	bool ok = (RegSetValueExA(hKey, SETTINGS_NOTIF_MODE, 0, REG_DWORD, (const BYTE *)&value, sizeof(value)) == ERROR_SUCCESS);
	RegCloseKey(hKey);
	return ok;
}

int get_disconnected_mode() {
	HKEY hKey;
	if(RegOpenKeyExA(HKEY_CURRENT_USER, SETTINGS_KEY, 0, KEY_QUERY_VALUE, &hKey) != ERROR_SUCCESS) return 0;
	DWORD value = 0;
	DWORD size = sizeof(value);
	DWORD type = 0;
	LONG result = RegQueryValueExA(hKey, "DisconnectedMode", NULL, &type, (LPBYTE)&value, &size);
	RegCloseKey(hKey);
	if(result != ERROR_SUCCESS || type != REG_DWORD) return 0;
	if(value > 2) return 0;
	return (int)value;
}

int get_disconnected_timeout() {
	HKEY hKey;
	if(RegOpenKeyExA(HKEY_CURRENT_USER, SETTINGS_KEY, 0, KEY_QUERY_VALUE, &hKey) != ERROR_SUCCESS) return 60;
	DWORD value = 60;
	DWORD size = sizeof(value);
	DWORD type = 0;
	LONG result = RegQueryValueExA(hKey, "DisconnectedTimeout", NULL, &type, (LPBYTE)&value, &size);
	RegCloseKey(hKey);
	if(result != ERROR_SUCCESS || type != REG_DWORD) return 60;
	return (int)value;
}

bool set_disconnected_mode(int mode) {
	HKEY hKey;
	if(RegCreateKeyExA(HKEY_CURRENT_USER, SETTINGS_KEY, 0, NULL, 0, KEY_SET_VALUE, NULL, &hKey, NULL) != ERROR_SUCCESS) {
		return false;
	}
	DWORD value = (DWORD)mode;
	bool ok = (RegSetValueExA(hKey, "DisconnectedMode", 0, REG_DWORD, (const BYTE *)&value, sizeof(value)) == ERROR_SUCCESS);
	RegCloseKey(hKey);
	return ok;
}

bool set_disconnected_timeout(int seconds) {
	HKEY hKey;
	if(RegCreateKeyExA(HKEY_CURRENT_USER, SETTINGS_KEY, 0, NULL, 0, KEY_SET_VALUE, NULL, &hKey, NULL) != ERROR_SUCCESS) {
		return false;
	}
	DWORD value = (DWORD)seconds;
	bool ok = (RegSetValueExA(hKey, "DisconnectedTimeout", 0, REG_DWORD, (const BYTE *)&value, sizeof(value)) == ERROR_SUCCESS);
	RegCloseKey(hKey);
	return ok;
}

static bool shortcut_matches_toast(const char *path) {
	bool match = false;
	IShellLinkA *psl = NULL;
	IPersistFile *ppf = NULL;
	IPropertyStore *pps = NULL;
	WCHAR wszPath[MAX_PATH];
	MultiByteToWideChar(CP_ACP, 0, path, -1, wszPath, MAX_PATH);

	if(FAILED(CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, IID_IShellLinkA, (void **)&psl))) return false;
	if(FAILED(psl->QueryInterface(IID_IPersistFile, (void **)&ppf))) {
		psl->Release();
		return false;
	}
	if(FAILED(ppf->Load(wszPath, STGM_READ))) {
		ppf->Release();
		psl->Release();
		return false;
	}
	if(FAILED(psl->QueryInterface(IID_IPropertyStore, (void **)&pps))) {
		ppf->Release();
		psl->Release();
		return false;
	}

	PROPVARIANT var;
	PropVariantInit(&var);
	if(SUCCEEDED(pps->GetValue(PKEY_AppUserModel_ID, &var)) && var.vt == VT_LPWSTR && var.pwszVal) {
		char aumid[256];
		if(WideCharToMultiByte(CP_ACP, 0, var.pwszVal, -1, aumid, sizeof(aumid), NULL, NULL) > 0) {
			if(strcmp(aumid, TOAST_AUMID) == 0) {
				char target[MAX_PATH];
				if(SUCCEEDED(psl->GetPath(target, MAX_PATH, NULL, SLGP_RAWPATH)) && target[0]) {
					DWORD attrs = GetFileAttributesA(target);
					if(attrs != INVALID_FILE_ATTRIBUTES) match = true;
				}
			}
		}
	}
	PropVariantClear(&var);
	pps->Release();
	ppf->Release();
	psl->Release();
	return match;
}

static bool shortcut_targets_exe(const char *path, const char *exePath) {
	bool match = false;
	IShellLinkA *psl = NULL;
	IPersistFile *ppf = NULL;
	WCHAR wszPath[MAX_PATH];
	MultiByteToWideChar(CP_ACP, 0, path, -1, wszPath, MAX_PATH);

	if(FAILED(CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, IID_IShellLinkA, (void **)&psl))) return false;
	if(FAILED(psl->QueryInterface(IID_IPersistFile, (void **)&ppf))) {
		psl->Release();
		return false;
	}
	if(FAILED(ppf->Load(wszPath, STGM_READ))) {
		ppf->Release();
		psl->Release();
		return false;
	}
	char target[MAX_PATH];
	if(SUCCEEDED(psl->GetPath(target, MAX_PATH, NULL, SLGP_RAWPATH)) && target[0]) {
		if(_stricmp(target, exePath) == 0) match = true;
	}
	ppf->Release();
	psl->Release();
	return match;
}

static bool scan_programs_for_toast(const char *dir, const char *startupDir, char *outPath, DWORD outSize) {
	char pattern[MAX_PATH];
	snprintf(pattern, sizeof(pattern), "%s\\*", dir);
	WIN32_FIND_DATAA ffd;
	HANDLE hFind = FindFirstFileA(pattern, &ffd);
	if(hFind == INVALID_HANDLE_VALUE) return false;
	do {
		if(strcmp(ffd.cFileName, ".") == 0 || strcmp(ffd.cFileName, "..") == 0) continue;
		char full[MAX_PATH];
		snprintf(full, sizeof(full), "%s\\%s", dir, ffd.cFileName);
		if(ffd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
		if(startupDir && _stricmp(full, startupDir) == 0) continue;
			if(scan_programs_for_toast(full, startupDir, outPath, outSize)) {
				FindClose(hFind);
				return true;
			}
		} else {
			const char *ext = strrchr(ffd.cFileName, '.');
			if(ext && _stricmp(ext, ".lnk") == 0) {
				if(shortcut_matches_toast(full)) {
					if(outPath && outSize > 0) {
						strncpy(outPath, full, outSize);
						outPath[outSize - 1] = '\0';
					}
					FindClose(hFind);
					return true;
				}
			}
		}
	} while(FindNextFileA(hFind, &ffd));
	FindClose(hFind);
	return false;
}

static bool scan_startup_for_exe(const char *dir, const char *exePath, char *outPath, DWORD outSize) {
	char pattern[MAX_PATH];
	snprintf(pattern, sizeof(pattern), "%s\\*", dir);
	WIN32_FIND_DATAA ffd;
	HANDLE hFind = FindFirstFileA(pattern, &ffd);
	if(hFind == INVALID_HANDLE_VALUE) return false;
	do {
		if(strcmp(ffd.cFileName, ".") == 0 || strcmp(ffd.cFileName, "..") == 0) continue;
		char full[MAX_PATH];
		snprintf(full, sizeof(full), "%s\\%s", dir, ffd.cFileName);
		if(ffd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
			if(scan_startup_for_exe(full, exePath, outPath, outSize)) {
				FindClose(hFind);
				return true;
			}
		} else {
			const char *ext = strrchr(ffd.cFileName, '.');
			if(ext && _stricmp(ext, ".lnk") == 0) {
				if(shortcut_targets_exe(full, exePath)) {
					if(outPath && outSize > 0) {
						strncpy(outPath, full, outSize);
						outPath[outSize - 1] = '\0';
					}
					FindClose(hFind);
					return true;
				}
			}
		}
	} while(FindNextFileA(hFind, &ffd));
	FindClose(hFind);
	return false;
}

bool has_toast_shortcut() {
	HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
	char programs[MAX_PATH];
	char startup[MAX_PATH];
	startup[0] = '\0';
	bool found = false;
	if(get_programs_dir(programs, sizeof(programs))) {
		get_startup_dir(startup, sizeof(startup));
		found = scan_programs_for_toast(programs, startup[0] ? startup : NULL, NULL, 0);
	}
	if(SUCCEEDED(hr)) CoUninitialize();
	return found;
}

bool create_toast_shortcut() {
	HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
	char programs[MAX_PATH];
	if(!get_programs_dir(programs, sizeof(programs))) {
		if(SUCCEEDED(hr)) CoUninitialize();
		return false;
	}

	char lnkPath[MAX_PATH];
	snprintf(lnkPath, sizeof(lnkPath), "%s\\ComPortNotify.lnk", programs);

	char exePath[MAX_PATH];
	if(!GetModuleFileNameA(NULL, exePath, MAX_PATH)) {
		if(SUCCEEDED(hr)) CoUninitialize();
		return false;
	}

	IShellLinkA *psl = NULL;
	IPersistFile *ppf = NULL;
	IPropertyStore *pps = NULL;
	if(FAILED(CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, IID_IShellLinkA, (void **)&psl))) {
		if(SUCCEEDED(hr)) CoUninitialize();
		return false;
	}
	psl->SetPath(exePath);

	if(FAILED(psl->QueryInterface(IID_IPropertyStore, (void **)&pps))) {
		psl->Release();
		if(SUCCEEDED(hr)) CoUninitialize();
		return false;
	}
	PROPVARIANT var;
	PropVariantInit(&var);
	WCHAR aumidW[256];
	MultiByteToWideChar(CP_ACP, 0, TOAST_AUMID, -1, aumidW, 256);
	InitPropVariantFromString(aumidW, &var);
	pps->SetValue(PKEY_AppUserModel_ID, var);
	pps->Commit();
	PropVariantClear(&var);
	pps->Release();

	if(FAILED(psl->QueryInterface(IID_IPersistFile, (void **)&ppf))) {
		psl->Release();
		if(SUCCEEDED(hr)) CoUninitialize();
		return false;
	}
	WCHAR wszPath[MAX_PATH];
	MultiByteToWideChar(CP_ACP, 0, lnkPath, -1, wszPath, MAX_PATH);
	HRESULT saveHr = ppf->Save(wszPath, TRUE);
	ppf->Release();
	psl->Release();
	if(SUCCEEDED(hr)) CoUninitialize();
	return SUCCEEDED(saveHr);
}

bool has_startup_shortcut() {
	HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
	char startup[MAX_PATH];
	char exePath[MAX_PATH];
	bool found = false;
	if(get_startup_dir(startup, sizeof(startup)) && GetModuleFileNameA(NULL, exePath, MAX_PATH)) {
		char foundPath[MAX_PATH];
		found = scan_startup_for_exe(startup, exePath, foundPath, sizeof(foundPath));
		if(found) {
			const char *name = strrchr(foundPath, '\\');
			set_startup_link_name(name ? name + 1 : foundPath);
		}
	}
	if(SUCCEEDED(hr)) CoUninitialize();
	return found;
}

bool create_startup_shortcut() {
	HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
	char startup[MAX_PATH];
	char exePath[MAX_PATH];
	char linkName[MAX_PATH];
	if(!get_startup_dir(startup, sizeof(startup)) || !GetModuleFileNameA(NULL, exePath, MAX_PATH)) {
		if(SUCCEEDED(hr)) CoUninitialize();
		return false;
	}
	if(!get_startup_link_name(linkName, sizeof(linkName))) {
		strncpy(linkName, DEFAULT_STARTUP_LINK, sizeof(linkName));
		linkName[sizeof(linkName) - 1] = '\0';
	}
	char lnkPath[MAX_PATH];
	snprintf(lnkPath, sizeof(lnkPath), "%s\\%s", startup, linkName);

	IShellLinkA *psl = NULL;
	IPersistFile *ppf = NULL;
	if(FAILED(CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, IID_IShellLinkA, (void **)&psl))) {
		if(SUCCEEDED(hr)) CoUninitialize();
		return false;
	}
	psl->SetPath(exePath);
	if(FAILED(psl->QueryInterface(IID_IPersistFile, (void **)&ppf))) {
		psl->Release();
		if(SUCCEEDED(hr)) CoUninitialize();
		return false;
	}
	WCHAR wszPath[MAX_PATH];
	MultiByteToWideChar(CP_ACP, 0, lnkPath, -1, wszPath, MAX_PATH);
	HRESULT saveHr = ppf->Save(wszPath, TRUE);
	ppf->Release();
	psl->Release();
	if(SUCCEEDED(saveHr)) set_startup_link_name(linkName);
	if(SUCCEEDED(hr)) CoUninitialize();
	return SUCCEEDED(saveHr);
}

bool delete_startup_shortcut() {
	char startup[MAX_PATH];
	char linkName[MAX_PATH];
	if(!get_startup_dir(startup, sizeof(startup))) return false;
	if(!get_startup_link_name(linkName, sizeof(linkName))) return false;
	char lnkPath[MAX_PATH];
	snprintf(lnkPath, sizeof(lnkPath), "%s\\%s", startup, linkName);
	if(DeleteFileA(lnkPath)) {
		set_startup_link_name(NULL);
		return true;
	}
	return false;
}

// Free ports linked list
void free_lports(lport_t *p) {
	while(p) {
		free(p->device);
		free(p->name);
		free(p->hwid);
		lport_t *_p = p;
		p = p->next;
		free(_p);
	}
}

// Previous and current port list, for change detection
lport_t * volatile _ports;
hport_t *history;

// Add new port to linked list
void add_lport(char *name, char *device, char *hwid) {
	lport_t * temp = (lport_t *)malloc(sizeof(lport_t));
	char *dev = (char *)malloc(strlen(device) + 1);
	char *nam = (char *)malloc(strlen(name) + 1);
	char *hid = hwid ? _strdup(hwid) : NULL;
	if(!temp || !dev || !nam || (hwid && !hid)) {
		if(hid) free(hid);
		if(nam) free(nam);
		if(dev) free(dev);
		if(temp) free(temp);
		return;
	}
	ZeroMemory(temp, sizeof(lport_t));
	strcpy(dev, device);
	strcpy(nam, name);
	temp->device = dev;
	temp->name = nam;
	temp->hwid = hid;
	temp->next = _ports;
	_ports = temp;
}

// Refresh the ports list
void refresh_ports(bool init = false) {
	// TODO: List should note time of new connections
	// TODO: Should show recent history and times on popup menu (disabled/grayed for disconnected ports, with timeout?)
	char szTooltip[sizeof(notifyIconData.szTip)] = {0};
	senum(add_lport); // List current ports to _ports
	{
		time_t now = time(NULL);
		hport_t *hp = history;
		while(hp) {
			hp->seen = false;
			hp = hp->next;
		}

		lport_t *a = _ports;
		while(a) {
		hport_t *found = find_hport(a->device);
		if(found) {
			found->seen = true;
			if(strcmp(found->name, a->name) != 0) {
				char *new_name = _strdup(a->name);
				if(new_name) {
					free(found->name);
					found->name = new_name;
				}
			}
			if((found->hwid && a->hwid && strcmp(found->hwid, a->hwid) != 0) || (!found->hwid && a->hwid)) {
				char *new_hwid = a->hwid ? _strdup(a->hwid) : NULL;
				if(a->hwid == NULL || new_hwid) {
					if(found->hwid) free(found->hwid);
					found->hwid = new_hwid;
				}
			}
			if(!found->connected) {
					found->connected = true;
					found->connected_at = now;
					found->disconnected_at = 0;
					if(found != history) {
						hport_t *prev = history;
						while(prev && prev->next != found) prev = prev->next;
						move_hport_to_head(prev, found);
					}
					if(!init) {
						char * text = mpprintf("Connected %s %s\n", found->device, found->name);
						if(text) {
							if(!szTooltip[0]) {
								strncpy(szTooltip, text, sizeof(szTooltip));
								szTooltip[sizeof(szTooltip) - 1] = '\0';
							}
							wchar_t wtext[512];
							MultiByteToWideChar(CP_ACP, 0, text, -1, wtext, 512);
							show_notification(L"ComPortNotify", wtext);
							free(text);
						}
					}
				}
			} else {
			hport_t *n = (hport_t *)malloc(sizeof(hport_t));
			if(n) {
				ZeroMemory(n, sizeof(hport_t));
				n->device = _strdup(a->device);
				n->name = _strdup(a->name);
				n->hwid = a->hwid ? _strdup(a->hwid) : NULL;
				if(n->device && n->name && (!a->hwid || n->hwid)) {
					n->connected = true;
					n->connected_at = init ? 0 : now;
					n->disconnected_at = 0;
					n->seen = true;
					n->next = history;
					history = n;
						if(!init) {
							char * text = mpprintf("Connected %s %s\n", n->device, n->name);
							if(text) {
								if(!szTooltip[0]) {
									strncpy(szTooltip, text, sizeof(szTooltip));
									szTooltip[sizeof(szTooltip) - 1] = '\0';
								}
								wchar_t wtext[512];
								MultiByteToWideChar(CP_ACP, 0, text, -1, wtext, 512);
								show_notification(L"ComPortNotify", wtext);
								free(text);
							}
						}
				} else {
					if(n->device) free(n->device);
					if(n->name) free(n->name);
					if(n->hwid) free(n->hwid);
					free(n);
				}
				}
			}
			a = a->next;
		}

		hport_t *prev = NULL;
		hp = history;
		while(hp) {
			if(hp->connected && !hp->seen) {
				hp->connected = false;
				hp->disconnected_at = now;
				if(!init) {
					if(hp != history) {
						hport_t *pp = history;
						while(pp && pp->next != hp) pp = pp->next;
						move_hport_to_head(pp, hp);
					}
					char * text = mpprintf("Removed %s %s\n", hp->device, hp->name);
					if(text) {
						if(!szTooltip[0]) {
							strncpy(szTooltip, text, sizeof(szTooltip));
							szTooltip[sizeof(szTooltip) - 1] = '\0';
						}
						wchar_t wtext[512];
						MultiByteToWideChar(CP_ACP, 0, text, -1, wtext, 512);
						show_notification(L"ComPortNotify", wtext);
						free(text);
					}
				}
			}

			int dmode = get_disconnected_mode();
			int dtimeout = get_disconnected_timeout();
			bool remove_now = false;
			if(!hp->connected) {
				if(dmode == 1) {
					remove_now = true;
				} else if(dmode == 2) {
					if(now - hp->disconnected_at >= dtimeout) remove_now = true;
				}
			}

			if(remove_now) {
				hport_t *next = hp->next;
				remove_hport(prev, hp);
				hp = next;
				continue;
			}

			prev = hp;
			hp = hp->next;
		}

		if(!init && szTooltip[0]) {
			strncpy(notifyIconData.szTip, szTooltip, sizeof(notifyIconData.szTip));
			notifyIconData.szTip[sizeof(notifyIconData.szTip) - 1] = '\0';
			Shell_NotifyIcon(NIM_MODIFY, &notifyIconData);
		}
	}

	free_lports(_ports);
	_ports = NULL;

}

// Populate the popup menu
typedef struct menu_text {
	char *prefix;
	char *desc;
	char *right;
	bool has_submenu;
	bool grayed;
	struct menu_text *next;
} menu_text_t;

typedef struct menu_clip {
	UINT id;
	char *text;
	struct menu_clip *next;
} menu_clip_t;

static HFONT g_menu_font = NULL;
static bool g_menu_nocheck = false;
static int g_menu_prefix_width = 0;
static int g_menu_desc_width = 0;

static HFONT get_menu_font() {
	if(!g_menu_font) {
		NONCLIENTMETRICS ncm;
		ZeroMemory(&ncm, sizeof(ncm));
		ncm.cbSize = sizeof(ncm);
		if(SystemParametersInfo(SPI_GETNONCLIENTMETRICS, sizeof(ncm), &ncm, 0)) {
			g_menu_font = CreateFontIndirect(&ncm.lfMenuFont);
		}
	}
	return g_menu_font;
}

static void reset_menu_font() {
	if(g_menu_font) {
		DeleteObject(g_menu_font);
		g_menu_font = NULL;
	}
}

static void free_menu_texts(menu_text_t *p) {
	while(p) {
		menu_text_t *n = p->next;
		free(p->prefix);
		free(p->desc);
		free(p->right);
		free(p);
		p = n;
	}
}

static void free_menu_clips(menu_clip_t *p) {
	while(p) {
		menu_clip_t *n = p->next;
		free(p->text);
		free(p);
		p = n;
	}
}

static bool copy_to_clipboard(const char *text) {
	if(!text) return false;
	if(!OpenClipboard(Hwnd)) return false;
	EmptyClipboard();
	size_t len = strlen(text) + 1;
	HGLOBAL hmem = GlobalAlloc(GMEM_MOVEABLE, len);
	if(!hmem) {
		CloseClipboard();
		return false;
	}
	void *dst = GlobalLock(hmem);
	if(!dst) {
		GlobalFree(hmem);
		CloseClipboard();
		return false;
	}
	memcpy(dst, text, len);
	GlobalUnlock(hmem);
	SetClipboardData(CF_TEXT, hmem);
	CloseClipboard();
	return true;
}

static char *format_time_label(time_t now, time_t t, bool just_now_allowed) {
	if(t == 0) return _strdup("Startup");
	if(t > now) t = now;

	int age = (int)difftime(now, t);
	if(age < 0) age = 0;

	if(just_now_allowed && age < 30) {
		return _strdup("Just now");
	}
	if(!just_now_allowed && age < 60) {
		char buf[32];
		snprintf(buf, sizeof(buf), "%ds", age);
		return _strdup(buf);
	}
	if(age < 60) {
		return _strdup("1 minute ago");
	}
	if(age < 3600) {
		int minutes = age / 60;
		char buf[64];
		snprintf(buf, sizeof(buf), "%d minute%s ago", minutes, minutes == 1 ? "" : "s");
		return _strdup(buf);
	}

	struct tm tm_now;
	struct tm tm_t;
	if(localtime_s(&tm_now, &now) != 0 || localtime_s(&tm_t, &t) != 0) return _strdup("");

	if(tm_now.tm_year == tm_t.tm_year && tm_now.tm_yday == tm_t.tm_yday) {
		SYSTEMTIME st = {0};
		st.wYear = (WORD)(tm_t.tm_year + 1900);
		st.wMonth = (WORD)(tm_t.tm_mon + 1);
		st.wDay = (WORD)tm_t.tm_mday;
		st.wHour = (WORD)tm_t.tm_hour;
		st.wMinute = (WORD)tm_t.tm_min;
		st.wSecond = (WORD)tm_t.tm_sec;
		char buf[64];
		if(GetTimeFormatA(LOCALE_USER_DEFAULT, 0, &st, NULL, buf, (int)sizeof(buf))) {
			return _strdup(buf);
		}
		return _strdup("");
	}

	if(tm_now.tm_year == tm_t.tm_year && tm_now.tm_yday == tm_t.tm_yday + 1) {
		return _strdup("Yesterday");
	}

	SYSTEMTIME st = {0};
	st.wYear = (WORD)(tm_t.tm_year + 1900);
	st.wMonth = (WORD)(tm_t.tm_mon + 1);
	st.wDay = (WORD)tm_t.tm_mday;
	char buf[64];
	if(GetDateFormatA(LOCALE_USER_DEFAULT, DATE_SHORTDATE, &st, NULL, buf, (int)sizeof(buf))) {
		return _strdup(buf);
	}
	return _strdup("");
}

void populate_menu(menu_text_t **allocs, menu_clip_t **clips, UINT *next_id) {
	hport_t * p = history;
	bool any = false;
	int dmode = get_disconnected_mode();
	time_t now = time(NULL);
	int just_now_count = 0;
	hport_t *scan = history;
	while(scan) {
		if(!scan->connected && dmode == 1) {
			scan = scan->next;
			continue;
		}
		time_t t = scan->connected ? scan->connected_at : scan->disconnected_at;
		if(t > 0 && difftime(now, t) >= 0 && difftime(now, t) < 30) {
			just_now_count++;
		}
		scan = scan->next;
	}
	bool just_now_allowed = (just_now_count <= 1);
	while(p) {
		if(!p->connected && dmode == 1) {
			p = p->next;
			continue;
		}
		any = true;
		char * prefix = NULL;
		char * desc = NULL;
		char * right = NULL;
		time_t t = p->connected ? p->connected_at : p->disconnected_at;
		prefix = _strdup(p->device);
		desc = _strdup(p->name);
		right = format_time_label(now, t, just_now_allowed);
		if(prefix && desc) {
			UINT flags = MF_OWNERDRAW | MF_POPUP;
			if(!p->connected) flags |= MF_GRAYED;
			menu_text_t *mt = (menu_text_t *)malloc(sizeof(menu_text_t));
			if(mt) {
				mt->prefix = prefix;
				mt->desc = desc;
				mt->right = right ? right : _strdup("");
				mt->has_submenu = true;
				mt->grayed = !p->connected;
				mt->next = *allocs;
				*allocs = mt;
				HMENU sub = CreatePopupMenu();
				if(p->hwid && p->hwid[0]) {
					UINT id = (*next_id)++;
					AppendMenuA(sub, MF_STRING, id, p->hwid);
					menu_clip_t *mc = (menu_clip_t *)malloc(sizeof(menu_clip_t));
					if(mc) {
						mc->id = id;
						mc->text = _strdup(p->hwid);
						mc->next = *clips;
						*clips = mc;
					}
				}
				if(p->device && p->device[0]) {
					UINT id = (*next_id)++;
					AppendMenuA(sub, MF_STRING, id, p->device);
					menu_clip_t *mc = (menu_clip_t *)malloc(sizeof(menu_clip_t));
					if(mc) {
						mc->id = id;
						mc->text = _strdup(p->device);
						mc->next = *clips;
						*clips = mc;
					}
				}
				if(!p->hwid || !p->hwid[0]) {
					// If no hardware ID, ensure submenu isn't empty.
					// (COM item above will typically exist.)
				}
				MENUITEMINFOA mii;
				ZeroMemory(&mii, sizeof(mii));
				mii.cbSize = sizeof(mii);
				mii.fMask = MIIM_FTYPE | MIIM_SUBMENU | MIIM_DATA | MIIM_STATE;
				mii.fType = MFT_OWNERDRAW;
				mii.fState = MFS_ENABLED;
				mii.dwItemData = (ULONG_PTR)mt;
				mii.hSubMenu = sub;
				InsertMenuItemA(Hmenu, (UINT)-1, TRUE, &mii);
			} else {
				free(prefix);
				free(desc);
				if(right) free(right);
			}
		}
		p = p->next;
	}
	if(!any) {
		menu_text_t *mt = (menu_text_t *)malloc(sizeof(menu_text_t));
		if(mt) {
			mt->prefix = _strdup("");
			mt->desc = _strdup("No serial ports detected");
			mt->right = _strdup("");
			mt->has_submenu = false;
			mt->grayed = true;
			mt->next = *allocs;
			*allocs = mt;
			MENUITEMINFOA mii;
			ZeroMemory(&mii, sizeof(mii));
			mii.cbSize = sizeof(mii);
			mii.fMask = MIIM_FTYPE | MIIM_DATA | MIIM_STATE;
			mii.fType = MFT_OWNERDRAW;
			mii.fState = MFS_DISABLED;
			mii.dwItemData = (ULONG_PTR)mt;
			InsertMenuItemA(Hmenu, (UINT)-1, TRUE, &mii);
		}
	}
}

// Application entry point
int WINAPI WinMain(HINSTANCE hThisInstance, HINSTANCE hPrevInstance, LPSTR lpszArgument, int nCmdShow) {
    MSG messages;            // Messages to the application are saved here
    WNDCLASSEX wincl;        // Data structure for the windowclass
    WM_TASKBAR = RegisterWindowMessageA("TaskbarCreated");
    
	// The Window structure
    wincl.hInstance = hThisInstance;
    wincl.lpszClassName = szClassName;
    wincl.lpfnWndProc = WindowProcedure;      // This function is called by windows
    wincl.style = CS_DBLCLKS;                 // Catch double-clicks
    wincl.cbSize = sizeof (WNDCLASSEX);

    // Use default icon and mouse-pointer
    wincl.hIcon = LoadIcon (GetModuleHandle(NULL), MAKEINTRESOURCE(ICO1));
    wincl.hIconSm = LoadIcon (GetModuleHandle(NULL), MAKEINTRESOURCE(ICO1));
    wincl.hCursor = LoadCursor (NULL, IDC_ARROW);
    wincl.lpszMenuName = NULL;                 // No menu
    wincl.cbClsExtra = 0;                      // No extra bytes after the window class
    wincl.cbWndExtra = 0;                      // structure or the window instance
    wincl.hbrBackground = (HBRUSH)(CreateSolidBrush(RGB(255, 255, 255)));

	// Register the window class, and if it fails quit the program
    if(!RegisterClassEx(&wincl)) return 0;

    // The class is registered, let's create the program
	// TODO: make message only window
    Hwnd = CreateWindowEx (
		0,                   // Extended possibilites for variation
		szClassName,         // Classname
		szClassName,         // Title Text
		WS_OVERLAPPEDWINDOW, // default window
		CW_USEDEFAULT,       // Windows decides the position
		CW_USEDEFAULT,       // where the window ends up on the screen
		544,                 // The programs width
		375,                 // and height in pixels
		HWND_DESKTOP,        // The window is a child-window to desktop
		//0,
		//0,
		//0,
		//0,
		//0,
		//HWND_MESSAGE,        // The window is a message-only window
		NULL,                // No menu
		hThisInstance,       // Program Instance handler
		NULL                 // No Window Creation data
    );
	
	 //CreateWindowEx( 0, class_name, "dummy_name", 0, 0, 0, 0, 0, HWND_MESSAGE, NULL, NULL, NULL );

	// Initialize the NOTIFYICONDATA structure only once
    InitNotifyIconData();

	// Sign up for device change notifications
	DEV_BROADCAST_DEVICEINTERFACE NotificationFilter;
    ZeroMemory( &NotificationFilter, sizeof(NotificationFilter) );
    NotificationFilter.dbcc_size = sizeof(DEV_BROADCAST_DEVICEINTERFACE);
    NotificationFilter.dbcc_devicetype = DBT_DEVTYP_DEVICEINTERFACE;
    NotificationFilter.dbcc_classguid = WceusbshGUID;
	HDEVNOTIFY hVolNotify = RegisterDeviceNotification(Hwnd, &NotificationFilter, DEVICE_NOTIFY_WINDOW_HANDLE);
	if(!hVolNotify) {
		MessageBox(Hwnd, TEXT("Failed to register for device notifications."), TEXT("ComPortNotify"), MB_OK | MB_ICONERROR);
		PostQuitMessage(1);
		die = true;
	}
    
	// Make the window visible on the screen
    //ShowWindow(Hwnd, SW_MINIMIZE);
	
	// Create the system tray icon
	Shell_NotifyIcon(NIM_ADD, &notifyIconData);
    
	// Initialize port list
	refresh_ports(true);
	
    // Message loop
    while(!die) {
		int result = GetMessage(&messages, NULL, 0, 0);
		if(result > 0) {
			// Translate virtual-key messages into character messages
			TranslateMessage(&messages);
			// Send message to WindowProcedure
			DispatchMessage(&messages);
		} else if(result < 0) {
			break;
		}
    }

    return messages.wParam;
}

//  This function is called by the Windows function DispatchMessage()
LRESULT CALLBACK WindowProcedure(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) {
    UINT temp;
	
	if (message == WM_TASKBAR && !IsWindowVisible(Hwnd)) {
        minimize();
        return 0;
    }

	// Handle the messages
	switch (message) {
		case WM_ACTIVATE:
			//Shell_NotifyIcon(NIM_ADD, &notifyIconData);
			//ShowWindow(Hwnd, SW_HIDE);
			break;

		case WM_CREATE:
			break;

		case WM_SYSCOMMAND:
			// System command
			switch(wParam & 0xFFF0) {
				case SC_MINIMIZE:
				case SC_CLOSE:  
					minimize() ;
					return 0;
					break;
			}
			break;

		case WM_DEVICECHANGE: {
			// Device list has changed
			PDEV_BROADCAST_DEVICEINTERFACE b = (PDEV_BROADCAST_DEVICEINTERFACE) lParam;

			// Output some messages to the window
			switch (wParam) {
				case DBT_DEVICEARRIVAL:
					//printf("[info] DBT_DEVICEARRIVAL\n");
					break;
				case DBT_DEVICEREMOVECOMPLETE:
					//printf("[info] DBT_DEVICEREMOVECOMPLETE\n");
					break;
				case DBT_DEVNODES_CHANGED:
					//printf("[info] DBT_DEVNODES_CHANGED\n");
					refresh_ports();
					break;
				default:
					//printf("[info] WM_DEVICECHANGE %d received\n", wParam);
					break;
			}
		} break;
			
		// Our user defined WM_SYSICON message
		case WM_SYSICON: 
			switch(wParam) {
				case ID_TRAY_APP_ICON:
					SetForegroundWindow(Hwnd);
					break;
			}
			if(lParam == WM_LBUTTONUP) {
				restore();
			} else if(lParam == WM_RBUTTONDOWN) {
				// Get current mouse position
				POINT curPoint ;
				GetCursorPos( &curPoint ) ;
				SetForegroundWindow(Hwnd);

				// Create menu
				Hmenu = CreatePopupMenu();
				MENUINFO mi;
				ZeroMemory(&mi, sizeof(mi));
				mi.cbSize = sizeof(mi);
				mi.fMask = MIM_STYLE;
				mi.dwStyle = MNS_NOCHECK;
				SetMenuInfo(Hmenu, &mi);
				g_menu_nocheck = true;
				menu_text_t *menu_texts = NULL;
				menu_clip_t *menu_clips = NULL;
				UINT next_id = 2000;
	g_menu_prefix_width = 0;
	g_menu_desc_width = 0;
	populate_menu(&menu_texts, &menu_clips, &next_id);
				if(menu_texts) {
					HDC hdc = GetDC(Hwnd);
					HFONT hfont = get_menu_font();
					HFONT old = hfont ? (HFONT)SelectObject(hdc, hfont) : NULL;
					menu_text_t *mt = menu_texts;
					while(mt) {
						if(mt->prefix && mt->prefix[0]) {
							SIZE sz = {0};
							GetTextExtentPoint32A(hdc, mt->prefix, (int)strlen(mt->prefix), &sz);
							if((int)sz.cx > g_menu_prefix_width) g_menu_prefix_width = (int)sz.cx;
						}
						if(mt->desc && mt->desc[0]) {
							SIZE sz = {0};
							GetTextExtentPoint32A(hdc, mt->desc, (int)strlen(mt->desc), &sz);
							if((int)sz.cx > g_menu_desc_width) g_menu_desc_width = (int)sz.cx;
						}
						mt = mt->next;
					}
					if(old) SelectObject(hdc, old);
					ReleaseDC(Hwnd, hdc);
				}
				HMENU Hsettings = CreatePopupMenu();
				UINT startupChecked = has_startup_shortcut() ? MF_CHECKED : MF_UNCHECKED;
				AppendMenu(Hsettings, MF_STRING | startupChecked, ID_TRAY_STARTUP, TEXT("Start with Windows"));
				HMENU Hnotify = CreatePopupMenu();
				int mode = get_notification_mode();
				UINT offFlag = (mode == NOTIF_MODE_OFF) ? MF_CHECKED : 0;
				UINT balloonFlag = (mode == NOTIF_MODE_BALLOON) ? MF_CHECKED : 0;
				UINT toastFlag = (mode == NOTIF_MODE_TOAST) ? MF_CHECKED : 0;
				AppendMenu(Hnotify, MF_STRING | offFlag, ID_TRAY_NOTIF_OFF, TEXT("Off"));
				AppendMenu(Hnotify, MF_STRING | balloonFlag, ID_TRAY_NOTIF_BALLOON, TEXT("Balloon"));
				AppendMenu(Hnotify, MF_STRING | toastFlag, ID_TRAY_NOTIF_TOAST, TEXT("Toast"));
				AppendMenu(Hnotify, MF_SEPARATOR, 0, NULL );
				UINT testFlags = (mode == NOTIF_MODE_OFF) ? MF_GRAYED : MF_ENABLED;
				AppendMenu(Hnotify, MF_STRING | testFlags, ID_TRAY_NOTIF_TEST, TEXT("Notification test"));
				AppendMenu(Hsettings, MF_POPUP, (UINT_PTR)Hnotify, TEXT("Notification"));
				HMENU Hdisc = CreatePopupMenu();
				int dmode = get_disconnected_mode();
				UINT showFlag = (dmode == 0) ? MF_CHECKED : 0;
				UINT hideFlag = (dmode == 1) ? MF_CHECKED : 0;
				UINT afterFlag = (dmode == 2) ? MF_CHECKED : 0;
				AppendMenu(Hdisc, MF_STRING | showFlag, ID_TRAY_DISC_SHOW, TEXT("Show"));
				AppendMenu(Hdisc, MF_STRING | hideFlag, ID_TRAY_DISC_HIDE, TEXT("Hide"));
				HMENU Hafter = CreatePopupMenu();
				int dt = get_disconnected_timeout();
				AppendMenu(Hafter, MF_STRING | ((dmode==2 && dt==10)?MF_CHECKED:0), ID_TRAY_DISC_AFTER_10, TEXT("10 seconds"));
				AppendMenu(Hafter, MF_STRING | ((dmode==2 && dt==60)?MF_CHECKED:0), ID_TRAY_DISC_AFTER_60, TEXT("1 minute"));
				AppendMenu(Hafter, MF_STRING | ((dmode==2 && dt==900)?MF_CHECKED:0), ID_TRAY_DISC_AFTER_900, TEXT("15 minutes"));
				AppendMenu(Hafter, MF_STRING | ((dmode==2 && dt==1800)?MF_CHECKED:0), ID_TRAY_DISC_AFTER_1800, TEXT("30 minutes"));
				AppendMenu(Hafter, MF_STRING | ((dmode==2 && dt==3600)?MF_CHECKED:0), ID_TRAY_DISC_AFTER_3600, TEXT("1 hour"));
				AppendMenu(Hdisc, MF_POPUP | afterFlag, (UINT_PTR)Hafter, TEXT("Hide after"));
				AppendMenu(Hsettings, MF_POPUP, (UINT_PTR)Hdisc, TEXT("Disconnected ports"));
				AppendMenu(Hmenu, MF_SEPARATOR, 0, NULL );
				AppendMenu(Hmenu, MF_POPUP, (UINT_PTR)Hsettings, TEXT("Settings"));
				AppendMenu(Hmenu, MF_SEPARATOR, 0, NULL );
				AppendMenu(Hmenu, MF_STRING, ID_TRAY_EXIT,  TEXT( "Exit" ) );

				// TrackPopupMenu blocks the app until TrackPopupMenu returns
				temp = TrackPopupMenu(Hmenu, TPM_RETURNCMD | TPM_NONOTIFY, curPoint.x, curPoint.y, 0, hwnd, NULL);

				DestroyMenu(Hmenu);
			
				SendMessage(hwnd, WM_NULL, 0, 0); // Send benign message to window to make sure the menu goes away
				if(temp == ID_TRAY_EXIT) {
					// Quit the application.
					Shell_NotifyIcon(NIM_DELETE, &notifyIconData);
					PostQuitMessage( 0 ) ;
					die = true;
				} else if(temp >= 2000) {
					menu_clip_t *mc = menu_clips;
					while(mc) {
						if(mc->id == (UINT)temp) {
							copy_to_clipboard(mc->text);
							break;
						}
						mc = mc->next;
					}
				} else if(temp == ID_TRAY_STARTUP) {
					bool checked = has_startup_shortcut();
					if(!checked) {
						create_startup_shortcut();
					} else {
						delete_startup_shortcut();
					}
				} else if(temp == ID_TRAY_NOTIF_OFF) {
					set_notification_mode(NOTIF_MODE_OFF);
				} else if(temp == ID_TRAY_NOTIF_BALLOON) {
					set_notification_mode(NOTIF_MODE_BALLOON);
				} else if(temp == ID_TRAY_NOTIF_TOAST) {
					if(!has_toast_shortcut()) {
						if(!create_toast_shortcut()) {
							MessageBox(Hwnd, TEXT("Failed to create Start Menu shortcut for toasts."), TEXT("ComPortNotify"), MB_OK | MB_ICONERROR);
							break;
						}
					}
					set_notification_mode(NOTIF_MODE_TOAST);
				} else if(temp == ID_TRAY_NOTIF_TEST) {
					int m = get_notification_mode();
					if(m == NOTIF_MODE_OFF) {
						break;
					}
					bool ok = show_notification(L"ComPortNotify", L"Test notification");
					char msg[160];
					long hr = toast_last_error();
					if(m == NOTIF_MODE_BALLOON) {
						snprintf(msg, sizeof(msg), "Notification test: Balloon");
					} else if(m == NOTIF_MODE_TOAST) {
						snprintf(msg, sizeof(msg), "Notification test: Toast (%s, 0x%08lX)", ok ? "OK" : "FAIL", hr);
					}
					MessageBoxA(Hwnd, msg, "ComPortNotify", MB_OK | MB_ICONINFORMATION);
				} else if(temp == ID_TRAY_DISC_SHOW) {
					set_disconnected_mode(0);
				} else if(temp == ID_TRAY_DISC_HIDE) {
					set_disconnected_mode(1);
				} else if(temp == ID_TRAY_DISC_AFTER_10) {
					set_disconnected_mode(2);
					set_disconnected_timeout(10);
				} else if(temp == ID_TRAY_DISC_AFTER_60) {
					set_disconnected_mode(2);
					set_disconnected_timeout(60);
				} else if(temp == ID_TRAY_DISC_AFTER_900) {
					set_disconnected_mode(2);
					set_disconnected_timeout(900);
				} else if(temp == ID_TRAY_DISC_AFTER_1800) {
					set_disconnected_mode(2);
					set_disconnected_timeout(1800);
				} else if(temp == ID_TRAY_DISC_AFTER_3600) {
					set_disconnected_mode(2);
					set_disconnected_timeout(3600);
				}
				free_menu_texts(menu_texts);
				free_menu_clips(menu_clips);
			}
			break;

		case WM_MEASUREITEM: {
			MEASUREITEMSTRUCT *mi = (MEASUREITEMSTRUCT *)lParam;
			if(mi->CtlType == ODT_MENU && mi->itemData) {
				HDC hdc = GetDC(Hwnd);
				HFONT hfont = get_menu_font();
				HFONT old = hfont ? (HFONT)SelectObject(hdc, hfont) : NULL;
				SIZE sz = {0};
				const menu_text_t *mt = (const menu_text_t *)mi->itemData;
				SIZE right = {0};
				SIZE desc = {0};
				if(mt->right) GetTextExtentPoint32A(hdc, mt->right, (int)strlen(mt->right), &right);
				if(mt->desc) GetTextExtentPoint32A(hdc, mt->desc, (int)strlen(mt->desc), &desc);
				int checkPad = g_menu_nocheck ? GetSystemMetrics(SM_CXEDGE) : GetSystemMetrics(SM_CXMENUCHECK);
				int submenuPad = mt->has_submenu ? GetSystemMetrics(SM_CXMENUSIZE) : 0;
				int sidePad = GetSystemMetrics(SM_CXEDGE);
				mi->itemWidth = g_menu_prefix_width + g_menu_desc_width + right.cx + (checkPad * 2) + (sidePad * 2) + submenuPad;
				if(g_menu_desc_width > 0) mi->itemWidth += 8;
				if(right.cx > 0) mi->itemWidth += 12; // gap
				int h = desc.cy;
				if(right.cy > h) h = right.cy;
				mi->itemHeight = h + 6;
				if(old) SelectObject(hdc, old);
				ReleaseDC(Hwnd, hdc);
				return TRUE;
			}
		} break;

		case WM_DRAWITEM: {
			DRAWITEMSTRUCT *di = (DRAWITEMSTRUCT *)lParam;
			if(di->CtlType == ODT_MENU && di->itemData) {
				HDC hdc = di->hDC;
				RECT rc = di->rcItem;
				bool selected = (di->itemState & ODS_SELECTED) != 0;
				HBRUSH bg = GetSysColorBrush(selected ? COLOR_HIGHLIGHT : COLOR_MENU);
				FillRect(hdc, &rc, bg);

				SetBkMode(hdc, TRANSPARENT);
				const menu_text_t *mt = (const menu_text_t *)di->itemData;
				COLORREF textColor;
				if(mt->grayed) {
					textColor = GetSysColor(COLOR_GRAYTEXT);
				} else {
					textColor = GetSysColor(selected ? COLOR_HIGHLIGHTTEXT : COLOR_MENUTEXT);
				}
				SetTextColor(hdc, textColor);

				HFONT hfont = get_menu_font();
				HFONT old = hfont ? (HFONT)SelectObject(hdc, hfont) : NULL;

				RECT tx = rc;
				int checkPad = g_menu_nocheck ? GetSystemMetrics(SM_CXEDGE) : GetSystemMetrics(SM_CXMENUCHECK);
				int submenuPad = mt->has_submenu ? GetSystemMetrics(SM_CXMENUSIZE) : 0;
				int sidePad = GetSystemMetrics(SM_CXEDGE);
				tx.left += checkPad + sidePad;
				tx.right -= sidePad + submenuPad;
				if(mt->prefix && mt->prefix[0]) {
					RECT px = tx;
					px.right = px.left + g_menu_prefix_width;
					DrawTextA(hdc, mt->prefix, -1, &px, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);
					if(mt->grayed) {
						SIZE psz = {0};
						GetTextExtentPoint32A(hdc, mt->prefix, (int)strlen(mt->prefix), &psz);
						int y = (px.top + px.bottom) / 2;
						HPEN pen = CreatePen(PS_SOLID, 1, textColor);
						HPEN oldPen = pen ? (HPEN)SelectObject(hdc, pen) : NULL;
						int inset = 0;
						int x1 = px.left + inset;
						int x2 = px.left + psz.cx - inset;
						if(x2 < x1) x2 = x1;
						MoveToEx(hdc, x1, y, NULL);
						LineTo(hdc, x2, y);
						if(oldPen) SelectObject(hdc, oldPen);
						if(pen) DeleteObject(pen);
					}
					tx.left = px.right + 8;
				}
				RECT dx = tx;
				dx.right = dx.left + g_menu_desc_width;
				if(mt->desc) {
					DrawTextA(hdc, mt->desc, -1, &dx, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);
				}
				if(mt->right && mt->right[0]) {
					DrawTextA(hdc, mt->right, -1, &tx, DT_RIGHT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);
				}

				if(old) SelectObject(hdc, old);
				return TRUE;
			}
		} break;

		case WM_NCHITTEST:
			// Intercept the hit-test message
			temp = DefWindowProc(hwnd, WM_NCHITTEST, wParam, lParam);
			return temp == HTCLIENT ? HTCAPTION : temp;

		case WM_SETTINGCHANGE:
		case WM_DPICHANGED:
			reset_menu_font();
			break;

		case WM_CLOSE:
			minimize();
			return 0;
			break;

		case WM_DESTROY:
			PostQuitMessage(0);
			die = true;
			break;

    }

    return DefWindowProc( hwnd, message, wParam, lParam ) ;
}

void minimize() {
    // Hide the main window
    //ShowWindow(Hwnd, SW_HIDE);
}

void restore() {
    //ShowWindow(Hwnd, SW_SHOW);
}

void show_balloon(const char *title, const char *body) {
	NOTIFYICONDATA data = notifyIconData;
	data.uFlags = NIF_INFO;
	strncpy(data.szInfoTitle, title ? title : "ComPortNotify", sizeof(data.szInfoTitle));
	data.szInfoTitle[sizeof(data.szInfoTitle) - 1] = '\0';
	strncpy(data.szInfo, body ? body : "", sizeof(data.szInfo));
	data.szInfo[sizeof(data.szInfo) - 1] = '\0';
	data.dwInfoFlags = NIIF_INFO;
	Shell_NotifyIcon(NIM_MODIFY, &data);
}

bool show_notification(const wchar_t *title, const wchar_t *body) {
	int mode = get_notification_mode();
	if(mode == NOTIF_MODE_OFF) return false;
	if(mode == NOTIF_MODE_BALLOON) {
		char atitle[128];
		char abody[256];
		WideCharToMultiByte(CP_ACP, 0, title ? title : L"", -1, atitle, sizeof(atitle), NULL, NULL);
		WideCharToMultiByte(CP_ACP, 0, body ? body : L"", -1, abody, sizeof(abody), NULL, NULL);
		show_balloon(atitle, abody);
		return true;
	}
	if(mode == NOTIF_MODE_TOAST) {
		return toast_show_winrt(title, body);
	}
	return false;
}

static hport_t *find_hport(const char *device) {
	hport_t *p = history;
	while(p) {
		if(strcmp(p->device, device) == 0) return p;
		p = p->next;
	}
	return NULL;
}

static void remove_hport(hport_t *prev, hport_t *cur) {
	if(prev) {
		prev->next = cur->next;
	} else {
		history = cur->next;
	}
	free(cur->device);
	free(cur->name);
	free(cur->hwid);
	free(cur);
}

static void move_hport_to_head(hport_t *prev, hport_t *cur) {
	if(!prev || !cur) return;
	prev->next = cur->next;
	cur->next = history;
	history = cur;
}

void InitNotifyIconData() {
    memset( &notifyIconData, 0, sizeof( NOTIFYICONDATA ) ) ;

    notifyIconData.cbSize = sizeof(NOTIFYICONDATA);
    notifyIconData.hWnd = Hwnd;
    notifyIconData.uID = ID_TRAY_APP_ICON;
    notifyIconData.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
    notifyIconData.uCallbackMessage = WM_SYSICON; // Set up our invented Windows Message
    notifyIconData.hIcon = (HICON)LoadIcon( GetModuleHandle(NULL),      MAKEINTRESOURCE(ICO1) ) ;
    strncpy(notifyIconData.szTip, szTIP, sizeof(notifyIconData.szTip));
    notifyIconData.szTip[sizeof(notifyIconData.szTip) - 1] = '\0';
}
