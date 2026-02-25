#include <stdarg.h>
#include <stdio.h>
#include "resource.h"
#include "serial.h"

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

// Port linked list
typedef struct lport {
	char * device;
	char * name;
	struct lport *next;
} lport_t;

// Free ports linked list
void free_lports(lport_t *p) {
	while(p) {
		free(p->device);
		free(p->name);
		lport_t *_p = p;
		p = p->next;
		free(_p);
	}
}

// Previous and current port list, for change detection
lport_t * volatile ports, * _ports;

// Add new port to linked list
void add_lport(char *name, char *device) {
	lport_t * temp = (lport_t *)malloc(sizeof(lport_t));
	char *dev = (char *)malloc(strlen(device) + 1);
	char *nam = (char *)malloc(strlen(name) + 1);
	if(!temp || !dev || !nam) {
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
	temp->next = _ports;
	_ports = temp;
}

// Show notification (each instance runs in its own thread)
// Assumes lpParam is heap-allocated and will free it
DWORD WINAPI NotifyThread( LPVOID lpParam ) {
    STARTUPINFO si;
    PROCESS_INFORMATION pi;
	LPSTR szCommand;
	
	// Generate command line
	szCommand = mpprintf("wintoast --appname \"%s\" --aumi \"DSp.Tools.CPNotify.1\" --image \"toast.png\" --expirems 1 --text \"%s\"",  szClassName, lpParam);
	if(!szCommand) {
		free(lpParam);
		return 0;
	}
	
	// Set up structs
	ZeroMemory( &si, sizeof(si) );
    si.cb = sizeof(si);
    ZeroMemory( &pi, sizeof(pi) );

    // Start the child process. 
    if( CreateProcess( NULL,   // No module name (use command line)
        szCommand,        // Command line
        NULL,           // Process handle not inheritable
        NULL,           // Thread handle not inheritable
        FALSE,          // Set handle inheritance to FALSE
        CREATE_NO_WINDOW ,              // No creation flags
        NULL,           // Use parent's environment block
        NULL,           // Use parent's starting directory 
        &si,            // Pointer to STARTUPINFO structure
        &pi )           // Pointer to PROCESS_INFORMATION structure
    ) {
        // Close handles if CreateProcess succeeded
        CloseHandle(pi.hThread);
        CloseHandle(pi.hProcess);
    }

	// Free up allocated memory
	free(szCommand);
	free(lpParam);
	return 0;
}

// Refresh the ports list
void refresh_ports(bool init = false) {
	// TODO: List should note time of new connections
	// TODO: Should show recent history and times on popup menu (disabled/grayed for disconnected ports, with timeout?)
	char szTooltip[sizeof(notifyIconData.szTip)] = {0};
	senum(add_lport); // List current ports to _ports
	if(!init) {		
		lport_t * b, * a;
		// Detect removals (naive search)
		a = ports;
		while(a) {
			b = _ports;
			while(b) {
				if(!strcmp(b->device, a->device)) break;
				b = b->next;
			}
			if(!b) {
				char * text = mpprintf("Removed %s %s\n", a->device, a->name);
				if(!text) {
					a = a->next;
					continue;
				}
				// Update notification
				if(!szTooltip[0]) {
					strncpy(szTooltip, text, sizeof(szTooltip));
					szTooltip[sizeof(szTooltip) - 1] = '\0';
				}
				// Send toast notification
				//CreateThread(NULL, 0, NotifyThread, text, 0, NULL);
				free(text);
			}
			a = a->next;
		}
		// Detect additions (naive search)
		a = _ports;
		while(a) {
			b = ports;
			while(b) {
				if(!strcmp(b->device, a->device)) break;
				b = b->next;
			}
			if(b) {
				/*
				if(strcmp(b->name, a->name)) {
					printf("Updated %s %s\n", a->device, a->name);
				}
				*/
			} else {
				char * text = mpprintf("Connected %s %s\n", a->device, a->name);
				if(!text) {
					a = a->next;
					continue;
				}
				// Update notification
				if(!szTooltip[0]) {
					strncpy(szTooltip, text, sizeof(szTooltip));
					szTooltip[sizeof(szTooltip) - 1] = '\0';
				}
				// Send toast notification
				HANDLE hThread = CreateThread(NULL, 0, NotifyThread, text, 0, NULL);
				if(hThread) {
					CloseHandle(hThread);
				} else {
					free(text);
				}
				
			}
			a = a->next;
		}
		free_lports(ports);

		if(szTooltip[0]) {
			strncpy(notifyIconData.szTip, szTooltip, sizeof(notifyIconData.szTip));
			notifyIconData.szTip[sizeof(notifyIconData.szTip) - 1] = '\0';
			Shell_NotifyIcon(NIM_MODIFY, &notifyIconData);
		}
	}

	ports = _ports;
	_ports = NULL;

}

// Populate the popup menu
void populate_menu() {
	lport_t * p = ports;
	if(p) {
		do {
			char * text = mpprintf("%s %s\n", p->device, p->name);
			if(text) {
				AppendMenu(Hmenu, MF_STRING, ID_TRAY_VOID, TEXT( text ));
				free(text);
			}
			p = p->next;
		} while(p);
	} else {
		AppendMenu(Hmenu, MF_STRING, ID_TRAY_VOID, TEXT( "No serial ports detected" ));
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
				populate_menu();
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
				}
			}
			break;

		case WM_NCHITTEST:
			// Intercept the hit-test message
			temp = DefWindowProc(hwnd, WM_NCHITTEST, wParam, lParam);
			return temp == HTCLIENT ? HTCAPTION : temp;

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
