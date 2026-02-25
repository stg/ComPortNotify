// Source file for  serial communications
//
// senseitg@gmail.com 2012-May-22

#include <stdint.h>
#include <stdbool.h>
#include <string.h>
#include "serial.h"

// windows headers
#include <stdio.h>
#include <windows.h>
#include <winnt.h>
#include <setupapi.h>

static HANDLE h_serial;
static COMMTIMEOUTS restore;

// GUID for serial ports class
//static const GUID GUID_SERENUM_BUS_ENUMERATOR={0x86E0D1E0L,0x8089,0x11D0,{0x9C,0xE4,0x08,0x00,0x3E,0x30,0x1F,0x73}};
#ifndef GUID_CLASS_COMPORT
static const GUID GUID_CLASS_COMPORT          = {0x86E0D1E0L, 0x8089, 0x11D0, {0x9C, 0xE4, 0x08, 0x00, 0x3E, 0x30,0x1F, 0x73}};
#endif
#ifndef GUID_SERENUM_BUS_ENUMERATOR
static const GUID GUID_SERENUM_BUS_ENUMERATOR = {0x4D36E978L, 0xE325, 0x11CE, {0xBF, 0xC1, 0x08, 0x00, 0x2B, 0xE1, 0x03, 0x18}};
#endif

  
BOOL RegQueryValueString(HKEY kKey, char *lpValueName, char **pszValue) {
  *pszValue = NULL;
  DWORD dwType = 0;
  DWORD dwDataSize = 0;
  LONG nError = RegQueryValueEx(kKey, lpValueName, NULL, &dwType, NULL, &dwDataSize);
  if(nError != ERROR_SUCCESS || dwType != REG_SZ) return false;
  DWORD dwAllocatedSize = dwDataSize + 1;
  *pszValue = malloc(dwAllocatedSize); 
  if(!*pszValue) return false;
  (*pszValue)[0] = 0;
  DWORD dwReturnedSize = dwAllocatedSize;
  nError = RegQueryValueEx(kKey, lpValueName, NULL, &dwType, *pszValue, &dwReturnedSize);
  if(nError != ERROR_SUCCESS || dwReturnedSize >= dwAllocatedSize) {
    free(*pszValue);
    *pszValue = NULL;
    return false;
  }
  if(dwReturnedSize == 0 || (*pszValue)[dwReturnedSize - 1] != 0) (*pszValue)[dwReturnedSize] = 0;
  return true;
}
  
  
bool IsNumeric(char *pszString, bool bIgnoreColon) {
  size_t nLen = strlen(pszString), i;
  if(nLen == 0) return false;
  bool bNumeric = true;
  for(i=0; i<nLen && bNumeric; i++) {
    bNumeric = (pszString[i] >= '0' &&  pszString[i] <= '9');
    if(bIgnoreColon && (pszString[i] == L':')) bNumeric = true;
  }
  return bNumeric;
}
  

bool QueryRegistryPortName(HKEY hDeviceKey, int *nPort) {
  bool bAdded = false;
  char *pszPortName = NULL;
  if (RegQueryValueString(hDeviceKey, "PortName", &pszPortName)) {
    size_t nLen = strlen(pszPortName);
    if (nLen > 3) {
      if ((memcmp(pszPortName, "COM", 3) == 0) && IsNumeric((pszPortName + 3), false)) {
        *nPort = atoi(pszPortName + 3);
        bAdded = true;
      }
    }
    free(pszPortName);
  }
  return bAdded;
}

  
// windows - enumerate serial ports
void senum(void (*fp_enum)(char *name, char *device)) {
  HDEVINFO h_devinfo;

  // First need to convert the name "Ports" to a GUID using SetupDiClassGuidsFromName
  DWORD dwGuids = 0;
  SetupDiClassGuidsFromName("Ports", NULL, 0, &dwGuids);
  if (dwGuids == 0) return;

  // Allocate the needed memory
  void *pGuids = malloc(dwGuids * sizeof(GUID));
  if(!pGuids) return;

  // Call the function again
  if (!SetupDiClassGuidsFromName("Ports", pGuids, dwGuids, &dwGuids)) {
    free(pGuids);
    return;
  }

  // Now create a "device information set" which is required to enumerate all the ports
  h_devinfo = SetupDiGetClassDevs(pGuids, NULL, NULL, DIGCF_PRESENT);
  free(pGuids);
  if(h_devinfo == INVALID_HANDLE_VALUE) return;

  // Finally do the enumeration
  bool bMoreItems = true;
  int nIndex = 0;
  SP_DEVINFO_DATA devInfo;
  while (bMoreItems) {
    // Enumerate the current device
    devInfo.cbSize = sizeof(SP_DEVINFO_DATA);
    bMoreItems = SetupDiEnumDeviceInfo(h_devinfo, nIndex, &devInfo);
    if (bMoreItems) {
      // Did we find a serial port for this device
      bool bAdded = false;
      // Get the registry key which stores the ports settings
      HKEY hDeviceKey = SetupDiOpenDevRegKey(h_devinfo, &devInfo, DICS_FLAG_GLOBAL, 0, DIREG_DEV, KEY_QUERY_VALUE);
      int nPort = 0;
      char szPortName[9];
      if (hDeviceKey != INVALID_HANDLE_VALUE) {
        if (QueryRegistryPortName(hDeviceKey, &nPort)) {
          bAdded = true;
          if (snprintf(szPortName, sizeof(szPortName), "COM%u:", nPort) >= (int)sizeof(szPortName)) {
            bAdded = false;
          }
        }
        // Close the key now that we are finished with it
        RegCloseKey(hDeviceKey);
      }
      // If the port was a serial port, then also try to get its friendly name
      if(bAdded) {
        char szFriendlyName[1024];
        szFriendlyName[0] = 0;
        DWORD dwSize = sizeof(szFriendlyName);
        DWORD dwType = 0;
        if(SetupDiGetDeviceRegistryProperty(h_devinfo, &devInfo, SPDRP_DEVICEDESC, &dwType, szFriendlyName, dwSize, &dwSize) && (dwType == REG_SZ)) {
          fp_enum(szFriendlyName, szPortName);
        }
      }
    }

    ++nIndex;
  }

  // Free up the "device information set" now that we are finished with it
  SetupDiDestroyDeviceInfoList(h_devinfo);  
  

  /*
  HDEVINFO h_devinfo;
  SP_DEVICE_INTERFACE_DATA ifdata;
  SP_DEVICE_INTERFACE_DETAIL_DATA *ifdetail;
  SP_DEVINFO_DATA ifinfo;
  int count=0;
  DWORD size=0;
  char friendlyname[MAX_PATH+1];
  char name[MAX_PATH+1];
  char *namecat;
  DWORD type;
  HKEY devkey;
  DWORD keysize;
  // initialize some stuff
  memset(&ifdata,0,sizeof(ifdata));
  ifdata.cbSize=sizeof(SP_DEVICE_INTERFACE_DATA);
  memset(&ifinfo,0,sizeof(ifinfo));
  ifinfo.cbSize=sizeof(ifinfo);
  // set up a device information set
  h_devinfo=SetupDiGetClassDevs(&GUID_SERENUM_BUS_ENUMERATOR,NULL,NULL,DIGCF_PRESENT|DIGCF_DEVICEINTERFACE);
  if(h_devinfo) {
    while(1) {
      // enumerate devices
      if(!SetupDiEnumDeviceInterfaces(h_devinfo,NULL,&GUID_SERENUM_BUS_ENUMERATOR,count,&ifdata)) break;
      size=0; 
      // fetch size required for "interface details" struct
      SetupDiGetDeviceInterfaceDetail(h_devinfo,&ifdata,NULL,0,&size,NULL);
      if(size) {
        // allocate and initialize "interface details" struct
        ifdetail=malloc(sizeof(SP_DEVICE_INTERFACE_DETAIL_DATA)+size);
        memset(ifdetail,0,sizeof(SP_DEVICE_INTERFACE_DETAIL_DATA)+size);
        ifdetail->cbSize=sizeof(SP_DEVICE_INTERFACE_DETAIL_DATA);
        // get "device interface details" and "device info"
        if(SetupDiGetDeviceInterfaceDetail(h_devinfo,&ifdata,ifdetail,size,&size,&ifinfo)) {
          // retrieve "friendly name" from the registry
          if(!SetupDiGetDeviceRegistryProperty(h_devinfo,&ifinfo,SPDRP_FRIENDLYNAME,&type,friendlyname,sizeof(friendlyname),&size)) {
						friendlyname[0]=0;
          }
          // open the registry key containing device information
          devkey=SetupDiOpenDevRegKey(h_devinfo,&ifinfo,DICS_FLAG_GLOBAL,0,DIREG_DEV,KEY_READ);
          if(devkey!=INVALID_HANDLE_VALUE) {
            keysize=sizeof(name);
            // retrieve the actual *short* port name from the registry
            if(RegQueryValueEx(devkey,"PortName",NULL,NULL,name,&keysize)==ERROR_SUCCESS) {
              // whoop-dee-fucking-doo, we finally have what we need!
              if(strlen(friendlyname)) {
              	// build a display name (because "friendly name" may not contain port name early enough)
              	namecat=malloc(strlen(name)+strlen(friendlyname)+2);
              	strcpy(namecat,name);
              	strcat(namecat," ");
              	strcat(namecat,friendlyname);
              	fp_enum(namecat,name);
              	free(namecat);
            	} else {
            		// no "friendly name" available, just use device name
              	fp_enum(name,name);
            	}
            }
            // close the registry key
            RegCloseKey(devkey);
          }
        }
        free(ifdetail);
      }
      count++;
    }
    // DESTROY! DESTROY!
    SetupDiDestroyDeviceInfoList(h_devinfo);
  }
  */
  
}

// windows - open serial port
// device has form "COMn"
bool sopen(char* device) {
  h_serial=CreateFile(device,GENERIC_READ|GENERIC_WRITE,0,0,OPEN_EXISTING,0,0);
  if(h_serial==INVALID_HANDLE_VALUE) {
  	// can't open port, verify format and...
  	if(strlen(device)>=4) {
  		if(memcmp(device,"COM",3)==0) {
  			if(device[3]>='1'&&device[3]<='9') {
  				// ..try alternate format (\\.\comN)
  				char *dev=malloc(strlen(device)+5);
				if(!dev) return false;
  				strcpy(dev,"\\\\.\\com");
  				strcat(dev,device+3);
  				h_serial=CreateFile(dev,GENERIC_READ|GENERIC_WRITE,0,0,OPEN_EXISTING,0,0);
  				free(dev);
  			}
  		}
  	}
  }
  return h_serial!=INVALID_HANDLE_VALUE;
}

// windows - configure serial port
bool sconfig(char* fmt) {
  DCB dcb;
  COMMTIMEOUTS cmt;
  // clear dcb  
  memset(&dcb,0,sizeof(DCB));
  dcb.DCBlength=sizeof(DCB);
  // configure serial parameters
  if(!BuildCommDCB(fmt,&dcb)) return false;
  dcb.fOutxCtsFlow=0;
  dcb.fOutxDsrFlow=0;
  dcb.fDtrControl=0;
  dcb.fOutX=0;
  dcb.fInX=0;
  dcb.fRtsControl=0;
  if(!SetCommState(h_serial,&dcb)) return false;
  // configure buffers
  if(!SetupComm(h_serial,1024,1024)) return false;
  // configure timeouts 
  GetCommTimeouts(h_serial,&cmt);
  memcpy(&restore,&cmt,sizeof(cmt));
  cmt.ReadIntervalTimeout=1;
  cmt.ReadTotalTimeoutMultiplier=1;
  cmt.ReadTotalTimeoutConstant=1;
  cmt.WriteTotalTimeoutConstant=1;
  cmt.WriteTotalTimeoutMultiplier=1;
  if(!SetCommTimeouts(h_serial,&cmt)) return false;
  return true;
}

// windows - read from serial port
int32_t sread(void *p_read,uint16_t i_read) {
  DWORD i_actual=0;
  if(!ReadFile(h_serial,p_read,i_read,&i_actual,NULL)) return -1;
  return (int32_t)i_actual;
}

// windows - write to serial port
int32_t swrite(void* p_write,uint16_t i_write) {
  DWORD i_actual=0;
  if(!WriteFile(h_serial,p_write,i_write,&i_actual,NULL)) return -1;
  return (int32_t)i_actual;
}

// windows - close serial port
bool sclose() {
  // politeness: restore (some) original configuration
  SetCommTimeouts(h_serial,&restore);
  return CloseHandle(h_serial)!=0;
}
