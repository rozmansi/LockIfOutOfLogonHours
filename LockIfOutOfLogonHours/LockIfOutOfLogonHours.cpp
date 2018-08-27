//////////////////////////////////////////////////////////////////////////////
//
// LockIfOutOfLogonHours - Locks current session if user is out of logon hours
//
//  Copyright © 2015 Simon Rozman (simon@rozman.si)
//
//////////////////////////////////////////////////////////////////////////////

#define SECURITY_WIN32

#include <Windows.h>
#include <Security.h>
#include <tchar.h>


extern "C" void WinMainCRTStartup()
{
    static const TCHAR pszURIPrefix[] = _T("LDAP://");
    static const LPOLESTR ppszUserObjectMembers[] = { L"logonHours", };
    static const DISPPARAMS paramsEmpty = {};

    UINT uiResult;
    HANDLE hHeap = GetProcessHeap();
    LPTSTR pszADsPath;
    IDispatch *pUserObj;
    DISPID dispidUserObjectMembers[_countof(ppszUserObjectMembers)];
    VARIANT varLogonHours = { VT_EMPTY };

    if (FAILED(CoInitialize(NULL))) {
        uiResult = 2; // COM initialization failed
        goto _cleanup1;
    }

    // Get user fully qualified DN.
    for (ULONG nLength = 1024; ;) {
        if ((pszADsPath = (LPTSTR)HeapAlloc(hHeap, 0, (nLength + _countof(pszURIPrefix) - 1)*sizeof(TCHAR))) == NULL) {
            uiResult = 3; // Out of memory
            goto _cleanup2;
        }
        if (GetUserNameEx(NameFullyQualifiedDN, pszADsPath + _countof(pszURIPrefix) - 1, &nLength))
            break;
        else {
            HeapFree(hHeap, 0, pszADsPath);
            if (GetLastError() != ERROR_MORE_DATA) {
                uiResult = 4; // GetUserNameEx() error
                goto _cleanup2;
            }
        }
    }

    // Add LDAP URI prefix.
    CopyMemory(pszADsPath, pszURIPrefix, (_countof(pszURIPrefix) - 1)*sizeof(TCHAR));

    if (FAILED(CoGetObject(pszADsPath, NULL, IID_IDispatch, (void**)&pUserObj))) {
        uiResult = 5; // Get user object failed
        goto _cleanup3;
    }

    if (FAILED(pUserObj->GetIDsOfNames(IID_NULL, (LPOLESTR*)ppszUserObjectMembers, _countof(ppszUserObjectMembers), LOCALE_SYSTEM_DEFAULT, dispidUserObjectMembers))) {
        uiResult = 6; // "logonHours" property not found
        goto _cleanup4;
    }

    if (FAILED(pUserObj->Invoke(dispidUserObjectMembers[0], IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET, (DISPPARAMS*)&paramsEmpty, &varLogonHours, NULL, NULL))) {
        uiResult = 6; // "logonHours" property get failed
        goto _cleanup4;
    }

    if (V_VT(&varLogonHours) == VT_EMPTY) {
        uiResult = 0; // User has no logon hours defined
        goto _cleanup5;
    } else if (V_VT(&varLogonHours) != (VT_ARRAY | VT_UI1)) {
        uiResult = 7; // Invalid logon hours
        goto _cleanup5;
    }

    {
        // Check user logon hours.
        SYSTEMTIME time;
        GetLocalTime(&time);
        TIME_ZONE_INFORMATION time_zone;
        GetTimeZoneInformation(&time_zone);

        WORD wIndex = ((time.wDayOfWeek * 24 + time.wHour) * 60 + time.wMinute + time_zone.Bias) / 60 % (21*8);
        LONG
            lIndexSA  = wIndex / 8,
            lIndexBit = wIndex % 8;
        UCHAR bData;

        if (FAILED(SafeArrayGetElement(V_ARRAY(&varLogonHours), &lIndexSA, &bData))) {
            uiResult = 8; // Logon hour index not found
            goto _cleanup5;
        }

        if (bData & (1 << lIndexBit)) {
            uiResult = 0; // Within logon hours
        } else {
            uiResult = 1; // Out of logon hours
            LockWorkStation();
        }
    }

_cleanup5:
    VariantClear(&varLogonHours);
_cleanup4:
    pUserObj->Release();
_cleanup3:
    HeapFree(hHeap, 0, pszADsPath);
_cleanup2:
    CoUninitialize();
_cleanup1:
    ExitProcess(uiResult);
}
