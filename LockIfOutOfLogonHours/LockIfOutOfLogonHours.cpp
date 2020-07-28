//////////////////////////////////////////////////////////////////////////////
//
// LockIfOutOfLogonHours - Locks current session if user is out of logon hours
//
//  Copyright © 2015 Simon Rozman (simon@rozman.si)
//
//////////////////////////////////////////////////////////////////////////////

#define SECURITY_WIN32

#include "LockIfOutOfLogonHours.h"
#include <Windows.h>
#include <Security.h>
#include <tchar.h>

static void GetLogonHoursIndices(_In_ const SYSTEMTIME *time, _In_ const TIME_ZONE_INFORMATION *tzi, _Out_ LONG *plIndexSA, _Out_ LONG *plIndexBit)
{
    LONG wIndex = (((LONG)time->wDayOfWeek * 24 + time->wHour) * 60 + time->wMinute + tzi->Bias) / 60 % (21*8);
    *plIndexSA  = wIndex / 8,
    *plIndexBit = wIndex % 8;
}

static LPTSTR FormatMsg(_In_z_ LPCTSTR lpFormat, ...)
{
    LPTSTR lpBuffer;
    va_list arg;
    va_start(arg, lpFormat);
    if (FormatMessage(
        FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_STRING,
        lpFormat,
        0,
        0,
        (LPTSTR)&lpBuffer,
        0,
        &arg)) {
        va_end(arg);
        return lpBuffer;
    } else {
        va_end(arg);
        return NULL;
    }
}

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, PWSTR pCmdLine, int nCmdShow)
{
    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(pCmdLine);
    UNREFERENCED_PARAMETER(nCmdShow);

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

    // Check user logon hours.
    SYSTEMTIME time;
    GetLocalTime(&time);
    TIME_ZONE_INFORMATION time_zone;
    GetTimeZoneInformation(&time_zone);

    LONG lIndexSA, lIndexBit;
    GetLogonHoursIndices(&time, &time_zone, &lIndexSA, &lIndexBit);

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

    SYSTEMTIME timeFuture = time;
    timeFuture.wMinute += 10;
    GetLogonHoursIndices(&timeFuture, &time_zone, &lIndexSA, &lIndexBit);
    if (FAILED(SafeArrayGetElement(V_ARRAY(&varLogonHours), &lIndexSA, &bData))) {
        uiResult = 8; // Logon hour index not found
        goto _cleanup5;
    }
    if (uiResult == 0 && !(bData & (1 << lIndexBit))) {
        // Within logon hours now, but out of logon hours in up to 10 minutes.
        HINSTANCE hShell32Res = LoadLibraryEx(_T("shell32.dll"), NULL, LOAD_LIBRARY_AS_IMAGE_RESOURCE);
        LPTSTR lpTitleFmt;
        LoadString(hInstance, IDS_WARNING_TITLE_FMT, (LPTSTR)&lpTitleFmt, 0);
        TCHAR szDate[0x100];
        GetDateFormat(LOCALE_USER_DEFAULT, DATE_LONGDATE, &time, NULL, szDate, _countof(szDate));
        szDate[_countof(szDate) - 1] = 0;
        TCHAR szTime[0x100];
        GetTimeFormat(LOCALE_USER_DEFAULT, 0, &time, NULL, szTime, _countof(szTime));
        szTime[_countof(szTime) - 1] = 0;
        LPTSTR lpTitle = FormatMsg(lpTitleFmt, szDate, szTime);
        LPTSTR lpMsgFmt;
        LoadString(hInstance, IDS_WARNING_MSG_FMT, (LPTSTR)&lpMsgFmt, 0);
        LPTSTR lpMsg = FormatMsg(lpMsgFmt, 60 - time.wMinute);
        MSGBOXPARAMS params = {
            sizeof(MSGBOXPARAMS),   // cbSize
            NULL,                   // hwndOwner
            hShell32Res,            // hInstance
            lpMsg,                  // lpszText
            lpTitle,                // lpszCaption
            MB_OK | MB_USERICON | MB_SYSTEMMODAL | MB_SETFOREGROUND | MB_TOPMOST, // dwStyle
            MAKEINTRESOURCE(16771), // lpszIcon
        };
        MessageBoxIndirect(&params);
        LocalFree(lpMsg);
        LocalFree(lpTitle);
        FreeLibrary(hShell32Res);
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
