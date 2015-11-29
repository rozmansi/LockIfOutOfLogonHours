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

#pragma comment(lib, "Secur32.lib")


int CALLBACK WinMain(_In_ HINSTANCE hInstance, _In_ HINSTANCE hPrevInstance, _In_ LPSTR lpCmdLine, _In_ int nCmdShow)
{
    UNREFERENCED_PARAMETER(hInstance);
    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);
    UNREFERENCED_PARAMETER(nCmdShow);

    static const WCHAR pszURIPrefix[] = L"LDAP://";
    static const LPOLESTR ppszUserObjectMembers[] = { L"logonHours", };
    static const DISPPARAMS paramsEmpty = {};

    int iResult;
    HANDLE hHeap = GetProcessHeap();
    LPTSTR pszADsPath;
    IDispatch *pUserObj;
    DISPID dispidUserObjectMembers[_countof(ppszUserObjectMembers)];
    VARIANT varLogonHours = { VT_EMPTY };

    if (FAILED(CoInitialize(NULL))) {
        iResult = 2; // COM initialization failed
        goto _cleanup1;
    }

    // Get user fully qualified DN.
    for (ULONG nLength = 1024; ;) {
        if ((pszADsPath = (LPTSTR)HeapAlloc(hHeap, 0, (nLength + _countof(pszURIPrefix) - 1)*sizeof(WCHAR))) == NULL) {
            iResult = 3; // Out of memory
            goto _cleanup2;
        }
        if (GetUserNameEx(NameFullyQualifiedDN, pszADsPath + _countof(pszURIPrefix) - 1, &nLength))
            break;
        else {
            HeapFree(hHeap, 0, pszADsPath);
            if (GetLastError() != ERROR_MORE_DATA) {
                iResult = 4; // GetUserNameEx() error
                goto _cleanup2;
            }
        }
    }

    // Add LDAP URI prefix.
    CopyMemory(pszADsPath, pszURIPrefix, (_countof(pszURIPrefix) - 1)*sizeof(WCHAR));

    if (FAILED(CoGetObject(pszADsPath, NULL, IID_IDispatch, (void**)&pUserObj))) {
        iResult = 5; // Get user object failed
        goto _cleanup3;
    }

    if (FAILED(pUserObj->GetIDsOfNames(IID_NULL, (LPOLESTR*)ppszUserObjectMembers, _countof(ppszUserObjectMembers), LOCALE_SYSTEM_DEFAULT, dispidUserObjectMembers))) {
        iResult = 6; // "logonHours" property not found
        goto _cleanup4;
    }

    if (FAILED(pUserObj->Invoke(dispidUserObjectMembers[0], IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET, (DISPPARAMS*)&paramsEmpty, &varLogonHours, NULL, NULL))) {
        iResult = 6; // "logonHours" property get failed
        goto _cleanup4;
    }

    if (V_VT(&varLogonHours) == VT_EMPTY) {
        iResult = 0; // User has no logon hours defined
        goto _cleanup5;
    } else if (V_VT(&varLogonHours) != (VT_ARRAY | VT_UI1)) {
        iResult = 7; // Invalid logon hours
        goto _cleanup5;
    }

    {
        // Check user logon hours.
        SYSTEMTIME time;
        GetSystemTime(&time);

        WORD wIndex = time.wDayOfWeek * 24 + time.wHour;
        LONG
            lIndexSA  = wIndex / 8,
            lIndexBit = wIndex % 8;
        UCHAR bData;

        if (FAILED(SafeArrayGetElement(V_ARRAY(&varLogonHours), &lIndexSA, &bData))) {
            iResult = 8; // Logon hour index not found
            goto _cleanup5;
        }

        if (bData & (1 << lIndexBit)) {
            iResult = 0; // Within logon hours
        } else {
            iResult = 1; // Out of logon hours
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
    return iResult;
}
