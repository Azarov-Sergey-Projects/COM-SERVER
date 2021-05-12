#include <objbase.h>
#include <assert.h>

#include "Register.h"

BOOL setKeyAndValue(const wchar_t* pszPath,
                     const wchar_t* szSubkey,
                     const wchar_t* szValue) ;

void CLSIDtochar(const CLSID& clsid, 
                  wchar_t* szCLSID,
                  int length) ;


LONG recursiveDeleteKey(HKEY hKeyParent, const wchar_t* szKeyChild) ;


const int CLSID_STRING_SIZE = 39 ;


HRESULT RegisterServer(HMODULE hModule,            // DLL module handle
                        const CLSID& clsid,         // Class ID
                        const wchar_t* szFriendlyName, // Friendly Name
                        const wchar_t* szVerIndProgID, // Programmatic
                        const wchar_t* szProgID)       //   IDs
{
 
    wchar_t szModule[512] ;
    DWORD dwResult =
        ::GetModuleFileName(hModule, 
                             szModule,
                             sizeof(szModule)/sizeof(char)) ;
    assert(dwResult != 0) ;

    // Convert the CLSID into a char.
    wchar_t szCLSID[CLSID_STRING_SIZE] ;
    CLSIDtochar(clsid, szCLSID, sizeof(szCLSID)) ;

    // Build the key CLSID\\{...}
    wchar_t szKey[64] ;
    //lstrcpy( szKey, TEXT( "WOW6432Node\\" ) );
    lstrcpy( szKey, TEXT( "CLSID\\" ) );
    lstrcatW(szKey, szCLSID) ;

    // Add the CLSID to the registry.
    setKeyAndValue(szKey, NULL, szFriendlyName) ;

    // Add the server filename subkey under the CLSID key.
    setKeyAndValue(szKey, TEXT("InprocServer32"), szModule) ;

    // Add the ProgID subkey under the CLSID key.
    setKeyAndValue(szKey, TEXT("ProgID"), szProgID) ;

    // Add the version-independent ProgID subkey under CLSID key.
    setKeyAndValue(szKey, TEXT("VersionIndependentProgID"),
                    szVerIndProgID) ;

    // Add the version-independent ProgID subkey under HKEY_CLASSES_ROOT.
    setKeyAndValue(szVerIndProgID, NULL, szFriendlyName) ; 
    setKeyAndValue(szVerIndProgID, TEXT("CLSID"), szCLSID) ;
    setKeyAndValue(szVerIndProgID, TEXT("CurVer"), szProgID) ;

    // Add the versioned ProgID subkey under HKEY_CLASSES_ROOT.
    setKeyAndValue(szProgID, NULL, szFriendlyName) ; 
    setKeyAndValue(szProgID, TEXT("CLSID"), szCLSID) ;

    return S_OK ;
}

//
// Remove the component from the registry.
//
LONG UnregisterServer(const CLSID& clsid,         // Class ID
                       const wchar_t* szVerIndProgID, // Programmatic
                       const wchar_t* szProgID)       //   IDs
{
    // Convert the CLSID into a char.
    wchar_t szCLSID[CLSID_STRING_SIZE] ;
    CLSIDtochar(clsid, szCLSID, sizeof(szCLSID)) ;

    // Build the key CLSID\\{...}
    wchar_t szKey[64] ;
    //lstrcpy(szKey, TEXT("WOW6432Node\\CLSID\\")) ;
    lstrcpy(szKey, TEXT("CLSID\\")) ;
    lstrcatW(szKey, szCLSID) ;

    // Delete the CLSID Key - CLSID\{...}
    LONG lResult = recursiveDeleteKey(HKEY_CLASSES_ROOT, szKey) ;
    assert((lResult == ERROR_SUCCESS) ||
            (lResult == ERROR_FILE_NOT_FOUND)) ; // Subkey may not exist.

                                                 // Delete the version-independent ProgID Key.
    lResult = recursiveDeleteKey(HKEY_CLASSES_ROOT, szVerIndProgID) ;
    assert((lResult == ERROR_SUCCESS) ||
            (lResult == ERROR_FILE_NOT_FOUND)) ; // Subkey may not exist.

                                                 // Delete the ProgID key.
    lResult = recursiveDeleteKey(HKEY_CLASSES_ROOT, szProgID) ;
    assert((lResult == ERROR_SUCCESS) ||
            (lResult == ERROR_FILE_NOT_FOUND)) ; // Subkey may not exist.

    return S_OK ;
}

///////////////////////////////////////////////////////////
//
// Internal helper functions
//

// Convert a CLSID to a char string.
void CLSIDtochar(const CLSID& clsid,
                  wchar_t* szCLSID,
                  int length)
{
    assert(length >= CLSID_STRING_SIZE) ;
    // Get CLSID
    LPOLESTR wszCLSID = NULL ;
    HRESULT hr = StringFromCLSID(clsid, &wszCLSID) ;
    assert(SUCCEEDED(hr)) ;
    size_t i;
    // Covert from wide characters to non-wide.
    wcstombs(reinterpret_cast<char*>(szCLSID), wszCLSID, length) ;
    // Free memory.
    CoTaskMemFree(wszCLSID) ;
}

//
// Delete a key and all of its descendents.
//
LONG recursiveDeleteKey(HKEY hKeyParent,           // Parent of key to delete
                         const wchar_t* lpszKeyChild)  // Key to delete
{
    // Open the child.
    HKEY hKeyChild ;
    LONG lRes = RegOpenKeyEx(hKeyParent, lpszKeyChild, 0,
                              KEY_ALL_ACCESS, &hKeyChild) ;
    if (lRes != ERROR_SUCCESS)
    {
        return lRes ;
    }

    // Enumerate all of the decendents of this child.
    FILETIME time ;
    wchar_t szBuffer[256] ;
    DWORD dwSize = 256 ;
    while (RegEnumKeyEx(hKeyChild, 0, szBuffer, &dwSize, NULL,
                         NULL, NULL, &time) == S_OK)
    {
        // Delete the decendents of this child.
        lRes = recursiveDeleteKey(hKeyChild, szBuffer) ;
        if (lRes != ERROR_SUCCESS)
        {
            // Cleanup before exiting.
            RegCloseKey(hKeyChild) ;
            return lRes;
        }
        dwSize = 256 ;
    }

    // Close the child.
    RegCloseKey(hKeyChild) ;

    // Delete this child.
    return RegDeleteKey(hKeyParent, lpszKeyChild) ;
}

//
// Create a key and set its value.
//   - This helper function was borrowed and modifed from
//     Kraig Brockschmidt's book Inside OLE.
//
BOOL setKeyAndValue(const wchar_t* szKey,
                     const wchar_t* szSubkey,
                     const wchar_t* szValue)
{
    HKEY hKey;
    wchar_t szKeyBuf[1024] ;

    // Copy keyname into buffer.
    lstrcpy(szKeyBuf, szKey) ;

    // Add subkey name to buffer.
    if (szSubkey != NULL)
    {
        lstrcatW(szKeyBuf, TEXT("\\")) ;
        lstrcatW(szKeyBuf, szSubkey ) ;
    }

    // Create and open key and subkey.
    long lResult = RegCreateKeyEx(HKEY_CLASSES_ROOT ,
                                   szKeyBuf, 
                                   0, NULL, REG_OPTION_NON_VOLATILE,
                                   KEY_ALL_ACCESS, NULL, 
                                   &hKey, NULL) ;
    if (lResult != ERROR_SUCCESS)
    {
        return FALSE ;
    }

    // Set the Value.
    if (szValue != NULL)
    {
        RegSetValueEx(hKey, NULL, 0, REG_SZ, 
                       (BYTE *)szValue, 
                       lstrlenW(szValue)+1) ;
    }

    RegCloseKey(hKey) ;
    return TRUE ;
}