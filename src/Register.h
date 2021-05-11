#pragma once
#include "pch.h"
HRESULT RegisterServer(HMODULE hModule, 
                        const CLSID& clsid, 
                        const wchar_t* szFriendlyName,
                        const wchar_t* szVerIndProgID,
                        const wchar_t* szProgID) ;

// This function will unregister a component.  Components
// call this function from their DllUnregisterServer function.
HRESULT UnregisterServer(const CLSID& clsid,
                          const wchar_t* szVerIndProgID,
                          const wchar_t* szProgID) ;