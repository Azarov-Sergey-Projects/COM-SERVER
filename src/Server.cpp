#include "pch.h"
#include <windows.h>

#include <initguid.h>
#include "SystemInfoFactory.h"
#include "registry.h"

long    g_lObjs = 0;
long    g_lLocks = 0;
extern const IID IID_ISystemInfo;

HMODULE g_hModule=NULL;
 CString friendlyName = TEXT( "Server.Friend.1" );
 CString progId = TEXT( "Server.Inproc.1" );
 CString noVerProgId = TEXT( "ServerCOM.1" );


STDAPI DllGetClassObject( REFCLSID rclsid, REFIID riid, void** ppv )
{
    HRESULT             hr;
    SystemInfoFactory    *pCF;

    pCF = nullptr;

    if ( rclsid != CLSID_SystemInfo )
        return( E_FAIL );

    pCF = new  SystemInfoFactory;

    if ( pCF == 0 )
        return( E_OUTOFMEMORY );

    hr = pCF->QueryInterface( riid, ppv );
    if ( FAILED( hr ) )
    {
        delete pCF;
        pCF = 0;
    }

    return hr;
}

STDAPI DllCanUnloadNow()
{
    if ( g_lObjs || g_lLocks )
        return( S_FALSE );
    else
        return( S_OK );
}

STDAPI DllRegisterServer()
{
    return RegisterServer(g_hModule, 
                          CLSID_SystemInfo,
                           friendlyName,
                           noVerProgId,
                          progId) ;
}

STDAPI DllUnregisterServer()
{
    return UnregisterServer(CLSID_SystemInfo,
                            noVerProgId,
                             progId) ;
}


BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved)
{

    if( ul_reason_for_call == DLL_PROCESS_ATTACH )
    {
        g_hModule = hModule;
    }
    return TRUE;
}

