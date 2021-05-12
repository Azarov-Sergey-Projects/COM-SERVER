#include <windows.h>

#include <initguid.h>
#include "CSystemInfoFactory.h"
#include "Register.h"
long    g_lObjs = 0;
long    g_lLocks = 0;
extern const IID IID_ISystemInfo;
//const GUID IID_ISystemInfo = { 0x32bb8320, 0xb41b, 0x11cf, 0xa6, 0xbb, 0x0, 0x80, 0xc7, 0xb2, 0xd6, 0x82 };

HMODULE g_hModule=NULL;
static CString friendlyName = TEXT( "Server.Friend.1" );
static CString progId = TEXT( "Server.Inproc.1" );
static CString noVerProgId = TEXT( "ServerCOM.1" );


STDAPI DllGetClassObject( REFCLSID rclsid, REFIID riid, void** ppv )
{
    HRESULT             hr;
    CSystemInfoFactory    *pCF;

    pCF = 0;

    if ( rclsid != IID_ISystemInfo )
        return( E_FAIL );

    pCF = new  CSystemInfoFactory;

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
                          IID_ISystemInfo,
                           friendlyName,
                           noVerProgId,
                          progId) ;
}

STDAPI DllUnregisterServer()
{
    return UnregisterServer(IID_ISystemInfo,
                            noVerProgId,
                             progId) ;
}


BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved)
{
    /*  switch (ul_reason_for_call)
    {
    case DLL_PROCESS_ATTACH:
    case DLL_THREAD_ATTACH:
    case DLL_THREAD_DETACH:
    case DLL_PROCESS_DETACH:
    break;
    }*/
    if( ul_reason_for_call == DLL_PROCESS_ATTACH )
    {
        g_hModule = hModule;
    }
    return TRUE;
}

