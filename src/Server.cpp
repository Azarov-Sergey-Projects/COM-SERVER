#include <windows.h>

#include <initguid.h>
#include "CSystemInfoFactory.h"
#include "Register.h"
long    g_lObjs = 0;
long    g_lLocks = 0;
extern const IID IID_ISystemInfo;
HMODULE g_hModule=NULL;
CString friendlyName = TEXT( "Server.Inproc.1" );
CString progId = TEXT( "Server.Inproc.1" );
CString noVerProgId = TEXT( "ServerCOM.1" );


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