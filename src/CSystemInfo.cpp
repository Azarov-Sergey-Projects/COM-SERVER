#include "pch.h"
#include <comdef.h>
#include <Wbemidl.h>
#include <dciman.h>
#include "CSystemInfo.h"
#pragma comment(lib, "wbemuuid.lib")
CSystemInfo::CSystemInfo()
{
    m_lRef = 0;

    // Increment the global object count
    InterlockedIncrement( &g_lObjs ); 
}

// The destructor
CSystemInfo::~CSystemInfo()
{
    // Decrement the global object count
    InterlockedDecrement( &g_lObjs ); 
}

STDMETHODIMP CSystemInfo::QueryInterface( REFIID riid, void** ppv )
{
    *ppv = nullptr;

    if ( riid == IID_IUnknown || riid == IID_ISystemInfo )
        *ppv = this;

    if ( *ppv )
    {
        AddRef();
        return( S_OK );
    }
    return (E_NOINTERFACE);
}

STDMETHODIMP_(ULONG) CSystemInfo::AddRef()
{
    return InterlockedIncrement( &m_lRef );
}

STDMETHODIMP_(ULONG) CSystemInfo::Release()
{
    if ( InterlockedDecrement( &m_lRef ) == 0 )
    {
        delete this;
        return 0;
    }

    return m_lRef;
}

STDMETHODIMP CSystemInfo::GetOS(CString* SystemInfo )
{
    HRESULT hres;
    hres = GetInfo( TEXT( "Win32_OperatingSystem" ), TEXT( "Version" ), SystemInfo );
    if( FAILED( hres ) )
    {
        return E_FAIL;
    }
    return S_OK;
}

STDMETHODIMP CSystemInfo::GetMBoardCreator( CString* info )
{
    HRESULT hres;
    hres=GetInfo( TEXT( "Win32_BaseBoard" ), TEXT( "Manufacturer" ), info );
    if( FAILED( hres ) )
    {
        return E_FAIL;
    }
    return S_OK;

}

STDMETHODIMP CSystemInfo::GetCPUINFO( UINT* clocks,UINT *frequency )
{
    HRESULT hres;
    hres = GetInfoUINT( TEXT( "Win32_Processor" ), TEXT( "NumberOfCores" ), clocks );
    if( FAILED( hres ) )
    {
        return E_FAIL;
    }
    hres = GetInfoUINT( TEXT( "Win32_Processor" ), TEXT( "MaxClockSpeed" ), frequency );
    if( FAILED( hres ) )
    {
        return E_FAIL;
    }
    return S_OK;
}

STDMETHODIMP_(long) CSystemInfo::MonitorInfo( CString* info, int* MonitorCount )
{
    HRESULT hres;
    *MonitorCount = GetSystemMetrics( SM_CMONITORS );
    if( *MonitorCount != 0 )
    {
        hres = GetInfo( TEXT( "Win32_DesktopMonitor" ), TEXT( "DeviceID" ), info );
    }
    else
    {
        info->SetString( TEXT( "There is no monitors connected to your PC" ) );
        hres = E_FAIL;
    }
    if( FAILED( hres ) )
    {
        return E_FAIL;
    }
    return S_OK;
}

STDMETHODIMP CSystemInfo::GetInfo( CString className, CString propertyName, CString* info )
{
    HRESULT hres;
    IWbemLocator* pLoc = NULL;
    IWbemServices* pSvc = NULL;
    IEnumWbemClassObject* pEnumerator = NULL;
    bool initialized = true;

    hres = CoInitialize( NULL );
    if( FAILED( hres ) )
    {
        return E_FAIL;
    }

    hres = CoInitializeSecurity(
        NULL,
        -1,                          // COM authentication
        NULL,                        // Authentication services
        NULL,                        // Reserved
        RPC_C_AUTHN_LEVEL_DEFAULT,   // Default authentication 
        RPC_C_IMP_LEVEL_IMPERSONATE, // Default Impersonation  
        NULL,                        // Authentication info
        EOAC_NONE,                   // Additional capabilities 
        NULL                         // Reserved
    );

    if( FAILED( hres ) && hres != RPC_E_TOO_LATE )
    {
    }

    hres = CoCreateInstance(
        CLSID_WbemAdministrativeLocator,
        0,
        CLSCTX_INPROC_SERVER,
        IID_IWbemLocator, reinterpret_cast< LPVOID* >(&pLoc) );

    if( FAILED( hres ) )
    {
        return E_FAIL;
    }

    hres = pLoc->ConnectServer(
        bstr_t( L"ROOT\\CIMV2" ),  // Object path of WMI namespace
        NULL,                    // User name. NULL = current user
        NULL,                    // User password. NULL = current
        0,                       // Locale. NULL indicates current
        NULL,                    // Security flags.
        0,                       // Authority (for example, Kerberos)
        0,                       // Context object 
        &pSvc                    // pointer to IWbemServices proxy
    );

    if( FAILED( hres ) )
    {
        return E_FAIL;
    }

    hres = CoSetProxyBlanket(
        pSvc,                        // Indicates the proxy to set
        RPC_C_AUTHN_WINNT,           // RPC_C_AUTHN_xxx
        RPC_C_AUTHZ_NONE,            // RPC_C_AUTHZ_xxx
        NULL,                        // Server principal name 
        RPC_C_AUTHN_LEVEL_CALL,      // RPC_C_AUTHN_LEVEL_xxx 
        RPC_C_IMP_LEVEL_IMPERSONATE, // RPC_C_IMP_LEVEL_xxx
        NULL,                        // client identity
        EOAC_NONE                    // proxy capabilities 
    );

    if( FAILED( hres ) )
    {
        return E_FAIL;
    }


    CString tmp = TEXT( "SELECT * FROM  ");
    tmp += className.GetString();
    hres = pSvc->ExecQuery(
        bstr_t( "WQL" ),
        bstr_t(tmp.GetString()),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
        NULL,
        &pEnumerator );

    if( FAILED( hres ) )
    {
        return E_FAIL;
    }

    IWbemClassObject* pclsObj = NULL;
    ULONG uReturn = 0;

    while( pEnumerator )
    {
        HRESULT hr = pEnumerator->Next( WBEM_INFINITE, 1, &pclsObj, &uReturn );

        if( uReturn == 0 )
        {
            break;
        }

        VARIANT vtProp;

        hr = pclsObj->Get(propertyName.GetString(), 0, &vtProp, 0, 0 );
        if( FAILED( hres ) )
        {
            return E_FAIL;
        }
        info->SetString( vtProp.bstrVal );
        VariantClear( &vtProp );
        pclsObj->Release();
    }
    pSvc->Release();
    pLoc->Release();
    pEnumerator->Release();
    CoUninitialize();
    return S_OK;
}

STDMETHODIMP CSystemInfo::GetInfoUINT( CString className, CString propertyName, uint32_t* info )
{
    HRESULT hres;
    IWbemLocator* pLoc = NULL;
    IWbemServices* pSvc = NULL;
    IEnumWbemClassObject* pEnumerator = NULL;
    bool initialized = true;

    hres = CoInitialize( NULL );
    if( FAILED( hres ) )
    {
        return E_FAIL;
    }

    hres = CoInitializeSecurity(
        NULL,
        -1,                          // COM authentication
        NULL,                        // Authentication services
        NULL,                        // Reserved
        RPC_C_AUTHN_LEVEL_DEFAULT,   // Default authentication 
        RPC_C_IMP_LEVEL_IMPERSONATE, // Default Impersonation  
        NULL,                        // Authentication info
        EOAC_NONE,                   // Additional capabilities 
        NULL                         // Reserved
    );

    if( FAILED( hres ) && hres != RPC_E_TOO_LATE )
    {
    }

    hres = CoCreateInstance(
        CLSID_WbemLocator,
        0,
        CLSCTX_INPROC_SERVER,
        IID_IWbemLocator, ( LPVOID* )&pLoc );

    if( FAILED( hres ) )
    {
        return E_FAIL;
    }

    hres = pLoc->ConnectServer(
        bstr_t( L"ROOT\\CIMV2" ),  // Object path of WMI namespace
        NULL,                    // User name. NULL = current user
        NULL,                    // User password. NULL = current
        0,                       // Locale. NULL indicates current
        NULL,                    // Security flags.
        0,                       // Authority (for example, Kerberos)
        0,                       // Context object 
        &pSvc                    // pointer to IWbemServices proxy
    );

    if( FAILED( hres ) )
    {
        return E_FAIL;
    }

    hres = CoSetProxyBlanket(
        pSvc,                        // Indicates the proxy to set
        RPC_C_AUTHN_WINNT,           // RPC_C_AUTHN_xxx
        RPC_C_AUTHZ_NONE,            // RPC_C_AUTHZ_xxx
        NULL,                        // Server principal name 
        RPC_C_AUTHN_LEVEL_CALL,      // RPC_C_AUTHN_LEVEL_xxx 
        RPC_C_IMP_LEVEL_IMPERSONATE, // RPC_C_IMP_LEVEL_xxx
        NULL,                        // client identity
        EOAC_NONE                    // proxy capabilities 
    );

    if( FAILED( hres ) )
    {
        return E_FAIL;
    }


    CString tmp = TEXT( "SELECT * FROM ");
    tmp += className.GetString();
    hres = pSvc->ExecQuery(
        bstr_t( "WQL" ),
        bstr_t(tmp.GetString()),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
        NULL,
        &pEnumerator );

    if( FAILED( hres ) )
    {
        return E_FAIL;
    }

    IWbemClassObject* pclsObj = NULL;
    ULONG uReturn = 0;

    while( pEnumerator )
    {
        HRESULT hr = pEnumerator->Next( WBEM_INFINITE, 1, &pclsObj, &uReturn );

        if( uReturn == 0 )
        {
            break;
        }

        VARIANT vtProp;

        hr = pclsObj->Get(propertyName.GetString(), 0, &vtProp, 0, 0 );
        if( FAILED( hres ) )
        {
            return E_FAIL;
        }
        *info = vtProp.uintVal;
        VariantClear( &vtProp );
        pclsObj->Release();
    }
    CoUninitialize();
    return S_OK;
}

