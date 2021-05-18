#include "pch.h"
#include <Windows.h>
#include "SystemInfoFactory.h"


SystemInfoFactory::SystemInfoFactory()
{
    m_lRef = 0;
}

SystemInfoFactory::~SystemInfoFactory()
{
}


STDMETHODIMP SystemInfoFactory::QueryInterface( REFIID riid, void** ppv )
{
    *ppv = nullptr;

    if ( riid == IID_IUnknown || riid == IID_IClassFactory )
        *ppv = this;

    if ( *ppv )
    {
        AddRef();
        return S_OK;
    }

    return(E_NOINTERFACE);
}


STDMETHODIMP_(ULONG) SystemInfoFactory::AddRef()
{
    return InterlockedIncrement( &m_lRef );
}

STDMETHODIMP_(ULONG) SystemInfoFactory::Release()
{
    if ( InterlockedDecrement( &m_lRef ) == 0 )
    {
        delete this;
        return 0;
    }

    return m_lRef;
}

STDMETHODIMP SystemInfoFactory::CreateInstance( LPUNKNOWN pUnkOuter, REFIID riid, void** ppvObj )
{
    SystemInfo*      pSystemInfo;
    HRESULT    hr;

    *ppvObj = nullptr;

    pSystemInfo = new SystemInfo;

    if ( pSystemInfo == 0 )
        return( E_OUTOFMEMORY );

    hr = pSystemInfo->QueryInterface( riid, ppvObj );

    if ( FAILED( hr ) )
        delete pSystemInfo;

    return hr;
}

STDMETHODIMP SystemInfoFactory::LockServer( BOOL fLock )
{
    if ( fLock )
        InterlockedIncrement( &g_lLocks ); 
    else
        InterlockedDecrement( &g_lLocks );

    return S_OK;
}