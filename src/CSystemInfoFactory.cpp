#include <Windows.h>
#include "CSystemInfoFactory.h"


CSystemInfoFactory::CSystemInfoFactory()
{
    m_lRef = 0;
}

CSystemInfoFactory::~CSystemInfoFactory()
{
}


STDMETHODIMP CSystemInfoFactory::QueryInterface( REFIID riid, void** ppv )
{
    *ppv = 0;

    if ( riid == IID_IUnknown || riid == IID_IClassFactory )
        *ppv = this;

    if ( *ppv )
    {
        AddRef();
        return S_OK;
    }

    return(E_NOINTERFACE);
}


STDMETHODIMP_(ULONG) CSystemInfoFactory::AddRef()
{
    return InterlockedIncrement( &m_lRef );
}

STDMETHODIMP_(ULONG) CSystemInfoFactory::Release()
{
    if ( InterlockedDecrement( &m_lRef ) == 0 )
    {
        delete this;
        return 0;
    }

    return m_lRef;
}

STDMETHODIMP CSystemInfoFactory::CreateInstance( LPUNKNOWN pUnkOuter, REFIID riid, void** ppvObj )
{
    CSystemInfo*      pSystemInfo;
    HRESULT    hr;

    *ppvObj = 0;

    pSystemInfo = new CSystemInfo;

    if ( pSystemInfo == 0 )
        return( E_OUTOFMEMORY );

    hr = pSystemInfo->QueryInterface( riid, ppvObj );

    if ( FAILED( hr ) )
        delete pSystemInfo;

    return hr;
}

STDMETHODIMP CSystemInfoFactory::LockServer( BOOL fLock )
{
    if ( fLock )
        InterlockedIncrement( &g_lLocks ); 
    else
        InterlockedDecrement( &g_lLocks );

    return S_OK;
}