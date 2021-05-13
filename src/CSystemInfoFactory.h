#pragma once
#include <Unknwn.h>
#include "CSystemInfo.h"
class CSystemInfoFactory :public IClassFactory
{
public:
    CSystemInfoFactory();
    ~CSystemInfoFactory();

    STDMETHOD( QueryInterface( REFIID, void** ) );
    STDMETHOD_( ULONG, AddRef() );
    STDMETHOD_( ULONG, Release() );

    STDMETHOD( CreateInstance( LPUNKNOWN, REFIID, void** ) );
    STDMETHOD( LockServer( BOOL ) );
protected:
    long m_lRef;
};