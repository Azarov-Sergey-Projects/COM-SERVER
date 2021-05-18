#pragma once
#include <Unknwn.h>
#include "SystemInfo.h"
class SystemInfoFactory :public IClassFactory
{
public:
    SystemInfoFactory();
    ~SystemInfoFactory();

    STDMETHOD( QueryInterface( REFIID, void** ) );
    STDMETHOD_( ULONG, AddRef() );
    STDMETHOD_( ULONG, Release() );

    STDMETHOD( CreateInstance( LPUNKNOWN, REFIID, void** ) );
    STDMETHOD( LockServer( BOOL ) );
protected:
    long m_lRef;
};