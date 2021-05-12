#pragma once

#include <objbase.h>
#include <INITGUID.H>
#include "ISystemInfo.h"
extern long g_lObjs;
extern long g_lLocks;
extern HMODULE g_hModule;

DEFINE_GUID( IID_ISystemInfo, 0x32bb8320, 0xb41b, 0x11cf, 0xa6, 0xbb, 0x0, 0x80, 0xc7, 0xb2, 0xd6, 0x82 );
DEFINE_GUID( CLSID_SystemInfo, 0x32bb8319, 0xb41b, 0x11cf, 0xa6, 0xbb, 0x0, 0x80, 0xc7, 0xb2, 0xd6, 0x82 );
/*extern "C"
{
   const IID IID_ISystemInfo = { 0x32bb8320, 0xb41b, 0x11cf,{0xa6, 0xbb, 0x0, 0x80, 0xc7, 0xb2, 0xd6, 0x82} };
}*/
class CSystemInfo :public ISystemInfo
{
public:
    CSystemInfo();
    ~CSystemInfo();
    STDMETHOD( QueryInterface( REFIID, void** ) );
    STDMETHOD_( ULONG, AddRef() );
    STDMETHOD_( ULONG, Release() );

    STDMETHOD( GetOS( CString* info ) );
    STDMETHOD( GetMBoardCreator(CString* info) );
    STDMETHOD( GetCPUINFO(CString* info) );
    STDMETHOD( MonitorInfo(CString* info) );
private:
    long m_lRef;
    STDMETHODIMP GetInfo( CString className,CString propertyName,CString*info );
};