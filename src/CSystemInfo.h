#pragma once

#include <objbase.h>
#include <INITGUID.H>
#include <iostream>
#include "IGetCPUInfo.h"
#include "IGetMonitorInfo.h"
#include "IGetOS.h"
#include "IMotherBoard.h"
extern long g_lObjs;
extern long g_lLocks;
extern HMODULE g_hModule;


// {9004C87E-F422-4DA0-B18B-BD48C81AB9BF}
DEFINE_GUID(CLSID_SystemInfo, 
             0x9004c87e, 0xf422, 0x4da0, 0xb1, 0x8b, 0xbd, 0x48, 0xc8, 0x1a, 0xb9, 0xbf);
// {2E91CA67-C4CC-4AF8-9236-7EADBD9317E8}
DEFINE_GUID(IID_ISystemInfo, 
             0x2e91ca67, 0xc4cc, 0x4af8, 0x92, 0x36, 0x7e, 0xad, 0xbd, 0x93, 0x17, 0xe8);


/*extern "C"
{
   const IID IID_ISystemInfo = { 0x32bb8320, 0xb41b, 0x11cf,{0xa6, 0xbb, 0x0, 0x80, 0xc7, 0xb2, 0xd6, 0x82} };
}*/
class CSystemInfo :public IMotherBoard,public ICPUInfo, public IMonitorInfo, public IOperationSystem
{
public:
    CSystemInfo();
    ~CSystemInfo();
    STDMETHOD( QueryInterface( REFIID, void** ) );
    STDMETHOD_( ULONG, AddRef() );
    STDMETHOD_( ULONG, Release() );

    STDMETHOD( GetOS( CString* info ) );
    STDMETHOD( GetMBoardCreator(CString* info) );
    STDMETHOD( GetCPUINFO(UINT* clocks,UINT *frequency) );
    STDMETHOD( MonitorInfo(CString* info,int *) );
private:
    long m_lRef;
    STDMETHODIMP GetInfo( CString className,CString propertyName,CString*info );
    STDMETHODIMP GetInfoUINT( CString className, CString propertyName, uint32_t* info );
};