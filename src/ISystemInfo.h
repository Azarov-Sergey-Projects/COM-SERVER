#pragma once
#include <Unknwn.h>
#include <atlstr.h>
#include "pch.h"
class ISystemInfo :public IUnknown
{
public:
    STDMETHOD( GetOS(CString* SystemInfo) ) PURE;
    STDMETHOD( GetMBoardCreator(CString* ) )PURE;
    STDMETHOD( GetCPUINFO(CString* ) ) PURE;
    STDMETHOD( MonitorInfo(CString* ) )PURE;
};