#pragma once
#include <Unknwn.h>
#include <atlstr.h>
class IMonitorInfo: public IUnknown
{
public:
    STDMETHOD( MonitorInfo(CString* ) )PURE;
};