#pragma once
#include <Unknwn.h>
#include <atlstr.h>
class IMonitorInfo: public IUnknown
{
public:
    STDMETHOD( MonitorInfo(CString*,int *,std::vector<uint32_t>*,std::vector<uint32_t>* ) )PURE;
};