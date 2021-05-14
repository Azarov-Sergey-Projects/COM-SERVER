#pragma once
#include <Unknwn.h>
#include <atlstr.h>
class ICPUInfo: public IUnknown
{
public:
    STDMETHOD( GetCPUINFO(UINT* clocks,UINT *frequency ) ) PURE;
};