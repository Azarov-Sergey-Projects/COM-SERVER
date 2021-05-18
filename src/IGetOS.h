#pragma once
#include <Unknwn.h>
#include <atlstr.h>
class IOperationSystem: public IUnknown
{
public:
    STDMETHOD( getOS(CString* SystemInfo) ) PURE;
};