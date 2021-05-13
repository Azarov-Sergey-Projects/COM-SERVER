#pragma once
#include <Unknwn.h>
#include <atlstr.h>
class IOperationSystem: public IUnknown
{
public:
    STDMETHOD( GetOS(CString* SystemInfo) ) PURE;
};