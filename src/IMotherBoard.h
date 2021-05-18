#pragma once
#include <Unknwn.h>
#include <atlstr.h>

class IMotherBoard: public IUnknown
{
public:
    STDMETHOD( getMBoardCreator(CString* ) )PURE;
};