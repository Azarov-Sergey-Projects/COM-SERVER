#include <objbase.h>
#include <iostream>


#include "CSystemInfoFactory.h"



int main()
{
    CoInitialize( NULL );
    CSystemInfoFactory* cSystem = NULL;
    HRESULT hr = CoCreateInstance( IID_ISystemInfo, NULL, CLSCTX_INPROC_SERVER, IID_ISystemInfo, ( void** )&cSystem );
    if( SUCCEEDED( hr ) )
    {
        std::cout << "SUCCEED" << std::endl;
    }
    return 0;
}