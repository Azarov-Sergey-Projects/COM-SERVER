#ifndef _REGISTRY_H_
#define _REGISTRY_H_
//
// Registry.h
//   - Helper functions registering and unregistering a component.
//

// This function will register a component in the Registry.
// The component calls this function from its DllRegisterServer function.
HRESULT RegisterServer( HMODULE hModule,
                        const CLSID& clsid,
                        const TCHAR* szFriendlyName,
                        const TCHAR* szVerIndProgID,
                        const TCHAR* szProgID );

// This function will unregister a component.  Components
// call this function from their DllUnregisterServer function.
HRESULT UnregisterServer( const CLSID& clsid,
                          const TCHAR* szVerIndProgID,
                          const TCHAR* szProgID );

#endif // !_REGISTRY_H_