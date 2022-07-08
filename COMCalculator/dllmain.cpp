// dllmain.cpp : Defines the entry point for the DLL application.
#include "pch.h"

#include "RegisterServer.h"
#include "ObjectsCalculator.h"

static HMODULE g_hModule;

BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved
                     )
{
    switch (ul_reason_for_call)
    {
    case DLL_PROCESS_ATTACH:
        g_hModule = hModule;
    case DLL_THREAD_ATTACH:
    case DLL_THREAD_DETACH:
    case DLL_PROCESS_DETACH:
        break;
    }
    return TRUE;
}

//////////////////////////////////////////////////////////////////
/* COM Calculator DLL functions */
long g_lObjs = 0;
long g_lLocks = 0;

// REGISTER COM SERVER
STDAPI DllRegisterServer(void)
{
    return RegisterCOMServer(g_GUID, g_hModule, g_versionName, g_name, L"Calculate COM objects in the system");
}

STDAPI DllUnregisterServer(void)
{
    return UnregisterCOMServer(g_GUID, g_versionName, g_name);
}

// CREATE INSTANCE
STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, void** ppv)
{
    // Make sure the CLSID is for our Expression component
    if (rclsid != CLSID_ObjectsCalculator)
    {
        return(E_FAIL);
    }

    ObjectsCalculatorClassFactory* pCF = 0;
    pCF = new ObjectsCalculatorClassFactory;
    if (pCF == 0)
    {
        return(E_OUTOFMEMORY);
    }

    const HRESULT hr = pCF->QueryInterface(riid, ppv);
    if (FAILED(hr))
    {
        delete pCF;
        pCF = 0;
    }

    return hr;
}

STDAPI DllCanUnloadNow(void)
{
    if (g_lObjs || g_lLocks) return(S_FALSE);
    else return(S_OK);
}
