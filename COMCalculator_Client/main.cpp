#include <iostream>
#include "../COMCalculator/IObjectsCalculator.h"

int main(void)
{
    // Initializing COM
    std::cout << "Initializing COM" << std::endl;
    if (FAILED(::CoInitialize(NULL)))
    {
        ::CoUninitialize();
        std::cout << "Unable to initialize COM" << std::endl;
        return -1;
    }

    // Get CLSID from Name
    HRESULT hr = E_FAIL;
    CLSID clsid;
    hr = ::CLSIDFromProgID(g_versionName.c_str(), &clsid);
    if (FAILED(hr))
    {
        std::cout << "Unable to get CLSID from ProgID. HR = " << hr << std::endl;
        return -1;
    }

    // Generate class
    IClassFactory* pCF = NULL;
    hr = ::CoGetClassObject(clsid,
        CLSCTX_INPROC,
        NULL,
        IID_IClassFactory,
        (void**)&pCF);
    if (FAILED(hr))
    {
        std::cout << "Failed to GetClassObject server instance. HR = " << hr << std::endl;
        return -1;
    }

    ///////////////////////////////////////////////////////////
    IUnknown* pUnk = NULL;;
    hr = pCF->CreateInstance(NULL, IID_IUnknown, (void**)&pUnk);
    pCF->Release();
    if (FAILED(hr))
    {
        std::cout << "Failed to create server instance. HR = " << hr << std::endl;
        return -1;
    }
    std::cout << "Instance created" << std::endl;

    ///////////////////////////////////////////////////////////
    IObjectsCalculator* pObjectsCalculator = NULL;
    hr = pUnk->QueryInterface(IID_ObjectsCalculator, (LPVOID*)&pObjectsCalculator);
    pUnk->Release();
    if (FAILED(hr))
    {
        std::cout << "QueryInterface() for IObjectCalculator failed" << std::endl;
        return -1;
    }

    // Using COM functions
    if (pObjectsCalculator->CalculateObjects() == S_FALSE)
    {
        std::wcout << "Error: Something goes wrong with COM Server." << std::endl;
    }
    else
    {
        std::wcout << "Enumerating was ended successfully" << std::endl;
    }
    pObjectsCalculator->Release();
    
    // Unitialize COM
    std::cout << "Shuting down COM" << std::endl;
    ::CoUninitialize();

    return 0;
}