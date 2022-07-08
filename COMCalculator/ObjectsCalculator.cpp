#include "pch.h"
#include "ObjectsCalculator.h"

/////////////////////////////////////////////////////////////////////////
/* ObjectsCalculator */
ObjectsCalculator::ObjectsCalculator()
{
    m_lRef = 0;
    InterlockedIncrement(&g_lObjs);
}

ObjectsCalculator::~ObjectsCalculator()
{
    InterlockedDecrement(&g_lObjs);
}

/////////////////////////////////////////////////////////////////////////
/* ObjectsCalculator::IUnknown */
STDMETHODIMP_(HRESULT __stdcall) ObjectsCalculator::QueryInterface(REFIID riid, void** ppv)
{
    *ppv = 0;

    if (riid == IID_IUnknown || riid == IID_ObjectsCalculator)
    {
        *ppv = this;
    }

    if (*ppv)
    {
        AddRef();
        return(S_OK);
    }
    return (E_NOINTERFACE);
}

STDMETHODIMP_(ULONG __stdcall) ObjectsCalculator::AddRef()
{
    return InterlockedIncrement(&m_lRef);
}

STDMETHODIMP_(ULONG __stdcall) ObjectsCalculator::Release()
{
    if (InterlockedDecrement(&m_lRef) == 0)
    {
        delete this;
        return 0;
    }

    return m_lRef;
}

/////////////////////////////////////////////////////////////////////////
/* ObjectsCalculator::ObjectsCalculator */
STDMETHODIMP_(HRESULT __stdcall) ObjectsCalculator::CalculateObjects()
{
    HKEY hKey = nullptr;
    if (RegOpenKeyExW(HKEY_CLASSES_ROOT, L"CLSID", 0, KEY_READ, &hKey) != ERROR_SUCCESS)
    {
        RegCloseKey(hKey);
        return S_FALSE;
    }
    
    EnumerateRegKey(hKey);
    RegCloseKey(hKey);
   
    return S_OK;
}

void ObjectsCalculator::EnumerateRegKey(HKEY hKey) const
{
    static const DWORD REG_MAX_KEY_LENGTH = 255;
    DWORD    cSubKeys = 0;               // number of subkeys 
    DWORD    cbMaxSubKey;              // longest subkey size 
    DWORD    cchMaxClass;              // longest class string 
    FILETIME ftLastWriteTime;      // last write time 
  
    DWORD retCode = -1;
    // Get the value count. 
    retCode = RegQueryInfoKeyW(
        hKey,                              // key handle 
        NULL,                       // buffer for class name 
        NULL,                     // size of class string 
        NULL,                     // reserved 
        &cSubKeys,                // number of subkeys 
        &cbMaxSubKey,            // longest subkey size 
        &cchMaxClass,            // longest class string 
        NULL,                // number of values for this key 
        NULL,            // longest value name 
        NULL,         // longest value data 
        NULL,   // security descriptor 
        &ftLastWriteTime);       // last write time
    if (retCode != ERROR_SUCCESS)
    {
        std::cout << "RegQueryInfoKeyW is failed. Error: " << ::GetLastError() << std::endl;
        return;
    }

    // Enumerate the subkeys, until RegEnumKeyEx fails.
    if (cSubKeys)
    {
        TCHAR    achKey[REG_MAX_KEY_LENGTH] = { 0 }; // buffer for subkey name
        DWORD    cbName;                             // size of name string 
        TCHAR    achKeyClass[REG_MAX_KEY_LENGTH] = { 0 }; // buffer for subkey class name
        DWORD    cbClassName;                             // size of class name string 
       
        std::cout << "Number of subkeys: " << cSubKeys << std::endl;
        
        DWORD i = -1;
        for (i = 0; i < cSubKeys; i++)
        {
            cbName = REG_MAX_KEY_LENGTH;
            cbClassName = REG_MAX_KEY_LENGTH;
            retCode = RegEnumKeyExW(hKey, 
                i,
                achKey,
                &cbName,
                NULL,
                achKeyClass,
                &cbClassName,
                &ftLastWriteTime);
            if (retCode == ERROR_SUCCESS)
            {
                std::wcout<< "[" << i + 1 << "] Key: " << achKey;
                if (cbClassName > 0) std::wcout << " Class: " << achKeyClass;
                std::cout << std::endl;
            }
            else if (retCode == ERROR_MORE_DATA)
            {
                std::wcout << "[" << i + 1 << "] " << "Buffer is small. More data is available." << std::endl;
            }
        }
    }
    else
    {
        std::cout << "Did not find any subkeys in CLSID"<< std::endl;
    }
}

/////////////////////////////////////////////////////////////////////////
/* ObjectsCalculatorClassFactory */
ObjectsCalculatorClassFactory::ObjectsCalculatorClassFactory()
{
    m_lRef = 0;
}

ObjectsCalculatorClassFactory::~ObjectsCalculatorClassFactory()
{
}

/* ObjectsCalculatorClassFactory::IUnknown */
STDMETHODIMP_(HRESULT __stdcall) ObjectsCalculatorClassFactory::QueryInterface(REFIID riid, void** ppv)
{
    *ppv = 0;

    if (riid == IID_IUnknown || riid == IID_IClassFactory)
    {
        *ppv = this;
    }

    if (*ppv)
    {
        AddRef();
        return S_OK;
    }

    return(E_NOINTERFACE);
}

STDMETHODIMP_(ULONG __stdcall) ObjectsCalculatorClassFactory::AddRef()
{
    return InterlockedIncrement(&m_lRef);
}

STDMETHODIMP_(ULONG __stdcall) ObjectsCalculatorClassFactory::Release()
{
    if (InterlockedDecrement(&m_lRef) == 0)
    {
        delete this;
        return 0;
    }

    return m_lRef;
}

/////////////////////////////////////////////////////////////////////////
/* ObjectsCalculatorClassFactory::IClassFactory */
STDMETHODIMP_(HRESULT __stdcall) ObjectsCalculatorClassFactory::CreateInstance(LPUNKNOWN pUnkOuter, REFIID riid, void** ppvObj)
{
    ObjectsCalculator* pObjectsCalculator;
    pObjectsCalculator = new ObjectsCalculator;
    if (pObjectsCalculator == 0) return(E_OUTOFMEMORY);

    HRESULT hr;
    *ppvObj = 0;
    hr = pObjectsCalculator->QueryInterface(riid, ppvObj);
    if (FAILED(hr)) delete pObjectsCalculator;

    return hr;
}

STDMETHODIMP_(HRESULT __stdcall) ObjectsCalculatorClassFactory::LockServer(BOOL fLock)
{
    if (fLock) InterlockedIncrement(&g_lLocks);
    else InterlockedDecrement(&g_lLocks);

    return S_OK;
}
