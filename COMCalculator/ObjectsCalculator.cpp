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
STDMETHODIMP_(HRESULT __stdcall) ObjectsCalculator::CalculateObjects(std::wstring _inBuffer, std::wstring* _outBuffer)
{
    *_outBuffer = _inBuffer + L" ADDED BY COM CALCULATOR";
    return S_OK;
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
