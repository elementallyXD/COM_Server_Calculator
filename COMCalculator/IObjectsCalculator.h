#pragma once
//
// IObjectsCalculator.h
//
#include "pch.h"
#include <atlbase.h> // CComBSTR

template<class S>
CLSID CreateCLSID(const S& hexString)
{
    CLSID clsid;
    CLSIDFromString(CComBSTR(hexString), &clsid);
    return clsid;
}

const std::wstring g_GUID = L"{8FA89B97-BF0C-4A56-9428-07CD71A7ACFF}"; // Randomly generated GUID
const CLSID CLSID_ObjectsCalculator = CreateCLSID(g_GUID.c_str());
const CLSID IID_ObjectsCalculator = CLSID_ObjectsCalculator;

const std::wstring g_name = L"COMObjectLib.Calculator";
const std::wstring g_versionName = L"COMObjectLib.Calculator.1";

class IObjectsCalculator : public IUnknown
{
public:
    STDMETHOD(CalculateObjects(std::wstring, std::wstring*))  PURE;
};
