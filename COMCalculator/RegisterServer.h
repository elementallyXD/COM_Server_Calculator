#pragma once
#include "pch.h"
#include <OleCtl.h> // SELFREG_E_CLASS

#define CHECK(a) if (a != ERROR_SUCCESS) return SELFREG_E_CLASS;

const HRESULT RegisterCOMServer(const std::wstring& _clsid, const HMODULE hModule, const std::wstring& _versionName, const std::wstring& _name, const std::wstring& _description);
const HRESULT UnregisterCOMServer(const std::wstring& _clsid, const std::wstring& _versionName, const std::wstring& _name);