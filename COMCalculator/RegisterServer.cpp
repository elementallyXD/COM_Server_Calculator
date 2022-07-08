#include "pch.h"
#include "RegisterServer.h"

const HRESULT RegisterCOMServer(const std::wstring& _clsid, const HMODULE _hModule, const std::wstring& _versionName, const std::wstring& _name, const std::wstring& _description)
{
    WCHAR szCOMPath[MAX_PATH + 1] = { 0 };
    if (GetModuleFileNameW(_hModule, szCOMPath, MAX_PATH) <= ERROR_SUCCESS)
        return ERROR_SUCCESS;

    HKEY hKey1, hKey2, hKey3;
    //////////////////////////////////////
    CHECK(RegCreateKeyW(HKEY_CLASSES_ROOT, _versionName.c_str(), &hKey1));
    CHECK(RegSetValueW(hKey1, NULL, REG_SZ, _description.c_str(), 15));

    CHECK(RegCreateKeyW(hKey1, L"CLSID", &hKey2));
    CHECK(RegSetValueW(hKey2, NULL, REG_SZ, _clsid.c_str(), 34));

    RegFlushKey(hKey1);
    RegFlushKey(hKey2);
    RegCloseKey(hKey1);
    RegCloseKey(hKey2);

    //////////////////////////////////////
    CHECK(RegCreateKeyW(HKEY_CLASSES_ROOT, _name.c_str(), &hKey1));
    CHECK(RegSetValueW(hKey1, NULL, REG_SZ, _description.c_str(), 30));

    CHECK(RegCreateKeyW(hKey1, L"CLSID", &hKey2));
    CHECK(RegSetValueW(hKey2, NULL, REG_SZ, _clsid.c_str(), 68));

    CHECK(RegCreateKeyW(hKey1, L"CurVer", &hKey2));
    CHECK(RegSetValueW(hKey2, NULL, REG_SZ, _versionName.c_str(), 38));

    RegFlushKey(hKey1);
    RegFlushKey(hKey2);
    RegCloseKey(hKey1);
    RegCloseKey(hKey2);

    //////////////////////////////////////
    /* CLSID\{CLSID} and sub keys */
    CHECK(RegOpenKeyW(HKEY_CLASSES_ROOT, L"CLSID", &hKey1));
    CHECK(RegCreateKeyW(hKey1, _clsid.c_str(), &hKey2));
    CHECK(RegSetValueW(hKey2, NULL, REG_SZ, _description.c_str(), 30));

    CHECK(RegCreateKeyW(hKey2, L"ProgID", &hKey3));
    CHECK(RegSetValueW(hKey3, NULL, REG_SZ, _versionName.c_str(), 38));

    RegFlushKey(hKey3);
    RegCloseKey(hKey3);

    //////////////////////////////////////
    CHECK(RegCreateKeyW(hKey2, L"VersionIndependentProgID", &hKey3));
    CHECK(RegSetValueW(hKey3, NULL, REG_SZ, _name.c_str(), 34));

    RegFlushKey(hKey3);
    RegCloseKey(hKey3);

    //////////////////////////////////////
    CHECK(RegCreateKeyW(hKey2, L"Programmable", &hKey3));

    RegFlushKey(hKey3);
    RegCloseKey(hKey3);

    //////////////////////////////////////
    CHECK(RegCreateKeyW(hKey2, L"InprocServer32", &hKey3));
    CHECK(RegSetValueW(hKey3, NULL, REG_SZ, szCOMPath, static_cast<DWORD>(wcslen(szCOMPath) * 2)));
    CHECK(RegSetValueExW(hKey3, L"ThreadingModel", 0, REG_SZ, (const unsigned char*)L"Apartment", 18));

    RegFlushKey(hKey3);
    RegCloseKey(hKey3);

    //////////////////////////////////////
    CHECK(RegCreateKeyW(hKey2, L"TypeLib", &hKey3));
    CHECK(RegSetValueW(hKey3, NULL, REG_SZ, _clsid.c_str(), 68));

    RegFlushKey(hKey3);
    RegCloseKey(hKey3);
    RegFlushKey(hKey2);
    RegCloseKey(hKey2);
    RegFlushKey(hKey1);
    RegCloseKey(hKey1);

    return S_OK;
}

const HRESULT UnregisterCOMServer(const std::wstring& _clsid, const std::wstring& _versionName, const std::wstring& _name)
{
    HKEY hKey1, hKey2;

    //////////////////////////////////////
    CHECK(RegOpenKeyExW(HKEY_CLASSES_ROOT, _versionName.c_str(), 0, KEY_ALL_ACCESS, &hKey1));
    CHECK(RegDeleteKeyW(hKey1, L"CLSID"));
    RegCloseKey(hKey1);
    CHECK(RegDeleteKeyW(HKEY_CLASSES_ROOT, _versionName.c_str()));

    //////////////////////////////////////
    CHECK(RegOpenKeyExW(HKEY_CLASSES_ROOT, _name.c_str(), 0, KEY_ALL_ACCESS, &hKey1));
    CHECK(RegDeleteKeyW(hKey1, L"CLSID"));
    CHECK(RegDeleteKeyW(hKey1, L"CurVer"));
    RegCloseKey(hKey1);
    CHECK(RegDeleteKeyW(HKEY_CLASSES_ROOT, _name.c_str()));

    ///////////////////////////////////
    CHECK(RegOpenKeyExW(HKEY_CLASSES_ROOT, L"CLSID", 0, KEY_ALL_ACCESS, &hKey1));
    CHECK(RegOpenKeyExW(hKey1, _clsid.c_str(), 0, KEY_ALL_ACCESS, &hKey2));
    CHECK(RegDeleteKeyW(hKey2, L"ProgID"));
    CHECK(RegDeleteKeyW(hKey2, L"VersionIndependentProgID"));
    CHECK(RegDeleteKeyW(hKey2, L"Programmable"));
    CHECK(RegDeleteKeyW(hKey2, L"InprocServer32"));
    CHECK(RegDeleteKeyW(hKey2, L"TypeLib"));
    RegCloseKey(hKey2);
    CHECK(RegDeleteKeyW(hKey1, _clsid.c_str()));
    RegCloseKey(hKey1);

    return S_OK;
}