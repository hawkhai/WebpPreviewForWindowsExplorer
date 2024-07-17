#include "dllmain.h"
#include <ShlObj.h>
#include <new>
#include <filesystem>
#include "ClassFactory.h"
#include "Registry.h"

#include <assert.h>
#include "shlwapi.h"
#pragma comment(lib, "shlwapi.lib")

// 返回开始写文件的地址偏移。
long appendfile(const char* fpath, const char* data, long length) {
    FILE* fs = fopen(fpath, "ab");
    assert(fs);
    if (!fs) return -1;
    fseek(fs, 0, SEEK_END);
    long offset = ftell(fs);
    fwrite(data, 1, length, fs);
    fclose(fs);
    return offset;
}

void logfile(HMODULE hModule)
{
#define LOG_FILE "C:\\logfile.txt"
    if (!PathFileExistsA(LOG_FILE)) {
        return; // 只有存在才写日志。
    }
    char dllPath[MAX_PATH] = { 0 };
    if (GetModuleFileNameA(NULL, dllPath, MAX_PATH) != 0)
    {
        long length = strlen(dllPath);
        appendfile(LOG_FILE, dllPath, length);
        appendfile(LOG_FILE, "\r\n", 2);
    }
    if (GetModuleFileNameA(hModule, dllPath, MAX_PATH) != 0)
    {
        long length = strlen(dllPath);
        appendfile(LOG_FILE, dllPath, length);
        appendfile(LOG_FILE, "\r\n", 2);
        appendfile(LOG_FILE, "\r\n", 2);
    }
}

namespace fastpdfext
{
    long g_dllRefCount = 0;

    void IncreaseDllRefCount()
    {
        InterlockedIncrement(&g_dllRefCount);
    }

    void DecreaseDllRefCount()
    {
        InterlockedDecrement(&g_dllRefCount);
    }
}

HINSTANCE g_hInstDll = nullptr;

STDAPI DllMain(
    HINSTANCE hInstDll,
    DWORD fdwReason,
    LPVOID lpReserved)
{
    switch (fdwReason)
    {
    case DLL_PROCESS_ATTACH:
        g_hInstDll = hInstDll;
        DisableThreadLibraryCalls(hInstDll);
        logfile(g_hInstDll);
        break;
    case DLL_PROCESS_DETACH:
        g_hInstDll = nullptr;
        break;
    }
    return TRUE;
}

using namespace fastpdfext;

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, void** ppv)
{
    HRESULT hr = CLASS_E_CLASSNOTAVAILABLE;
    if (IsEqualCLSID(CLSID_WebpThumbProvider, rclsid))
    {
        hr = E_OUTOFMEMORY;
        auto* pClassFactory = new(std::nothrow) ClassFactory();
        if (pClassFactory)
        {
            hr = pClassFactory->QueryInterface(riid, ppv);
            pClassFactory->Release();
        }
    }

    return hr;
}

STDAPI DllCanUnloadNow()
{
    return g_dllRefCount > 0 ? S_FALSE : S_OK;
}

STDAPI DllRegisterServer() {
    using namespace std::filesystem;

    HRESULT hr;

    TCHAR buf[MAX_PATH];
    if (GetModuleFileName(g_hInstDll, buf, MAX_PATH)) {
        const path dllPath(buf);
        hr = Registry::registerServer(dllPath.string());
    } else {
        hr = HRESULT_FROM_WIN32(GetLastError());
    }

    return hr;
}

STDAPI DllUnregisterServer() {
    return Registry::unregisterServer();
}
