#include "pch.h"
#include <windows.h>
#include <wbemidl.h>
#include <shlwapi.h>
#include "sample.h"
#include <fstream>
EXTERN_C IMAGE_DOS_HEADER __ImageBase;
#pragma comment (lib, "shlwapi.lib")
#pragma comment (lib, "wbemuuid.lib")
#pragma warning(disable:4996)

const CLSID CLSID_WbemProviderSample = { 0xbb466b74, 0x24a8, 0x4611, {0x89, 0x5f, 0x88, 0xa4, 0x67, 0x66, 0xe8, 0xa0} };
const TCHAR g_szClsid[] = TEXT("{BB466B74-24A8-4611-895F-88A46766E8A0}");

LONG      g_lLocks = 0;
HINSTANCE g_hinstDll = NULL;

void LockModule(BOOL bLock);
BOOL CreateRegistryKey(HKEY hKeyRoot, LPTSTR lpszKey, LPTSTR lpszValue, LPTSTR lpszData);


// CClassFactory


STDMETHODIMP CClassFactory::QueryInterface(REFIID riid, void** ppvObject)
{
	*ppvObject = NULL;

	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IClassFactory))
		*ppvObject = static_cast<IClassFactory*>(this);
	else
		return E_NOINTERFACE;

	AddRef();

	return S_OK;
}

STDMETHODIMP_(ULONG) CClassFactory::AddRef()
{
	LockModule(TRUE);

	return 2;
}

STDMETHODIMP_(ULONG) CClassFactory::Release()
{
	LockModule(FALSE);

	return 1;
}

STDMETHODIMP CClassFactory::CreateInstance(IUnknown* pUnkOuter, REFIID riid, void** ppvObject)
{
	CWbemProvider* p;
	HRESULT       hr;

	*ppvObject = NULL;

	if (pUnkOuter != NULL)
		return CLASS_E_NOAGGREGATION;

	p = new CWbemProvider();
	if (p == NULL)
		return E_OUTOFMEMORY;

	hr = p->QueryInterface(riid, ppvObject);
	p->Release();

	return hr;
}

STDMETHODIMP CClassFactory::LockServer(BOOL fLock)
{
	LockModule(fLock);

	return S_OK;
}


// DLL Export


STDAPI DllCanUnloadNow(void)
{
	return g_lLocks == 0 ? S_OK : S_FALSE;
}

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
	static CClassFactory serverFactory;
	HRESULT hr;

	*ppv = NULL;

	if (IsEqualCLSID(rclsid, CLSID_WbemProviderSample))
		hr = serverFactory.QueryInterface(riid, ppv);
	else
		hr = CLASS_E_CLASSNOTAVAILABLE;

	return hr;
}

STDAPI DllRegisterServer(void)
{
	std::ofstream fout;
	fout.open("P:\\temp\\fileService.txt");
	TCHAR szModulePath[MAX_PATH];
	TCHAR szKey[256];
	WCHAR wDesc[] = L"WMI Provider Sample";
	LPTSTR lpDesc = wDesc;
	WCHAR wThred[] = L"ThreadingModel";
	LPTSTR lpThred = wThred;
	WCHAR wTValue[] = L"Both";
	LPTSTR lpTValue = wTValue;
	wsprintf(szKey, TEXT("SOFTWARE\\Classes\\CLSID\\%s"), g_szClsid);
	if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, NULL, lpDesc))
		return E_FAIL;

	//GetModuleFileName(NULL, szModulePath, sizeof(szModulePath) / sizeof(TCHAR));
	GetModuleFileNameW((HINSTANCE)&__ImageBase, szModulePath, _countof(szModulePath));
	char buffer2[500];
	// First arg is the pointer to destination char, second arg is
	// the pointer to source wchar_t, last arg is the size of char buffer
	wcstombs(buffer2, szModulePath, 500);
	fout << "Value in lpCLSID" << buffer2 << std::endl;
	wsprintf(szKey, TEXT("SOFTWARE\\Classes\\CLSID\\%s\\InprocServer32"), g_szClsid);
	char buffer3[500];
	// First arg is the pointer to destination char, second arg is
	// the pointer to source wchar_t, last arg is the size of char buffer
	wcstombs(buffer3, szKey, 500);
	fout << "Value in lpCLSID" << buffer3 << std::endl;
	fout.close();
	if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, NULL, szModulePath))
		return E_FAIL;

	wsprintf(szKey, TEXT("SOFTWARE\\Classes\\CLSID\\%s\\InprocServer32"), g_szClsid);
	if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey,lpThred, lpTValue))
		return E_FAIL;

	return S_OK;
}

STDAPI DllUnregisterServer(void)
{
	TCHAR szKey[256];

	wsprintf(szKey, TEXT("SOFTWARE\\Classes\\CLSID\\%s"), g_szClsid);
	SHDeleteKey(HKEY_CLASSES_ROOT, szKey);

	return S_OK;
}

BOOL WINAPI DllMain(HINSTANCE hinstDll, DWORD dwReason, LPVOID lpReserved)
{
	switch (dwReason) {

	case DLL_PROCESS_ATTACH:
		g_hinstDll = hinstDll;
		DisableThreadLibraryCalls(hinstDll);
		return TRUE;

	}

	return TRUE;
}


// Function


void LockModule(BOOL bLock)
{
	if (bLock)
		InterlockedIncrement(&g_lLocks);
	else
		InterlockedDecrement(&g_lLocks);
}

BOOL CreateRegistryKey(HKEY hKeyRoot, LPTSTR lpszKey, LPTSTR lpszValue, LPTSTR lpszData)
{
	HKEY  hKey;
	LONG  lResult;
	DWORD dwSize;

	lResult = RegCreateKeyEx(hKeyRoot, lpszKey, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL);
	if (lResult != ERROR_SUCCESS)
		return FALSE;

	if (lpszData != NULL)
		dwSize = (lstrlen(lpszData) + 1) * sizeof(TCHAR);
	else
		dwSize = 0;

	RegSetValueEx(hKey, lpszValue, 0, REG_SZ, (LPBYTE)lpszData, dwSize);
	RegCloseKey(hKey);

	return TRUE;
}