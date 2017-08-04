#include "stdafx.h"
#include "Sysutils.h"
#include <atlbase.h>
#include <atlstr.h>
#include <iostream>
#include <windows.h>
#include <tchar.h>
#include <comdef.h>
#include <Wbemidl.h>
#include <vector>
#include <fstream>
#include <winsock2.h>
#include <iphlpapi.h>
#include <stdio.h>
#include <stdlib.h>

#pragma comment(lib, "wbemuuid.lib")

using namespace std;

typedef BOOL(WINAPI *LPFN_ISWOW64PROCESS) (HANDLE, PBOOL);

LPFN_ISWOW64PROCESS fnIsWow64Process;


#define WORKING_BUFFER_SIZE 15000
#define MAX_TRIES 3

#define MALLOC(x) HeapAlloc(GetProcessHeap(), 0, (x))
#define FREE(x) HeapFree(GetProcessHeap(), 0, (x))

Sysutils::Sysutils(): gBitness(0)
{
}

CString Sysutils::GetHMS32WOW64VersionInfo()
{
	LPCWSTR pszValueName = L"DisplayName";
	HKEY hKey = nullptr;

		LPCTSTR pszSubkey = L"SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\hMailServer_is1";

	if (RegOpenKey(HKEY_LOCAL_MACHINE, pszSubkey, &hKey) != ERROR_SUCCESS) {
		AtlThrowLastWin32();
	}

	// Buffer to store string read from registry
	TCHAR szValue[1024] = { 0 };
	DWORD cbValueLength = sizeof(szValue);

	// Query string value
	if (RegQueryValueEx(
		hKey,
		pszValueName,
		nullptr,
		nullptr,
		reinterpret_cast<LPBYTE>(&szValue),
		&cbValueLength)
		!= ERROR_SUCCESS) {
		AtlThrowLastWin32();
	}
	return CString(szValue);
}

CString Sysutils::GetHMS64VersionInfo()
{
	LPCWSTR pszValueName = L"DisplayName";
	HKEY hKey = nullptr;

	LPCTSTR pszSubkey = L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\hMailServer_is1";

	if (RegOpenKey(HKEY_LOCAL_MACHINE, pszSubkey, &hKey) != ERROR_SUCCESS) {
		AtlThrowLastWin32();
	}

	// Buffer to store string read from registry
	TCHAR szValue[1024] = { 0 };
	DWORD cbValueLength = sizeof(szValue);

	// Query string value
	if (RegQueryValueEx(
		hKey,
		pszValueName,
		nullptr,
		nullptr,
		reinterpret_cast<LPBYTE>(&szValue),
		&cbValueLength)
		!= ERROR_SUCCESS) {
		AtlThrowLastWin32();
	}
	return CString(szValue);
}

int Sysutils::GetTotalAvaiableMemory()
{	
	HRESULT hres;
	CString result;

	hres = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
	if (FAILED(hres)) {	}

	hres = CoInitializeSecurity(
		nullptr,
		-1, 
		nullptr, 
		nullptr, 
		RPC_C_AUTHN_LEVEL_DEFAULT, 
		RPC_C_IMP_LEVEL_IMPERSONATE, 
		nullptr,  
		EOAC_NONE,
		nullptr
	);


	if (FAILED(hres)) {
		CoUninitialize();
	}

	IWbemLocator *pLoc = nullptr;

	hres = CoCreateInstance(
		CLSID_WbemLocator,
		nullptr,
		CLSCTX_INPROC_SERVER,
		IID_IWbemLocator, (LPVOID *)&pLoc);

	if (FAILED(hres)) {
		CoUninitialize();
	}

	IWbemServices *pSvc = nullptr;	
	hres = pLoc->ConnectServer(
		_bstr_t(L"ROOT\\CIMV2"),
		nullptr,
		nullptr,
		nullptr,
		NULL,
		nullptr, 
		nullptr,  
		&pSvc  
	);

	if (FAILED(hres)) {
		pLoc->Release();
		CoUninitialize();
	}	
	
	// Set security flags
	hres = CoSetProxyBlanket(
		pSvc,
		RPC_C_AUTHN_WINNT,
		RPC_C_AUTHZ_NONE,
		nullptr, 
		RPC_C_AUTHN_LEVEL_CALL, 
		RPC_C_IMP_LEVEL_IMPERSONATE,
		nullptr,  
		EOAC_NONE 
	);

	if (FAILED(hres)) {
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();		
	}

	IEnumWbemClassObject* pEnumerator = nullptr;
	hres = pSvc->ExecQuery(
		bstr_t("WQL"),
		bstr_t("SELECT * FROM Win32_OperatingSystem"),
		WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
		nullptr,
		&pEnumerator);

	if (FAILED(hres)) {
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
	}

	// Get the data 
	IWbemClassObject *pclsObj = nullptr;
	ULONG uReturn = 0;	
	
	HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);
	if (SUCCEEDED(hr)) {

		VARIANT vtProp;

		// Get the value of the Name property
		pclsObj->Get(L"TotalVisibleMemorySize", 0, &vtProp, nullptr, nullptr);
		result = vtProp.bstrVal;
		VariantClear(&vtProp);
		pclsObj->Release();

		// Cleanup
		pSvc->Release();
		pLoc->Release();
		pEnumerator->Release();
		CoUninitialize();
	}
	
	int totalmemory = _wtoi(result);	
	return totalmemory;	
}

int Sysutils::GetFreeMemory()
{
	HRESULT hres;
	CString result;

	hres = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
	if (FAILED(hres)) {	}

	hres = CoInitializeSecurity(
		nullptr,
		-1,
		nullptr,  
		nullptr, 
		RPC_C_AUTHN_LEVEL_DEFAULT, 
		RPC_C_IMP_LEVEL_IMPERSONATE, 
		nullptr,
		EOAC_NONE,
		nullptr 
	);


	if (FAILED(hres)) {		
		CoUninitialize();
	}

	IWbemLocator *pLoc = nullptr;

	hres = CoCreateInstance(
		CLSID_WbemLocator,
		nullptr,
		CLSCTX_INPROC_SERVER,
		IID_IWbemLocator, (LPVOID *)&pLoc);

	if (FAILED(hres)) {
		CoUninitialize();
	}
	
	IWbemServices *pSvc = nullptr;
	hres = pLoc->ConnectServer(
		_bstr_t(L"ROOT\\CIMV2"), 
		nullptr, 
		nullptr, 
		nullptr,
		NULL, 
		nullptr,
		nullptr,
		&pSvc 
	);

	if (FAILED(hres)) 	{
		pLoc->Release();
		CoUninitialize();	
	}

	// Set security flags
	hres = CoSetProxyBlanket(
		pSvc,                     
		RPC_C_AUTHN_WINNT, 
		RPC_C_AUTHZ_NONE, 
		nullptr, 
		RPC_C_AUTHN_LEVEL_CALL, 
		RPC_C_IMP_LEVEL_IMPERSONATE, 
		nullptr, 
		EOAC_NONE 
	);

	if (FAILED(hres)) {
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
	}
		
	IEnumWbemClassObject* pEnumerator = nullptr;
	hres = pSvc->ExecQuery(
		bstr_t("WQL"),
		bstr_t("SELECT * FROM Win32_OperatingSystem"),
		WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
		nullptr,
		&pEnumerator);

	if (FAILED(hres)) {
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();		
	}

	// Get the data from the query 
	IWbemClassObject *pclsObj = nullptr;
	ULONG uReturn = 0;
	HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);

	if (SUCCEEDED(hr)) {

		VARIANT vtProp;

		// Get the value of the Name property
		pclsObj->Get(L"FreePhysicalMemory", 0, &vtProp, nullptr, nullptr);
		result = vtProp.bstrVal;

		VariantClear(&vtProp);
		pclsObj->Release();

		// Cleanup
		// ========

		pSvc->Release();
		pLoc->Release();
		pEnumerator->Release();
		CoUninitialize();
	}
	int totalmemory = _wtoi(result);

	return totalmemory;
}

// Get OS Info
CString Sysutils::GetOSDesc()
{
	HRESULT hres;

	hres = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
	if (FAILED(hres)) { }

	hres = CoInitializeSecurity(
		nullptr,
		-1,
		nullptr,
		nullptr, 
		RPC_C_AUTHN_LEVEL_DEFAULT,
		RPC_C_IMP_LEVEL_IMPERSONATE, 
		nullptr,
		EOAC_NONE,
		nullptr
	);
	
	if (FAILED(hres))	{
		CoUninitialize();
	}

	IWbemLocator *pLoc = nullptr;

	hres = CoCreateInstance(
		CLSID_WbemLocator,
		nullptr,
		CLSCTX_INPROC_SERVER,
		IID_IWbemLocator, (LPVOID *)&pLoc);

	if (FAILED(hres)) {
		CoUninitialize();		
	}

	IWbemServices *pSvc = nullptr;

	hres = pLoc->ConnectServer(
		_bstr_t(L"ROOT\\CIMV2"), // Object path of WMI namespace
		nullptr,                 // User name. NULL = current user
		nullptr,                 // User password. NULL = current
		nullptr,                 // Locale. NULL indicates current
		0,                       // Security flags.
		nullptr,                 // Authority (for example, Kerberos)
		nullptr,                 // Context object 
		&pSvc                    // pointer to IWbemServices proxy
	);

	if (FAILED(hres)) {
		pLoc->Release();
		CoUninitialize();		
	}

	hres = CoSetProxyBlanket(
		pSvc,                        // Indicates the proxy to set
		RPC_C_AUTHN_WINNT,           // RPC_C_AUTHN_xxx
		RPC_C_AUTHZ_NONE,            // RPC_C_AUTHZ_xxx
		nullptr,                     // Server principal name 
		RPC_C_AUTHN_LEVEL_CALL,      // RPC_C_AUTHN_LEVEL_xxx 
		RPC_C_IMP_LEVEL_IMPERSONATE, // RPC_C_IMP_LEVEL_xxx
		nullptr,                     // client identity
		EOAC_NONE                    // proxy capabilities 
	);

	if (FAILED(hres)) {
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
	}

	IEnumWbemClassObject* pEnumerator = nullptr;
	hres = pSvc->ExecQuery(
		bstr_t("WQL"),
		bstr_t("SELECT * FROM Win32_OperatingSystem"),
		WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
		nullptr,
		&pEnumerator);

	if (FAILED(hres)) {
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
	}

	IWbemClassObject *pclsObj = nullptr;
	ULONG uReturn = 0;

	pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);

		VARIANT vtProp;		
		pclsObj->Get(L"Caption", 0, &vtProp, nullptr, nullptr);		
		CString result = vtProp.bstrVal;
		VariantClear(&vtProp);
		pclsObj->Release();

	pSvc->Release();
	pLoc->Release();
	pEnumerator->Release();
	CoUninitialize();
	return result;
}

// Get CPU Infos
CString Sysutils::GetCPUInfo()
{
	HRESULT hres;

	hres = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
	if (FAILED(hres)) {}

	hres = CoInitializeSecurity(
		nullptr,
		-1,
		nullptr,
		nullptr,
		RPC_C_AUTHN_LEVEL_DEFAULT,
		RPC_C_IMP_LEVEL_IMPERSONATE,
		nullptr,
		EOAC_NONE,
		nullptr
	);

	if (FAILED(hres)) {
		CoUninitialize();
	}

	IWbemLocator *pLoc = nullptr;

	hres = CoCreateInstance(
		CLSID_WbemLocator,
		nullptr,
		CLSCTX_INPROC_SERVER,
		IID_IWbemLocator, (LPVOID *)&pLoc);

	if (FAILED(hres)) {
		CoUninitialize();
	}

	IWbemServices *pSvc = nullptr;

	hres = pLoc->ConnectServer(
		_bstr_t(L"ROOT\\CIMV2"), // Object path of WMI namespace
		nullptr,                 // User name. NULL = current user
		nullptr,                 // User password. NULL = current
		nullptr,                 // Locale. NULL indicates current
		0,                       // Security flags.
		nullptr,                 // Authority (for example, Kerberos)
		nullptr,                 // Context object 
		&pSvc                    // pointer to IWbemServices proxy
	);

	if (FAILED(hres)) {
		pLoc->Release();
		CoUninitialize();
	}

	hres = CoSetProxyBlanket(
		pSvc,                        // Indicates the proxy to set
		RPC_C_AUTHN_WINNT,           // RPC_C_AUTHN_xxx
		RPC_C_AUTHZ_NONE,            // RPC_C_AUTHZ_xxx
		nullptr,                     // Server principal name 
		RPC_C_AUTHN_LEVEL_CALL,      // RPC_C_AUTHN_LEVEL_xxx 
		RPC_C_IMP_LEVEL_IMPERSONATE, // RPC_C_IMP_LEVEL_xxx
		nullptr,                     // client identity
		EOAC_NONE                    // proxy capabilities 
	);

	if (FAILED(hres)) {
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
	}

	IEnumWbemClassObject* pEnumerator = nullptr;
	hres = pSvc->ExecQuery(
		bstr_t("WQL"),
		bstr_t("SELECT name, numberofcores FROM Win32_Processor"),
		WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
		nullptr,
		&pEnumerator);

	if (FAILED(hres)) {
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
	}

	IWbemClassObject *pclsObj = nullptr;

	ULONG uReturn = 0;

	pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);

	VARIANT vtProp;
	pclsObj->Get(L"Name", 0, &vtProp, nullptr, nullptr);
	
	CString result = vtProp.bstrVal;
	VariantClear(&vtProp);
	pclsObj->Release();


	pSvc->Release();
	pLoc->Release();
	pEnumerator->Release();
	CoUninitialize();
	return result;
}

// Get Ethernet NIC Info
CString Sysutils::GetNicInfo()
{
	HRESULT hres;

	hres = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
	if (FAILED(hres)) {}

	hres = CoInitializeSecurity(
		nullptr,
		-1,
		nullptr,
		nullptr,
		RPC_C_AUTHN_LEVEL_DEFAULT,
		RPC_C_IMP_LEVEL_IMPERSONATE,
		nullptr,
		EOAC_NONE,
		nullptr
	);

	if (FAILED(hres)) {
		CoUninitialize();
	}

	IWbemLocator *pLoc = nullptr;

	hres = CoCreateInstance(
		CLSID_WbemLocator,
		nullptr,
		CLSCTX_INPROC_SERVER,
		IID_IWbemLocator, (LPVOID *)&pLoc);

	if (FAILED(hres)) {
		CoUninitialize();
	}

	IWbemServices *pSvc = nullptr;

	hres = pLoc->ConnectServer(
		_bstr_t(L"ROOT\\CIMV2"), // Object path of WMI namespace
		nullptr,                 // User name. NULL = current user
		nullptr,                 // User password. NULL = current
		nullptr,                 // Locale. NULL indicates current
		0,                       // Security flags.
		nullptr,                 // Authority (for example, Kerberos)
		nullptr,                 // Context object 
		&pSvc                    // pointer to IWbemServices proxy
	);

	if (FAILED(hres)) {
		pLoc->Release();
		CoUninitialize();
	}

	hres = CoSetProxyBlanket(
		pSvc,                        // Indicates the proxy to set
		RPC_C_AUTHN_WINNT,           // RPC_C_AUTHN_xxx
		RPC_C_AUTHZ_NONE,            // RPC_C_AUTHZ_xxx
		nullptr,                     // Server principal name 
		RPC_C_AUTHN_LEVEL_CALL,      // RPC_C_AUTHN_LEVEL_xxx 
		RPC_C_IMP_LEVEL_IMPERSONATE, // RPC_C_IMP_LEVEL_xxx
		nullptr,                     // client identity
		EOAC_NONE                    // proxy capabilities 
	);

	if (FAILED(hres)) {
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
	}

	IEnumWbemClassObject* pEnumerator = nullptr;
	hres = pSvc->ExecQuery(
		bstr_t("WQL"),
		// bstr_t("SELECT name FROM Win32_NetworkAdapter WHERE Manufacturer != 'Microsoft' "),
		bstr_t("SELECT name FROM Win32_NetworkAdapter WHERE AdapterType ='Ethernet 802.3'"),

		WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
		nullptr,
		&pEnumerator);

	if (FAILED(hres)) {
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
	}

	IWbemClassObject *pclsObj = nullptr;

	ULONG uReturn = 0;

	pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);

	VARIANT vtProp;
	pclsObj->Get(L"Name", 0, &vtProp, nullptr, nullptr);

	CString result = vtProp.bstrVal;
	VariantClear(&vtProp);
	pclsObj->Release();


	pSvc->Release();
	pLoc->Release();
	pEnumerator->Release();
	CoUninitialize();
	return result;
}

// Get Ethernet NIC speed 
int Sysutils::GetNicSpeed()
{
	HRESULT hres;

	hres = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
	if (FAILED(hres)) {}

	hres = CoInitializeSecurity(
		nullptr,
		-1,
		nullptr,
		nullptr,
		RPC_C_AUTHN_LEVEL_DEFAULT,
		RPC_C_IMP_LEVEL_IMPERSONATE,
		nullptr,
		EOAC_NONE,
		nullptr
	);

	if (FAILED(hres)) {
		CoUninitialize();
	}

	IWbemLocator *pLoc = nullptr;

	hres = CoCreateInstance(
		CLSID_WbemLocator,
		nullptr,
		CLSCTX_INPROC_SERVER,
		IID_IWbemLocator, (LPVOID *)&pLoc);

	if (FAILED(hres)) {
		CoUninitialize();
	}

	IWbemServices *pSvc = nullptr;

	hres = pLoc->ConnectServer(
		_bstr_t(L"ROOT\\CIMV2"), // Object path of WMI namespace
		nullptr,                 // User name. NULL = current user
		nullptr,                 // User password. NULL = current
		nullptr,                 // Locale. NULL indicates current
		0,                       // Security flags.
		nullptr,                 // Authority (for example, Kerberos)
		nullptr,                 // Context object 
		&pSvc                    // pointer to IWbemServices proxy
	);

	if (FAILED(hres)) {
		pLoc->Release();
		CoUninitialize();
	}

	hres = CoSetProxyBlanket(
		pSvc,                        // Indicates the proxy to set
		RPC_C_AUTHN_WINNT,           // RPC_C_AUTHN_xxx
		RPC_C_AUTHZ_NONE,            // RPC_C_AUTHZ_xxx
		nullptr,                     // Server principal name 
		RPC_C_AUTHN_LEVEL_CALL,      // RPC_C_AUTHN_LEVEL_xxx 
		RPC_C_IMP_LEVEL_IMPERSONATE, // RPC_C_IMP_LEVEL_xxx
		nullptr,                     // client identity
		EOAC_NONE                    // proxy capabilities 
	);

	if (FAILED(hres)) {
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
	}

	IEnumWbemClassObject* pEnumerator = nullptr;
	hres = pSvc->ExecQuery(
		bstr_t("WQL"),
		bstr_t("SELECT speed FROM Win32_NetworkAdapter WHERE AdapterType ='Ethernet 802.3'"),
		WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
		nullptr,
		&pEnumerator);

	if (FAILED(hres)) {
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
	}

	IWbemClassObject *pclsObj = nullptr;

	ULONG uReturn = 0;

	pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);

	VARIANT vtProp;
	pclsObj->Get(L"speed", 0, &vtProp, nullptr, nullptr);

	CString result = vtProp.bstrVal;
	int iSpeed = _wtoi(result);

	VariantClear(&vtProp);
	pclsObj->Release();


	pSvc->Release();
	pLoc->Release();
	pEnumerator->Release();
	CoUninitialize();
	return iSpeed;
}

// Get Ethernet NIC is enabled (true/false)
CString Sysutils::GetNicIsEnabled()
{
	HRESULT hres;

	hres = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
	if (FAILED(hres)) {}

	hres = CoInitializeSecurity(
		nullptr,
		-1,
		nullptr,
		nullptr,
		RPC_C_AUTHN_LEVEL_DEFAULT,
		RPC_C_IMP_LEVEL_IMPERSONATE,
		nullptr,
		EOAC_NONE,
		nullptr
	);

	if (FAILED(hres)) {
		CoUninitialize();
	}

	IWbemLocator *pLoc = nullptr;

	hres = CoCreateInstance(
		CLSID_WbemLocator,
		nullptr,
		CLSCTX_INPROC_SERVER,
		IID_IWbemLocator, (LPVOID *)&pLoc);

	if (FAILED(hres)) {
		CoUninitialize();
	}

	IWbemServices *pSvc = nullptr;

	hres = pLoc->ConnectServer(
		_bstr_t(L"ROOT\\CIMV2"), // Object path of WMI namespace
		nullptr,                 // User name. NULL = current user
		nullptr,                 // User password. NULL = current
		nullptr,                 // Locale. NULL indicates current
		0,                       // Security flags.
		nullptr,                 // Authority (for example, Kerberos)
		nullptr,                 // Context object 
		&pSvc                    // pointer to IWbemServices proxy
	);

	if (FAILED(hres)) {
		pLoc->Release();
		CoUninitialize();
	}

	hres = CoSetProxyBlanket(
		pSvc,                        // Indicates the proxy to set
		RPC_C_AUTHN_WINNT,           // RPC_C_AUTHN_xxx
		RPC_C_AUTHZ_NONE,            // RPC_C_AUTHZ_xxx
		nullptr,                     // Server principal name 
		RPC_C_AUTHN_LEVEL_CALL,      // RPC_C_AUTHN_LEVEL_xxx 
		RPC_C_IMP_LEVEL_IMPERSONATE, // RPC_C_IMP_LEVEL_xxx
		nullptr,                     // client identity
		EOAC_NONE                    // proxy capabilities 
	);

	if (FAILED(hres)) {
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
	}

	IEnumWbemClassObject* pEnumerator = nullptr;
	hres = pSvc->ExecQuery(
		bstr_t("WQL"),
		bstr_t("SELECT netenabled FROM Win32_NetworkAdapter"),
		//bstr_t("SELECT NetEnabled FROM Win32_NetworkAdapter"),

		WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
		nullptr,
		&pEnumerator);

	if (FAILED(hres)) {
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
	}

	IWbemClassObject *pclsObj = nullptr;

	ULONG uReturn = 0;

	pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);

	VARIANT vtProp;
	pclsObj->Get(L"netenabled", 0, &vtProp, nullptr, nullptr);

	CString result;

	if (vtProp.bstrVal == VARIANT_FALSE) {
		result = "Enabled";
	} else {
		result = "Disabled";
	}
		
	VariantClear(&vtProp);
	pclsObj->Release();
	
	pSvc->Release();
	pLoc->Release();
	pEnumerator->Release();
	CoUninitialize();
	return result;
}

// Get Ethernet Connection (customizeable description)
CString Sysutils::GetNicConnectionID()
{
	HRESULT hres;

	hres = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
	if (FAILED(hres)) {}

	hres = CoInitializeSecurity(
		nullptr,
		-1,
		nullptr,
		nullptr,
		RPC_C_AUTHN_LEVEL_DEFAULT,
		RPC_C_IMP_LEVEL_IMPERSONATE,
		nullptr,
		EOAC_NONE,
		nullptr
	);

	if (FAILED(hres)) {
		CoUninitialize();
	}

	IWbemLocator *pLoc = nullptr;

	hres = CoCreateInstance(
		CLSID_WbemLocator,
		nullptr,
		CLSCTX_INPROC_SERVER,
		IID_IWbemLocator, (LPVOID *)&pLoc);

	if (FAILED(hres)) {
		CoUninitialize();
	}

	IWbemServices *pSvc = nullptr;

	hres = pLoc->ConnectServer(
		_bstr_t(L"ROOT\\CIMV2"), // Object path of WMI namespace
		nullptr,                 // User name. NULL = current user
		nullptr,                 // User password. NULL = current
		nullptr,                 // Locale. NULL indicates current
		0,                       // Security flags.
		nullptr,                 // Authority (for example, Kerberos)
		nullptr,                 // Context object 
		&pSvc                    // pointer to IWbemServices proxy
	);

	if (FAILED(hres)) {
		pLoc->Release();
		CoUninitialize();
	}

	hres = CoSetProxyBlanket(
		pSvc,                        // Indicates the proxy to set
		RPC_C_AUTHN_WINNT,           // RPC_C_AUTHN_xxx
		RPC_C_AUTHZ_NONE,            // RPC_C_AUTHZ_xxx
		nullptr,                     // Server principal name 
		RPC_C_AUTHN_LEVEL_CALL,      // RPC_C_AUTHN_LEVEL_xxx 
		RPC_C_IMP_LEVEL_IMPERSONATE, // RPC_C_IMP_LEVEL_xxx
		nullptr,                     // client identity
		EOAC_NONE                    // proxy capabilities 
	);

	if (FAILED(hres)) {
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
	}

	IEnumWbemClassObject* pEnumerator = nullptr;
	hres = pSvc->ExecQuery(
		bstr_t("WQL"),
		bstr_t("SELECT NetConnectionID FROM Win32_NetworkAdapter WHERE AdapterType ='Ethernet 802.3'"),
		WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
		nullptr,
		&pEnumerator);

	if (FAILED(hres)) {
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
	}

	IWbemClassObject *pclsObj = nullptr;

	ULONG uReturn = 0;

	pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);

	VARIANT vtProp;
	pclsObj->Get(L"NetConnectionID", 0, &vtProp, nullptr, nullptr);
	CString result = vtProp.bstrVal;
	
	VariantClear(&vtProp);
	pclsObj->Release();

	pSvc->Release();
	pLoc->Release();
	pEnumerator->Release();
	CoUninitialize();
	return result;
}

// Get DNS-Info
CString Sysutils::GetDNSInfo()
{
	HRESULT hres;

	hres = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
	if (FAILED(hres)) {}

	hres = CoInitializeSecurity(
		nullptr,
		-1,
		nullptr,
		nullptr,
		RPC_C_AUTHN_LEVEL_DEFAULT,
		RPC_C_IMP_LEVEL_IMPERSONATE,
		nullptr,
		EOAC_NONE,
		nullptr
	);

	if (FAILED(hres)) {
		CoUninitialize();
	}

	IWbemLocator *pLoc = nullptr;

	hres = CoCreateInstance(
		CLSID_WbemLocator,
		nullptr,
		CLSCTX_INPROC_SERVER,
		IID_IWbemLocator, (LPVOID *)&pLoc);

	if (FAILED(hres)) {
		CoUninitialize();
	}

	IWbemServices *pSvc = nullptr;

	hres = pLoc->ConnectServer(
		_bstr_t(L"ROOT\\CIMV2"), // Object path of WMI namespace
		nullptr,                 // User name. NULL = current user
		nullptr,                 // User password. NULL = current
		nullptr,                 // Locale. NULL indicates current
		0,                       // Security flags.
		nullptr,                 // Authority (for example, Kerberos)
		nullptr,                 // Context object 
		&pSvc                    // pointer to IWbemServices proxy
	);

	if (FAILED(hres)) {
		pLoc->Release();
		CoUninitialize();
	}

	hres = CoSetProxyBlanket(
		pSvc,                        // Indicates the proxy to set
		RPC_C_AUTHN_WINNT,           // RPC_C_AUTHN_xxx
		RPC_C_AUTHZ_NONE,            // RPC_C_AUTHZ_xxx
		nullptr,                     // Server principal name 
		RPC_C_AUTHN_LEVEL_CALL,      // RPC_C_AUTHN_LEVEL_xxx 
		RPC_C_IMP_LEVEL_IMPERSONATE, // RPC_C_IMP_LEVEL_xxx
		nullptr,                     // client identity
		EOAC_NONE                    // proxy capabilities 
	);

	if (FAILED(hres)) {
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
	}

	IEnumWbemClassObject* pEnumerator = nullptr;
	hres = pSvc->ExecQuery(
		bstr_t("WQL"),
		bstr_t("SELECT name, numberofcores FROM Win32_Processor"),
		WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
		nullptr,
		&pEnumerator);

	if (FAILED(hres)) {
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
	}

	IWbemClassObject *pclsObj = nullptr;
	ULONG uReturn = 0;
	pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);

	VARIANT vtProp;
	pclsObj->Get(L"Name", 0, &vtProp, nullptr, nullptr);

	CString result = vtProp.bstrVal;
	VariantClear(&vtProp);
	pclsObj->Release();
	
	pSvc->Release();
	pLoc->Release();
	pEnumerator->Release();
	CoUninitialize();
	return result;
}

// Get Mac-Address 
CString Sysutils::GetMacAddress()
{
	HRESULT hres;

	hres = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
	if (FAILED(hres)) {}

	hres = CoInitializeSecurity(
		nullptr,
		-1,
		nullptr,
		nullptr,
		RPC_C_AUTHN_LEVEL_DEFAULT,
		RPC_C_IMP_LEVEL_IMPERSONATE,
		nullptr,
		EOAC_NONE,
		nullptr
	);

	if (FAILED(hres)) {
		CoUninitialize();
	}

	IWbemLocator *pLoc = nullptr;

	hres = CoCreateInstance(
		CLSID_WbemLocator,
		nullptr,
		CLSCTX_INPROC_SERVER,
		IID_IWbemLocator, (LPVOID *)&pLoc);

	if (FAILED(hres)) {
		CoUninitialize();
	}

	IWbemServices *pSvc = nullptr;

	hres = pLoc->ConnectServer(
		_bstr_t(L"ROOT\\CIMV2"), // Object path of WMI namespace
		nullptr,                 // User name. NULL = current user
		nullptr,                 // User password. NULL = current
		nullptr,                 // Locale. NULL indicates current
		0,                       // Security flags.
		nullptr,                 // Authority (for example, Kerberos)
		nullptr,                 // Context object 
		&pSvc                    // pointer to IWbemServices proxy
	);

	if (FAILED(hres)) {
		pLoc->Release();
		CoUninitialize();
	}

	hres = CoSetProxyBlanket(
		pSvc,                        // Indicates the proxy to set
		RPC_C_AUTHN_WINNT,           // RPC_C_AUTHN_xxx
		RPC_C_AUTHZ_NONE,            // RPC_C_AUTHZ_xxx
		nullptr,                     // Server principal name 
		RPC_C_AUTHN_LEVEL_CALL,      // RPC_C_AUTHN_LEVEL_xxx 
		RPC_C_IMP_LEVEL_IMPERSONATE, // RPC_C_IMP_LEVEL_xxx
		nullptr,                     // client identity
		EOAC_NONE                    // proxy capabilities 
	);

	if (FAILED(hres)) {
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
	}

	IEnumWbemClassObject* pEnumerator = nullptr;
	hres = pSvc->ExecQuery(
		bstr_t("WQL"),
		bstr_t("SELECT MACAddress FROM Win32_NetworkAdapter WHERE AdapterType ='Ethernet 802.3'"),
		WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
		nullptr,
		&pEnumerator);

	if (FAILED(hres)) {
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
	}

	IWbemClassObject *pclsObj = nullptr;
	ULONG uReturn = 0;
	pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);

	VARIANT vtProp;
	pclsObj->Get(L"MACAddress", 0, &vtProp, nullptr, nullptr);

	CString result = vtProp.bstrVal;
	VariantClear(&vtProp);
	pclsObj->Release();

	pSvc->Release();
	pLoc->Release();
	pEnumerator->Release();
	CoUninitialize();
	return result;
}

// Get Default Gateway
CString Sysutils::GetDefaultGateWay()
{
	HRESULT hres;

	hres = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
	if (FAILED(hres)) {}

	hres = CoInitializeSecurity(
		nullptr,
		-1,
		nullptr,
		nullptr,
		RPC_C_AUTHN_LEVEL_DEFAULT,
		RPC_C_IMP_LEVEL_IMPERSONATE,
		nullptr,
		EOAC_NONE,
		nullptr
	);

	if (FAILED(hres)) {
		CoUninitialize();
	}

	IWbemLocator *pLoc = nullptr;

	hres = CoCreateInstance(
		CLSID_WbemLocator,
		nullptr,
		CLSCTX_INPROC_SERVER,
		IID_IWbemLocator, (LPVOID *)&pLoc);

	if (FAILED(hres)) {
		CoUninitialize();
	}

	IWbemServices *pSvc = nullptr;

	hres = pLoc->ConnectServer(
		_bstr_t(L"ROOT\\CIMV2"), // Object path of WMI namespace
		nullptr,                 // User name. NULL = current user
		nullptr,                 // User password. NULL = current
		nullptr,                 // Locale. NULL indicates current
		0,                       // Security flags.
		nullptr,                 // Authority (for example, Kerberos)
		nullptr,                 // Context object 
		&pSvc                    // pointer to IWbemServices proxy
	);

	if (FAILED(hres)) {
		pLoc->Release();
		CoUninitialize();
	}

	hres = CoSetProxyBlanket(
		pSvc,                        // Indicates the proxy to set
		RPC_C_AUTHN_WINNT,           // RPC_C_AUTHN_xxx
		RPC_C_AUTHZ_NONE,            // RPC_C_AUTHZ_xxx
		nullptr,                     // Server principal name 
		RPC_C_AUTHN_LEVEL_CALL,      // RPC_C_AUTHN_LEVEL_xxx 
		RPC_C_IMP_LEVEL_IMPERSONATE, // RPC_C_IMP_LEVEL_xxx
		nullptr,                     // client identity
		EOAC_NONE                    // proxy capabilities 
	);

	if (FAILED(hres)) {
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
	}

	IEnumWbemClassObject* pEnumerator = nullptr;
	hres = pSvc->ExecQuery(
		bstr_t("WQL"),
		bstr_t("SELECT Nexthop FROM Win32_IP4RouteTable WHERE (Mask='0.0.0.0')"),
		WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
		nullptr,
		&pEnumerator);

	if (FAILED(hres)) {
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
	}

	IWbemClassObject *pclsObj = nullptr;
	ULONG uReturn = 0;
	pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);

	VARIANT vtProp;
	pclsObj->Get(L"Nexthop", 0, &vtProp, nullptr, nullptr);

	CString result = vtProp.bstrVal;
	VariantClear(&vtProp);
	pclsObj->Release();

	pSvc->Release();
	pLoc->Release();
	pEnumerator->Release();
	CoUninitialize();
	return result;
}

// Get DNSServers
CString Sysutils::GetDNSServers()
{
	
	return (L"No Inpletemeted");
}



CString Sysutils::GetOSArchitecture()
{
	HRESULT hr;
	CString result;

	hr = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
	if (FAILED(hr))	{ }

	hr = CoInitializeSecurity(
		nullptr,
		-1,                          // COM authentication
		nullptr,                        // Authentication services
		nullptr,                        // Reserved
		RPC_C_AUTHN_LEVEL_DEFAULT,   // Default authentication 
		RPC_C_IMP_LEVEL_IMPERSONATE, // Default Impersonation  
		nullptr,                        // Authentication info
		EOAC_NONE,                   // Additional capabilities 
		nullptr                         // Reserved
	);

	if (FAILED(hr)) {
		CoUninitialize();
	}

	IWbemLocator *pLoc = nullptr;

	hr = CoCreateInstance(
		CLSID_WbemLocator,
		nullptr,
		CLSCTX_INPROC_SERVER,
		IID_IWbemLocator, (LPVOID *)&pLoc);

	if (FAILED(hr))	{
		CoUninitialize();		
	}
	
	IWbemServices *pSvc = nullptr;

	hr = pLoc->ConnectServer(
		_bstr_t(L"ROOT\\CIMV2"), // Object path of WMI namespace
		nullptr,                    // User name. NULL = current user
		nullptr,                    // User password. NULL = current
		nullptr,                       // Locale. NULL indicates current
		NULL,                    // Security flags.
		nullptr,                       // Authority (for example, Kerberos)
		nullptr,                       // Context object 
		&pSvc                    // pointer to IWbemServices proxy
	);

	if (FAILED(hr))	{
		pLoc->Release();
		CoUninitialize();
	}
	
	// Set security flags
	hr = CoSetProxyBlanket(
		pSvc,                        // Indicates the proxy to set
		RPC_C_AUTHN_WINNT,           // RPC_C_AUTHN_xxx
		RPC_C_AUTHZ_NONE,            // RPC_C_AUTHZ_xxx
		nullptr,                        // Server principal name 
		RPC_C_AUTHN_LEVEL_CALL,      // RPC_C_AUTHN_LEVEL_xxx 
		RPC_C_IMP_LEVEL_IMPERSONATE, // RPC_C_IMP_LEVEL_xxx
		nullptr,                        // client identity
		EOAC_NONE                    // proxy capabilities 
	);

	if (FAILED(hr)) {
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
	}

	IEnumWbemClassObject* pEnumerator = nullptr;
	hr = pSvc->ExecQuery(
		bstr_t("WQL"),
		bstr_t("SELECT OSArchitecture FROM Win32_OperatingSystem"),
		WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
		nullptr,
		&pEnumerator);

	if (FAILED(hr)){
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
	}
	
	IWbemClassObject *pclsObj = nullptr;
	ULONG uReturn = 0;
	hr = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);

	if (SUCCEEDED(hr)) {
		VARIANT vtProp;
		pclsObj->Get(L"OSArchitecture", 0, &vtProp, nullptr, nullptr);
		result = vtProp.bstrVal;
		VariantClear(&vtProp);
		pclsObj->Release();

		// Cleanup
		pSvc->Release();
		pLoc->Release();
		pEnumerator->Release();
		CoUninitialize();
	}
	result.MakeLower();
	return result;
}

// Get Major Servicepack no.
int Sysutils::GetOSServicePack()
{
	OSVERSIONINFOEX osvi;
	SYSTEM_INFO si;

		ZeroMemory(&si, sizeof(SYSTEM_INFO));
		ZeroMemory(&osvi, sizeof(OSVERSIONINFOEX));

		osvi.dwOSVersionInfoSize = sizeof(OSVERSIONINFOEX);
		GetVersionEx((OSVERSIONINFO*)&osvi);
		GetSystemInfo(&si);
	
	return osvi.wServicePackMajor;
}

// Get Major Servicepack no.
int Sysutils::GetOSBuild()
{
	OSVERSIONINFOEX osvi;
	SYSTEM_INFO si;

	ZeroMemory(&si, sizeof(SYSTEM_INFO));
	ZeroMemory(&osvi, sizeof(OSVERSIONINFOEX));

	osvi.dwOSVersionInfoSize = sizeof(OSVERSIONINFOEX);
	GetVersionEx((OSVERSIONINFO*)&osvi);
	GetSystemInfo(&si);	
	
	return osvi.dwBuildNumber;
}

CString Sysutils::GetOSBuildEx()
{
	HRESULT hr;
	CString result;

	hr = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
	if (FAILED(hr)) {}

	hr = CoInitializeSecurity(
		nullptr,
		-1,                          // COM authentication
		nullptr,                        // Authentication services
		nullptr,                        // Reserved
		RPC_C_AUTHN_LEVEL_DEFAULT,   // Default authentication 
		RPC_C_IMP_LEVEL_IMPERSONATE, // Default Impersonation  
		nullptr,                        // Authentication info
		EOAC_NONE,                   // Additional capabilities 
		nullptr                         // Reserved
	);

	if (FAILED(hr)) {
		CoUninitialize();
	}

	IWbemLocator *pLoc = nullptr;

	hr = CoCreateInstance(
		CLSID_WbemLocator,
		nullptr,
		CLSCTX_INPROC_SERVER,
		IID_IWbemLocator, (LPVOID *)&pLoc);

	if (FAILED(hr)) {
		CoUninitialize();
	}

	IWbemServices *pSvc = nullptr;

	hr = pLoc->ConnectServer(
		_bstr_t(L"ROOT\\CIMV2"), // Object path of WMI namespace
		nullptr,                    // User name. NULL = current user
		nullptr,                    // User password. NULL = current
		nullptr,                       // Locale. NULL indicates current
		NULL,                    // Security flags.
		nullptr,                       // Authority (for example, Kerberos)
		nullptr,                       // Context object 
		&pSvc                    // pointer to IWbemServices proxy
	);

	if (FAILED(hr)) {
		pLoc->Release();
		CoUninitialize();
	}

	// Set security flags
	hr = CoSetProxyBlanket(
		pSvc,                        // Indicates the proxy to set
		RPC_C_AUTHN_WINNT,           // RPC_C_AUTHN_xxx
		RPC_C_AUTHZ_NONE,            // RPC_C_AUTHZ_xxx
		nullptr,                        // Server principal name 
		RPC_C_AUTHN_LEVEL_CALL,      // RPC_C_AUTHN_LEVEL_xxx 
		RPC_C_IMP_LEVEL_IMPERSONATE, // RPC_C_IMP_LEVEL_xxx
		nullptr,                        // client identity
		EOAC_NONE                    // proxy capabilities 
	);

	if (FAILED(hr)) {
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
	}

	IEnumWbemClassObject* pEnumerator = nullptr;
	hr = pSvc->ExecQuery(
		bstr_t("WQL"),
		bstr_t("SELECT version FROM Win32_OperatingSystem"),
		WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
		nullptr,
		&pEnumerator);

	if (FAILED(hr)) {
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
	}

	IWbemClassObject *pclsObj = nullptr;
	ULONG uReturn = 0;
	hr = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);

	if (SUCCEEDED(hr)) {
		VARIANT vtProp;
		pclsObj->Get(L"version", 0, &vtProp, nullptr, nullptr);
		result = vtProp.bstrVal;
		VariantClear(&vtProp);
		pclsObj->Release();

		// Cleanup
		pSvc->Release();
		pLoc->Release();
		pEnumerator->Release();
		CoUninitialize();
	}
	result.MakeLower();
	return result;
}

// Get CPU-Cores
int Sysutils::GetCPUCores()
{
	OSVERSIONINFOEX osvi;
	SYSTEM_INFO si;

		ZeroMemory(&si, sizeof(SYSTEM_INFO));
		ZeroMemory(&osvi, sizeof(OSVERSIONINFOEX));

		osvi.dwOSVersionInfoSize = sizeof(OSVERSIONINFOEX);
		GetVersionEx((OSVERSIONINFO*)&osvi);
		GetSystemInfo(&si);
		

	return si.dwNumberOfProcessors;
}

int Sysutils::MaxProcessMemorySize()
{
	HRESULT hres;

	hres = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
	if (FAILED(hres))	{}

	hres = CoInitializeSecurity(
		nullptr,
		-1, 
		nullptr,
		nullptr, 
		RPC_C_AUTHN_LEVEL_DEFAULT,  
		RPC_C_IMP_LEVEL_IMPERSONATE,
		nullptr, 
		EOAC_NONE, 
		nullptr
	);


	if (FAILED(hres)) {	}

	IWbemLocator *pLoc = nullptr;
	hres = CoCreateInstance(
		CLSID_WbemLocator,
		nullptr,
		CLSCTX_INPROC_SERVER,
		IID_IWbemLocator, (LPVOID *)&pLoc);

	if (FAILED(hres)) {
		CoUninitialize();
	}

	IWbemServices *pSvc = nullptr;
	hres = pLoc->ConnectServer(
		_bstr_t(L"ROOT\\CIMV2"), 
		nullptr,
		nullptr,
		nullptr,
		NULL,
		nullptr,
		nullptr,
		&pSvc
	);

	if (FAILED(hres)) {
		pLoc->Release();
		CoUninitialize();
	}
	
	// Set security flags
	hres = CoSetProxyBlanket(
		pSvc,                       
		RPC_C_AUTHN_WINNT,          
		RPC_C_AUTHZ_NONE,         
		nullptr,                      
		RPC_C_AUTHN_LEVEL_CALL,      
		RPC_C_IMP_LEVEL_IMPERSONATE, 
		nullptr, 
		EOAC_NONE 
	);

	if (FAILED(hres)) {
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();		
	}

	IEnumWbemClassObject* pEnumerator = nullptr;
	hres = pSvc->ExecQuery(
		bstr_t("WQL"),
		bstr_t("SELECT * FROM Win32_OperatingSystem"),
		WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
		nullptr,
		&pEnumerator);

	if (FAILED(hres))	{
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
	}

	IWbemClassObject *pclsObj = nullptr;
	ULONG uReturn = 0;
	pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);

			VARIANT vtProp;
			pclsObj->Get(L"MaxProcessMemorySize", 0, &vtProp, nullptr, nullptr);
			CString result = vtProp.bstrVal;
			int maxprocmemory = _wtoi(result);

			VariantClear(&vtProp);
			pclsObj->Release();

	// Cleanup
	pSvc->Release();
	pLoc->Release();
	pEnumerator->Release();
	CoUninitialize();

  return maxprocmemory;
}

CString Sysutils::GetTimestamp()
{
	time_t     now = time(nullptr);
	char       buf[80];
	struct tm tstruct = *localtime(&now);
	strftime(buf, sizeof(buf), "%Y-%m-%d.%X", &tstruct);

	return (CString) buf;
}

CString Sysutils::GetMemoryInUse()
{
	MEMORYSTATUSEX statex;

		statex.dwLength = sizeof(statex);
		GlobalMemoryStatusEx(&statex);

		CString memoryfree = nullptr;
		memoryfree.Format(L"%d", statex.dwMemoryLoad);

	return memoryfree;
}



int Sysutils::GetHMSBitness()
{
	return this->gBitness;
}

CString Sysutils::GetBinaryBitness(LPCTSTR lpApplicationName)
{
	DWORD dwBinaryType;
	CString strResult = nullptr;

	if (GetBinaryType(lpApplicationName, &dwBinaryType)
		&& dwBinaryType == SCS_64BIT_BINARY) {
 		strResult = L"64-Bit (x64)";
		this->gBitness = 64;
	}

	if (GetBinaryType(lpApplicationName, &dwBinaryType)
		&& dwBinaryType == SCS_32BIT_BINARY) {
		strResult = L"32-Bit (x86)";		
		gBitness = 32;
	}
	return strResult;
}


int Sysutils::GetMyBitness()
{
	DWORD dwBinaryType;
	int bitness = 0;

	if (GetBinaryType(GetHMSInfoPath(), &dwBinaryType)
		&& dwBinaryType == SCS_64BIT_BINARY) {		
		bitness = 64;
	}

	if (GetBinaryType(GetHMSInfoPath(), &dwBinaryType)
		&& dwBinaryType == SCS_32BIT_BINARY) {
		bitness = 32;		
	}
	return bitness;
}

// Get WOW32 Installpath of hMailServer 32-Bit on Windows 64-Bit
CString Sysutils::GetHMS32WOW64InstallPath()
{
	LPCWSTR pszValueName = L"InstallLocation";
	LPCTSTR pszSubkey = L"SOFTWARE\\Wow6432Node\\hMailServer";	
	HKEY hKey = nullptr;	
	
	if (RegOpenKey(HKEY_LOCAL_MACHINE, pszSubkey, &hKey) != ERROR_SUCCESS) {
		AtlThrowLastWin32();
	}

	// Buffer to store string read from registry
	TCHAR szValue[1024];
	DWORD cbValueLength = sizeof(szValue);
	
	// Query string value
	if (RegQueryValueEx(
		hKey,
		pszValueName,
		nullptr,
		nullptr,
		reinterpret_cast<LPBYTE>(&szValue),
		&cbValueLength)
		!= ERROR_SUCCESS) {
		AtlThrowLastWin32();
	}
	return CString(szValue);

}

// Get Win64 Installpath of hMailServer 64-Bit, native
CString Sysutils::GetHMS64InstallPath()
{
	LPCTSTR pszSubkey = L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\hMailServer_is1";
	HKEY hKey = HKEY_LOCAL_MACHINE;
	TCHAR szValue[1024] = { 0 };

	HKEY subKey = nullptr;
	if (!(RegOpenKeyEx(hKey, pszSubkey, 0, KEY_READ, &subKey) != ERROR_SUCCESS)) {

		LPCWSTR pszValueName = L"InstallLocation";
		pszSubkey = L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\hMailServer_is1";

		if (RegOpenKey(HKEY_LOCAL_MACHINE, pszSubkey, &hKey) != ERROR_SUCCESS) {
			AfxMessageBox(L"Error on Querying Registry HIVE Key");
		}

		// Buffer to store string 
		DWORD cbValueLength = sizeof(szValue);

		if (RegQueryValueEx(
			hKey,
			pszValueName,
			nullptr,
			nullptr,
			reinterpret_cast<LPBYTE>(&szValue),
			&cbValueLength)
			!= ERROR_SUCCESS) {			
			   AfxMessageBox(L"Error on Querying Registry for HMS-Installinfo");
			}
		}	
	return szValue;
}

// Get Win32 Installpath of hMailServer 32-Bit, native
CString Sysutils::GetHMS32InstallPath()
{
	LPCTSTR pszSubkey = L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\hMailServer_is1";
	HKEY hKey = HKEY_LOCAL_MACHINE;
	TCHAR szValue[1024] = { 0 };

	HKEY subKey = nullptr;
	if (!(RegOpenKeyEx(hKey, pszSubkey, 0, KEY_READ, &subKey) != ERROR_SUCCESS)) {

		LPCWSTR pszValueName = L"InstallLocation";
		pszSubkey = L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\hMailServer_is1";

		if (RegOpenKey(HKEY_LOCAL_MACHINE, pszSubkey, &hKey) != ERROR_SUCCESS) {
			AfxMessageBox(L"Error on Querying Registry HIVE Key");
		}

		// Buffer to store string 
		DWORD cbValueLength = sizeof(szValue);

		if (RegQueryValueEx(
			hKey,
			pszValueName,
			nullptr,
			nullptr,
			reinterpret_cast<LPBYTE>(&szValue),
			&cbValueLength)
			!= ERROR_SUCCESS) {
			AfxMessageBox(L"Error on Querying Registry for HMS-Installinfo");
		}
	}
	return szValue;
}

// Get Win32 Installdate of hMailServer 32-Bit, native
CString Sysutils::GetHMS32InstallDate()
{
	LPCTSTR pszSubkey = L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\hMailServer_is1";
	HKEY hKey = HKEY_LOCAL_MACHINE;
	TCHAR szValue[1024];

	HKEY subKey = nullptr;
	if (!(RegOpenKeyEx(hKey, pszSubkey, 0, KEY_READ, &subKey) != ERROR_SUCCESS)) {

		LPCWSTR pszValueName = L"InstallDate";
		pszSubkey = L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\hMailServer_is1";

		if (RegOpenKey(HKEY_LOCAL_MACHINE, pszSubkey, &hKey) != ERROR_SUCCESS) {
			AfxMessageBox(L"Error on Querying Registry HIVE Key");
		}

		// Buffer to store string 
		DWORD cbValueLength = sizeof(szValue);

		if (RegQueryValueEx(
			hKey,
			pszValueName,
			nullptr,
			nullptr,
			reinterpret_cast<LPBYTE>(&szValue),
			&cbValueLength)
			!= ERROR_SUCCESS) {
			AfxMessageBox(L"Error on Querying Registry for HMS-Installinfo");
		}
	}
	return szValue;
}

CString Sysutils::GetHMS64NativeVersionInfo()
{
	LPCTSTR pszSubkey = L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\hMailServer_is1";
	HKEY hKey = HKEY_LOCAL_MACHINE;
	TCHAR szValue[1024];

	HKEY subKey = nullptr;
	if (!(RegOpenKeyEx(hKey, pszSubkey, 0, KEY_READ, &subKey) != ERROR_SUCCESS)) {

		LPCWSTR pszValueName = L"DisplayName";
		pszSubkey = L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\hMailServer_is1";

		if (RegOpenKey(HKEY_LOCAL_MACHINE, pszSubkey, &hKey) != ERROR_SUCCESS) {
			AfxMessageBox(L"Error on Querying Registry HIVE Key");
		}

		// Buffer to store string 
		DWORD cbValueLength = sizeof(szValue);

		if (RegQueryValueEx(
			hKey,
			pszValueName,
			nullptr,
			nullptr,
			reinterpret_cast<LPBYTE>(&szValue),
			&cbValueLength)
			!= ERROR_SUCCESS) {
			AfxMessageBox(L"Error on Querying Registry for HMS-Installinfo");
		}
	}
	return szValue;
}

CString Sysutils::GetHMS32NativeVersionInfo()
{
	LPCTSTR pszSubkey = L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\hMailServer_is1";
	HKEY hKey = HKEY_LOCAL_MACHINE;
	TCHAR szValue[1024];

	HKEY subKey = nullptr;
	if (!(RegOpenKeyEx(hKey, pszSubkey, 0, KEY_READ, &subKey) != ERROR_SUCCESS)) {

		LPCWSTR pszValueName = L"DisplayName";
		pszSubkey = L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\hMailServer_is1";

		if (RegOpenKey(HKEY_LOCAL_MACHINE, pszSubkey, &hKey) != ERROR_SUCCESS) {
			AfxMessageBox(L"Error on Querying Registry HIVE Key");
		}

		// Buffer to store string 
		DWORD cbValueLength = sizeof(szValue);

		if (RegQueryValueEx(
			hKey,
			pszValueName,
			nullptr,
			nullptr,
			reinterpret_cast<LPBYTE>(&szValue),
			&cbValueLength)
			!= ERROR_SUCCESS) {
			AfxMessageBox(L"Error on Querying Registry for HMS-Installinfo");
		}
	}
	return szValue;
}

// Get the current path off HMSInfo.exe
CString Sysutils::GetHMSInfoPath()
{
	//char buf[1024] = { 0 };
	LPSTR buf = nullptr;
	GetModuleFileNameA(nullptr, buf, sizeof(buf));
	char* wrkdir = (char*)buf;
	return((CString)wrkdir);
}

// Checking if we are running on a 32 or a 64-Bit Operating System
BOOL Sysutils::IsWow64()
{
	BOOL bIsWow64 = FALSE;

	fnIsWow64Process = (LPFN_ISWOW64PROCESS)GetProcAddress(
		GetModuleHandle(TEXT("kernel32")), "IsWow64Process");

	if (nullptr != fnIsWow64Process) {
		if (!fnIsWow64Process(GetCurrentProcess(), &bIsWow64)) 	{
			//handle error - not yet.
		}
	}
	return bIsWow64;
}

bool Sysutils::IsNative64Bit()
{
	SYSTEM_INFO System;
	GetNativeSystemInfo(&System);
	bool result = false;
	if (System.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64)
	{
		result = true;
	}

	return result;
}

// check if a 32-Bit Version of hMailServer is allready on a 64-Bit OS installed.
bool Sysutils::IsHMS32_WOW64_Installed() {

	bool rst = false;	

	if (this->GetOSArchitecture() == "64-bit") {

		LPCTSTR pszSubkey = L"SOFTWARE\\Wow6432Node\\hMailServer";
		HKEY hKey = HKEY_LOCAL_MACHINE;

		HKEY subKey = nullptr;
		if (!(RegOpenKeyEx(hKey, pszSubkey, 0, KEY_READ, &subKey) != ERROR_SUCCESS)) {

			LPCWSTR pszValueName = L"InstallLocation";			
			pszSubkey = L"SOFTWARE\\Wow6432Node\\hMailServer";

			if (RegOpenKey(HKEY_LOCAL_MACHINE, pszSubkey, &hKey) != ERROR_SUCCESS) {
				AtlThrowLastWin32();
			}

			// Buffer to store string read from registry
			TCHAR szValue[1024];
			DWORD cbValueLength = sizeof(szValue);

			// Query string value
			if (RegQueryValueEx(
				hKey,
				pszValueName,
				nullptr,
				nullptr,
				reinterpret_cast<LPBYTE>(&szValue),
				&cbValueLength)
				!= ERROR_SUCCESS) {
				AtlThrowLastWin32();
			}

			CString wrkstr = szValue;
			wrkstr = wrkstr +L"\\Bin\\hMailServer.exe";

			if (this->GetBinaryBitness(wrkstr) == "32-Bit (x86)")
				rst = true;
		}	
	}
	
	return rst;
}

bool Sysutils::IsHMS32_native_installed() {		

	bool rst = false;
	if (this->GetOSArchitecture() == "32-bit") {
		LPCTSTR pszSubkey = L"SOFTWARE\\hMailServer";
		HKEY hKey = HKEY_LOCAL_MACHINE;

		HKEY subKey = nullptr;
		LONG result = RegOpenKeyEx(hKey, pszSubkey, 0, KEY_READ, &subKey);
		if (result != ERROR_SUCCESS)
			rst = false;
		else
			rst = true;
	}
	return rst;

}

bool Sysutils::IsHMS64_native_installed() {

	bool rst = false;

	if (this->GetOSArchitecture() == "64-bit") {

		LPCTSTR pszSubkey = L"SOFTWARE\\Wow6432Node\\hMailServer";
		HKEY hKey = HKEY_LOCAL_MACHINE;

		HKEY subKey = nullptr;
		if (!(RegOpenKeyEx(hKey, pszSubkey, 0, KEY_READ, &subKey) != ERROR_SUCCESS)) {

			LPCWSTR pszValueName = L"InstallLocation";
			pszSubkey = L"SOFTWARE\\Wow6432Node\\hMailServer";

			if (RegOpenKey(HKEY_LOCAL_MACHINE, pszSubkey, &hKey) != ERROR_SUCCESS) {
				AtlThrowLastWin32();
			}

			// Buffer to store string read from registry
			TCHAR szValue[1024];
			DWORD cbValueLength = sizeof(szValue);

			// Query string value
			if (RegQueryValueEx(
				hKey,
				pszValueName,
				nullptr,
				nullptr,
				reinterpret_cast<LPBYTE>(&szValue),
				&cbValueLength)
				!= ERROR_SUCCESS) {
				AtlThrowLastWin32();
			}

			CString wrkstr = szValue;
			wrkstr = wrkstr + L"\\Bin\\hMailServer.exe";

			if (this->GetBinaryBitness(wrkstr) == "64-Bit (x64)")
				rst = true;
		}
	}
	return rst;
}

// Get the used type of Database
CString Sysutils::GetSQLDatabaseType(CString iniFilepath)
{
	string cfilename = iniFilepath+"\\Bin\\hMailServer.INI";
	ifstream fileInput;
	string line;
	char* search = "Type="; 
	string result;
						
	fileInput.open(cfilename.c_str());
	if (fileInput.is_open()) {
		unsigned int curLine = 0;
		while (getline(fileInput, line)) { 
			curLine++;
			if (line.find(search, 0) != string::npos) {		
				result = line;
			}
		}
		fileInput.close();
	}
	
	CString wrkx = CA2T(result.c_str());
	return wrkx;
}

// Check if we are running inside VirtualBox Guestsystem or not
CString Sysutils::GetIsVirtualBox()
{
	if (CreateFile(L"\\\\.\\VBoxMiniRdrDN", GENERIC_READ, FILE_SHARE_READ,
		nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr) != INVALID_HANDLE_VALUE) {
		return (L"Virtualbox Guest Environment detected!");
	} else	{
		return (L"No Virtualmachine detected");
	}
		
}
// Check Windows Service Status
char* Sysutils::getServiceStatus()
{
		TCHAR szSvcName[80] = L"hMailServer";
		SERVICE_STATUS_PROCESS ssp;
		DWORD dwBytesNeeded;
		char* iReturn = nullptr;

		SC_HANDLE schSCManager = OpenSCManager(
			nullptr,						// local computer
			nullptr,						// ServicesActive database 
			SC_MANAGER_QUERY_LOCK_STATUS);  // full access rights 

		if (nullptr == schSCManager) {
		}

		// Get a handle to the service.
		SC_HANDLE schService = OpenService(
			schSCManager,         // SCM database 
			szSvcName,            // name of service 			
			SERVICE_QUERY_STATUS );

		if (schService == nullptr)
		{
			CloseServiceHandle(schSCManager);
		}

		// Make sure the service is not already stopped.
		if (!QueryServiceStatusEx(
			schService,
			SC_STATUS_PROCESS_INFO,
			(LPBYTE)&ssp,
			sizeof(SERVICE_STATUS_PROCESS),
			&dwBytesNeeded))
		{
			iReturn = "Error! Cannot access Servicedata";
		}

		if (ssp.dwCurrentState == SERVICE_STOPPED) {
			iReturn = "stopped";
		}

		if (ssp.dwCurrentState == SERVICE_RUNNING) {
			iReturn = "running";
		}

		CloseServiceHandle(schService);
		CloseServiceHandle(schSCManager);

		return iReturn;
}

// Check Windows the Username for the Service Status
CString Sysutils::GetServiceAccountName()
{
	HRESULT hres;

	hres = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
	if (FAILED(hres)) {}

	hres = CoInitializeSecurity(
		nullptr,
		-1,
		nullptr,
		nullptr,
		RPC_C_AUTHN_LEVEL_DEFAULT,
		RPC_C_IMP_LEVEL_IMPERSONATE,
		nullptr,
		EOAC_NONE,
		nullptr
	);

	if (FAILED(hres)) {
		CoUninitialize();
	}

	IWbemLocator *pLoc = nullptr;

	hres = CoCreateInstance(
		CLSID_WbemLocator,
		nullptr,
		CLSCTX_INPROC_SERVER,
		IID_IWbemLocator, (LPVOID *)&pLoc);

	if (FAILED(hres)) {
		CoUninitialize();
	}

	IWbemServices *pSvc = nullptr;

	hres = pLoc->ConnectServer(
		_bstr_t(L"ROOT\\CIMV2"), // Object path of WMI namespace
		nullptr,                 // User name. NULL = current user
		nullptr,                 // User password. NULL = current
		nullptr,                 // Locale. NULL indicates current
		0,                       // Security flags.
		nullptr,                 // Authority (for example, Kerberos)
		nullptr,                 // Context object 
		&pSvc                    // pointer to IWbemServices proxy
	);

	if (FAILED(hres)) {
		pLoc->Release();
		CoUninitialize();
	}

	hres = CoSetProxyBlanket(
		pSvc,                        // Indicates the proxy to set
		RPC_C_AUTHN_WINNT,           // RPC_C_AUTHN_xxx
		RPC_C_AUTHZ_NONE,            // RPC_C_AUTHZ_xxx
		nullptr,                     // Server principal name 
		RPC_C_AUTHN_LEVEL_CALL,      // RPC_C_AUTHN_LEVEL_xxx 
		RPC_C_IMP_LEVEL_IMPERSONATE, // RPC_C_IMP_LEVEL_xxx
		nullptr,                     // client identity
		EOAC_NONE                    // proxy capabilities 
	);

	if (FAILED(hres)) {
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
	}

	IEnumWbemClassObject* pEnumerator = nullptr;
	hres = pSvc->ExecQuery(
		bstr_t("WQL"),
		bstr_t("SELECT startname FROM Win32_Service"),
		WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
		nullptr,
		&pEnumerator);

	if (FAILED(hres)) {
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
	}

	IWbemClassObject *pclsObj = nullptr;
	ULONG uReturn = 0;

	pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);

	VARIANT vtProp;
	pclsObj->Get(L"startname", 0, &vtProp, nullptr, nullptr);
	CString result = vtProp.bstrVal;
	VariantClear(&vtProp);
	pclsObj->Release();

	pSvc->Release();
	pLoc->Release();
	pEnumerator->Release();
	CoUninitialize();
	return result;
}

// Check Windows hMailServer Service Startmode
CString Sysutils::GetServiceStartMode()
{
	HRESULT hres;

	hres = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
	if (FAILED(hres)) {}

	hres = CoInitializeSecurity(
		nullptr,
		-1,
		nullptr,
		nullptr,
		RPC_C_AUTHN_LEVEL_DEFAULT,
		RPC_C_IMP_LEVEL_IMPERSONATE,
		nullptr,
		EOAC_NONE,
		nullptr
	);

	if (FAILED(hres)) {
		CoUninitialize();
	}

	IWbemLocator *pLoc = nullptr;

	hres = CoCreateInstance(
		CLSID_WbemLocator,
		nullptr,
		CLSCTX_INPROC_SERVER,
		IID_IWbemLocator, (LPVOID *)&pLoc);

	if (FAILED(hres)) {
		CoUninitialize();
	}

	IWbemServices *pSvc = nullptr;

	hres = pLoc->ConnectServer(
		_bstr_t(L"ROOT\\CIMV2"), // Object path of WMI namespace
		nullptr,                 // User name. NULL = current user
		nullptr,                 // User password. NULL = current
		nullptr,                 // Locale. NULL indicates current
		0,                       // Security flags.
		nullptr,                 // Authority (for example, Kerberos)
		nullptr,                 // Context object 
		&pSvc                    // pointer to IWbemServices proxy
	);

	if (FAILED(hres)) {
		pLoc->Release();
		CoUninitialize();
	}

	hres = CoSetProxyBlanket(
		pSvc,                        // Indicates the proxy to set
		RPC_C_AUTHN_WINNT,           // RPC_C_AUTHN_xxx
		RPC_C_AUTHZ_NONE,            // RPC_C_AUTHZ_xxx
		nullptr,                     // Server principal name 
		RPC_C_AUTHN_LEVEL_CALL,      // RPC_C_AUTHN_LEVEL_xxx 
		RPC_C_IMP_LEVEL_IMPERSONATE, // RPC_C_IMP_LEVEL_xxx
		nullptr,                     // client identity
		EOAC_NONE                    // proxy capabilities 
	);

	if (FAILED(hres)) {
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
	}

	IEnumWbemClassObject* pEnumerator = nullptr;
	hres = pSvc->ExecQuery(
		bstr_t("WQL"),
		bstr_t("SELECT startmode FROM Win32_Service"),
		WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
		nullptr,
		&pEnumerator);

	if (FAILED(hres)) {
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
	}

	IWbemClassObject *pclsObj = nullptr;
	ULONG uReturn = 0;

	pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);

	VARIANT vtProp;
	pclsObj->Get(L"startmode", 0, &vtProp, nullptr, nullptr);
	CString result = vtProp.bstrVal;
	VariantClear(&vtProp);
	pclsObj->Release();

	pSvc->Release();
	pLoc->Release();
	pEnumerator->Release();
	CoUninitialize();
	return result;
}

Sysutils::~Sysutils()
{
}
