// hMailServerInfo.cpp : implementation file

#include "stdafx.h"
#include "hMailServer.h"
#include <hMailServer_i.c>
#include "HMSInfo.h"
#include "hMailServerInfo.h"
#include <stdio.h>
#include <iostream>
#include <comutil.h>

IMPLEMENT_DYNAMIC(hMailServerInfo, CWnd)

hMailServerInfo::hMailServerInfo()
{
}

hMailServerInfo::~hMailServerInfo()
{
}

BEGIN_MESSAGE_MAP(hMailServerInfo, CWnd)
END_MESSAGE_MAP()

/* COM - Get hMailServer Version */
CString hMailServerInfo::Version()
{

	HRESULT hr;
	IInterfaceApplication *app;
	BSTR pVal = nullptr;

	hr = CoInitialize(nullptr);
	if (SUCCEEDED(hr)) {
		CoCreateInstance(
			CLSID_Application,
			nullptr,
			CLSCTX_LOCAL_SERVER,
			IID_IInterfaceApplication,
			(void**)&app);		

		app->get_Version(&pVal);
		app->Release();
	}
	
	return (pVal);
}

// COM - Returns Max Delivery Threads
int hMailServerInfo::MaxDeliveryThreads(CString pUsr, CString pPwd)
{

	long pVal = 0;

	try  {
	
	HRESULT hr;
	IInterfaceApplication *oApp = nullptr;

	hr = CoInitialize(nullptr);
	
	if (SUCCEEDED(hr)) {
		CoCreateInstance(
			CLSID_Application,
			nullptr,
			CLSCTX_LOCAL_SERVER,
			IID_IInterfaceApplication,
			(void**)&oApp);

		BSTR usr = SysAllocString((_bstr_t)pUsr);
		BSTR pwd = SysAllocString((_bstr_t)pPwd);
		IInterfaceAccount *pAcc;


			oApp->Authenticate(usr, pwd, &pAcc);

			if (pAcc) {
				IInterfaceSettings *pSettings;
				oApp->get_Settings(&pSettings);

				pSettings->get_MaxDeliveryThreads(&pVal);

			}
			else {
				AfxMessageBox(L"Login wrong, try again");
			}


		}
		
		oApp->Release();

	}

	catch (std::exception & stdExp)	 // sample exception handler
	{
		// Show error message
		::MessageBoxA(nullptr, stdExp.what(), "hMailServer reported a internal error", MB_ICONEXCLAMATION | MB_OK);

		return 1;   // return to exit the application

	}
	catch (...)	   // catches exception of types not caught above
	{
		pVal = 0;
		::MessageBoxA(nullptr, "hMailServer reported a fatal error", "Fatal Error", MB_ICONEXCLAMATION | MB_OK);

		return 1;   // return to exit the application

	}

	return pVal;
}

// Query processed messages
int hMailServerInfo::ProcessedMessages(CString pUsr, CString pPwd)
{
	long pVal = 0;
	HRESULT hr;
	IInterfaceApplication *oApp = nullptr;

	hr = CoInitialize(nullptr);

	if (SUCCEEDED(hr)) {
		CoCreateInstance(
			CLSID_Application,
			nullptr,
			CLSCTX_LOCAL_SERVER,
			IID_IInterfaceApplication,
			(void**)&oApp);

		BSTR usr = SysAllocString((_bstr_t)pUsr);
		BSTR pwd = SysAllocString((_bstr_t)pPwd);
		IInterfaceAccount *pAcc;

		oApp->Authenticate(usr, pwd, &pAcc);

		if (pAcc) {
			IInterfaceStatus *pStatus;
			oApp->get_Status(&pStatus);
			pStatus->get_ProcessedMessages(&pVal);

			
		}	
	}
	oApp->Release();
	return pVal;	
}

// Query processed messages
int hMailServerInfo::getUndeliveredMessages(CString pUsr, CString pPwd)
{
	BSTR pVal;
	HRESULT hr;
	IInterfaceApplication *oApp = nullptr;

	hr = CoInitialize(nullptr);

	if (SUCCEEDED(hr)) {
		CoCreateInstance(
			CLSID_Application,
			nullptr,
			CLSCTX_LOCAL_SERVER,
			IID_IInterfaceApplication,
			(void**)&oApp);

		BSTR usr = SysAllocString((_bstr_t)pUsr);
		BSTR pwd = SysAllocString((_bstr_t)pPwd);
		IInterfaceAccount *pAcc;

		oApp->Authenticate(usr, pwd, &pAcc);

		if (pAcc) {
			IInterfaceStatus *pStatus;
			oApp->get_Status(&pStatus);			
			pStatus->get_UndeliveredMessages(&pVal);

		}
	}
	oApp->Release();
	return 0;
}

// COM Query hMailServer Startuptime
CString hMailServerInfo::StartTime(CString pUsr, CString pPwd)
{
	BSTR pVal = nullptr;
	HRESULT hr;
	IInterfaceApplication *oApp = nullptr;

	hr = CoInitialize(nullptr);

	if (SUCCEEDED(hr)) {
		CoCreateInstance(
			CLSID_Application,
			nullptr,
			CLSCTX_LOCAL_SERVER,
			IID_IInterfaceApplication,
			(void**)&oApp);

		BSTR usr = SysAllocString((_bstr_t)pUsr);
		BSTR pwd = SysAllocString((_bstr_t)pPwd);
		IInterfaceAccount *pAcc;

		oApp->Authenticate(usr, pwd, &pAcc);

		if (pAcc) {
			IInterfaceStatus *pStatus;
			oApp->get_Status(&pStatus);
			pStatus->get_StartTime(&pVal);
		}
	}
	oApp->Release();
	return pVal;
}

// COM Query connect SMTP-Clients
int hMailServerInfo::UserSessions(CString pUsr, CString pPwd)
{
	long pVal = 0;
	HRESULT hr;
	IInterfaceApplication *oApp = nullptr;

	hr = CoInitialize(nullptr);

	if (SUCCEEDED(hr)) {
		CoCreateInstance(
			CLSID_Application,
			nullptr,
			CLSCTX_LOCAL_SERVER,
			IID_IInterfaceApplication,
			(void**)&oApp);

		BSTR usr = SysAllocString((_bstr_t)pUsr);
		BSTR pwd = SysAllocString((_bstr_t)pPwd);
		IInterfaceAccount *pAcc;

		oApp->Authenticate(usr, pwd, &pAcc);

		if (pAcc) {
			IInterfaceStatus *pStatus;
			oApp->get_Status(&pStatus);
			pStatus->get_SessionCount(eSTSMTPClient, &pVal);
		}
	}
	oApp->Release();
	return pVal;
}


// COM Query hMailServer Startuptime
CString hMailServerInfo::IsDBConnected(CString pUsr, CString pPwd)
{
	VARIANT_BOOL pVal = VARIANT_FALSE;
	CString iState;
	HRESULT hr;
	IInterfaceApplication *oApp = nullptr;

	hr = CoInitialize(nullptr);

	if (SUCCEEDED(hr)) {
		CoCreateInstance(
			CLSID_Application,
			nullptr,
			CLSCTX_LOCAL_SERVER,
			IID_IInterfaceApplication,
			(void**)&oApp);

		BSTR usr = SysAllocString((_bstr_t)pUsr);
		BSTR pwd = SysAllocString((_bstr_t)pPwd);
		IInterfaceAccount *pAcc;

		oApp->Authenticate(usr, pwd, &pAcc);

		if (pAcc) {
			IInterfaceDatabase *pDatabase;
			oApp->get_Database(&pDatabase);
			pDatabase->get_IsConnected(&pVal);
			
		
		if (pVal == VARIANT_TRUE) {
			  iState = "Connected to SQL-Database";
			}
			else {				
		   	  iState = "NOT connected to SQL-Database";
			}
		}

		oApp->Release();
		
	}
	 return iState;
}

// COM Query hMailServer Startuptime
CString hMailServerInfo::getClientConnectedName(CString pUsr, CString pPwd)
{
	
	

	try {

		HRESULT hr;
		IInterfaceApplication *oApp = nullptr;

		hr = CoInitialize(nullptr);

		if (SUCCEEDED(hr)) {
			CoCreateInstance(
				CLSID_Application,
				nullptr,
				CLSCTX_LOCAL_SERVER,
				IID_IInterfaceApplication,
				(void**)&oApp);

			BSTR usr = SysAllocString((_bstr_t)pUsr);
			BSTR pwd = SysAllocString((_bstr_t)pPwd);
			IInterfaceAccount *pAcc;

			oApp->Authenticate(usr, pwd, &pAcc);

			if (pAcc) {
				IInterfaceUtilities * pUtilities;				
				oApp->get_Utilities(&pUtilities);


				

			
				CString tmp;
				//tmp.Format(foo);
				
				//AfxMessageBox(foo);

	
				
			/*	BSTR a = nullptr;*/
	/*			pClient->get_Username(&a);*/

				// pSettings->get_MaxDeliveryThreads(&pVal);

			}
			else {
				AfxMessageBox(L"Login wrong, try again");
			}

		}
		oApp->Release();
	}

	catch (std::exception & stdExp)	 // sample exception handler
	{
		// Show error message
		::MessageBoxA(nullptr, stdExp.what(), "hMailServer reported a internal error", MB_ICONEXCLAMATION | MB_OK);

	}
	catch (...)	   // catches exception of types not caught above
	{
		::MessageBoxA(nullptr, "hMailServer reported a fatal error", "Fatal Error", MB_ICONEXCLAMATION | MB_OK);		

	}
	
	// return (CString)pVal);
	return nullptr;
}

