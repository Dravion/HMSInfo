#include "stdafx.h"   
#include "versioninfo.h"   

#ifndef _WIN32_WCE   
#pragma comment (lib, "Version.lib")   
#endif   


CVersionInfo::CVersionInfo(LPCTSTR lpAppName)
{
	// Initialize member variables:   
	ASSERT(lpAppName);
	Initialize(lpAppName);
}
 
CVersionInfo::CVersionInfo(HINSTANCE hInst)
{
	// Define buffer for application name:   
	TCHAR szAppName[_MAX_PATH + 1];

	// Get the application name (name of .exe file):   
	if (!GetModuleFileName(hInst, szAppName, _countof(szAppName)))
	{
		// Failed to get name.  Set string to NULL and return:   
		TRACE(_T("GetModuleFileName () failed, OS Error == %08X\n"), GetLastError());
		ZeroMemory(&m_stFixedInfo, sizeof(m_stFixedInfo));
		return;
	}

	// Initialize member variables:   
	Initialize(szAppName);
}
 
void CVersionInfo::Initialize(LPCTSTR szAppName)
{
	ASSERT(szAppName);

	// Initialize fixed info structure:   
	ZeroMemory(&m_stFixedInfo, sizeof(m_stFixedInfo));

	// Load the Version information   
	// Load signon Version information   
	DWORD dwHandle;
	DWORD dwSize;

	// Determine the size of the VERSIONINFO resource:   
	if (!((dwSize = GetFileVersionInfoSize((LPTSTR)szAppName, &dwHandle))))
	{
		TRACE(_T("GetFileVesionInfoSize () failed on %s"), szAppName);
		return;
	}

	// Declare pointer to Version info resource:   
	LPBYTE lpVerInfo = nullptr;

#ifdef _WIN32_WCE   
	TRY
#else   
	try
#endif   
	{
		// Allocate memory to hold Version info:   
		lpVerInfo = new BYTE[dwSize];

		// Read the VERSIONINFO resource from the file:   
		if (GetFileVersionInfo((LPTSTR)szAppName, dwHandle, dwSize, lpVerInfo))
		{
			VS_FIXEDFILEINFO *lpFixedInfo;
			LPCTSTR lpText;
			UINT uSize;

			// Read the FILE DESCRIPTION:   
			if (VerQueryValue(lpVerInfo, _T("\\StringFileInfo\\040904B0\\FileDescription"), (void **)&lpText, &uSize))
				m_strDescription = lpText;

			// Read additional comment:   
			if (VerQueryValue(lpVerInfo, _T("\\StringFileInfo\\040904B0\\Comments"), (void **)&lpText, &uSize))
				m_strComments = lpText;

			// Read company name:   
			if (VerQueryValue(lpVerInfo, _T("\\StringFileInfo\\040904B0\\CompanyName"), (void **)&lpText, &uSize))
				m_strCompany = lpText;

			// Read product name:   
			if (VerQueryValue(lpVerInfo, _T("\\StringFileInfo\\040904B0\\ProductName"), (void **)&lpText, &uSize))
				m_strProductName = lpText;

			// Read internal name:   
			if (VerQueryValue(lpVerInfo, _T("\\StringFileInfo\\040904B0\\InternalName"), (void **)&lpText, &uSize))
				m_strInternalName = lpText;

			// Read legal copyright:   
			if (VerQueryValue(lpVerInfo, _T("\\StringFileInfo\\040904B0\\LegalCopyright"), (void **)&lpText, &uSize))
				m_strLegalCopyright = lpText;

			// Read file Version 
			if (VerQueryValue(lpVerInfo, _T("\\StringFileInfo\\040904B0\\FileVersion"), (void **)&lpText, &uSize))
				m_strFileVersion = lpText;

			// Read product Version 
			if (VerQueryValue(lpVerInfo, _T("\\StringFileInfo\\040904B0\\ProductVersion"), (void **)&lpText, &uSize))
				m_strProductVersion = lpText;

			

			// Read the FIXEDINFO portion:   
			if (VerQueryValue(lpVerInfo, _T("\\"), (void **)&lpFixedInfo, &uSize))
				m_stFixedInfo = *lpFixedInfo;
		}
	}

	// Handle exceptions:   
#ifdef _WIN32_WCE   
	CATCH(CException, e)
	{
		e->Delete();
#else   
	catch (...)
	{
#endif   
		// EXCEPTIONMSG();
	}

#ifdef _WIN32_WCE   
	END_CATCH
#endif   

		// Free memory allocated for Version info:   
		delete[] lpVerInfo;
	}

void CVersionInfo::Format(CString &strOutput)
{
	// Append Version number information to the sign on string:   
	strOutput.Format(_T("V%d.%d%d.%d"),
		HIWORD(m_stFixedInfo.dwFileVersionMS),
		LOWORD(m_stFixedInfo.dwFileVersionMS),
		HIWORD(m_stFixedInfo.dwFileVersionLS),
		LOWORD(m_stFixedInfo.dwFileVersionLS));

	// Add UNICODE tag:   
#ifndef _WIN32_WCE   
#ifdef _UNICODE   
	strOutput += _T(" - U");
#endif   
#endif   

	// include a debug indication:   
	if (m_stFixedInfo.dwFileFlags & VS_FF_DEBUG)
		strOutput += _T(" (Debug)");
}