#ifndef _VERSIONINFO_H 
#define _VERSIONINFO_H 
#include "stdafx.h"
//#include <bemapiset.h>


// ************************************************************************** 
class CVersionInfo
{
public:
	CVersionInfo(LPCTSTR lpszAppName);	// If you know the filename 
	CVersionInfo(HINSTANCE hInst);			// If the file is in memory 

	WORD GetMajorVersion() { return (HIWORD(m_stFixedInfo.dwFileVersionMS)); }
	WORD GetMinorVersion() { return (LOWORD(m_stFixedInfo.dwFileVersionMS)); }
	WORD GetBuildNumber() { return (LOWORD(m_stFixedInfo.dwFileVersionLS)); }

	LPCTSTR GetDescription() { return (m_strDescription); }
	LPCTSTR GetComments() { return (m_strComments); }
	LPCTSTR GetCompany() { return (m_strCompany); }
	LPCTSTR GetProductName() { return (m_strProductName); }
	LPCTSTR GetInternalName() { return (m_strInternalName); }
	LPCTSTR GetLegalCopyright() { return (m_strLegalCopyright); }
	LPCTSTR GetFileversion() { return (m_strFileVersion); }
	LPCTSTR GetProductVersion() { return (m_strProductVersion); }


	void Format(CString &strOutput);

	const VS_FIXEDFILEINFO &GetFixedInfo() { return (m_stFixedInfo); }

protected:
	void Initialize(LPCTSTR lpszAppName);

	VS_FIXEDFILEINFO m_stFixedInfo;
	CString m_strDescription;
	CString m_strComments;
	CString m_strCompany;
	CString m_strProductName;
	CString m_strInternalName;
	CString m_strLegalCopyright;
	CString m_strFileVersion;
	CString m_strProductVersion;
};


#endif	// _VERSIONINFO_H 
