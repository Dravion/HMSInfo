#pragma once


// hMailServerInfo

class hMailServerInfo : public CWnd
{
	DECLARE_DYNAMIC(hMailServerInfo)

public:
	hMailServerInfo();
	virtual ~hMailServerInfo();
	CString Version();

protected:
	DECLARE_MESSAGE_MAP()
	
public:
	// Prototypes
	int MaxDeliveryThreads(CString pUsr, CString pPwd);
	int ProcessedMessages(CString pUsr, CString pPwd);
	int getUndeliveredMessages(CString pUsr, CString pPwd);
	CString StartTime(CString pUsr, CString pPwd);
	int UserSessions(CString pUsr, CString pPwd);
	CString IsDBConnected(CString pUsr, CString pPwd);
	CString getClientConnectedName(CString pUsr, CString pPwd);
};


