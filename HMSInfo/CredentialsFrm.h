#pragma once
#include "afxwin.h"


// CredentialsFrm dialog

class CredentialsFrm : public CDialogEx
{
	DECLARE_DYNAMIC(CredentialsFrm)

public:
	CredentialsFrm(CWnd* pParent = nullptr);   // standard constructor
	virtual ~CredentialsFrm();

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_CredentialsFrm };
#endif

protected:
	void DoDataExchange(CDataExchange* pDX) override;    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	CEdit ed_username;
	CEdit ed_password;
	CString hPwd;
	CString hUsr;
	//afx_msg void bt_scan();
	//afx_msg void OnEnChangeedusername();
	void OnOK() override;
	afx_msg void OnBnClickedOk();
};
