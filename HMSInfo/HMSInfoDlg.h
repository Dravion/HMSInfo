#pragma once
#include "afxwin.h"
#include "afxcmn.h"
#include <netfw.h>

// CHMSInfoDlg dialog
class CHMSInfoDlg : public CDialogEx
{
// Construction
public:
	void Get_FirewallSettings_PerProfileType(NET_FW_PROFILE_TYPE2 ProfileTypePassed, INetFwPolicy2* pNetFwPolicy2);
	CHMSInfoDlg(CWnd* pParent = nullptr);	// standard constructor

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_HMSINFO_DIALOG };
#endif

	protected:
	void DoDataExchange(CDataExchange* pDX) override;	// DDX/DDV support
	
// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	BOOL OnInitDialog() override;
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon() const;
	DECLARE_MESSAGE_MAP()
public:

	afx_msg void OnBnClickedBtQuit();
	afx_msg void bt_scan();
	CListBox VAR_INFO_LIST;
	afx_msg void OnLbnSelchangeInfoList();
	afx_msg void OnBnClickedBtSave();	
	CTabCtrl tab_main;
	CTabCtrl TabCtrlMain;
	CTabCtrl MyTabCtrl;
	afx_msg void OnSelchangeTab1(NMHDR *pNMHDR, LRESULT *pResult);
	CListBox LIST_VAR_hMailServer;
	CStatic lb_status;
};
