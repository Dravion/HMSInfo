#include "stdafx.h"
#include "HMSInfo.h"
#include "CredentialsFrm.h"
#include "afxdialogex.h"


IMPLEMENT_DYNAMIC(CredentialsFrm, CDialogEx)

CredentialsFrm::CredentialsFrm(CWnd* pParent /*=NULL*/)
	: CDialogEx(IDD_CredentialsFrm, pParent)
{
}

CredentialsFrm::~CredentialsFrm()
{
}

void CredentialsFrm::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	SetDlgItemText(IDC_ed_username, L"Administrator");
	DDX_Control(pDX, IDC_ed_username, ed_username);
	DDX_Control(pDX, IDC_ed_password, ed_password);

	HWND hWnd;
	GetDlgItem(IDC_ed_password, &hWnd);
	::PostMessage(hWnd, WM_SETFOCUS, 0, 0);

}

BEGIN_MESSAGE_MAP(CredentialsFrm, CDialogEx)	
	ON_COMMAND(IDOK, OnOK)	
	ON_BN_CLICKED(IDOK, &CredentialsFrm::OnBnClickedOk)
END_MESSAGE_MAP()

void CredentialsFrm::OnOK()
{
	this->ed_username.GetWindowTextW(hUsr);
	this->ed_password.GetWindowTextW(hPwd);
	CDialogEx::OnOK();
}

void CredentialsFrm::OnBnClickedOk()
{	
	CDialogEx::OnOK();
}
