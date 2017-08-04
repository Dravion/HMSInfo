#include "stdafx.h"
#include "HMSInfo.h"
#include "HMSInfoDlg.h"
#include "afxdialogex.h"
#include <windows.h>
#include <Winver.h>
#include <iostream>
#include <vector>
#include "versioninfo.h"
#include "sysutils.h"
#include <ctime>
#include <fstream>
#include "ipenum.h"
#include "hMailServerInfo.h"
#include "CredentialsFrm.h"
#include <netfw.h>

using namespace std;

// Globals
CString aIps;
CString hmsUsername = L"Administrator";
CString hmsPassword = nullptr;

// Firewall info forward declarations
HRESULT WFCOMInitialize(INetFwPolicy2** ppNetFwPolicy2);

// Get Firewall info 
void CHMSInfoDlg::Get_FirewallSettings_PerProfileType(NET_FW_PROFILE_TYPE2 ProfileTypePassed, INetFwPolicy2* pNetFwPolicy2)
{
	VARIANT_BOOL bIsEnabled = FALSE;
	NET_FW_ACTION action;

	if (SUCCEEDED(pNetFwPolicy2->get_FirewallEnabled(ProfileTypePassed, &bIsEnabled)))
	{
		CString fw;
		fw.Format(L"%s %hs", L"Firewall is:", bIsEnabled?"Enabled":"Disabled");
		this->LIST_VAR_hMailServer.AddString(fw);
	}

	if (SUCCEEDED(pNetFwPolicy2->get_BlockAllInboundTraffic(ProfileTypePassed, &bIsEnabled)))
	{		
		CString fw;
		fw.Format(L"%s %hs", L"Block all inbound traffic is:", bIsEnabled ? "Enabled" : "Disabled");
		this->LIST_VAR_hMailServer.AddString(fw);
	}

	if (SUCCEEDED(pNetFwPolicy2->get_NotificationsDisabled(ProfileTypePassed, &bIsEnabled)))
	{	
		CString fw;
		fw.Format(L"%s %hs", L"Notifications are:", bIsEnabled ?  "disabled" : "enabled");
		this->LIST_VAR_hMailServer.AddString(fw);
	}

	if (SUCCEEDED(pNetFwPolicy2->get_UnicastResponsesToMulticastBroadcastDisabled(ProfileTypePassed, &bIsEnabled)))
	{
		CString fw;
		fw.Format(L"%s %hs", L"UnicastResponsesToMulticastBroadcast is:", bIsEnabled ? "disabled" : "enabled");
		this->LIST_VAR_hMailServer.AddString(fw);
	}

	if (SUCCEEDED(pNetFwPolicy2->get_DefaultInboundAction(ProfileTypePassed, &action)))
	{		
		CString fw;
		fw.Format(L"%s %hs", L"Default inbound action is:", action != NET_FW_ACTION_BLOCK ? "Allow" : "Block");
		this->LIST_VAR_hMailServer.AddString(fw);
	}

	if (SUCCEEDED(pNetFwPolicy2->get_DefaultOutboundAction(ProfileTypePassed, &action)))
	{		
		CString fw;
		fw.Format(L"%s %hs", L"Default outbound action is:", action != NET_FW_ACTION_BLOCK ? "Allow" : "Block");
		this->LIST_VAR_hMailServer.AddString(fw);
	}
	
}

// Instantiate INetFwPolicy2
HRESULT WFCOMInitialize(INetFwPolicy2** ppNetFwPolicy2)
{
	HRESULT hr;

	hr = CoCreateInstance(
		__uuidof(NetFwPolicy2),
		nullptr,
		CLSCTX_INPROC_SERVER,
		__uuidof(INetFwPolicy2),
		(void**)ppNetFwPolicy2);

	if (FAILED(hr))
	{
		printf("CoCreateInstance for INetFwPolicy2 failed: 0x%08lx\n", hr);
		goto Cleanup;
	}

Cleanup:

	return hr;
}

class CMyIPEnum : public CIPEnum
{
	BOOL EnumCallbackFunction(int nAdapter, const in_addr& address) override;
};

BOOL CMyIPEnum::EnumCallbackFunction(int nAdapter, const in_addr& address)
{
	CString ips = nullptr;
	ips.Format(L"%d.%d.%d.%d", address.S_un.S_un_b.s_b1,
		address.S_un.S_un_b.s_b2, address.S_un.S_un_b.s_b3, address.S_un.S_un_b.s_b4);
		aIps = aIps + ips + L"\t/\t";
	
	return TRUE;
}

// CAboutDlg dialog used for App About
class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX) override;    // DDX/DDV support

// Implementation
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(IDD_ABOUTBOX)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()

CHMSInfoDlg::CHMSInfoDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(IDD_HMSINFO_DIALOG, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CHMSInfoDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_INFO_LIST, VAR_INFO_LIST);
	DDX_Control(pDX, IDC_TAB1, MyTabCtrl);
	DDX_Control(pDX, IDC_LIST_hMailServer, LIST_VAR_hMailServer);

}

BEGIN_MESSAGE_MAP(CHMSInfoDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()	
	ON_BN_CLICKED(ID_BT_QUIT, &CHMSInfoDlg::OnBnClickedBtQuit)
	ON_BN_CLICKED(IDC_BUTTON1, &CHMSInfoDlg::bt_scan)	
	ON_LBN_SELCHANGE(IDC_INFO_LIST, &CHMSInfoDlg::OnLbnSelchangeInfoList)
	ON_BN_CLICKED(IDC_BT_SAVE, &CHMSInfoDlg::OnBnClickedBtSave)
ON_NOTIFY(TCN_SELCHANGE, IDC_TAB1, &CHMSInfoDlg::OnSelchangeTab1)
END_MESSAGE_MAP()


// CHMSInfoDlg message handlers
BOOL CHMSInfoDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// Add "Info" menu item to system menu.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != nullptr)
	{
		
		CString strAboutMenu;
		strAboutMenu.LoadString(IDS_ABOUTBOX);

		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

    // Extra initialization

	MyTabCtrl.InsertItem(0, _T("General"), 1);
	MyTabCtrl.InsertItem(1, _T("COM && Firewall"), 2);
	
	// Listbox, larger fonts
	CFont m_font;
	m_font.CreateFont(25,              // Height
		2,                             // Width
		0,                             // Escapement
		0,                             // Orientation
		FW_BLACK,                      // Weight
		FALSE,                         // Italic
		FALSE,                         // Underline
		0,                             // StrikeOut
		DEFAULT_CHARSET,               // CharSet
		OUT_DEFAULT_PRECIS,            // OutPrecision
		CLIP_DEFAULT_PRECIS,           // ClipPrecision
		CLEARTYPE_QUALITY,             // Quality
		DEFAULT_PITCH | FF_SWISS,      // PitchAndFamily
		L"Arial");                     // Facename

	VAR_INFO_LIST.SetFont(&m_font);
	LIST_VAR_hMailServer.SetFont(&m_font);
	
	// Adjust size                           //width //height
	VAR_INFO_LIST.SetWindowPos(nullptr,20,51,575,    461,     0);
	
	

	

	// Adjust size
	this->LIST_VAR_hMailServer.SetWindowPos(nullptr, 20, 51, 575, 451, 0);
	
	this->LIST_VAR_hMailServer.AddString(L"No Data");
	
	#ifdef _DEBUG
		#ifdef _M_IX86						
			this->SetWindowTextW(L"HMSinfo (Debug) - 32-Bit");
		#endif 

		#ifdef _M_X64
			this->SetWindowTextW(L"HMSinfo (Debug) - 64-Bit");
		#endif 
	#else
		#ifdef _M_IX86						
			this->SetWindowTextW(L"HMSinfo (Release) - 32-Bit");
		#endif 

		#ifdef _M_X64
			this->SetWindowTextW(L"HMSinfo (Release) - 64-Bit");
		#endif 
	#endif

	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CHMSInfoDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

void CHMSInfoDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else {
		CDialogEx::OnPaint();
	}
}

HCURSOR CHMSInfoDlg::OnQueryDragIcon() const
{
	return static_cast<HCURSOR>(m_hIcon);
}

void CHMSInfoDlg::OnBnClickedBtQuit()
{
	CDialogEx::OnOK();
}

// Scan Button, collect all data
void CHMSInfoDlg::bt_scan()
{				
	Sysutils *obj = new Sysutils();	
	this->VAR_INFO_LIST.ResetContent();
	  
	// ##### Start WOW32 #####
	if (obj->IsHMS32_WOW64_Installed() == true) {
		
		CString varInstallDir = obj->GetHMS32WOW64InstallPath();
		CString releasedesc = obj->GetHMS32WOW64VersionInfo();

		if (varInstallDir.GetLength() < 2) {
			AfxMessageBox(L"No hMailServer found, finishing");
			AfxPostQuitMessage(0);
		}

		VAR_INFO_LIST.AddString(L"Version: " + releasedesc +" (32-Bit on WOW64)");	
		VAR_INFO_LIST.AddString(L"Installfolder: "+varInstallDir);
		
		CVersionInfo *verinfo
			= new CVersionInfo(varInstallDir + "\\Bin\\libeay32.dll");			
			CString v_fileversion = verinfo->GetFileversion();
			CString v_productversion = verinfo->GetProductVersion();
			int v_build = verinfo->GetBuildNumber();
			delete verinfo;

		CString msg;
		msg.Format(L"%s %d", static_cast<const wchar_t*>(v_fileversion), v_build);
		VAR_INFO_LIST.AddString(L"OpenSSL Version: " + msg + " (click for details...)");

		CString osbuild = nullptr;		
		osbuild.Format(L"Build: "+ obj->GetOSBuildEx());

		CString bitreparse;
		if (obj->GetOSArchitecture() == "32-bit")
			bitreparse = L"32-Bit";

		if (obj->GetOSArchitecture() == "64-bit")
			bitreparse = L"64-Bit";
		
		CString sp = nullptr;

		if (obj->GetOSServicePack() > 0) {
			sp.Format(L"%d", obj->GetOSServicePack());			
			VAR_INFO_LIST.AddString(L"OS: "+ obj->GetOSDesc() +" "+ osbuild +" "+ bitreparse +" SP "+ sp);
		}
		else {
			VAR_INFO_LIST.AddString(L"OS: "+ obj->GetOSDesc() +" "+ osbuild +" "+ bitreparse +" RTM");
		}

		CString cpucores = nullptr;
		cpucores.Format(L"%d", obj->GetCPUCores());

		CString totalmemory = nullptr;
		totalmemory.Format(L"%d", obj->GetTotalAvaiableMemory() / 1024);

		CString freememory = nullptr;
		freememory.Format(L"%d", obj->GetFreeMemory() / 1024);

		CString maxorocessmemory = nullptr;
		maxorocessmemory.Format(L"%d", obj->MaxProcessMemorySize() / 1024 / 1024);

		CString nicspeed = nullptr;
		nicspeed.Format(L"%d %s", obj->GetNicSpeed() / 1000 / 1000,L" MBit/s");
		
		CString nicenabled = nullptr;
		nicenabled.Format(L"%s", static_cast<const wchar_t*>(obj->GetNicIsEnabled()));

		VAR_INFO_LIST.AddString(L"CPU: " + obj->GetCPUInfo() +" "+ cpucores + " Cores");			
		VAR_INFO_LIST.AddString(L"Total Memory avaiable: " + totalmemory + " MBytes");
		VAR_INFO_LIST.AddString(L"Free Memory avaiable: " + freememory + " MBytes ("+obj->GetMemoryInUse() +"% in use)");
		VAR_INFO_LIST.AddString(L"Granted hMailServer Memory: " + maxorocessmemory + " MBytes");		
		VAR_INFO_LIST.AddString(L"Networkadapter: " + obj->GetNicInfo() +" "+nicspeed);
		VAR_INFO_LIST.AddString(L"Networkadapter: " + nicenabled + " / Description: "+obj->GetNicConnectionID());
		VAR_INFO_LIST.AddString(L"Networkadapter-MAC-Address: " + obj->GetMacAddress());
		VAR_INFO_LIST.AddString(L"Default Gateway: " + obj->GetDefaultGateWay());
			
		CMyIPEnum ip;
		ip.Enumerate();

		VAR_INFO_LIST.AddString(L"IP-Addresses: "+ aIps);		
		VAR_INFO_LIST.AddString(L"Virtualization: " + obj->GetIsVirtualBox());
		VAR_INFO_LIST.AddString(L"Database-" + obj->GetSQLDatabaseType(varInstallDir));
				
		CString svc_status = nullptr;
		svc_status.Format(L"hMailServer status: %hs", +obj->getServiceStatus());																				
		VAR_INFO_LIST.AddString(svc_status);
		
		CString svc_account = nullptr;
		svc_account.Format(L"Service account: %s", +obj->GetServiceAccountName());
		VAR_INFO_LIST.AddString(svc_account);
		
		CString svc_startmode = nullptr;
		svc_startmode.Format(L"Service Startmode: %s", +obj->GetServiceStartMode());
		VAR_INFO_LIST.AddString(svc_startmode);				
		VAR_INFO_LIST.AddString(L"Scan-Timestamp: " + obj->GetTimestamp());

	} // ##### Stop WOW32 #####
	
	// ##### Start Win32-Native #####
	if (obj->IsHMS32_native_installed() == true) {
		
		CString releasedesc = obj->GetHMS32NativeVersionInfo();
		CString vInstallDir = obj->GetHMS32InstallPath();
	
		if (vInstallDir.GetLength() < 2) {
			AfxMessageBox(L"No hMailServer found, finishing");
			AfxPostQuitMessage(0);			
		}

		VAR_INFO_LIST.AddString(L"Version: " + releasedesc + " (32-Bit, native)");
		VAR_INFO_LIST.AddString(L"Installfolder: " + vInstallDir);

		CString vInstallDate = obj->GetHMS32InstallDate();

		   vInstallDate.Insert(6, _T("."));
		   vInstallDate.Insert(4, _T("."));	
			 _tprintf(L"1: %s\n", (LPCTSTR)vInstallDate);

		VAR_INFO_LIST.AddString(L"HMS-Installdate: " + vInstallDate);

		CVersionInfo *verinfo
			= new CVersionInfo(vInstallDir + "\\Bin\\libeay32.dll");
			CString v_fileversion = verinfo->GetFileversion();
			CString v_productversion = verinfo->GetProductVersion();
			int v_build = verinfo->GetBuildNumber();
			delete verinfo;

			CString msg;
			msg.Format(L"%s %d", static_cast<const wchar_t*>(v_fileversion), v_build);
			VAR_INFO_LIST.AddString(L"OpenSSL Version: " + msg + " (click for details...)");

			CString cpucores = nullptr;
			cpucores.Format(L"%d", obj->GetCPUCores());

			CString totalmemory = nullptr;
			totalmemory.Format(L"%d", obj->GetTotalAvaiableMemory() / 1024);

			CString freememory = nullptr;
			freememory.Format(L"%d", obj->GetFreeMemory() / 1024);

			CString maxorocessmemory = nullptr;
			maxorocessmemory.Format(L"%d", obj->MaxProcessMemorySize() / 1024 / 1024);

			CString nicspeed = nullptr;
			nicspeed.Format(L"%d %s", obj->GetNicSpeed() / 1000 / 1000, L" MBit/s");
			
			CString nicenabled = nullptr;
			nicenabled.Format(L"%s", static_cast<const wchar_t*>(obj->GetNicIsEnabled()));

			CString osbuild = nullptr;
			osbuild.Format(L"Build %d", obj->GetOSBuild());

			CString bitreparse;
			if (obj->GetOSArchitecture() == "32-bit")
				bitreparse = L"32-Bit";

			if (obj->GetOSArchitecture() == "64-bit")
				bitreparse = L"64-Bit";

			CString sp = nullptr;

			if (obj->GetOSServicePack() > 0) {
				sp.Format(L"%d", obj->GetOSServicePack());
				VAR_INFO_LIST.AddString(L"OS: " + obj->GetOSDesc() + " " + osbuild + " " + bitreparse + " SP " + sp);
			}
			else {
				VAR_INFO_LIST.AddString(L"OS: " + obj->GetOSDesc() + " " + osbuild + " " + bitreparse + " RTM");
			}

			VAR_INFO_LIST.AddString(L"CPU: " + obj->GetCPUInfo() +cpucores+" Cores");
			VAR_INFO_LIST.AddString(L"Total RAM Memory installed: " + totalmemory + " MBytes");
			VAR_INFO_LIST.AddString(L"Free Memory avaiable: " + freememory + " MBytes ("
									+ obj->GetMemoryInUse() + "% in use)");
			VAR_INFO_LIST.AddString(L"Granted hMailServer Memory: " + maxorocessmemory + " MBytes");				
			VAR_INFO_LIST.AddString(L"Ethernet: " + obj->GetNicInfo() + " " + nicspeed);
			VAR_INFO_LIST.AddString(L"Ethernet: " + nicenabled);
			VAR_INFO_LIST.AddString(L"Virtualization: " + obj->GetIsVirtualBox());
			VAR_INFO_LIST.AddString(L"Database-" + obj->GetSQLDatabaseType(vInstallDir));			
			
	} // ##### End Win32-Native #####
		
	// ##### Start Win64-Native #####
	if (obj->IsHMS64_native_installed() == true) {		

		CString releasedesc = obj->GetHMS32NativeVersionInfo();
		CString vInstallDir = obj->GetHMS32InstallPath();

		if (vInstallDir.GetLength() < 2) {
			AfxMessageBox(L"No hMailServer found, finishing");
			AfxPostQuitMessage(0);
		}

		VAR_INFO_LIST.AddString(L"Version: " + releasedesc + " (32-Bit, native)");
		VAR_INFO_LIST.AddString(L"Installfolder: " + vInstallDir);

		CString vInstallDate = obj->GetHMS32InstallDate();

		vInstallDate.Insert(6, _T("."));
		vInstallDate.Insert(4, _T("."));
		_tprintf(L"1: %s\n", (LPCTSTR)vInstallDate);

		VAR_INFO_LIST.AddString(L"HMS-Installdate: " + vInstallDate);

		CVersionInfo *verinfo
			= new CVersionInfo(vInstallDir + "\\Bin\\libeay32.dll");
		CString v_fileversion = verinfo->GetFileversion();
		CString v_productversion = verinfo->GetProductVersion();
		int v_build = verinfo->GetBuildNumber();
		delete verinfo;

		CString msg;
		msg.Format(L"%s %d", static_cast<const wchar_t*>(v_fileversion), v_build);
		VAR_INFO_LIST.AddString(L"OpenSSL Version: " + msg + " (click for details...)");

		CString cpucores = nullptr;
		cpucores.Format(L"%d", obj->GetCPUCores());

		CString totalmemory = nullptr;
		totalmemory.Format(L"%d", obj->GetTotalAvaiableMemory() / 1024);

		CString freememory = nullptr;
		freememory.Format(L"%d", obj->GetFreeMemory() / 1024);

		CString maxorocessmemory = nullptr;
		maxorocessmemory.Format(L"%d", obj->MaxProcessMemorySize() / 1024 / 1024);

		CString nicspeed = nullptr;
		nicspeed.Format(L"%d %s", obj->GetNicSpeed() / 1000 / 1000, L" MBit/s");

		CString nicenabled = nullptr;
		nicenabled.Format(L"%s", static_cast<const wchar_t*>(obj->GetNicIsEnabled()));

		CString osbuild = nullptr;
		osbuild.Format(L"Build %d", obj->GetOSBuild());

		CString bitreparse;
		if (obj->GetOSArchitecture() == "32-bit")
			bitreparse = L"32-Bit";

		if (obj->GetOSArchitecture() == "64-bit")
			bitreparse = L"64-Bit";

		CString sp = nullptr;

		if (obj->GetOSServicePack() > 0) {
			sp.Format(L"%d", obj->GetOSServicePack());
			VAR_INFO_LIST.AddString(L"OS: " + obj->GetOSDesc() + " " + osbuild + " " + bitreparse + " SP " + sp);
		}
		else {
			VAR_INFO_LIST.AddString(L"OS: " + obj->GetOSDesc() + " " + osbuild + " " + bitreparse + " RTM");
		}

		VAR_INFO_LIST.AddString(L"CPU: " + obj->GetCPUInfo() + cpucores + " Cores");
		VAR_INFO_LIST.AddString(L"Total RAM Memory installed: " + totalmemory + " MBytes");
		VAR_INFO_LIST.AddString(L"Free Memory avaiable: " + freememory + " MBytes ("
			+ obj->GetMemoryInUse() + "% in use)");
		VAR_INFO_LIST.AddString(L"Granted hMailServer Memory: " + maxorocessmemory + " MBytes");
		VAR_INFO_LIST.AddString(L"Ethernet: " + obj->GetNicInfo() + " " + nicspeed);
		VAR_INFO_LIST.AddString(L"Ethernet: " + nicenabled);
		VAR_INFO_LIST.AddString(L"Virtualization: " + obj->GetIsVirtualBox());
		VAR_INFO_LIST.AddString(L"Database-" + obj->GetSQLDatabaseType(vInstallDir));

	}		
	delete obj;
} // ##### End - Win64-Native #####

void CHMSInfoDlg::OnLbnSelchangeInfoList() {

	try {
		CString strText;

		int index = this->VAR_INFO_LIST.GetCurSel();
		this->VAR_INFO_LIST.GetText(index, strText);
		
		int found = strText.Find(L"OpenSSL");
		if (found != -1) {
			const int result = MessageBoxA(m_hWnd, "Review known OpenSSL vulnerabilities online? ", "Confirm", MB_ICONQUESTION | MB_YESNO | MB_SYSTEMMODAL);

			switch (result)	{
			case IDYES:
				ShellExecute(nullptr, L"open", L"https://www.openssl.org/news/vulnerabilities.html", nullptr, nullptr, SW_SHOWNORMAL);
				break;
			case IDNO:		
				break;
			default:; // do nothing;
			}
		}
	}
	catch (exception e) {
		MessageBox(L"Not in list");
	}
}

void CHMSInfoDlg::OnBnClickedBtSave() {
	char strFilter[] = { "HMS Info Logfiles (*.log)|*.log|" };
	CFileDialog FileDlg(FALSE, CString(".log"), nullptr, OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT, CString(strFilter));

	if (FileDlg.DoModal() == IDOK)	{
	
		CString vFile = FileDlg.GetFolderPath() +"\\"+ FileDlg.GetFileName();
		CString listitems;

		USES_CONVERSION;
		
		char const* fFilename = T2CA(vFile);
		FILE *f = fopen(fFilename, "w");
		
		if (f == nullptr) {
			MessageBoxW(L"Error cannot create file!",L"Error",MB_ICONERROR);
			exit(1);
		}

		fprintf(f, "%S\n",L"### HMS-Info - Logfile ###");
		for (int i = 0; i < VAR_INFO_LIST.GetCount(); i++) {
			int n = VAR_INFO_LIST.GetTextLen(i);
			VAR_INFO_LIST.GetText(i, listitems.GetBuffer(n));
			listitems.ReleaseBuffer();
			fprintf(f, "%S\n", listitems.GetBuffer(0));
		}
		fclose(f);
	}
}

// Event fired if tabs are selected and changed
void CHMSInfoDlg::OnSelchangeTab1(NMHDR *pNMHDR, LRESULT *pResult)
{
	int iSel = MyTabCtrl.GetCurSel();

	if (iSel == 0)
	{		
		GetDlgItem(IDC_INFO_LIST)->ShowWindow(SW_SHOW);		
		GetDlgItem(IDC_LIST_hMailServer)->ShowWindow(SW_HIDE);
	}

	else if (iSel == 1)
	{
		GetDlgItem(IDC_INFO_LIST)->ShowWindow(SW_HIDE);		

		// Delete all the items
		LIST_VAR_hMailServer.ResetContent();
		hMailServerInfo *obj = new hMailServerInfo();		
		
		if (hmsPassword.GetLength() == 0) {
			CredentialsFrm dlg = new CredentialsFrm();

			dlg.DoModal();
			hmsPassword = dlg.hPwd;
			hmsUsername = dlg.hUsr;
		}
		
		   int result = obj->MaxDeliveryThreads(hmsUsername, hmsPassword);
		   int  i_processedMessages = obj->ProcessedMessages(hmsUsername, hmsPassword);
		   CString i_startuptime = obj->StartTime(hmsUsername, hmsPassword);
		   CString dbconnections = obj->IsDBConnected(hmsUsername, hmsPassword);
		   int unCount = obj->getUndeliveredMessages(hmsUsername, hmsPassword);		   		   

		   if (result != 0) {
				
				CString hVersion;
				hVersion.Format(L"Build: "+ obj->Version());

				LIST_VAR_hMailServer.AddString(hVersion);

				CString maxth;
				maxth.Format(L"%d", result);

				CString i_mproc;
				i_mproc.Format(L"%d", i_processedMessages);				
				
				CString umCount;
			    umCount.Format(L"%d", unCount);															
				
				LIST_VAR_hMailServer.AddString(L"MaxThreads: " + maxth);
				LIST_VAR_hMailServer.AddString(L"ProcessedMessages: " + i_mproc);
				LIST_VAR_hMailServer.AddString(L"Undelivered Messages: " + umCount);
				LIST_VAR_hMailServer.AddString(L"Startuptime: " + i_startuptime);				
				LIST_VAR_hMailServer.AddString(L"DB-Connection: " + dbconnections);
																	
				HRESULT hrComInit;
				HRESULT hr;

				INetFwPolicy2 *pNetFwPolicy2 = nullptr;

				// Initialize COM.
				hrComInit = CoInitializeEx(
					nullptr,
					COINIT_APARTMENTTHREADED
				);

				if (hrComInit != RPC_E_CHANGED_MODE)
				{
					if (FAILED(hrComInit))
					{
						printf("CoInitializeEx failed: 0x%08lx\n", hrComInit);
						goto Cleanup;
					}
				}

				// Retrieve INetFwPolicy2
				hr = WFCOMInitialize(&pNetFwPolicy2);
				if (FAILED(hr))
				{
					goto Cleanup;
				}

				printf("Settings for the firewall public profile:\n");
				Get_FirewallSettings_PerProfileType(NET_FW_PROFILE2_PUBLIC, pNetFwPolicy2);

			Cleanup:

				// Release INetFwPolicy2
				if (pNetFwPolicy2 != nullptr)
				{
					pNetFwPolicy2->Release();
				}

				// Uninitialize COM.
				if (SUCCEEDED(hrComInit))
				{
					CoUninitialize();
				}
											
			} else	{
				LIST_VAR_hMailServer.AddString(L"Error: Username/Password wrong");
			}
		}				
		
		GetDlgItem(IDC_LIST_hMailServer)->ShowWindow(SW_SHOW);				

	*pResult = 0;
}
