#pragma once
#include <string>
#include <iostream>

using namespace std;

class Sysutils
{
public:
	Sysutils();
	int gBitness; 
	CString GetHMS32WOW64VersionInfo();
	CString GetHMS64VersionInfo();
	int GetTotalAvaiableMemory();

	int GetFreeMemory();
	CString GetOSDesc();
	CString GetCPUInfo();
	CString GetNicInfo();
	int GetNicSpeed();
	CString GetNicIsEnabled();
	CString GetNicConnectionID();
	CString GetDNSInfo();
	CString GetMacAddress();
	CString GetDefaultGateWay();
	CString GetDNSServers();
	CString GetTimestamp();
	CString GetMemoryInUse();
	CString GetOSArchitecture();	
	CString GetBinaryBitness(LPCTSTR lpApplicationName);
	int GetMyBitness();
	int GetOSServicePack();
	int GetOSBuild();
	CString GetOSBuildEx();
	int GetCPUCores();
	int MaxProcessMemorySize();
	CString GetHMS32WOW64InstallPath();
	CString GetHMS64InstallPath();
	CString GetHMS32InstallPath();
	CString GetHMS32InstallDate();
	CString GetHMS64NativeVersionInfo();
	CString GetHMS32NativeVersionInfo();
	CString GetHMSInfoPath();
	BOOL IsWow64();
	bool IsNative64Bit();
	bool IsHMS32_WOW64_Installed();
	bool IsHMS32_native_installed();
	bool IsHMS64_native_installed();
	CString GetSQLDatabaseType(CString iniFilepath);
	CString GetIsVirtualBox();
	int GetHMSBitness();
	char* getServiceStatus();
	CString GetServiceAccountName();
	CString GetServiceStartMode();

	~Sysutils();
};

