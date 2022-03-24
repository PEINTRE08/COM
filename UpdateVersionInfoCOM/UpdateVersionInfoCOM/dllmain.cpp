// dllmain.cpp: DllMain 的实现。

#include "stdafx.h"
#include "resource.h"
#include "UpdateVersionInfoCOM_i.h"
#include "dllmain.h"
#include "compreg.h"

CUpdateVersionInfoCOMModule _AtlModule;

class CUpdateVersionInfoCOMApp : public CWinApp
{
public:

// 重写
	virtual BOOL InitInstance();
	virtual int ExitInstance();

	DECLARE_MESSAGE_MAP()
};

BEGIN_MESSAGE_MAP(CUpdateVersionInfoCOMApp, CWinApp)
END_MESSAGE_MAP()

CUpdateVersionInfoCOMApp theApp;

BOOL CUpdateVersionInfoCOMApp::InitInstance()
{
	return CWinApp::InitInstance();
}

int CUpdateVersionInfoCOMApp::ExitInstance()
{
	return CWinApp::ExitInstance();
}
