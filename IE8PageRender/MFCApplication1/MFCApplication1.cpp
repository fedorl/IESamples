
// MFCApplication1.cpp : Defines the class behaviors for the application.
//

#include "stdafx.h"
#include "MFCApplication1.h"
#include "MFCApplication1Dlg.h"
#include "afxpriv.h"
//#include <..\src\occimpl.h>
#include "ControlSite.h"
#include <sstream>

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CMFCApplication1App

BEGIN_MESSAGE_MAP(CMFCApplication1App, CWinApp)
	//ON_COMMAND(ID_HELP, &CWinApp::OnHelp)
END_MESSAGE_MAP()


// CMFCApplication1App construction

CMFCApplication1App::CMFCApplication1App()
{
	// support Restart Manager
	m_dwRestartManagerSupportFlags = AFX_RESTART_MANAGER_SUPPORT_RESTART;

	// TODO: add construction code here,
	// Place all significant initialization in InitInstance
}


// The one and only CMFCApplication1App object

CMFCApplication1App theApp;


// CMFCApplication1App initialization


BOOL SetupIEVersion()
{
	wchar_t pFileName[2048];
	memset(pFileName, 0, sizeof(pFileName));
	if (!GetModuleFileName(NULL, pFileName, 2048))
	{
		OutputDebugStringA("failed to get file name");
		return FALSE;
	}
	std::wstring fileName(pFileName);
	std::size_t found = fileName.find_last_of(_T("/\\"));
	if (found < 0)
		found = 0;
	found++;
	fileName = fileName.substr(found, fileName.size());
	OutputDebugString((_T("file name: ") + fileName + _T("\n")).c_str());

	//set IE8 emu
	{
		const wchar_t* _MS_FEATURE_BROWSER_EMULATION =
			_T("Software\\Microsoft\\Internet Explorer\\Main\\FeatureControl\\FEATURE_BROWSER_EMULATION");
		CRegKey regKey;
		LSTATUS regStatus =
			regKey.Open(HKEY_CURRENT_USER, _MS_FEATURE_BROWSER_EMULATION);
		if (regStatus == ERROR_FILE_NOT_FOUND)
		{
			regStatus = regKey.Create(HKEY_CURRENT_USER, _MS_FEATURE_BROWSER_EMULATION);
		}
		if (regStatus != ERROR_SUCCESS)
		{
			OutputDebugStringA("error opening registry key code");
			return FALSE;
		}
		regStatus = regKey.SetDWORDValue(fileName.c_str(), 11000);
		if (regStatus != ERROR_SUCCESS)
		{
			OutputDebugStringA("error setting registry value code=");
			return FALSE;
		}
	}
	//force GDI to get vector EMF
	{
		const wchar_t* _MS_FEATURE_BROWSER_DRAWMODE =
			_T("Software\\Microsoft\\Internet Explorer\\Main\\FeatureControl\\FEATURE_IVIEWOBJECTDRAW_DMLT9_WITH_GDI");
		CRegKey regKey;
		LSTATUS regStatus =
			regKey.Open(HKEY_CURRENT_USER, _MS_FEATURE_BROWSER_DRAWMODE);
		if (regStatus == ERROR_FILE_NOT_FOUND)
		{
			regStatus = regKey.Create(HKEY_CURRENT_USER, _MS_FEATURE_BROWSER_DRAWMODE);
		}
		if (regStatus != ERROR_SUCCESS)
		{
			OutputDebugStringA("error opening registry key code");
			return FALSE;
		}
		regStatus = regKey.SetDWORDValue(fileName.c_str(), 00000001);
		if (regStatus != ERROR_SUCCESS)
		{
			OutputDebugStringA("error setting registry value code=");
			return FALSE;
		}
	}

	return TRUE;
}



BOOL CMFCApplication1App::InitInstance()
{
	if (!SetupIEVersion())
		return FALSE;

	// InitCommonControlsEx() is required on Windows XP if an application
	// manifest specifies use of ComCtl32.dll version 6 or later to enable
	// visual styles.  Otherwise, any window creation will fail.
	INITCOMMONCONTROLSEX InitCtrls;
	InitCtrls.dwSize = sizeof(InitCtrls);
	// Set this to include all the common control classes you want to use
	// in your application.
	InitCtrls.dwICC = ICC_WIN95_CLASSES;
	InitCommonControlsEx(&InitCtrls);

	CWinApp::InitInstance();


	CCustomOccManager *pMgr = new CCustomOccManager;

	// Create an IDispatch class for
	// extending the Dynamic HTML Object Model 
	m_pDispatch.Attach( new CImpIDispatch() );

	// Set our control containment up
	// but using our control container 
	// management class instead of MFC's default
	AfxEnableControlContainer(pMgr);

	//AfxEnableControlContainer();

	// Create the shell manager, in case the dialog contains
	// any shell tree view or shell list view controls.
	CShellManager *pShellManager = new CShellManager;

	// Activate "Windows Native" visual manager for enabling themes in MFC controls
	CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManagerWindows));

	// Standard initialization
	// If you are not using these features and wish to reduce the size
	// of your final executable, you should remove from the following
	// the specific initialization routines you do not need
	// Change the registry key under which our settings are stored
	// TODO: You should modify this string to be something appropriate
	// such as the name of your company or organization
	SetRegistryKey(_T("Local AppWizard-Generated Applications"));

	CMFCApplication1Dlg dlg;
	m_pMainWnd = &dlg;
	INT_PTR nResponse = dlg.DoModal();
	if (nResponse == IDOK)
	{
	}
	else if (nResponse == IDCANCEL)
	{
		// TODO: Place code here to handle when the dialog is
		//  dismissed with Cancel
	}
	else if (nResponse == -1)
	{
		TRACE(traceAppMsg, 0, "Warning: dialog creation failed, so application is terminating unexpectedly.\n");
		TRACE(traceAppMsg, 0, "Warning: if you are using MFC controls on the dialog, you cannot #define _AFX_NO_MFC_CONTROLS_IN_DIALOGS.\n");
	}

	// Delete the shell manager created above.
	if (pShellManager != NULL)
	{
		delete pShellManager;
	}

	// Since the dialog has been closed, return FALSE so that we exit the
	//  application, rather than start the application's message pump.
	return FALSE;
}

