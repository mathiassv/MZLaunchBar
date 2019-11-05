
// MZLaunchBarDemoDlg.cpp : implementation file
//

#include "stdafx.h"
#include "MZLaunchBarDemo.h"
#include "MZLaunchBarDemoDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CAboutDlg dialog used for App About

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// Dialog Data
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Implementation
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
END_MESSAGE_MAP()


// CMZLaunchBarDemoDlg dialog




CMZLaunchBarDemoDlg::CMZLaunchBarDemoDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CMZLaunchBarDemoDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CMZLaunchBarDemoDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
  DDX_Control(pDX, IDC_LAUNCHBAR, m_LaunchBar);
}

BEGIN_MESSAGE_MAP(CMZLaunchBarDemoDlg, CDialog)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
  ON_BN_CLICKED(IDC_CHECK1, &CMZLaunchBarDemoDlg::OnBnClickedCheckAny)
  ON_BN_CLICKED(IDC_CHECK2, &CMZLaunchBarDemoDlg::OnBnClickedCheckAny)
  ON_BN_CLICKED(IDC_CHECK3, &CMZLaunchBarDemoDlg::OnBnClickedCheckAny)
  ON_BN_CLICKED(IDC_CHECK4, &CMZLaunchBarDemoDlg::OnBnClickedCheckAny)
  ON_BN_CLICKED(IDC_CHECK5, &CMZLaunchBarDemoDlg::OnBnClickedCheckAny)
  ON_BN_CLICKED(IDC_CHECK6, &CMZLaunchBarDemoDlg::OnBnClickedCheckAny)
  ON_BN_CLICKED(IDC_CHECK7, &CMZLaunchBarDemoDlg::OnBnClickedCheckAny)
  ON_BN_CLICKED(IDC_CHECK8, &CMZLaunchBarDemoDlg::OnBnClickedCheckAny)
  ON_BN_CLICKED(IDC_CHECK9, &CMZLaunchBarDemoDlg::OnBnClickedCheckAny)
  ON_BN_CLICKED(IDC_CHECK10,&CMZLaunchBarDemoDlg::OnBnClickedCheckAny)
  ON_BN_CLICKED(IDC_CHECK11,&CMZLaunchBarDemoDlg::OnBnClickedCheckAny)
  ON_BN_CLICKED(IDC_CHECK12,&CMZLaunchBarDemoDlg::OnBnClickedCheckAny)
  ON_BN_CLICKED(IDC_CHECK13,&CMZLaunchBarDemoDlg::OnBnClickedCheckAny)
  ON_BN_CLICKED(IDC_CHECK14,&CMZLaunchBarDemoDlg::OnBnClickedCheckAny)
END_MESSAGE_MAP()


// CMZLaunchBarDemoDlg message handlers

BOOL CMZLaunchBarDemoDlg::OnInitDialog()
{
  // Set what options should be enabled by default.
  m_LaunchBar.Style(MZBS_DROPADD | MZBS_DRAGREARRANGE | MZBS_LAUNCHADMIN_SHIFT | MZBS_CONTEXTMENUWS | MZBS_CONTEXTMENU 
    | MZBS_ALLOW_CUSTOMIZE | MZBS_ALLOW_REMOVE| MZBS_LAUNCH_CLICK | MZBS_DRAW_HLIGHT_HOVER | MZBS_LAUNCH_DROPASPARM);
  
	CDialog::OnInitDialog();

  m_LaunchBar.SetIconSize(CSize(32,32), false);
  m_LaunchBar.Init(true);
//  // force NCCALCSIZE
//  m_LaunchBar.SetWindowPos(NULL, 0, 0, 0, 0, SWP_FRAMECHANGED| SWP_NOSIZE | SWP_NOMOVE);

  // Add "About..." menu item to system menu.
  // IDM_ABOUTBOX must be in the system command range.
  ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
  ASSERT(IDM_ABOUTBOX < 0xF000);

  CMenu* pSysMenu = GetSystemMenu(FALSE);
  if (pSysMenu != NULL)
  {
    BOOL bNameValid;
    CString strAboutMenu;
    bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
    ASSERT(bNameValid);
    if (!strAboutMenu.IsEmpty())
    {
      pSysMenu->AppendMenu(MF_SEPARATOR);
      pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
    }
  }

  SetWindowText(_T("MZLaunchBar Demp v1.0 - Created by - Mathias Svensson - Result42.com"));

  // Set the icon for this dialog.  The framework does this automatically
  //  when the application's main window is not a dialog
  SetIcon(m_hIcon, TRUE);     // Set big icon
  SetIcon(m_hIcon, FALSE);    // Set small icon

  m_LaunchBar.AddItem(_T("%WINDIR%\\System32\\notepad.exe"), _T("Notepad"), true);
  m_LaunchBar.AddItem(_T("%WINDIR%\\System32\\perfmon.exe"), _T("PerfMon"), true);
  m_LaunchBar.AddItem(_T("%WINDIR%\\System32\\winver.exe"), _T("WinVer"), true);

  // Use alternative icon for for cmd.exe.
  auto p = m_LaunchBar.AddItem(_T("%WINDIR%\\System32\\cmd.exe"), _T("Cmd"), false);
  p->altIconPath = _T("%WINDIR%\\System32\\shell32.dll");
  p->altIconIdx= 32;
  m_LaunchBar.UpdateIcon(p.get());
  m_LaunchBar.EnableTooltip(true);

  // Update GUI option checkboxes
  UpdateCheckBoxes(m_LaunchBar.Style());
  return TRUE;  // return TRUE  unless you set the focus to a control
}

void CMZLaunchBarDemoDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CMZLaunchBarDemoDlg::OnPaint()
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
	else
	{
		CDialog::OnPaint();
	}
}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CMZLaunchBarDemoDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

void CMZLaunchBarDemoDlg::IsOptionChecked(DWORD& dwOptions, UINT nID, DWORD dwOption)
{
  if(IsDlgButtonChecked(nID))
    dwOptions |= dwOption;
}

void CMZLaunchBarDemoDlg::SetOption(DWORD dwOptions, UINT nID, DWORD dwFlag)
{
  CheckDlgButton(nID, (((dwOptions & dwFlag) > 0) ? BST_CHECKED : BST_UNCHECKED));
}

void CMZLaunchBarDemoDlg::UpdateCheckBoxes(DWORD dwOptions)
{
  SetOption(dwOptions, IDC_CHECK1, MZBS_DROPADD);
  SetOption(dwOptions, IDC_CHECK2, MZBS_DRAGREARRANGE);
  SetOption(dwOptions, IDC_CHECK3, MZBS_LAUNCH_CLICK);
  SetOption(dwOptions, IDC_CHECK4, MZBS_LAUNCH_DBLCLICK);
  SetOption(dwOptions, IDC_CHECK5, MZBS_LAUNCHADMIN_SHIFT);
  SetOption(dwOptions, IDC_CHECK6, MZBS_LAUNCHADMIN_CTRL);
  SetOption(dwOptions, IDC_CHECK7, MZBS_CONTEXTMENU);
  SetOption(dwOptions, IDC_CHECK8, MZBS_SHELLMENU);
  SetOption(dwOptions, IDC_CHECK9, MZBS_ALLOW_REMOVE);
  SetOption(dwOptions, IDC_CHECK10,MZBS_ALLOW_CUSTOMIZE);
  SetOption(dwOptions, IDC_CHECK11,MZBS_LAUNCH_DROPASPARM);
  SetOption(dwOptions, IDC_CHECK13,MZBS_DRAW_HLIGHT_DOWN);
  SetOption(dwOptions, IDC_CHECK14,MZBS_DRAW_HLIGHT_HOVER);
  SetOption(dwOptions, IDC_CHECK15,MZBS_CONTEXTMENUWS);

  CheckDlgButton(IDC_CHECK12, (m_LaunchBar.GetIconSize().cx > 16) ? BST_CHECKED : BST_UNCHECKED);
}

void CMZLaunchBarDemoDlg::OnBnClickedCheckAny()
{
  DWORD dwOptions = 0;
  IsOptionChecked(dwOptions, IDC_CHECK1, MZBS_DROPADD);
  IsOptionChecked(dwOptions, IDC_CHECK2, MZBS_DRAGREARRANGE);
  IsOptionChecked(dwOptions, IDC_CHECK3, MZBS_LAUNCH_CLICK);
  IsOptionChecked(dwOptions, IDC_CHECK4, MZBS_LAUNCH_DBLCLICK);
  IsOptionChecked(dwOptions, IDC_CHECK5, MZBS_LAUNCHADMIN_SHIFT);
  IsOptionChecked(dwOptions, IDC_CHECK6, MZBS_LAUNCHADMIN_CTRL);
  IsOptionChecked(dwOptions, IDC_CHECK7, MZBS_CONTEXTMENU);
  IsOptionChecked(dwOptions, IDC_CHECK8, MZBS_SHELLMENU);
  IsOptionChecked(dwOptions, IDC_CHECK9, MZBS_ALLOW_REMOVE);
  IsOptionChecked(dwOptions, IDC_CHECK10, MZBS_ALLOW_CUSTOMIZE);
  IsOptionChecked(dwOptions, IDC_CHECK11, MZBS_LAUNCH_DROPASPARM);
  IsOptionChecked(dwOptions, IDC_CHECK13, MZBS_DRAW_HLIGHT_DOWN);
  IsOptionChecked(dwOptions, IDC_CHECK14, MZBS_DRAW_HLIGHT_HOVER);
  IsOptionChecked(dwOptions, IDC_CHECK15, MZBS_CONTEXTMENUWS);
  
  m_LaunchBar.Style(dwOptions);

  if(IsDlgButtonChecked(IDC_CHECK12))
  {
    if(m_LaunchBar.GetIconSize().cx < 32)
    {
      m_LaunchBar.SetIconSize(CSize(32,32), true);
      m_LaunchBar.AdjustHeight();
    }
  }
  else
  {
    // Set Large icons
    if(m_LaunchBar.GetIconSize().cx > 16)
    {
      m_LaunchBar.SetIconSize(CSize(16,16), true);
      m_LaunchBar.AdjustHeight();
    }
  }

}

