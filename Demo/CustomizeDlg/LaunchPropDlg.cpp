// source\LaunchPropDlg.cpp : implementation file
//

#include <stdafx.h>
#include "LaunchPropDlg.h"
#include <MZLaunchBar.h>

// LaunchPropDlg dialog

IMPLEMENT_DYNAMIC(LaunchPropDlg, CDialog)

UINT LaunchPropDlg::s_WM_REFRESH_ICONS = ::RegisterWindowMessage( _T("LaunchProp_REFRESH_ICONS") );

namespace
{
  CString ExpandEnvironmentStr(const CString& str)
  {
    TCHAR szTmp[16*1024];
    if(ExpandEnvironmentStrings(str, szTmp, _countof(szTmp)) == 0)
      return str;

    return szTmp;
  }

}
LaunchPropDlg::LaunchPropDlg(MZLaunchItem* pItem, CWnd* pParent /*=NULL*/)
	: CDialog(LaunchPropDlg::IDD, pParent)
  , m_pLaunchItem(pItem)
{
  m_nIconIndex = 0;
  m_dwChange = 0;

  if(m_pLaunchItem)
  {
    m_strProgram = m_pLaunchItem->programPath.c_str();
    m_strProgramParam = m_pLaunchItem->programParameters.c_str();
    m_strStartInFolder = m_pLaunchItem->startFolder.c_str();
    m_strDescription = m_pLaunchItem->tooltip.c_str();
    m_strIconFilename = m_pLaunchItem->altIconPath.c_str();
    m_nIconIndex = m_pLaunchItem->altIconIdx;
  }

}

LaunchPropDlg::~LaunchPropDlg()
{
}

void LaunchPropDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);

  DDX_Control(pDX, IDC_ICONLIST, m_wndIconList);
  DDX_Text(pDX, IDC_PROGRAM, m_strProgram);
  DDX_Text(pDX, IDC_PROGRAMPARAM, m_strProgramParam);
  DDX_Text(pDX, IDC_STARTIN, m_strStartInFolder);
  DDX_Text(pDX, IDC_DESC, m_strDescription);
  DDX_Text(pDX, IDC_ICONFILE, m_strIconFilename);
  DDX_Text(pDX, IDC_PROGRAM, m_strProgram);
  DDX_Text(pDX, IDC_ICONIDX, m_nIconIndex);
  DDX_LBIndex(pDX, IDC_ICONLIST, m_nIconIndex);

  DDX_FileEditCtrl(pDX, IDC_ICONFILE, m_wndIconFilename, FEC_FILEOPEN);
  DDX_FileEditCtrl(pDX, IDC_PROGRAM, m_wndProgram, FEC_FILEOPEN);
  DDX_FileEditCtrl(pDX, IDC_STARTIN, m_wndStartFolder, FEC_FOLDER);
}

BEGIN_MESSAGE_MAP(LaunchPropDlg, CDialog)
  ON_WM_MEASUREITEM()
  ON_WM_DRAWITEM()
  ON_WM_DESTROY()
  ON_LBN_SELCHANGE(IDC_ICONLIST, OnIconIdxChanged)
  ON_LBN_DBLCLK(IDC_ICONLIST, OnDblclkIconlist)
  ON_BN_CLICKED(IDC_BTN_REFRESHICONS, &LaunchPropDlg::OnBnClickedBtnRefreshicons)
  ON_NOTIFY(FEC_NM_POSTBROWSE, IDC_ICONFILE, OnPostBrowseIconFile)
  ON_REGISTERED_MESSAGE(s_WM_REFRESH_ICONS, OnRefreshIcons)
  ON_BN_CLICKED(IDOK, &LaunchPropDlg::OnBnClickedOk)
END_MESSAGE_MAP()

// LaunchPropDlg message handlers

void LaunchPropDlg::CleanupIcons()
{
  int nCount = m_wndIconList.GetCount();
  for (int i=0; i<nCount; i++)
  {
    HICON hIcon = (HICON) m_wndIconList.GetItemData(i);
    DestroyIcon(hIcon);
  }
  m_wndIconList.ResetContent();
}

void LaunchPropDlg::UpdateIconList() 
{
  //Free up any existing HICON's used for drawing
  CleanupIcons();

  //Do we have a valid filename
  if(m_strIconFilename.GetLength() == 0)
    return;

  CString iconFilePath = ExpandEnvironmentStr(m_strIconFilename);
  
  if(GetFileAttributes(iconFilePath) == INVALID_FILE_ATTRIBUTES)
    return;

  int nNum = (int) ExtractIcon(AfxGetInstanceHandle(), iconFilePath, (UINT) -1);
  m_wndIconList.ResetContent();
  if (nNum == 0)
  {
    CString msg;
    msg = _T("No icons found in file");
    MessageBox(msg, _T("No Icons"), MB_OK | MB_ICONEXCLAMATION);
    return;
  }

  for(int i = 0; i < nNum; ++i)
  {
    HICON hIcon = ExtractIcon(AfxGetInstanceHandle(), iconFilePath, i);
    m_wndIconList.InsertString(i, _T(""));
    m_wndIconList.SetItemData(i, (LPARAM) hIcon);
  }

  m_wndIconList.SetCurSel(0);
  
}


void LaunchPropDlg::OnMeasureItem(int nIDCtl, LPMEASUREITEMSTRUCT lpMeasureItemStruct) 
{
  if (nIDCtl == IDC_ICONLIST)
    lpMeasureItemStruct->itemHeight = GetSystemMetrics(SM_CYICON) + 12;
  else
    CDialog::OnMeasureItem(nIDCtl, lpMeasureItemStruct);
}

void LaunchPropDlg::OnDrawItem(int nIDCtl, LPDRAWITEMSTRUCT lpDIS) 
{
  if (nIDCtl == IDC_ICONLIST)
  {
    ASSERT(lpDIS->CtlType == ODT_LISTBOX);	

    CDC* pDC = CDC::FromHandle(lpDIS->hDC);
    COLORREF cr = (COLORREF)lpDIS->itemData; // RGB in item data

    if ((lpDIS->itemState & ODS_SELECTED) && (lpDIS->itemAction & (ODA_SELECT | ODA_DRAWENTIRE)))
    {
      // item has been selected - draw selection rectangle
      COLORREF crHilite = GetSysColor(COLOR_HIGHLIGHT);
      CBrush br(crHilite);
      pDC->FillRect(&lpDIS->rcItem, &br);
    }

    if (!(lpDIS->itemState & ODS_SELECTED) && (lpDIS->itemAction & ODA_SELECT))
    {
      // Item has been de-selected -- remove selection rectangle
      CBrush br(RGB(255, 255, 255));
      pDC->FillRect(&lpDIS->rcItem, &br);
    }

    //Draw the icon
    pDC->DrawIcon(lpDIS->rcItem.left+2, lpDIS->rcItem.top+4, (HICON) lpDIS->itemData);
  }
  else
    CDialog::OnDrawItem(nIDCtl, lpDIS);
}

void LaunchPropDlg::OnDestroy() 
{
  CleanupIcons();

  //Let the parent do its stuff
  CDialog::OnDestroy();
}

void LaunchPropDlg::OnDblclkIconlist() 
{
}

void LaunchPropDlg::OnIconIdxChanged()
{
  int nIdx = m_wndIconList.GetCurSel();
  m_nIconIndex = nIdx;
  SetDlgItemInt(IDC_ICONIDX, m_nIconIndex, FALSE);
}

BOOL LaunchPropDlg::OnInitDialog() 
{
  //Let the parent do its stuff
  CDialog::OnInitDialog();

  //Update the icon list
  m_wndIconList.SetColumnWidth(GetSystemMetrics(SM_CXICON) + 6);  
  UpdateIconList();
  m_wndIconList.SetCurSel(m_nIconIndex);

  return TRUE;
}

void LaunchPropDlg::OnBnClickedBtnRefreshicons()
{
  UpdateData(TRUE);
  UpdateIconList();
  m_wndIconList.SetCurSel(m_nIconIndex);
}

void LaunchPropDlg::OnPostBrowseIconFile(NMHDR* pNMHDR, LRESULT* pResult)
{
  FEC_NOTIFY* pFEC = (FEC_NOTIFY*)pNMHDR;

  if(pFEC->pNewText)
  {
    // reset 
    m_nIconIndex = 0;
    // Post Message to start Refresh Icons,
    ::PostMessage(GetSafeHwnd(), s_WM_REFRESH_ICONS, 0 , 0);
  }

  *pResult = FALSE;
}

LRESULT LaunchPropDlg::OnRefreshIcons(WPARAM /*wParam*/, LPARAM /*lParam*/)
{
  OnBnClickedBtnRefreshicons();
  return 0;
} //


void LaunchPropDlg::OnBnClickedOk()
{
  UpdateData(TRUE);
  
  m_dwChange = 0;
  if(m_strProgram.Compare(m_pLaunchItem->programPath.c_str()) != 0)
  {
    m_dwChange |= LIPCH_PROGRAM;
    m_pLaunchItem->programPath = m_strProgram;
  }

  if(m_strProgramParam.Compare(m_pLaunchItem->programParameters.c_str()) != 0)
  {
    m_dwChange |= LIPCH_PARAMETERS;
    m_pLaunchItem->programParameters = m_strProgramParam;
  }

  if(m_strStartInFolder.Compare(m_pLaunchItem->startFolder.c_str()) != 0)
  {
    m_dwChange |= LIPCH_STARTFOLDER;
    m_pLaunchItem->startFolder = m_strStartInFolder;
  }

  if(m_strDescription.Compare(m_pLaunchItem->tooltip.c_str()) != 0)
  {
    m_dwChange |= LIPCH_DESC;
    m_pLaunchItem->tooltip = m_strDescription;
  }

  if(m_strIconFilename.Compare(m_pLaunchItem->altIconPath.c_str()) != 0)
  {
    m_dwChange |= LIPCH_ALTICONFILE;
    m_pLaunchItem->altIconPath = m_strIconFilename;
  }

  if(m_nIconIndex != m_pLaunchItem->altIconIdx)
  {
    m_dwChange |= LIPCH_ALTICONIDX;
    m_pLaunchItem->altIconIdx = m_nIconIndex;
  }

  CDialog::OnOK();
}
