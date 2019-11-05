#pragma once

#include "FileEditCtrl\FileEditCtrl.h"
#include "..\resource.h"

class MZLaunchItem;

// LaunchPropDlg dialog
#define LIPCH_PROGRAM     0x00000010
#define LIPCH_PARAMETERS  0x00000020
#define LIPCH_STARTFOLDER 0x00000040
#define LIPCH_DESC        0x00000040
#define LIPCH_ALTICONFILE 0x00000100
#define LIPCH_ALTICONIDX  0x00000200
class LaunchPropDlg : public CDialog
{
	DECLARE_DYNAMIC(LaunchPropDlg)

public:
	LaunchPropDlg(MZLaunchItem* pItem, CWnd* pParent = NULL);   // standard constructor
	virtual ~LaunchPropDlg();

// Dialog Data
	enum { IDD = IDD_LAUNCHPROP };

  DWORD GetChanges() { return m_dwChange; }
protected:
  void CleanupIcons();
  void UpdateIconList();
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
  virtual BOOL OnInitDialog();

  afx_msg void OnMeasureItem(int nIDCtl, LPMEASUREITEMSTRUCT lpMeasureItemStruct);
  afx_msg void OnDrawItem(int nIDCtl, LPDRAWITEMSTRUCT lpDrawItemStruct);
  afx_msg void OnDblclkIconlist();
  afx_msg void OnIconIdxChanged();
  afx_msg void OnDestroy();
  afx_msg LRESULT OnRefreshIcons(WPARAM wParam, LPARAM lParam);
  
  afx_msg void OnPostBrowseIconFile(NMHDR* pNMHDR, LRESULT* pResult);
	DECLARE_MESSAGE_MAP()

  static UINT s_WM_REFRESH_ICONS;

  CFileEditCtrl m_wndIconFilename;
  CFileEditCtrl m_wndProgram;
  CFileEditCtrl m_wndStartFolder;

  CString  m_strProgram;
  CString  m_strProgramParam;
  CString  m_strStartInFolder;
  CString  m_strDescription;
  CString  m_strIconFilename;
  int      m_nIconIndex;
  CListBox m_wndIconList;

  DWORD m_dwChange;
  MZLaunchItem* m_pLaunchItem;
public:
  afx_msg void OnBnClickedBtnRefreshicons();
  afx_msg void OnBnClickedOk();
};
