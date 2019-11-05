
// MZLaunchBarDemoDlg.h : header file
//

#pragma once
#include <MZLaunchBar.h>

// CMZLaunchBarDemoDlg dialog
class CMZLaunchBarDemoDlg : public CDialog
{
// Construction
public:
	CMZLaunchBarDemoDlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
	enum { IDD = IDD_MZLAUNCHBARDEMO_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support


// Implementation
protected:
  afx_msg void OnBnClickedCheckAny();
  void IsOptionChecked(DWORD& dwOptions, UINT nID, DWORD dwOption);
  void SetOption(DWORD dwOptions, UINT nID, DWORD dwFlag);
  void UpdateCheckBoxes(DWORD dwOptions);

  HICON m_hIcon;

  MZLaunchBar m_LaunchBar;

	// Generated message map functions
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
  
};
