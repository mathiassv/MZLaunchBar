/////////////////////////////////////////////////////////////////////////////////////////////
//
//  MZLaunchBar v1.0 - Copyright 2012 - Mathias Svensson
//
//  Created by : Mathias Svensson - http://result42.com/projects/mzlunchbar
//
//  License
//  You are allowed to use this source code in any product (commercial, shareware, freeware or otherwise) 
//  Except if you product is a file manager like product, or other product that is a competitor to Multi Commander. Then you are not allowed to use this source code.
//  You may include this source code when/if you release your products source code, as long as the the original author information is not change.
//  If source code is modified, then own modification must be clearly marked and a link to the original source code must be provided in the header.
//  You may NOT modify the original copyright details at the top of each module
//  If you fix any bug or add enhanchments please send the changes to the author so they may be included in the original source.
//
/////////////////////////////////////////////////////////////////////////////////////////////
//
// See MZLaunchBar.cpp for history

#pragma once

// To use ShellContextMenu option. MZLUNCHBAR_USE_SHELLCONTEXTMENU must be defined.
// You also need to include ShellContextMenu.cpp/.h into you project, (Created by R. Engels)
// You find ShellContextMenu class here : http://www.codeproject.com/Articles/4025/Use-Shell-ContextMenu-in-your-applications
// #define MZLUNCHBAR_USE_SHELLCONTEXTMENU

// If MZLUNCHBAR_USE_SHELLCONTEXTMENU is used or if option MZBS_DROPADD is used. 
// Then make sure AfxOleInit(); is called in InitInstance of your application

// To use the LaunchPropDlg to customize item use define MZLUNCHBAR_USE_LANUCHPROPDLG
// #define MZLUNCHBAR_USE_LANUCHPROPDLG

// OR define them in stdafx.h

#include <memory>
#include <vector>
#include <string>

#ifndef STL_string
  #define STL_string std::string
  #ifndef _UNICODE
    typedef STL_string	 STLString;
  #endif
#endif

#ifndef STL_wstring
  #define STL_wstring std::wstring
  #ifdef _UNICODE
    typedef STL_wstring  STLString;
  #endif
#endif

class MZLaunchBar;
//////////////////////////////////////////////////////////////////////////
  class MZLaunchBarDropTarget
    : public COleDropTarget
  {
  public:
    MZLaunchBarDropTarget(MZLaunchBar* pOwner)
      : m_pOwner(pOwner)
      , m_bRegistered(false)
    {
    }

    virtual DROPEFFECT	OnDragEnter(CWnd* pWnd, COleDataObject* pDataObject, DWORD dwKeyState, CPoint point);
    virtual DROPEFFECT	OnDragOver(CWnd* pWnd, COleDataObject* pDataObject, DWORD dwKeyState, CPoint point);
    virtual BOOL	      OnDrop(CWnd* pWnd, COleDataObject* pDataObject, DROPEFFECT dropEffect, CPoint point);
    virtual DROPEFFECT  OnDropEx(CWnd* pWnd, COleDataObject* pDataObject, DROPEFFECT dropDefault, DROPEFFECT dropList, CPoint point);
    virtual void		    OnDragLeave(CWnd* pWnd);

    bool Register();
    void Unregister();

  private:
    bool m_bRegistered;
    MZLaunchBar*  m_pOwner;
  };
//////////////////////////////////////////////////////////////////////////
// MZLaunchBar Context menu commands
#define MZLBC_OPEN       10
#define MZLBC_OPENADMIN  11
#define MZLBC_REMOVE     15
#define MZLBC_CUSTOMIZE  16

#define MZLBC_ADDSEP          20
#define MZLBC_ALLOWDROPADD    21
#define MZLBC_ALLOWREARRANGE  22

#define MZLB_MAX      30

// MZLaunchItem State
#define MZLIS_HOVEROVER    0x00001000L // Hovering over item
#define MZLIS_BUTTONDOWN   0x00002000L // Icon Button Pressed down


    // MZLaunchBar
class MZLaunchItem
{
public:
  MZLaunchItem()
    : width(0)
    , altIconIdx(-1)
    , nIconIdx(-1)
    , bSeparator(false)
    , dwState(0)
    , dwCustomParam(0)
    , dwCustomData(0)
    , rcItemArea(0,0,0,0)
    , rcItemAreaDragging(0,0,0,0)
  {

  }
  //////////////////////////////////////////////////////////////////////////
  STLString programPath;       // file
  STLString programParameters; // Parameter to be sent to file
  STLString startFolder;       // Start folder when file is opened
  STLString tooltip;           // Tooltip for icon

  //////////////////////////////////////////////////////////////////////////
  // If it should take icon from alternative file,
  STLString altIconPath;
  int       altIconIdx;
  //////////////////////////////////////////////////////////////////////////

  int width;    // Width of the item.
  int nIconIdx; // Icon index in the ImageList
  bool bSeparator; // true if item is a separatoritem

  DWORD dwState; // MZLIS_ state

  //////////////////////////////////////////////////////////////////////////
  DWORD     dwCustomParam;
  DWORD_PTR dwCustomData;

  //////////////////////////////////////////////////////////////////////////
  // Cache
  CRect rcItemArea;         // Draw area
  CRect rcItemAreaDragging; // Draw area while item is being dragged
};

//////////////////////////////////////////////////////////////////////////
//
// MZLaunchBar Style and Options flags.
//
#define MZBS_DROPADD            0x00000010 // Allow files to be added by dropping them in the MZLaunchBar
#define MZBS_DRAGREARRANGE      0x00000020 // Allow items to be rearrange by dragging them
#define MZBS_CONTEXTMENUWS      0x00000080 // Show a context menu when pressing on whitespace (outside of an icon), Allowing user to Add Separator, and change some options.
#define MZBS_CONTEXTMENU        0x00000100 // Show Item Context Menu
#define MZBS_SHELLMENU          0x00000200 // Show Shell context menu when right clicking on an item. (Must be combined with MZBS_CONTEXTMENU)
#define MZBS_ALLOW_CUSTOMIZE    0x00000400 // Allow user to customize items
#define MZBS_ALLOW_REMOVE       0x00000800 // Allow user to remove items
#define MZBS_LAUNCH_CLICK       0x00002000 // Launch item if icon is single clicked
#define MZBS_LAUNCH_DBLCLICK    0x00004000 // Launch item if icon is double clicked
#define MZBS_LAUNCH_DROPASPARM  0x00008000 // Drop files to be lanuch as parameters to program
#define MZBS_LAUNCHADMIN_SHIFT  0x00010000 // Launch program as Admin if Shift key was pressed while item was clicked/dblclicked
#define MZBS_LAUNCHADMIN_CTRL   0x00020000 // Launch program as Admin if Ctrl key was pressed while item was clicked/dblclicked
#define MZBS_DRAW_HLIGHT_DOWN   0x01000000  // Draw Highlight frame around icon when mouse button is down on icon
#define MZBS_DRAW_HLIGHT_HOVER  0x02000000  // Draw Highlight frame around icon when hover over icon
#define MZBS_ALIGNVCENTER       0x04000000  // Align to vertical center of client area
#define MZBS_ALIGNBOTTOM        0x08000000  // Align to bottom of client area (default is TOP align)

class MZLaunchBar : public CWnd
{
  DECLARE_DYNAMIC(MZLaunchBar)

public:
  MZLaunchBar();
  virtual ~MZLaunchBar();

  virtual BOOL Create(DWORD dwStyle, DWORD dwLBStyle , const RECT& rect, CWnd* pParentWnd, UINT nID, CCreateContext* pContext = NULL);

  virtual std::shared_ptr<MZLaunchItem> AddItem(const STLString& programPath, const STLString& tooltip, bool bGetIcon);
  virtual std::shared_ptr<MZLaunchItem> AddSeparatorItem();
  virtual void RemoveItem(MZLaunchItem* pItem);

  DWORD GetItemCount();

  // After you added items call this.. (This is also called after dropadd)
  bool RebuildTooltips();

  // Save if item are rearrange/added/changed then auto call Save()
  void SaveOnChange(bool b) { m_bSaveOnChange = b; }

  // MZBS_ flags
  void Style(DWORD dwStyle);
  DWORD Style() { return m_dwStyle; }

  // Extra margin around button (width)
  short MarginX() const { return m_nMarginX;  }
  void MarginX(short val) { m_nMarginX = val; }

  // Extra margin around button (height)
  short MarginY() const { return m_nMarginY;  }
  void MarginY(short val) { m_nMarginY = val; }

  // Space between icons/buttons
  int Spacing() const { return m_nSpacing; }
  void Spacing(int val) { m_nSpacing = val; }

  void SetIconSize(CSize cz, bool bRefreshIcons = true);
  CSize GetIconSize() { return m_IconSize; }

  void Init(bool bRefreshIcons);
  virtual void InitPaintResources();

  // If you use an separator a separator icon must be used.
  bool AddSeparatorIcon(HICON hIcon);

  bool RegisterDropTarget(bool bEnable);

  // COleDropTarget
  virtual DROPEFFECT  OnDragEnter(CWnd* pWnd, COleDataObject* pDataObject, DWORD dwKeyState, CPoint point);
  virtual DROPEFFECT  OnDragOver(CWnd* pWnd, COleDataObject* pDataObject, DWORD dwKeyState, CPoint point);
  virtual BOOL        OnDrop(CWnd* pWnd, COleDataObject* pDataObject, DROPEFFECT dropEffect, CPoint point);
  virtual DROPEFFECT  OnDropEx(CWnd* pWnd, COleDataObject* pDataObject, DROPEFFECT dropDefault, DROPEFFECT dropList, CPoint point);
  virtual void        OnDragLeave(CWnd* pWnd);
  virtual DROPEFFECT  GetDropEffect(MZLaunchItem* pOverItem);

  void UpdateIcon(MZLaunchItem* pItem);

  void EnableTooltip(bool bEnable);

  // Not a good solution for this. But depending on where you place the controller it might
  // be needed. so the height on the controller is correct.
  void AdjustHeight();
    
  virtual BOOL PreTranslateMessage(MSG* pMsg);

  // Used by external source that what to know the size if there is X items. 
  // FOr example if an area needs to be reserved
  virtual bool CalcSize(CSize& sz, int nItems);

  // Return how many items the controller can show at current size. 
         DWORD MaxVisibleItems();

  // Overload to handle your own load/save format.
  virtual void Load(const STLString& filename);
  virtual void Save();
protected:

  DECLARE_MESSAGE_MAP()

  afx_msg int  OnCreate(LPCREATESTRUCT lpCreateStruct);
  afx_msg void OnNcCalcSize(BOOL bCalcValidRects, NCCALCSIZE_PARAMS FAR* lpncsp);
  
  afx_msg BOOL OnEraseBkgnd(CDC* pDC);
  afx_msg void OnPaint();
  
  afx_msg void    OnMouseMove(UINT nFlags, CPoint point);
  afx_msg LRESULT OnMouseLeave(WPARAM wParam, LPARAM lParam);
  
  afx_msg void OnRButtonDown(UINT nFlags, CPoint point);
  afx_msg void OnContextMenu(CWnd* /*pWnd*/, CPoint /*point*/);
  afx_msg void OnRButtonUp(UINT nFlags, CPoint point);
  
  afx_msg void OnLButtonDblClk(UINT nFlags, CPoint point);
  afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
  afx_msg void OnLButtonUp(UINT nFlags, CPoint point);

  afx_msg BOOL OnToolTipNotify(UINT id, NMHDR* pNMHDR, LRESULT* pResult);
  
  afx_msg BOOL OnSetCursor(CWnd* pWnd, UINT nHitTest, UINT message);

  void OnLButtonClick(const CPoint& pt);
  void OnLButtonDblClick(const CPoint& pt);

  void EnterDragMode(MZLaunchItem* pItem, const CPoint& pt);
  void RecalculateDragPositions(const CPoint& pt);
  void InitDragPositions();
  void CommitDragPositions();

  void Draw(CDC* pDC);
  void DrawBackground(CDC* pDC, const CRect& rc);
  void DrawItems(CDC* pDC, const CRect& rc);
  void DrawItem(CDC* pDC, const CRect& rc, const MZLaunchItem* pItem);
  void DrawItemBackground(CDC* pDC,const CRect& rc,const MZLaunchItem* pItem);
  void DrawInsertLine(CDC* pDC, const CRect& rc);

  // Including margins
  int GetItemWith(MZLaunchItem* pItem);

  void UpdateItemAreas(CDC* pDC, const CRect& rcItemsArea);
  void InvalidateItem(MZLaunchItem* pItem);

  void OnHandleContextMenuCommand(int nCmd, MZLaunchItem* pItem);

  MZLaunchItem*  GetItemAtPos(const CPoint& pt);
  MZLaunchItem*  GetItemAfterPos(const CPoint& pt);
  MZLaunchItem*  GetItemBeforePos(const CPoint& pt);
  MZLaunchItem*  GetItemBefore(const MZLaunchItem* pItem);

  virtual void PreSubclassWindow();
  virtual BOOL Create( DWORD dwLBStyle = 0 );

  void InitToolTipCtrl();

  virtual void InitImageList();
  
  // Expand Environment tags (Override if you want to support own specialized tags)
  virtual STLString ExpandStringTags(const STLString& str);

  std::shared_ptr<MZLaunchItem> CreateSeparatorItem();
  //////////////////////////////////////////////////////////////////////////
  //
  // Override function below to change look and behavior

  // Create new item 
  virtual std::shared_ptr<MZLaunchItem> MZLaunchBar::CreateItem( const STLString& programPath, const STLString& tooltip);

  // Get icon for an item. Override to use custom icon handling
  virtual HICON GetIcon(MZLaunchItem* pItem);
  virtual HICON GetIcon(const STLString& programPath, bool bExpand = true);
  virtual HICON GetIcon(const STLString& file, int nIconIdx, DWORD dwCustomParam = 0);
  
  virtual int  AddIcon(HICON hIcon);
  virtual void RefreshIcons();

  virtual void OnItemOverLeave(MZLaunchItem* pItemLeft);
  virtual void OnItemOverEnter(MZLaunchItem* pItem);
  virtual void OnItemOver(MZLaunchItem* pItem);
  
  // user clicked on item using left mouse button
  virtual void OnLButtonClick(MZLaunchItem* pItem, const CPoint& pt);
  // user double clicked on item using left mouse button
  virtual void OnLButtonDblClick(MZLaunchItem* pItem, const CPoint& pt);

  // User Requested to show the context menu for an item
  virtual void OnContextMenu(MZLaunchItem* pItem, const CPoint& pt);
  // user Requested to show the ShellContext menu for an item.  ( must be compiled with #define MZLUNCHBAR_USE_SHELLCONTEXTMENU to work )
  virtual void OnShellContextMenu(MZLaunchItem* pItem, const CPoint& pt);

  // Launch Item using shell execute
  virtual void OnLaunchItem(MZLaunchItem* pItem, bool bAsAdmin, std::vector<STLString>& vFileParameters);

  // User Requested to remove item
  virtual void OnRemoveItem(MZLaunchItem* pItem);
  // user Requested to add a separator
  virtual void OnAddSeparator(MZLaunchItem* pAfterItem);
  // User Requested to Customize item
  virtual void OnCustomizeItem(MZLaunchItem* pItem);

  // Show Remove Item confirmation dialog. Override to do your own. (for MultiLanguage support?)
  virtual bool ConfirmRemoveItem(MZLaunchItem* pItem);

  // When Item is dropped on the Launchbar to be added. if this return false the item will not be added
  virtual bool AcceptDroppedItem(const STLString& filepath);

  // Item has changed/added
  enum MZLIChange
  {
    Change_Added,
    Change_AddedSep,
    Change_Rearranged,
    Change_Modified,
    Change_Removed
  };

  virtual void OnChanged(MZLaunchItem* pItem, MZLIChange change);
  // if bInsert == false then pItem is the item that the files was dropped on
  // if bInsert == true then pItem is the item we and to insert the new items BEFORE.
  // if bInsert == true and pItem == nullptr then insert item at end
  virtual void OnDropFiles(std::vector<STLString>& vDroppedFiles, MZLaunchItem* pItem, bool bInsert);
  virtual std::shared_ptr<MZLaunchItem> CreateNewLaunchItemFromDroppedPath(const STLString& filepath);
  //////////////////////////////////////////////////////////////////////////

  bool InsertBefore(std::vector<std::shared_ptr<MZLaunchItem>>& vItems, MZLaunchItem* pItem, std::shared_ptr<MZLaunchItem>& pInsertItem);
  bool InsertAfter(std::vector<std::shared_ptr<MZLaunchItem>>& vItems, MZLaunchItem* pItem, std::shared_ptr<MZLaunchItem>& pInsertItem);
  bool InsertFirst(std::vector<std::shared_ptr<MZLaunchItem>>& vItems, std::shared_ptr<MZLaunchItem>& pInsertItem);
  bool RemoveItem(std::vector<std::shared_ptr<MZLaunchItem>>& vItems, MZLaunchItem* pItem);

  
  //////////////////////////////////////////////////////////////////////////
  //
  // variables
  //
  bool    m_bCreatedByCreate; // If Create(..) was called to create controller.. else it was create using subclassing
  DWORD   m_dwStyle;
  CSize   m_IconSize;
  short   m_nMarginX;
  short   m_nMarginY;
  int     m_nSpacing; // Space between "buttons"

  CBrush   m_bkBrush;
  CBrush   m_bkItemBrush;
  CPen     m_penItemFrame;
  COLORREF m_crDropMark; // color for the DropMark

  bool    m_bSaveOnChange;
  bool    m_bDrawStateDirty;
  bool		m_bMouseOver;			// use to keep track if we have register event for MouseLeave when we got a MouseOver
  bool    m_bDblClk;        // So we can separate single and double click in OnLButtonUp
  bool    m_bPrepareDragging;
  bool    m_bDragging;
  bool    m_bDropping;
  bool    m_bRegisteredForDropTarget;
  bool    m_bToolTip;
  bool    m_bRebuildTooltip;

  // Option that can be turn on/off in context menu.. 
  bool    m_bAllowDropAdd;
  bool    m_bAllowRearrange;

  STLString m_CurrentTooltip;

  CPoint  m_ptDragPos;
  int     m_SeparatorIconIdx;

  CImageList    m_Images;
  CToolTipCtrl  m_ToolTip;

  std::vector<std::shared_ptr<MZLaunchItem>>  m_vItems;
  std::vector<std::shared_ptr<MZLaunchItem>>  m_vDragItems; // Tmp storage when dragging items

  MZLaunchItem* m_pOverItem;
  MZLaunchItem* m_pClickItem; // also used as Drop target 

  MZLaunchBarDropTarget m_DropTarget;       // OLE Drop target 
  HCURSOR  m_hCursor;
};
