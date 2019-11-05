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
/*
  Changes

  v1.0 2012-05-01 First version

*/

#include <stdafx.h>
#include "MZLaunchBar.h"

#ifdef MZLUNCHBAR_USE_SHELLCONTEXTMENU
  #include "..\Demo\ShellContextMenu\ShellContextMenu.h"
#endif

#ifdef MZLUNCHBAR_USE_LANUCHPROPDLG
#include "..\Demo\CustomizeDlg\LaunchPropDlg.h"
#endif


#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

#define TOOLTIP_ID 1

#define MZLAUNCHBARCTRL_CLASSNAME _T("MZLaunchBarCtrl")

IMPLEMENT_DYNAMIC(MZLaunchBar, CWnd)


namespace
{
  #ifndef IsSHIFTpressed
  #define IsSHIFTpressed() ( (GetKeyState(VK_SHIFT) & (1 << (sizeof(SHORT)*8-1))) != 0   )
  #endif

  #ifndef IsCTRLpressed
  #define IsCTRLpressed()  ( (GetKeyState(VK_CONTROL) & (1 << (sizeof(SHORT)*8-1))) != 0 )
  #endif

  typedef UINT (WINAPI* pSHExtractIconsW)(LPCWSTR, int, int, int, HICON*, UINT*, UINT, UINT); 
  static pSHExtractIconsW g_pSHExtractIconsW = nullptr;
    
  static void InitSHExtractIconsW(void) 
  {
    if(g_pSHExtractIconsW == nullptr)
    {
      HMODULE hmod = GetModuleHandleA("Shell32.dll");
      if (hmod) 
      {
        g_pSHExtractIconsW = (pSHExtractIconsW)GetProcAddress(hmod, "SHExtractIconsW");
      }
    }
  }
   


  class CMemoryDC: public CDC
  {
  public:
    CMemoryDC(CDC* pDC) : CDC()
    {
      ASSERT(pDC != NULL);

      m_pDC = pDC;
      m_pOldBitmap = NULL;
      m_bMemDC = !pDC->IsPrinting();

      if (m_bMemDC)    // Create a Memory DC
      {
        pDC->GetClipBox(&m_rect);
        CreateCompatibleDC(pDC);
        m_bitmap.CreateCompatibleBitmap(pDC, m_rect.Width(), m_rect.Height());
        m_pOldBitmap = SelectObject(&m_bitmap);
        SetWindowOrg(m_rect.left, m_rect.top);
        FillSolidRect(m_rect, pDC->GetBkColor());
      }
      else        
      {
        // Make a copy of the relevant parts of the current DC for printing
        m_hDC       = pDC->m_hDC;
        m_hAttribDC = pDC->m_hAttribDC;
      }

    }

    ~CMemoryDC()
    {
      if (m_bMemDC)
      {
        // Copy the offscreen bitmap onto the screen.
        m_pDC->BitBlt(m_rect.left, m_rect.top, m_rect.Width(), m_rect.Height(), this, m_rect.left, m_rect.top, SRCCOPY);

        //Swap back the original bitmap.
        SelectObject(m_pOldBitmap);
      } 
      else 
      {
        m_hDC = m_hAttribDC = NULL;
      }
    }

    CMemoryDC* operator->() { return this;  }
    operator CMemoryDC*()   { return this;  }

  private:
    CBitmap  m_bitmap;      // Offscreen bitmap
    CBitmap* m_pOldBitmap;  // bitmap originally found in CMemDC
    CDC*     m_pDC;         // Saves CDC passed in constructor
    CRect    m_rect;        // Rectangle of drawing area.
    BOOL     m_bMemDC;      // TRUE if CDC really is a Memory DC.
  };
}

//////////////////////////////////////////////////////////////////////////
//
//  MZLaunchBarDropTarget
//
//////////////////////////////////////////////////////////////////////////
DROPEFFECT MZLaunchBarDropTarget::OnDragEnter( CWnd* pWnd, COleDataObject* pDataObject, DWORD dwKeyState, CPoint point )
{
  if(m_pOwner)
    return m_pOwner->OnDragEnter(pWnd, pDataObject, dwKeyState, point);

  return DROPEFFECT_NONE;
}

DROPEFFECT MZLaunchBarDropTarget::OnDragOver( CWnd* pWnd, COleDataObject* pDataObject, DWORD dwKeyState, CPoint point )
{
  if(m_pOwner)
    return m_pOwner->OnDragOver(pWnd, pDataObject, dwKeyState, point);

  return DROPEFFECT_NONE;
}

BOOL MZLaunchBarDropTarget::OnDrop( CWnd* pWnd, COleDataObject* pDataObject, DROPEFFECT dropEffect, CPoint point )
{
  if(m_pOwner)
    return m_pOwner->OnDrop(pWnd, pDataObject, dropEffect, point);

  return FALSE;
}

DROPEFFECT MZLaunchBarDropTarget::OnDropEx( CWnd* pWnd, COleDataObject* pDataObject, DROPEFFECT dropDefault, DROPEFFECT dropList, CPoint point )
{
  if(m_pOwner)
    return m_pOwner->OnDropEx(pWnd, pDataObject, dropDefault, dropList, point);

  return DROPEFFECT_NONE;
}

void MZLaunchBarDropTarget::OnDragLeave( CWnd* pWnd )
{
  if(m_pOwner)
    m_pOwner->OnDragLeave(pWnd);
}

bool MZLaunchBarDropTarget::Register()
{
  static bool bInProcedure = false;
  if (bInProcedure)
    return false;

  if(m_bRegistered)
    return true; // already registered

  bInProcedure = true;
  m_bRegistered = COleDropTarget::Register(m_pOwner) > 0;
  bInProcedure = false;
  return m_bRegistered;
}

void MZLaunchBarDropTarget::Unregister()
{
  m_bRegistered = false;
  COleDropTarget::Revoke();
}

//////////////////////////////////////////////////////////////////////////
MZLaunchBar::MZLaunchBar()
  : m_dwStyle(0)
  , m_nMarginY(2)
  , m_nMarginX(1)
  , m_nSpacing(3)
  , m_bSaveOnChange(true)
  , m_bDrawStateDirty(true)
  , m_bMouseOver(false)
  , m_bDblClk(false)
  , m_bCreatedByCreate(false)
  , m_bDragging(false)
  , m_bDropping(false)
  , m_bPrepareDragging(false)
  , m_bRegisteredForDropTarget(false)
  , m_bAllowDropAdd(false)
  , m_bAllowRearrange(false)
  , m_bToolTip(false)
  , m_pOverItem(nullptr)
  , m_pClickItem(nullptr)
  , m_DropTarget(this)
  , m_SeparatorIconIdx(-1)
  , m_hCursor(0)
{
  WNDCLASS wndclass;
  HINSTANCE hInst = AfxGetInstanceHandle();

  if(!(::GetClassInfo(hInst, MZLAUNCHBARCTRL_CLASSNAME, &wndclass)))
  {
    wndclass.style = CS_DBLCLKS|CS_HREDRAW|CS_VREDRAW;
    wndclass.lpfnWndProc = ::DefWindowProc;
    wndclass.cbClsExtra = wndclass.cbWndExtra = 0;
    wndclass.hInstance = hInst;
    wndclass.hIcon = NULL;
    wndclass.hCursor = LoadCursor(hInst, IDC_ARROW);
    wndclass.hbrBackground = (HBRUSH)COLOR_BTNFACE;
    wndclass.lpszMenuName = NULL;
    wndclass.lpszClassName = MZLAUNCHBARCTRL_CLASSNAME;

    if (!AfxRegisterClass(&wndclass))
      AfxThrowResourceException();
  }

  m_IconSize = CSize(16,16);
}

void MZLaunchBar::InitPaintResources()
{
  m_bkBrush.CreateSysColorBrush(COLOR_BTNFACE);
  m_bkItemBrush.CreateSysColorBrush(COLOR_BTNFACE);
  m_penItemFrame.CreatePen(PS_SOLID, 1, GetSysColor(COLOR_BTNSHADOW));

  m_crDropMark = GetSysColor(COLOR_HOTLIGHT);
}

MZLaunchBar::~MZLaunchBar()
{
  if(m_hCursor)
    ::DestroyCursor(m_hCursor);
}


BEGIN_MESSAGE_MAP(MZLaunchBar, CWnd)
  ON_WM_CREATE()
  ON_WM_NCCALCSIZE()
  ON_WM_PAINT()
  ON_WM_ERASEBKGND()
  ON_WM_MOUSEMOVE()
  ON_MESSAGE(WM_MOUSELEAVE, OnMouseLeave)
  ON_WM_RBUTTONDOWN()
  ON_WM_RBUTTONUP()
  ON_WM_CONTEXTMENU()
  ON_WM_LBUTTONDBLCLK()
  ON_WM_LBUTTONDOWN()
  ON_WM_LBUTTONUP()
  ON_NOTIFY_EX( TTN_NEEDTEXT, 0, OnToolTipNotify)
  ON_WM_SETCURSOR()
END_MESSAGE_MAP()

void MZLaunchBar::Init(bool bRefreshIcons)
{
  if(GetSafeHwnd() == 0)
    return;

  if(m_dwStyle & MZBS_DROPADD || m_dwStyle & MZBS_LAUNCH_DROPASPARM)
    RegisterDropTarget(true);
  else
    RegisterDropTarget(false);   

  if(bRefreshIcons)
  {
    InitImageList();
    RefreshIcons();
    AdjustHeight();
  }
}

void MZLaunchBar::Style( DWORD dwStyle )
{
  m_bAllowDropAdd = (m_dwStyle & MZBS_DROPADD) > 0;
  m_bAllowRearrange = (m_dwStyle & MZBS_DRAGREARRANGE) > 0;

  m_dwStyle = dwStyle; 
  Init(true);
}

int MZLaunchBar::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
  if (CWnd::OnCreate(lpCreateStruct) == -1)
    return -1;

  if(m_bToolTip)
  {
    InitToolTipCtrl();
  }
  return 0;
}

void MZLaunchBar::OnNcCalcSize(BOOL /*bCalcValidRects*/, NCCALCSIZE_PARAMS FAR* /*lpncsp*/) 
{
//   int nHeight = m_IconSize.cy + (m_nMargin * 2);
//   lpncsp->rgrc[0].bottom = lpncsp->rgrc[0].top + nHeight;
}

void MZLaunchBar::AdjustHeight()
{
  int nHeight = m_IconSize.cy + (m_nMarginY * 2); // +2 for the extra minmimum margin (1px up and down)

  // If WS_BOARDER then we need more height - However GetStyle does not return WS_BORDER even if it is set
  if(CWnd::GetStyle() & WS_BORDER) // DO NOT WORK
    nHeight += 6;

  CRect rc;
  GetWindowRect(&rc);
  CRect rcNewSize = rc;
  rcNewSize.bottom = rcNewSize.top + nHeight;

  if(rc.Height() != rcNewSize.Height())
    SetWindowPos(NULL, rcNewSize.left, rcNewSize.top, rcNewSize.Width(), rcNewSize.Height(), SWP_NOMOVE | SWP_NOZORDER);
}

void MZLaunchBar::PreSubclassWindow()
{
  CWnd::PreSubclassWindow();

  if(m_bCreatedByCreate == false)
  {
    if(!Create(m_dwStyle) )
      AfxThrowMemoryException();

    AdjustHeight();
  }
}

BOOL MZLaunchBar::Create( DWORD dwStyle, DWORD dwLBStyle, const RECT& rect, CWnd* pParentWnd, UINT nID, CCreateContext* pContext /*= NULL*/ )
{
  m_bCreatedByCreate = true;

  if(!CWnd::Create(MZLAUNCHBARCTRL_CLASSNAME, NULL, dwStyle, rect, pParentWnd, nID, pContext))
    return FALSE;

  return Create( dwLBStyle );
}

BOOL MZLaunchBar::Create( DWORD dwLBStyle )
{
  // Create Internal resources font, bitmap and so on
  m_dwStyle = dwLBStyle;

  m_bAllowDropAdd = (m_dwStyle & MZBS_DROPADD) > 0;
  m_bAllowRearrange = (m_dwStyle & MZBS_DRAGREARRANGE) > 0;


  if(m_bToolTip)
  {
    InitToolTipCtrl();
  }

  InitPaintResources();

  /*
  // Create Default Font here.
  CWnd* pWnd = GetParent();
  if(pWnd)
  {
    CFont* pFont = pWnd->GetFont();
    if(pFont)
    {
      LOGFONT lf;
      pFont->GetLogFont(&lf);

      m_fontDefault.DeleteObject();
      m_fontDefault.CreateFontIndirect(&lf);

      lf.lfUnderline = 1;
      m_fontHot.DeleteObject();
      m_fontHot.CreateFontIndirect(&lf);
    }
  }*/


//   if(m_dwStyle & MZBS_ACCEPTDROP)
//     ModifyStyleEx(0, WS_EX_ACCEPTFILES, 0);

  if(m_bCreatedByCreate)
    Init(true);

  return TRUE;
}
void MZLaunchBar::InitToolTipCtrl()
{
  CRect rect; 
  GetClientRect(rect);
  if(m_ToolTip.GetSafeHwnd() == 0)
  {
    m_ToolTip.Create(this, TTS_NOPREFIX);
    m_ToolTip.SetMaxTipWidth(450);

    m_bDrawStateDirty = true;
  }
}
bool MZLaunchBar::AddSeparatorIcon(HICON hIcon)
{
  m_SeparatorIconIdx = AddIcon(hIcon);
  return true;
}
BOOL MZLaunchBar::PreTranslateMessage(MSG* pMsg) 
{
  if(m_bToolTip)
    m_ToolTip.RelayEvent(pMsg);

  return __super::PreTranslateMessage(pMsg);
}

void MZLaunchBar::EnableTooltip(bool bEnable)
{
  m_bToolTip = bEnable;

  if(m_bToolTip)
    InitToolTipCtrl();

  if(m_ToolTip.GetSafeHwnd())
    m_ToolTip.Activate(m_bToolTip);
}

bool MZLaunchBar::RebuildTooltips()
{
  if(!(m_bToolTip && m_ToolTip.GetSafeHwnd()))
    return false;

  // Remove All
  int nCount = m_ToolTip.GetToolCount();
  for(int n = 0; n < nCount; ++n)
  {
    m_ToolTip.DelTool(this, n);
  }

  int nIDTool = 1;
  for(auto it = begin(m_vItems); it != end(m_vItems); ++it)
  {
    auto p = (*it);
    m_ToolTip.AddTool(this, LPSTR_TEXTCALLBACK , p->rcItemArea , nIDTool );
    ++nIDTool;
  }
  
  m_ToolTip.Activate(TRUE);
  return TRUE;
}

// Calculate the size needed for x Items
bool MZLaunchBar::CalcSize(CSize& sz, int nItems)
{
  sz.cy = m_IconSize.cy + (m_nMarginY * 2);
  
  int nItemWidth = (m_IconSize.cx + 2 ) + (m_nMarginX * 2);
  nItemWidth += m_nSpacing; // extra space between 'buttons'

  sz.cx = nItemWidth * nItems;
  return true;
}

DWORD MZLaunchBar::MaxVisibleItems()
{
  CRect rc;
  GetClientRect(&rc);

  CSize sz;
  CalcSize(sz, 1); // get width for 1 item

  return rc.Width() / sz.cx;
}

void MZLaunchBar::Load(const STLString& /*filename*/)
{
  m_bSaveOnChange = false;

  // Load and insert items

  m_bSaveOnChange = true;
}

void MZLaunchBar::Save()
{

}

BOOL MZLaunchBar::OnEraseBkgnd(CDC* /*pDC*/)
{
  return FALSE;
}
void MZLaunchBar::OnPaint()
{
  CPaintDC dc(this); // device context for painting

  CMemoryDC MemDC(&dc);
  Draw(&MemDC);
}

void MZLaunchBar::Draw(CDC* pDC)
{
//  CFont *pOldFont = pDC->SelectObject(&m_fontDefault);

  CRect rectClip;
  if (pDC->GetClipBox(&rectClip) == ERROR)
    return;

  CRect rcClientRect;
  GetClientRect( &rcClientRect );

  //////////////////////////////////////////////////////////////////////////
  // Draw Background 
  DrawBackground(pDC, rcClientRect);

  UpdateItemAreas(pDC, rcClientRect);

  DrawItems(pDC, rcClientRect);
}

void MZLaunchBar::DrawBackground(CDC* pDC, const CRect& rc)
{
  pDC->FillRect(rc,&m_bkBrush);
}

int MZLaunchBar::GetItemWith(MZLaunchItem* pItem)
{
  return pItem->width + (m_nMarginX * 2); // *2 for before and after margin
}

void MZLaunchBar::UpdateItemAreas(CDC* /*pDC*/, const CRect& rcItemsArea)
{
  if(m_bDrawStateDirty == false)
    return;

  CRect rcItem = rcItemsArea;

  if(m_dwStyle & MZBS_ALIGNBOTTOM)
  {
    rcItem.top = rcItem.bottom - (m_IconSize.cy + ((m_nMarginY) * 2));
  }
  else if(m_dwStyle & MZBS_ALIGNVCENTER)
  {
    int nHeight = m_IconSize.cy + ((m_nMarginY) * 2);
    if(rcItem.Height() > nHeight)
    {
      int nOffset = (rcItem.Height() - nHeight) / 2;
      rcItem.top += nOffset;
      rcItem.bottom = rcItem.top + m_IconSize.cy + ((m_nMarginY) * 2);
    }
    else
    {
      // top 
      rcItem.bottom = rcItem.top + m_IconSize.cy + ((m_nMarginY) * 2) ;
    }
  }
  else // Top Align
    rcItem.bottom = rcItem.top + m_IconSize.cy + ((m_nMarginY) * 2) ;

  for(auto it = begin(m_vItems); it != end(m_vItems); ++it)
  {
    MZLaunchItem* pItem = (*it).get();
    rcItem.left += m_nSpacing;
    rcItem.right = rcItem.left + GetItemWith(pItem);
    
    pItem->rcItemArea = rcItem;

    rcItem.left = rcItem.right;
  }

  // area changed. so tooltips must be updated.
  RebuildTooltips();

  m_bDrawStateDirty = false;
}

void MZLaunchBar::DrawItems(CDC* pDC, const CRect& rc)
{
  if(m_vItems.empty())
    return; // No items to draw

  CRect rcItem = rc;
  
  bool bDrawInsertLine = false;
  if(m_bDropping)
  {
    if(m_pOverItem == nullptr)
      bDrawInsertLine = true;
  }
  
  DWORD nItem = 0;
  for(auto it = begin(m_vItems); it != end(m_vItems); ++it, ++nItem)
  {
    auto pItem = (*it).get();
    if(m_bDragging)
      rcItem = pItem->rcItemAreaDragging;
    else
      rcItem = pItem->rcItemArea;

    // draw insert line BEFORE an item.
    if(m_bDropping && bDrawInsertLine)
    {
      if(m_ptDragPos.x < rcItem.left && m_ptDragPos.x >= (rcItem.left - m_nSpacing))
      {
        CRect rcLine = rcItem;
        rcLine.left -= m_nSpacing;
        rcLine.right = rcLine.left + m_nSpacing;
        DrawInsertLine(pDC, rcLine);
        bDrawInsertLine = false;
      }
    }
    
    if(m_bDragging)
    {
      if(m_pClickItem != pItem)
      {
        if(rcItem.right > rc.right) // item is out of draw area, so do not draw it
          continue;

        DrawItem(pDC, rcItem, pItem);
      }
    }
    else
    {
      if(rcItem.right > rc.right) // item is out of draw area, so do not draw it
        continue;

      DrawItem(pDC, rcItem, pItem);
    }
  }

  // Draw Insert line at the end of the list of icons
  if(m_bDropping && bDrawInsertLine)
  {
    MZLaunchItem* pItem = m_vItems[m_vItems.size()-1].get();
    if(pItem)
    {
      CRect rcLine = rcItem;
      rcLine.left = rcLine.right;
      rcLine.right = rcLine.left + m_nSpacing;
      DrawInsertLine(pDC, rcLine);
    }
  }

  // when dragging always draw the drag item LAST. so it is drawn over existing item if they overlap
  if(m_bDragging && m_pClickItem)
    DrawItem(pDC, m_pClickItem->rcItemAreaDragging, m_pClickItem);
}

void MZLaunchBar::DrawInsertLine(CDC* pDC, const CRect& rc)
{
  pDC->FillSolidRect(rc, m_crDropMark);
}

void MZLaunchBar::DrawItem(CDC* pDC, const CRect& rc, const MZLaunchItem* pItem)
{
  if(pItem == nullptr || (pItem->nIconIdx == -1 && pItem->bSeparator == false))
    return;

  if(m_bDragging == false)
  {
    // Draw Item Background and frame
    DrawItemBackground(pDC, rc, pItem);
  }

  CPoint ptCenter = rc.CenterPoint();
  CRect rcIcon = rc;

  if(pItem->bSeparator && pItem->nIconIdx == -1)
  {
    rcIcon.left = ptCenter.x - 1;
    rcIcon.right = ptCenter.x + 1;

    rcIcon.top = ptCenter.y - m_IconSize.cy/2;
    rcIcon.bottom = ptCenter.y + m_IconSize.cy/2;

    // No Separator icon used so draw a separator
    pDC->DrawEdge(&rcIcon, BDR_RAISEDINNER, BF_RECT|BF_LEFT|BF_RIGHT); 
  }
  else
  {
    // If button is down. move the icon down a small amount
    if(pItem->dwState & MZLIS_BUTTONDOWN && m_bDragging == false)
      ptCenter.y += 1;

    rcIcon.left = ptCenter.x - m_IconSize.cx/2;
    rcIcon.right = ptCenter.x + m_IconSize.cx/2;
    rcIcon.top = ptCenter.y - m_IconSize.cy/2;
    rcIcon.bottom = ptCenter.y + m_IconSize.cy/2;

    m_Images.Draw(pDC, pItem->nIconIdx, CPoint(rcIcon.left, rcIcon.top), ILD_TRANSPARENT);
  }

  /*
  // Draw Rect around item bounds
  COLORREF clrBack = RGB(100,100,222);
  COLORREF clrBack2= RGB(100,200,222);
  CRect rcEdge = rc;
  pDC->Draw3dRect(rcEdge, clrBack, clrBack2);

  // Draw Rect without margin
  rcEdge.DeflateRect(m_nMargin,0,m_nMargin,0);
  pDC->Draw3dRect(rcEdge, clrBack, clrBack2);
  */
  
}

void MZLaunchBar::DrawItemBackground(CDC* pDC,const CRect& rc,const MZLaunchItem* pItem)
{
  bool bDrawHiglightFrame = false;

  if(pItem->dwState & MZLIS_BUTTONDOWN && m_dwStyle & MZBS_DRAW_HLIGHT_DOWN)
    bDrawHiglightFrame = true;

  if(pItem->dwState & MZLIS_HOVEROVER && m_dwStyle & MZBS_DRAW_HLIGHT_HOVER)
    bDrawHiglightFrame = true;

  if(pItem->bSeparator)
    bDrawHiglightFrame = false;

  if(bDrawHiglightFrame)
  {
    CRect rcEdge = rc;

    CBrush* pOldBrush = pDC->SelectObject(&m_bkItemBrush);
    CPen* pOldPen = pDC->SelectObject(&m_penItemFrame);

    CPoint pt(3,3);
    pDC->RoundRect(rcEdge, pt);

    // restore pen and brush
    pDC->SelectObject(pOldPen);
    pDC->SelectObject(pOldBrush);
  }
}

void MZLaunchBar::SetIconSize(CSize cz, bool bRefreshIcons)
{
  bool bIconSizeChanged = false;
  if(m_IconSize != cz)
    bIconSizeChanged = true;

  m_IconSize = cz;
  
  if(GetSafeHwnd() && m_Images.GetSafeHandle())
  {
    if(bRefreshIcons)
      RefreshIcons();
  }
}

void MZLaunchBar::InitImageList()
{
  if(m_Images.GetSafeHandle())
    m_Images.DeleteImageList();

  m_Images.Create(m_IconSize.cx, m_IconSize.cy, ILC_MASK|ILC_COLOR32, 8, 4);
}

void MZLaunchBar::RefreshIcons()
{
  if(m_vItems.empty() == false)
  {
    m_Images.DeleteImageList();
    InitImageList();
  }

  for(auto it = begin(m_vItems); it != end(m_vItems); ++it)
  {
    MZLaunchItem* pItem = (*it).get();

    pItem->width = m_IconSize.cx + 2;

    HICON hIcon = GetIcon(pItem);
    pItem->nIconIdx = AddIcon(hIcon);
    DestroyIcon(hIcon);
  }

  if(!m_vItems.empty())
  {
    m_bDrawStateDirty = true;
    Invalidate();
  }
}

std::shared_ptr<MZLaunchItem> MZLaunchBar::CreateSeparatorItem()
{
  std::shared_ptr<MZLaunchItem> pItem = std::make_shared<MZLaunchItem>();
  pItem->width = m_IconSize.cx + 2;
  pItem->bSeparator = true;
  pItem->nIconIdx = m_SeparatorIconIdx;
  return pItem;
}

std::shared_ptr<MZLaunchItem> MZLaunchBar::CreateItem( const STLString& programPath, const STLString& tooltip)
{
  std::shared_ptr<MZLaunchItem> pItem = std::make_shared<MZLaunchItem>();
  pItem->programPath = programPath;
  pItem->tooltip = tooltip;

  pItem->width = m_IconSize.cx + 2;
  
  return pItem;
}

std::shared_ptr<MZLaunchItem> MZLaunchBar::AddItem( const STLString& programPath, const STLString& tooltip, bool bGetIcon )
{
  std::shared_ptr<MZLaunchItem> pItem = CreateItem(programPath, tooltip);
  
  if(bGetIcon)
    UpdateIcon(pItem.get());
  
  m_vItems.push_back(pItem);
  OnChanged(pItem.get(), Change_Added);

  return pItem;
}

std::shared_ptr<MZLaunchItem> MZLaunchBar::AddSeparatorItem()
{
  auto p = CreateSeparatorItem();
  m_vItems.push_back(p);

  OnChanged(p.get(), Change_Added);
  return p;
}

void MZLaunchBar::UpdateIcon(MZLaunchItem* pItem)
{
  HICON hIcon = GetIcon(pItem);
  pItem->nIconIdx = AddIcon(hIcon);
  DestroyIcon(hIcon); // A Copy was added in AddIcon(...)
}

void MZLaunchBar::RemoveItem(MZLaunchItem* pItem)
{
  for(auto it = begin(m_vItems); it != end(m_vItems); ++it)
  {
    if((*it).get() == pItem)
    {
      m_vItems.erase(it);
      OnChanged(nullptr, Change_Removed);
      Invalidate(TRUE);
      m_bDrawStateDirty = true;
      return;
    }
  }
}

DWORD MZLaunchBar::GetItemCount()
{
  return (DWORD)m_vItems.size();
}

int MZLaunchBar::AddIcon(HICON hIcon)
{
  if(hIcon == 0)
    return -1;

  return m_Images.Add(hIcon);
}

HICON MZLaunchBar::GetIcon(MZLaunchItem* pItem)
{
  if(!pItem)
    return 0;

  if(pItem->altIconIdx != -1)
  {
    HICON hIcon = GetIcon(pItem->altIconPath, pItem->altIconIdx, pItem->dwCustomParam);
    if(hIcon)
      return hIcon;
  }

  return GetIcon(pItem->programPath);
}

// Remember to DESTORYICON when done
HICON MZLaunchBar::GetIcon(const STLString& file, int nIconIdx, DWORD /*dwCustomParam*/)
{
  if(file.find('%') != STLString::npos)
  {
    return GetIcon(ExpandStringTags(file), nIconIdx);
  }

  HICON hIcon = 0;
  UINT nIconIds = 0;

  InitSHExtractIconsW();
  if(g_pSHExtractIconsW)
  {
    /*UINT nResult = */g_pSHExtractIconsW(file.c_str(), nIconIdx, m_IconSize.cx, m_IconSize.cx, &hIcon, &nIconIds, 1, 0);
    return hIcon;
  }
 
  return 0;
}

HICON MZLaunchBar::GetIcon(const STLString& programPath, bool bExpand)
{
  if(bExpand && programPath.find('%') != STLString::npos)
  {
    return GetIcon(ExpandStringTags(programPath), false);
  }

  SHFILEINFO shfi;
  ZeroMemory(&shfi,sizeof(SHFILEINFO));
  DWORD_PTR dOK = 0;
//  BOOL bShared = FALSE;

  long lFileAttribute = FILE_ATTRIBUTE_NORMAL;

  DWORD dwAttribute = GetFileAttributes(programPath.c_str());
  if(dwAttribute != INVALID_FILE_ATTRIBUTES && dwAttribute & FILE_ATTRIBUTE_DIRECTORY)
    lFileAttribute = FILE_ATTRIBUTE_DIRECTORY;

  // Get Icon From .EXE File
  if( programPath.length() > 0 )
  {
    HICON hIcon = 0;
    UINT nIconIds = 0;

    // Try to Get File Icon. (for folder we use SHGetFileInfo)
    if(!(lFileAttribute & FILE_ATTRIBUTE_DIRECTORY))
    {
      InitSHExtractIconsW();
      if(g_pSHExtractIconsW)
      {
        UINT nResult = g_pSHExtractIconsW(programPath.c_str(), 0, m_IconSize.cx, m_IconSize.cx, &hIcon, &nIconIds, 1, 0);
        if(nResult != -1 && hIcon != 0)
          return hIcon;
      }
    }

    // Failed to get File icon or path was a folder.

    if(m_IconSize.cx > 16)
      dOK = SHGetFileInfo(programPath.c_str(), lFileAttribute , &shfi , sizeof(SHFILEINFO), SHGFI_USEFILEATTRIBUTES | SHGFI_DISPLAYNAME | SHGFI_TYPENAME | SHGFI_ICON | SHGFI_LARGEICON |SHGFI_ICONLOCATION);
    else
      dOK = SHGetFileInfo(programPath.c_str(), lFileAttribute , &shfi , sizeof(SHFILEINFO), SHGFI_USEFILEATTRIBUTES | SHGFI_DISPLAYNAME | SHGFI_TYPENAME | SHGFI_ICON | SHGFI_SMALLICON | SHGFI_ICONLOCATION);
      
    if( dOK > 0 && shfi.hIcon)
      return shfi.hIcon;
  }
  
  return 0;
}

MZLaunchItem*  MZLaunchBar::GetItemAtPos(const CPoint& pt)
{
  for(auto it = begin(m_vItems); it != end(m_vItems); ++it)
  {
    MZLaunchItem* pItem = (*it).get();

    if(pItem->rcItemArea.PtInRect(pt))
      return pItem;
  }

  return nullptr;
}

MZLaunchItem*  MZLaunchBar::GetItemBeforePos(const CPoint& pt)
{
  MZLaunchItem* pPrevItem = nullptr;

  for(auto it = begin(m_vItems); it != end(m_vItems); ++it)
  {
    MZLaunchItem* pItem = (*it).get();

    if(pItem->rcItemArea.left > pt.x)
      return pPrevItem;  

    pPrevItem = pItem;
  }

  return nullptr;
}

MZLaunchItem*  MZLaunchBar::GetItemAfterPos(const CPoint& pt)
{
  for(auto it = begin(m_vItems); it != end(m_vItems); ++it)
  {
    MZLaunchItem* pItem = (*it).get();

    if(pItem->rcItemArea.right >= pt.x)
      return pItem;      
  }

  return nullptr;
}

MZLaunchItem*  MZLaunchBar::GetItemBefore(const MZLaunchItem* pItem)
{
  MZLaunchItem* pPrevItem = nullptr;
  for(auto it = begin(m_vItems); it != end(m_vItems); ++it)
  {
    MZLaunchItem* pCurItem = (*it).get();

    if(pCurItem == pItem)
      return pPrevItem;

    pPrevItem = pCurItem;
  }

  return nullptr;
}

void MZLaunchBar::InvalidateItem(MZLaunchItem* pItem)
{
  if(pItem)
    InvalidateRect(pItem->rcItemArea);
  else
    InvalidateRect(NULL);
}


void MZLaunchBar::OnMouseMove(UINT nFlags, CPoint point)
{
  if(!m_bMouseOver)
  {
    m_bMouseOver = true;
    TRACKMOUSEEVENT tme;
    tme.cbSize = sizeof(tme);
    tme.dwFlags = TME_LEAVE;
    tme.hwndTrack = m_hWnd;
    _TrackMouseEvent(&tme);
    Invalidate();
  }

  if(m_bDropping)
  {
    CWnd::OnMouseMove(nFlags, point);
    return;
  }

  bool bOverItemChanged = false;
  MZLaunchItem* pItem = GetItemAtPos(point);
  if(pItem != m_pOverItem)
  {
    bOverItemChanged = true;
    if(m_pOverItem)
    {
      OnItemOverLeave(m_pOverItem);

      if(m_pOverItem->dwState & MZLIS_HOVEROVER)
        m_pOverItem->dwState &= ~MZLIS_HOVEROVER;
    }

    // Set new over item
    m_pOverItem = pItem;
    if(m_pOverItem)
    {
      m_pOverItem->dwState |= MZLIS_HOVEROVER;
      OnItemOverEnter(m_pOverItem);
    }
  }

  if(m_pOverItem)
    OnItemOver(m_pOverItem);
  
  if(bOverItemChanged)
    InvalidateRect(NULL);
  
  if(m_bPrepareDragging && m_bDragging == false)
  {
    CSize szSensitivity(5,5);
    CPoint ptHotDelta = m_ptDragPos - point;

    CPoint ptMouse;
    GetCursorPos(&ptMouse);
    ScreenToClient(&ptMouse);
    if (abs(ptHotDelta.x) > szSensitivity.cx && abs(ptHotDelta.y) < szSensitivity.cy)
    {
      TRACE(_T("Start Dragging Tab\n"));
      m_bDragging = true;
      m_bPrepareDragging = false;
      InitDragPositions();
    }
  }

  if(m_bDragging && m_ptDragPos != point)
  {
    RecalculateDragPositions(point);
    m_ptDragPos = point;
  }

  CWnd::OnMouseMove(nFlags, point);
}

void MZLaunchBar::InitDragPositions()
{
  m_vDragItems.clear();
  for(auto it = begin(m_vItems); it != end(m_vItems); ++it)
  {
    (*it)->rcItemAreaDragging = (*it)->rcItemArea;
    m_vDragItems.push_back((*it));
  }
}

void MZLaunchBar::CommitDragPositions()
{
  m_vItems.clear();
  m_vItems = m_vDragItems;
  OnChanged(m_pClickItem, Change_Rearranged);
  
  m_bDrawStateDirty = true; // so that positions are recalculated

  m_vDragItems.clear();
}

namespace
{
  bool IsOverlapping(const CRect& rcDragItem, const CRect& rcItem, int percOverlapping)
  {
    int nMinOverlapSize = min(rcDragItem.Width(), rcItem.Width());
    nMinOverlapSize = (int)(nMinOverlapSize * ((double)percOverlapping/100));

    bool bLeftSideOverItem = false;
    bool bRightSideOverItem = false;

    // Overlapping on the left size ?
    if(rcDragItem.left < rcItem.right && rcDragItem.left > rcItem.left)
    {
      int overlapping = min(rcItem.right,rcDragItem.right) - rcDragItem.left;
      if(overlapping >= nMinOverlapSize && rcDragItem.left < (rcItem.left + overlapping) + rcItem.Width())
         bLeftSideOverItem = true;
    }
    
    // right side of drag item is over item.
    if(rcDragItem.right > rcItem.left && rcDragItem.right < rcItem.right)
    {
      int overlapping = rcDragItem.right - rcItem.left;
      // rcDragitem should NOT be allow to be located inside the overitem
      if(overlapping >= nMinOverlapSize && rcDragItem.right > (rcItem.left - overlapping) + rcItem.Width())
        bRightSideOverItem = true;
    }

    if(bLeftSideOverItem || bRightSideOverItem)
      return true;

    return false;
  }


  int FindPosOfItem(std::vector<std::shared_ptr<MZLaunchItem>>& vItems, MZLaunchItem* pItem)
  {
    size_t nMax = vItems.size();
    for(size_t n = 0; n < nMax; ++n)
    {
      if(vItems[n].get() == pItem)
        return (int)n;
    }
    return -1; // not found
  }

  
  std::shared_ptr<MZLaunchItem> GetSharedItem(std::vector<std::shared_ptr<MZLaunchItem>>& vItems, MZLaunchItem* pItem)
  {
    for(auto it = begin(vItems); it != end(vItems); ++it)
    {
      if((*it).get() == pItem)
        return (*it);
    }
    return nullptr;
  }

 


}
//  When dragging an icon, it updated the layout while dragging. But if dragging is aborted it should revert to original state.
//  so all the temporary locations need to be stored separate
void MZLaunchBar::RecalculateDragPositions(const CPoint& pt)
{
  CPoint distance = pt - m_ptDragPos;

  // New position of the item we are dragging.
  if(m_pClickItem)
  {
    m_pClickItem->rcItemAreaDragging.left += distance.x;
    m_pClickItem->rcItemAreaDragging.right += distance.x;
  }

  
  bool bIsOverItem = false;

  CRect rcDragItem = m_pClickItem->rcItemAreaDragging;

  // if any item is located under the dragged item to more then 55% then Move insert moved item into that location.
  MZLaunchItem* pItem = nullptr;

  int nPos = FindPosOfItem(m_vDragItems, m_pClickItem);
  if(nPos == -1)
  {
    TRACE(_T("Drag item not found\n"));
    return;
  }

  bool bRecalc = false;
  // find overlapped item
  for(auto it = begin(m_vDragItems); it != end(m_vDragItems); ++it)
  {
    pItem = (*it).get();
    if(pItem == m_pClickItem)
      continue;

    CRect rcItem = pItem->rcItemAreaDragging;

    // If the drag item is overlapping another tab, with either 40% if the drag tab or the overlapped tab.
    if(IsOverlapping(rcDragItem, rcItem, 50))
    {
      TRACE(_T("Item Overlapping\n"));
      bIsOverItem = true;
                  
      std::shared_ptr<MZLaunchItem> pDragItem = GetSharedItem(m_vDragItems, m_pClickItem);

      RemoveItem(m_vDragItems, m_pClickItem);

      if(rcItem.left < rcDragItem.left)
        InsertBefore(m_vDragItems, pItem, pDragItem);
      else
        InsertAfter(m_vDragItems, pItem, pDragItem);
        
      bRecalc = true;
      break;
    }
  }

  if(bRecalc)
  {
    CRect rcItem = m_pClickItem->rcItemAreaDragging;
    rcItem.left = 0;
    for(auto it = begin(m_vDragItems); it != end(m_vDragItems); ++it)
    {
      MZLaunchItem* pItem = (*it).get();
      
      rcItem.left += m_nSpacing;
      rcItem.right = rcItem.left + GetItemWith(pItem);

      if(pItem != m_pClickItem) // Do not update current drag item
        pItem->rcItemAreaDragging = rcItem;

      rcItem.left = rcItem.right;
    }
  }

  Invalidate();
}

LRESULT MZLaunchBar::OnMouseLeave(WPARAM /*wParam*/, LPARAM /*lParam*/)
{
  if (m_bMouseOver)
  {
    m_bMouseOver = false;
    Invalidate();
  }
  
  if(m_pOverItem)
  {
    OnItemOverLeave(m_pOverItem);

    if(m_pOverItem->dwState & MZLIS_HOVEROVER)
      m_pOverItem->dwState &= ~MZLIS_HOVEROVER;
  }

  m_pOverItem = nullptr;
  return 0;
}

void MZLaunchBar::OnRButtonDown(UINT nFlags, CPoint point)
{
  m_bDblClk = false;
  m_pClickItem = GetItemAtPos(point);

  if(m_pClickItem)
  {
    m_pClickItem->dwState |= MZLIS_BUTTONDOWN;
    Invalidate();
    SetCapture();
  }

  CWnd::OnRButtonDown(nFlags, point);
}

void MZLaunchBar::OnRButtonUp(UINT nFlags, CPoint point)
{
  if(m_pClickItem)
  {
    ReleaseCapture();
    if(m_pClickItem->dwState & MZLIS_BUTTONDOWN)
    {
      m_pClickItem->dwState &= ~MZLIS_BUTTONDOWN;
      Invalidate();
    }
  }

  CWnd::OnRButtonDown(nFlags, point);
}

void MZLaunchBar::OnContextMenu(CWnd* /*pWnd*/, CPoint point)
{
  ReleaseCapture();

  CPoint ptClient = point;
  ScreenToClient(&ptClient);

  // Need the shared item. because if the item is removed using the context menu
  // we will crash below when removing flags.
  auto pItem = GetSharedItem(m_vItems, GetItemAtPos(ptClient));
  if(pItem && m_dwStyle & MZBS_CONTEXTMENU)
  {
    if(m_dwStyle & MZBS_SHELLMENU)
      OnShellContextMenu(pItem.get(), point);
    else
      OnContextMenu(pItem.get(), point);
  }
  else
  {
    if(pItem == nullptr && m_dwStyle & MZBS_CONTEXTMENUWS)
      OnContextMenu((MZLaunchItem*)nullptr, point);
  }

  if(m_pClickItem && m_pClickItem->dwState & MZLIS_BUTTONDOWN)
  {
    m_pClickItem->dwState &= ~MZLIS_BUTTONDOWN;
    Invalidate();
  }

  m_pClickItem = nullptr;
}

void MZLaunchBar::OnShellContextMenu(MZLaunchItem* pItem, const CPoint& pt)
{
  
#ifdef MZLUNCHBAR_USE_SHELLCONTEXTMENU
  CShellContextMenu scm;
    
  STLString strProgramPath = ExpandStringTags(pItem->programPath);
  scm.SetObjects(strProgramPath.c_str());
  scm.ShowContextMenu(this, pt);
#else
  UNREFERENCED_PARAMETER(pt);
  UNREFERENCED_PARAMETER(pItem);
#endif
}

void MZLaunchBar::OnLButtonDblClk(UINT nFlags, CPoint point)
{
  m_bDblClk = true;
  m_pClickItem = GetItemAtPos(point);
  CWnd::OnLButtonDblClk(nFlags, point);
}

void MZLaunchBar::OnLButtonDown(UINT nFlags, CPoint point)
{
  m_bDblClk = false;
  m_pClickItem = GetItemAtPos(point);

  if(m_pClickItem && m_pClickItem->bSeparator == false)
  {
    m_pClickItem->dwState |= MZLIS_BUTTONDOWN;
    Invalidate();
  }
  
  CWnd::OnLButtonDown(nFlags, point);

  EnterDragMode(m_pClickItem, point);
}

void MZLaunchBar::OnLButtonUp(UINT nFlags, CPoint point)
{
  // Make sure that mouse was not moved away from item during click
  MZLaunchItem* pItem = GetItemAtPos(point);
  
  if(pItem && pItem == m_pClickItem && m_bDragging == false)
  {
    if(m_bDblClk)
    {
      OnLButtonDblClick(point);
      m_bDblClk = false;
    }
    else
    {
      OnLButtonClick(point);
    }
  }

  if(m_pClickItem && m_pClickItem->dwState & MZLIS_BUTTONDOWN)
    m_pClickItem->dwState &= ~MZLIS_BUTTONDOWN;

  if(m_bDragging)
  {
    TRACE(_T("Stop Dragging\n"));
    m_bPrepareDragging = false;
    m_bDragging = false;

    if(m_pClickItem)
    {
      CommitDragPositions();
    }

    ReleaseCapture();
    Invalidate();
  }
  else if(m_bPrepareDragging)
  {
    TRACE(_T("Cleanup after Prepare Dragging\n"));
    m_bPrepareDragging = false;
    m_bDragging = false;
    ReleaseCapture();
    Invalidate();
  }
  
  m_pClickItem = nullptr;

  CWnd::OnLButtonUp(nFlags, point);
}

void MZLaunchBar::OnLButtonClick(const CPoint& pt)
{
  MZLaunchItem* pItem = GetItemAtPos(pt);
  if(pItem)
    OnLButtonClick(pItem, pt);
}

void MZLaunchBar::OnLButtonDblClick(const CPoint& pt)
{
   MZLaunchItem* pItem = GetItemAtPos(pt);
   if(pItem)
    OnLButtonDblClick(pItem, pt);
}

void MZLaunchBar::OnLButtonClick(MZLaunchItem* pItem, const CPoint& /*pt*/)
{
  bool bAsAdmin = false;

  if(!(m_dwStyle & MZBS_LAUNCH_CLICK))
    return;

  if(m_dwStyle & MZBS_LAUNCHADMIN_SHIFT && IsSHIFTpressed())
    bAsAdmin = true;

  if(m_dwStyle & MZBS_LAUNCHADMIN_CTRL && IsCTRLpressed())
    bAsAdmin = true;

  std::vector<STLString> vEmpty;
  OnLaunchItem(pItem, bAsAdmin, vEmpty);
}

void MZLaunchBar::OnLButtonDblClick(MZLaunchItem* pItem, const CPoint& /*pt*/)
{
  bool bAsAdmin = false;

  if(!(m_dwStyle & MZBS_LAUNCH_DBLCLICK))
    return;

  if(m_dwStyle & MZBS_LAUNCHADMIN_SHIFT && IsSHIFTpressed())
    bAsAdmin = true;

  if(m_dwStyle & MZBS_LAUNCHADMIN_CTRL && IsCTRLpressed())
    bAsAdmin = true;

  std::vector<STLString> vEmpty;
  OnLaunchItem(pItem, bAsAdmin, vEmpty);
}

void MZLaunchBar::OnContextMenu(MZLaunchItem* pItem, const CPoint& pt)
{
  CMenu popupMenu;
  if(!popupMenu.CreatePopupMenu())
    return;

  CPoint ptClient;
  ptClient = pt;
  ScreenToClient(&ptClient);

  if(pItem)
  {
    if(pItem->bSeparator  == false)
    {
      popupMenu.AppendMenu(MF_STRING, MZLBC_OPEN, _T("Open"));
      popupMenu.AppendMenu(MF_STRING, MZLBC_OPENADMIN, _T("Open as Administrator"));

      if(m_dwStyle & MZBS_ALLOW_REMOVE || m_dwStyle & MZBS_ALLOW_CUSTOMIZE )
        popupMenu.AppendMenu(MF_SEPARATOR, 12, _T(""));
    }

    if(m_dwStyle & MZBS_ALLOW_REMOVE)
      popupMenu.AppendMenu(MF_STRING, MZLBC_REMOVE, _T("Remove"));

    if(m_dwStyle & MZBS_ALLOW_CUSTOMIZE && pItem->bSeparator == false)
      popupMenu.AppendMenu(MF_STRING, MZLBC_CUSTOMIZE, _T("Customize..."));
  }
  else
  {
    if(m_dwStyle & MZBS_DRAGREARRANGE || m_dwStyle & MZBS_DROPADD)
      popupMenu.AppendMenu(MF_STRING, MZLBC_ADDSEP, _T("Add Separator"));

    bool bMenuSeparatorAdded = false;
    if(m_dwStyle & MZBS_DROPADD)
    {
      popupMenu.AppendMenu(MF_SEPARATOR, 0 , _T(""));
      bMenuSeparatorAdded = true;

      DWORD dwFlags = MF_STRING;
      if(m_bAllowDropAdd)
        dwFlags |= MF_CHECKED;
      popupMenu.AppendMenu(dwFlags, MZLBC_ALLOWDROPADD, _T("Allow - Add button by dropping file/folder"));
    }
    if(m_dwStyle & MZBS_DRAGREARRANGE)
    {
      if(!bMenuSeparatorAdded)
        popupMenu.AppendMenu(MF_SEPARATOR, 0 , _T(""));

      DWORD dwFlags = MF_STRING;
      if(m_bAllowRearrange)
        dwFlags |= MF_CHECKED;
      popupMenu.AppendMenu(dwFlags, MZLBC_ALLOWREARRANGE, _T("Allow - Rearrange of buttons"));
    }
  }
  int nCmdResult = popupMenu.TrackPopupMenu(TPM_LEFTALIGN|TPM_LEFTBUTTON|TPM_VERTICAL|TPM_RETURNCMD, pt.x, pt.y, this);

  if(nCmdResult == MZLBC_ADDSEP)
  {
    // Find Item after Point
    if(pItem)
      pItem = GetItemBefore(pItem);
    else
      pItem = GetItemBeforePos(ptClient);
  }

  OnHandleContextMenuCommand(nCmdResult, pItem);
}

void MZLaunchBar::OnHandleContextMenuCommand(int nCmd, MZLaunchItem* pItem)
{
  if(nCmd == MZLBC_OPEN)
  {
    std::vector<STLString> vEmpty;
    OnLaunchItem(pItem, false, vEmpty);
  }
  else if(nCmd == MZLBC_OPENADMIN)
  {
    std::vector<STLString> vEmpty;
    OnLaunchItem(pItem, true, vEmpty);
  }
  else if(nCmd == MZLBC_REMOVE)
    OnRemoveItem(pItem);
  else if(nCmd == MZLBC_CUSTOMIZE)
    OnCustomizeItem(pItem);
  else if(nCmd == MZLBC_ADDSEP)
    OnAddSeparator(pItem);
  else if(nCmd == MZLBC_ALLOWDROPADD)
    m_bAllowDropAdd = !m_bAllowDropAdd;
  else if(nCmd == MZLBC_ALLOWREARRANGE)
    m_bAllowRearrange = !m_bAllowRearrange;
}

void MZLaunchBar::OnItemOverLeave(MZLaunchItem* /*pItemLeft*/)
{}

void MZLaunchBar::OnItemOverEnter(MZLaunchItem* /*pItem*/)
{}

void MZLaunchBar::OnItemOver(MZLaunchItem* /*pItem*/)
{}

// if bInsert == false then pItem is the item that the files was dropped on
// if bInsert == true then pItem is the item we and to insert the new items BEFORE.
// if bInsert == true and pItem == nullptr then insert item at end
void MZLaunchBar::OnDropFiles(std::vector<STLString>& vDroppedFiles, MZLaunchItem* pItem, bool bInsert)
{

  if(bInsert)
  {
    if(!(m_dwStyle & MZBS_DROPADD))
      return;

    if(m_bAllowDropAdd == false)
      return;

    // Only Add the first dropped file
    STLString filepath = vDroppedFiles.at(0);

    if(!AcceptDroppedItem(filepath))
      return;

    auto pNewItem = CreateNewLaunchItemFromDroppedPath(filepath);
    if(pNewItem)
    {
      if(pItem)
        InsertBefore(m_vItems, pItem, pNewItem);
      else
        m_vItems.push_back(pNewItem);

      m_bDrawStateDirty = true;
      OnChanged(pNewItem.get(), Change_Added);
      Invalidate();
    }
  }
  else
  {
    if(!(m_dwStyle & MZBS_LAUNCH_DROPASPARM))
      return;

    bool bAsAdmin = false;
    if(m_dwStyle & MZBS_LAUNCHADMIN_SHIFT && IsSHIFTpressed())
      bAsAdmin = true;

    if(m_dwStyle & MZBS_LAUNCHADMIN_CTRL && IsCTRLpressed())
      bAsAdmin = true;

    OnLaunchItem(pItem, bAsAdmin, vDroppedFiles);
  }
}

std::shared_ptr<MZLaunchItem> MZLaunchBar::CreateNewLaunchItemFromDroppedPath(const STLString& filepath)
{
  std::shared_ptr<MZLaunchItem> pItem = CreateItem(filepath, filepath);
  
  if(pItem)
    UpdateIcon(pItem.get());

  return pItem;
}

void MZLaunchBar::OnLaunchItem(MZLaunchItem* pItem, bool bAsAdmin, std::vector<STLString>& vFileParameters)
{
  if(!pItem)
    return;

  CString strOperation;
  if(bAsAdmin)
    strOperation = _T("runas");
  else
    strOperation = _T("open");

  STLString param;
  STLString parameters;

  // only first file is supported. Most external programs only support 1 file
  if(vFileParameters.size() > 0)
  {
    // quote parameter
    parameters += _T("\"");
    parameters += vFileParameters[0];
    parameters += _T("\"");
  }
    
  // add item parameters first
  if(pItem->programParameters.empty() == false)
    param = pItem->programParameters;
  
  // add file parameters to the end of the parameter list
  if(parameters.empty() == false)
  {
    if(param.empty() == false)
      param += _T(" ");

    param += parameters;
  }

  CString strParameters;
  int nShowCmd = SW_SHOWNORMAL;
  
  param = ExpandStringTags(param);
  STLString programPath = ExpandStringTags(pItem->programPath);
  STLString startFolder = ExpandStringTags(pItem->startFolder);

  ShellExecute(0, strOperation, programPath.c_str(), param.c_str(), startFolder.c_str(), nShowCmd);
}

void MZLaunchBar::OnAddSeparator(MZLaunchItem* pAfterItem)
{
  std::shared_ptr<MZLaunchItem> pItem = CreateSeparatorItem();

  if(pAfterItem == nullptr)
    InsertFirst(m_vItems, pItem);
  else
    InsertAfter(m_vItems, pAfterItem, pItem);

  OnChanged(pItem.get(), Change_AddedSep);
  m_bDrawStateDirty = true;
  Invalidate();
}

void MZLaunchBar::OnRemoveItem(MZLaunchItem* pItem)
{
  if(ConfirmRemoveItem(pItem))
    RemoveItem(pItem);
}

bool MZLaunchBar::ConfirmRemoveItem(MZLaunchItem* pItem)
{
  CString msg;
  if(pItem->bSeparator)
  {
    msg.Format(_T("Remove separator ?"), pItem->programPath.c_str());
  }
  else
  {
    msg.Format(_T("Remove : \"%s\""), pItem->programPath.c_str());
  }
  
  if(MessageBox(msg, _T("Remove item ?"), MB_YESNOCANCEL|MB_ICONQUESTION) == IDYES)
    return true;

  return false;
}
void MZLaunchBar::OnCustomizeItem(MZLaunchItem* pItem)
{
#ifdef MZLUNCHBAR_USE_LANUCHPROPDLG
  // Show Configuration dialog for item pItem
  if(!pItem)
    return;

  LaunchPropDlg dlg(pItem, this);

  if(dlg.DoModal() == IDOK)
  {
    DWORD dwChanges = dlg.GetChanges();
    if(dwChanges & LIPCH_ALTICONFILE || dwChanges & LIPCH_ALTICONIDX || dwChanges & LIPCH_PROGRAM) 
      RefreshIcons();

    if(dwChanges != 0)
      Save();
  }
#else
  UNREFERENCED_PARAMETER(pItem);
#endif
}

bool MZLaunchBar::AcceptDroppedItem(const STLString& /*filepath*/)
{
  return true;
}

void MZLaunchBar::EnterDragMode(MZLaunchItem* pItem, const CPoint& pt)
{
  if(pItem == nullptr)
    pItem = GetItemAtPos(pt);

  if(pItem && pItem->bSeparator)
    return; // can not drag separator item

  if(m_dwStyle & MZBS_DRAGREARRANGE && pItem)
  {
    if(m_bAllowRearrange == false)
      return;

    TRACE(_T("Prepare Dragging\n"));
    m_bPrepareDragging = true;
    SetCapture();
    m_ptDragPos = pt;
  }
}


bool MZLaunchBar::RegisterDropTarget(bool bEnable)
{
  if(bEnable)
  {
    return m_DropTarget.Register();
  }
  else
  {
    m_DropTarget.Unregister();
    return true;
  }
}

DROPEFFECT	MZLaunchBar::OnDragEnter(CWnd* /*pWnd*/, COleDataObject* pDataObject, DWORD /*dwKeyState*/, CPoint point)
{
  TRACE("(%p) - OnDragEnter, point %d.%d\n", this, point.x, point.y);

  if( pDataObject->IsDataAvailable(CF_HDROP) )
  {
    if(m_dwStyle & MZBS_DROPADD || m_dwStyle & MZBS_LAUNCH_DROPASPARM)
    {
      m_bDropping = true;
      return DROPEFFECT_COPY;
    }
  }

  return DROPEFFECT_NONE;
}

void		MZLaunchBar::OnDragLeave(CWnd* /*pWnd*/)
{
  TRACE("(%p) - OnDragLeave\n", this);

   m_pOverItem = nullptr;
   m_bDropping = false;
}

DROPEFFECT MZLaunchBar::GetDropEffect(MZLaunchItem* pOverItem)
{
  if(pOverItem)
  {
    if(pOverItem->bSeparator)
      return DROPEFFECT_NONE;

    if(m_dwStyle & MZBS_LAUNCH_DROPASPARM)
      return DROPEFFECT_MOVE;
  }
  else
  {
    if(m_dwStyle & MZBS_DROPADD && m_bAllowDropAdd)
        return DROPEFFECT_COPY;
  }
  return DROPEFFECT_NONE;
}

DROPEFFECT	MZLaunchBar::OnDragOver(CWnd* /*pWnd*/, COleDataObject* pDataObject, DWORD /*dwKeyState*/, CPoint point)
{
  MZLaunchItem* pItem = GetItemAtPos(point);

  if( pDataObject->IsDataAvailable(CF_HDROP) )
  {
    if(m_ptDragPos != point)
    {
      m_ptDragPos = point;
      Invalidate();
    }

    if(m_pOverItem != pItem)
      m_pOverItem = pItem;
    
    return GetDropEffect(m_pOverItem);
  }

  return DROPEFFECT_NONE;
}

namespace
{
  int  GetFilePathFromHDROP(std::vector<STLString>& vPaths, HDROP hDrop)
  {
    TCHAR szNextFile[_MAX_PATH];

    vPaths.clear();
    UINT uNumFiles = DragQueryFile( hDrop , (UINT)-1 , NULL , 0 );
    for( UINT uFile = 0; uFile < uNumFiles; ++uFile )
    {
      // Get the next filename from the HDROP info.
      if( DragQueryFile( hDrop, uFile, szNextFile, MAX_PATH ) > 0 )
      {
        STLString path = szNextFile;
        vPaths.push_back(path);
      }

    }
    return (int)vPaths.size();
  }
}
BOOL MZLaunchBar::OnDrop(CWnd* /*pWnd*/, COleDataObject* pDataObject, DROPEFFECT dropEffect, CPoint point)
{
  TRACE("(%p) - OnDrop - dropeffect : %d, point %d.%d\n", this, dropEffect, point.x, point.y);

  m_bDropping = false;

  if(pDataObject->IsDataAvailable(CF_HDROP) == FALSE)
    return FALSE;

  std::vector<STLString> vFilePaths; 
  BOOL bDropped = FALSE;
  HGLOBAL hGlobal = pDataObject->GetGlobalData(CF_HDROP);
  if(hGlobal)
  {
    HDROP hDrop = (HDROP)hGlobal;
    GetFilePathFromHDROP(vFilePaths, hDrop);
    DragFinish( hDrop );
  }

  if(bDropped)
    GlobalFree(hGlobal); // Really ?

  if(vFilePaths.empty() == false)
  {
    MZLaunchItem* pRefItem = m_pOverItem; 
    bool bInsert = false;

    if(pRefItem == nullptr)
    {
      bInsert = true;
      pRefItem = GetItemAfterPos(m_ptDragPos);
    }

    OnDropFiles(vFilePaths, pRefItem, bInsert);
    m_pOverItem = nullptr;
    return TRUE;
  }

  return bDropped;
}

DROPEFFECT MZLaunchBar::OnDropEx(CWnd* /*pWnd*/, COleDataObject* /*pDataObject*/, DROPEFFECT dropDefault, DROPEFFECT /*dropList*/, CPoint point)
{
  TRACE("(%p) - OnDropEx - dropeffect : %d, point %d.%d\n", this, dropDefault, point.x, point.y);
  return (DROPEFFECT)-1;
}

BOOL MZLaunchBar::OnToolTipNotify(UINT /*id*/, NMHDR* pNMHDR, LRESULT* /*pResult*/)
{ 
  TOOLTIPTEXT *pTTT = (TOOLTIPTEXT *)pNMHDR; 
  UINT_PTR nID = pNMHDR->idFrom;

  if(nID > 0)
    --nID;

  if(nID >= m_vItems.size())
    return FALSE;

  auto p = m_vItems[nID];
  m_CurrentTooltip = p->tooltip;

  pTTT->lpszText = (LPWSTR)m_CurrentTooltip.c_str();
  return TRUE;
}

void MZLaunchBar::OnChanged(MZLaunchItem* /*pItem*/, MZLIChange /*change*/)
{
  if(m_bSaveOnChange)
    Save();
}

STLString MZLaunchBar::ExpandStringTags(const STLString& str)
{
  TCHAR szTmp[16*1024];
  if(ExpandEnvironmentStrings(str.c_str(), szTmp, _countof(szTmp)) == 0)
    return str;

  return szTmp;
}

bool  MZLaunchBar::RemoveItem(std::vector<std::shared_ptr<MZLaunchItem>>& vItems, MZLaunchItem* pItem)
{
  for(auto it = begin(vItems); it != end(vItems); ++it)
  {
    if((*it).get() == pItem)
    {
      vItems.erase(it);
      return true;
    }
  }
  return true;
}

bool  MZLaunchBar::InsertAfter(std::vector<std::shared_ptr<MZLaunchItem>>& vItems, MZLaunchItem* pItem, std::shared_ptr<MZLaunchItem>& pInsertItem)
{
  for(auto it = begin(vItems); it != end(vItems); ++it)
  {
    if((*it).get() == pItem)
    {
      vItems.insert(it+1, pInsertItem);
      return true;
    }
  }
  return true;
}
bool  MZLaunchBar::InsertBefore(std::vector<std::shared_ptr<MZLaunchItem>>& vItems, MZLaunchItem* pItem, std::shared_ptr<MZLaunchItem>& pInsertItem)
{
  for(auto it = begin(vItems); it != end(vItems); ++it)
  {
    if((*it).get() == pItem)
    {
      vItems.insert(it, pInsertItem);
      return true;
    }
  }
  return true;
}

bool  MZLaunchBar::InsertFirst(std::vector<std::shared_ptr<MZLaunchItem>>& vItems, std::shared_ptr<MZLaunchItem>& pInsertItem)
{
  vItems.insert(begin(vItems), pInsertItem);
  return true;
}

BOOL MZLaunchBar::OnSetCursor(CWnd* pWnd, UINT nHitTest, UINT message)
{

  if( nHitTest == HTCLIENT )
  {
    if(m_hCursor == 0)
      m_hCursor = ::LoadCursor(NULL, MAKEINTRESOURCE(IDC_ARROW) );

    if(m_hCursor)
      ::SetCursor( m_hCursor );

    return TRUE;
  }

  return CWnd::OnSetCursor(pWnd, nHitTest, message);
}
