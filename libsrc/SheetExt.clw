!* SHEET control extender
!* mikeduglas@yandex.ru
!* 2022

                              MEMBER

  INCLUDE('sheetext.inc'), ONCE

  MAP
    MODULE('win api')
      winapi::GetSysColor(LONG nIndex),UNSIGNED,PASCAL,PROC,NAME('GetSysColor')
    END

    sht_SubclassProc(HWND hWnd, ULONG wMsg, UNSIGNED wParam, LONG lParam, ULONG subclassId, UNSIGNED dwRefData), LONG, PASCAL, PRIVATE
    tab_SubclassProc(HWND hWnd, ULONG wMsg, UNSIGNED wParam, LONG lParam, ULONG subclassId, UNSIGNED dwRefData), LONG, PASCAL, PRIVATE

    LOWORD(LONG pLongVal), LONG, PRIVATE
    HIWORD(LONG pLongVal), LONG, PRIVATE
    GET_X_LPARAM(LONG pLongVal), SHORT, PRIVATE
    GET_Y_LPARAM(LONG pLongVal), SHORT, PRIVATE
  
    INCLUDE('printf.inc'), ONCE
  END

!!!region comctl32.dll
COMCTL32_DLL                  CSTRING('comctl32.dll'), STATIC
DLLGETVERSION_NAME            CSTRING('DllGetVersion'), STATIC
MINCOMCTL32_MAJORVERSION      EQUATE(6)

typComCtl32DllVersionInfo     GROUP, TYPE
cbSize                          ULONG
dwMajorVersion                  ULONG
dwMinorVersion                  ULONG
dwBuildNumber                   ULONG
dwPlatformID                    ULONG
                              END

paComCtl32DllGetVersion       LONG, NAME('fptr_ComCtl32DllGetVersion')
szComCtl32DllGetVersion       CSTRING('DllGetVersion'), STATIC

  MAP
    MODULE('win api')
      winapi::GetModuleHandle(*CSTRING szModuleName),HMODULE,PASCAL,RAW,NAME('GetModuleHandleA')
      winapi::LoadLibrary(*CSTRING szLibFileName),HINSTANCE,PASCAL,RAW,NAME('LoadLibraryA')
      winapi::FreeLibrary(LONG hModule),BOOL,PASCAL,PROC,NAME('FreeLibrary')
      winapi::GetProcAddress(LONG hModule, *CSTRING szProcName),LONG,PASCAL,RAW,NAME('GetProcAddress')
      winapi::GetProcAddress(LONG hModule, LONG pOrdinalValue),LONG,PASCAL,RAW,NAME('GetProcAddress')
      winapi::GetLastError(),LONG,PASCAL,NAME('GetLastError')
    END
    MODULE('comctl32 api')
      comctl32::DllGetVersion(LONG pDvi),HRESULT,PASCAL,NAME('fptr_ComCtl32DllGetVersion'),DLL
    END

    !- check for ComCtl32.dll v6.x.x.x is used (manifested exe)
    IsComCtl32v6Enabled(), BOOL, PRIVATE
  END
!!!endregion

!!!region Unicode
UTF16::BOMLE                  STRING('<0FFh,0FEh>')
UTF16::BOMBE                  STRING('<0FEh,0FFh>')
UTF16::Char0                  STRING('<0h,0h>')
UTF8::BOM                     STRING('<0EFh,0BBh,0BFh>')
!!!endregion

!!!region Custom colors
COLOR:WINDOWGRAY              EQUATE(0F0F0F0H)    !- default TAB background 
COLOR:GainsboroE3             EQUATE(00E3E3E3h)   !- default TAB background if TabStyle:BlackAndWhite and TabStyle:Colored
!!!endregion

!!!region TAB's properties
_TAB_IsVisible_               EQUATE('_TAB_IsVisible_')
_TAB_Subclassed_              EQUATE('_TAB_Subclassed_')
_TAB_left                     EQUATE('_TAB_left')
_TAB_top                      EQUATE('_TAB_top')
_TAB_right                    EQUATE('_TAB_right')
_TAB_bottom                   EQUATE('_TAB_bottom')
_X_exists                     EQUATE('_X_exists')
_X_hovered                    EQUATE('_X_hovered')
_X_left                       EQUATE('_X_left')
_X_top                        EQUATE('_X_top')
_X_right                      EQUATE('_X_right')
_X_bottom                     EQUATE('_X_bottom')
!!!endregion

!!!region TAB values found empirically
TAB::PadWidth                 EQUATE(6)
TAB::BorderWidth              EQUATE(4)
TAB::StylePad                 EQUATE(4)
TAB::SelVertOffest            EQUATE(-2)
!!!endregion

!!!region Tab info
typTabInfo                    GROUP, TYPE
text                            STRING(256)
icon                            STRING(256)
fontName                        STRING(31)
fontSize                        LONG
fontColor                       LONG
fontStyle                       LONG
fontCharset                     LONG
                              END
typTabInfos                   QUEUE(typTabInfo), TYPE
feq                             SIGNED
                              END

typTabChildren                QUEUE, TYPE
feq                             SIGNED
                              END
!!!endregion

!!!region Callbacks
sht_SubclassProc              PROCEDURE(HWND hWnd, ULONG wMsg, UNSIGNED wParam, LONG lParam, ULONG subclassId, UNSIGNED dwRefData)
win                             TWnd
ctrl                            &TSheetExtBase
  CODE
  win.SetHandle(hWnd)
  !- get TSheetExtBase instance
  ctrl &= (dwRefData)
  IF ctrl &= NULL
    !- not our window
    RETURN win.DefSubclassProc(wMsg, wParam, lParam)
  END

  CASE wMsg
  OF WM_MOUSEMOVE
    ctrl.OnMouseMove(wParam, lParam)
    
  OF WM_LBUTTONDOWN
    IF ctrl.OnLButtonDown(wParam, lParam) = FALSE
      RETURN FALSE
    END
    
  OF WM_RBUTTONUP
    IF ctrl.OnRButtonUp(wParam, lParam) = FALSE
      RETURN FALSE
    END
    
  OF WM_PAINT
    ctrl.OnPaint()
    RETURN FALSE
  END
  
  !- call original window proc
  RETURN ctrl.DefSubclassProc(wMsg, wParam, lParam)

tab_SubclassProc              PROCEDURE(HWND hWnd, ULONG wMsg, UNSIGNED wParam, LONG lParam, ULONG subclassId, UNSIGNED dwRefData)
win                             TWnd  !- TAB control
ctrl                            &TSheetExtBase
  CODE
  win.SetHandle(hWnd)
  !- get TWnd instance
  ctrl &= (dwRefData)
  IF ctrl &= NULL
    !- not our window
    RETURN win.DefSubclassProc(wMsg, wParam, lParam)
  END

  CASE wMsg
  OF WM_GETTEXT
    !- notify our SHEET that TAB received WM_GETTEXT
    ctrl.NotifyTabGetText(win)
  END
  
  !- call original window proc
  RETURN win.DefSubclassProc(wMsg, wParam, lParam)
!!!endregion
  
!!!region Helpers
IsComCtl32v6Enabled           PROCEDURE()
hComCtlDll                      HMODULE, AUTO
dvi                             LIKE(typComCtl32DllVersionInfo)
hr                              HRESULT, AUTO
bRet                            BOOL(FALSE)
  CODE
  hComCtlDll = winapi::LoadLibrary(COMCTL32_DLL)
  IF hComCtlDll
    paComCtl32DllGetVersion = winapi::GetProcAddress(hComCtlDll, szComCtl32DllGetVersion)
    IF paComCtl32DllGetVersion
      dvi.cbSize = SIZE(dvi)
      hr = comctl32::DllGetVersion(ADDRESS(dvi))
      IF hr >= 0  !- success
!        printd('%s v%i.%i.%i.%i', COMCTL32_DLL, dvi.dwMajorVersion, dvi.dwMinorVersion, dvi.dwBuildNumber, dvi.dwPlatformID)
        IF dvi.dwMajorVersion >= MINCOMCTL32_MAJORVERSION
          !- ComCtl32.dll version at least 6.x.x.x
          bRet = TRUE
        END
      END
    END
    winapi::FreeLibrary(hComCtlDll)
  END
  RETURN bRet
  
LOWORD                        PROCEDURE(LONG pLongVal)
  CODE
  RETURN BAND(pLongVal, 0FFFFh)

HIWORD                        PROCEDURE(LONG pLongVal)
  CODE
  RETURN BSHIFT(BAND(pLongVal, 0FFFF0000h), -16)

GET_X_LPARAM                  PROCEDURE(LONG pLongVal)
  CODE
  RETURN LOWORD(pLongVal)

GET_Y_LPARAM                  PROCEDURE(LONG pLongVal)
  CODE
  RETURN HIWORD(pLongVal)
!!!endregion
  
!!!region TSheetExtBase
TSheetExtBase.Construct       PROCEDURE()
  CODE
  SELF.bComCtl32Enabled = IsComCtl32v6Enabled()
  
TSheetExtBase.Destruct        PROCEDURE()
  CODE
  
TSheetExtBase.Init            PROCEDURE(SIGNED pFeq)
i                               LONG, AUTO
tabFeq                          SIGNED, AUTO
  CODE
  ASSERT(pFeq{PROP:Type} = CREATE:sheet)
  IF pFeq{PROP:Type} <> CREATE:sheet
    RETURN
  END
  
  PARENT.Init(pFeq)
  
  !- remoce unsupported SHEET attributes
  SELF.FEQ{PROP:Spread} = FALSE
  SELF.FEQ{PROP:Above} = TRUE
!  SELF.FEQ{PROP:Join} = TRUE
  COMPILE('New sheet properties', _C80_)
  SELF.FEQ{PROP:Up} = FALSE
  SELF.FEQ{PROP:Down} = FALSE
  !'New sheet properties'
  
  !- set default "X" button properties
  SELF.SetCustomButtonAsClose()
    
  !- append extra space on the right side
  LOOP i=1 TO SELF.FEQ{PROP:NumTabs}
    tabFeq = SELF.FEQ{PROP:Child, i}
    SELF.PrepareTab(tabFeq)
  END

  !- overwrite default subclass proc
  SELF.SetWindowSubclass(ADDRESS(sht_SubclassProc), 0, ADDRESS(SELF))

TSheetExtBase.SetCustomButtonInfo PROCEDURE(<STRING pText>, <UNSIGNED pExtraSpaces>, | 
                                    <STRING pFontName>, <REAL pFontSize>, <UNSIGNED pFontStyle>, <LONG pFontCharset>, | 
                                    <LONG pFontColor>, <LONG pBackColor>, <LONG pHoverFontColor>, <LONG pHoverBackColor>, |
                                    <BOOL pIsUnicode>)
  CODE
  IF pText
    SELF.btnInfo.ButtonText = pText
  END
  IF NOT OMITTED(pExtraSpaces)
    SELF.btnInfo.ExtraSpaces = pExtraSpaces
  END
  IF pFontName
    SELF.btnInfo.FontName = pFontName
  END
  IF NOT OMITTED(pFontSize)
    SELF.btnInfo.FontSize = pFontSize
  END
  IF NOT OMITTED(pFontStyle)
    SELF.btnInfo.FontStyle = pFontStyle
  END
  IF NOT OMITTED(pFontCharset)
    SELF.btnInfo.FontCharset = pFontCharset
  END
  IF NOT OMITTED(pFontColor)
    SELF.btnInfo.FontColor = pFontColor
  END
  IF NOT OMITTED(pBackColor)
    SELF.btnInfo.BackColor = pBackColor
  END
  IF NOT OMITTED(pHoverFontColor)
    SELF.btnInfo.HoverFontColor = pHoverFontColor
  END
  IF NOT OMITTED(pHoverBackColor)
    SELF.btnInfo.HoverBackColor = pHoverBackColor
  END
  IF NOT OMITTED(pIsUnicode)
    SELF.btnInfo.IsUnicode = pIsUnicode
  END
  
TSheetExtBase.SetCustomButtonInfo PROCEDURE(typCustomTabButtonInfo pInfo)
  CODE
  SELF.btnInfo :=: pInfo
  
TSheetExtBase.SetCustomButtonText PROCEDURE(<STRING pText>, <BOOL pIsUnicode>)
  CODE
  IF pText
    SELF.btnInfo.ButtonText = pText
  END
  IF NOT OMITTED(pIsUnicode)
    SELF.btnInfo.IsUnicode = pIsUnicode
    IF pIsUnicode
      !- add UTF16::BOMLE
      IF SUB(SELF.btnInfo.ButtonText, 1, 2) <> UTF16::BOMLE
        SELF.btnInfo.ButtonText = UTF16::BOMLE & SELF.btnInfo.ButtonText
      END
    END
  END

TSheetExtBase.SetCustomButtonFont PROCEDURE(<STRING pFontName>, <REAL pFontSize>, <UNSIGNED pFontStyle>, <LONG pFontCharset>)
  CODE
  IF pFontName
    SELF.btnInfo.FontName = pFontName
  END
  IF NOT OMITTED(pFontSize)
    SELF.btnInfo.FontSize = pFontSize
  END
  IF NOT OMITTED(pFontStyle)
    SELF.btnInfo.FontStyle = pFontStyle
  END
  IF NOT OMITTED(pFontCharset)
    SELF.btnInfo.FontCharset = pFontCharset
  END

TSheetExtBase.SetCustomButtonColors   PROCEDURE(<LONG pFontColor>, <LONG pBackColor>, <LONG pHoverFontColor>, <LONG pHoverBackColor>)
  CODE
  IF NOT OMITTED(pFontColor)
    SELF.btnInfo.FontColor = pFontColor
  END
  IF NOT OMITTED(pBackColor)
    SELF.btnInfo.BackColor = pBackColor
  END
  IF NOT OMITTED(pHoverFontColor)
    SELF.btnInfo.HoverFontColor = pHoverFontColor
  END
  IF NOT OMITTED(pHoverBackColor)
    SELF.btnInfo.HoverBackColor = pHoverBackColor
  END

TSheetExtBase.SetCustomButtonAsClose  PROCEDURE()
  CODE
  SELF.SetCustomButtonInfo('r', 6, 'Webdings', 5/6, FONT:regular, CHARSET:SYMBOL, COLOR:NONE, COLOR:NONE, COLOR:White, COLOR:Gray)

TSheetExtBase.SetCustomButtonAsDropDown   PROCEDURE()
  CODE
  SELF.SetCustomButtonInfo('6', 6, 'Webdings', 5/6, FONT:regular, CHARSET:SYMBOL, COLOR:NONE, COLOR:NONE, COLOR:White, COLOR:Gray)

TSheetExtBase.SetCustomButtonAsHelp   PROCEDURE()
  CODE
  SELF.SetCustomButtonInfo('s', 6, 'Webdings', 5/6, FONT:regular, CHARSET:SYMBOL, COLOR:NONE, COLOR:NONE, COLOR:White, COLOR:Gray)

TSheetExtBase.PrepareTab      PROCEDURE(SIGNED pTabFeq)
tabCtrl                         TWnd
  CODE
  ASSERT(pTabFeq{PROP:Type} = CREATE:tab)
  IF pTabFeq{PROP:Type} <> CREATE:tab
    RETURN
  END
  
  !- append extra space
  pTabFeq{PROP:Text} = pTabFeq{PROP:Text} & ALL(' ', SELF.btnInfo.ExtraSpaces) &'<0>'
  
  !- subclass TAB control to catch WM_GETTEXT messages
  tabCtrl.Init(pTabFeq)
  IF NOT tabCtrl.GetPropA(_TAB_Subclassed_)
    tabCtrl.SetWindowSubclass(ADDRESS(tab_SubclassProc), 0, ADDRESS(SELF))
    tabCtrl.SetPropA(_TAB_Subclassed_, TRUE)
  END
  
TSheetExtBase.GetOriginalTabText  PROCEDURE(SIGNED pTabFeq)
  CODE
  RETURN CLIP(SUB(pTabFeq{PROP:Text}, 1, LEN(pTabFeq{PROP:Text})-LEN(CLIP(SELF.btnInfo.ButtonText))))
  
TSheetExtBase.ClearTabProps   PROCEDURE()
tabFeq                          SIGNED, AUTO
tabCtrl                         TWnd
i                               LONG, AUTO
  CODE
  LOOP i=1 TO SELF.FEQ{PROP:NumTabs}
    tabFeq = SELF.FEQ{PROP:Child, i}
    tabCtrl.Init(tabFeq)
    
    tabCtrl.SetPropA(_TAB_IsVisible_, 0)

    tabCtrl.SetPropA(_X_exists, FALSE)
    tabCtrl.SetPropA(_X_hovered, FALSE)
    tabCtrl.SetPropA(_X_left, 0)
    tabCtrl.SetPropA(_X_top, 0)
    tabCtrl.SetPropA(_X_right, -1)
    tabCtrl.SetPropA(_X_bottom, -1)
  END
  
TSheetExtBase.NotifyTabGetText    PROCEDURE(TWnd pTabCtrl)
  CODE
  !- in WM_PAINT sheet control sends WM_GETTEXT to its VISIBLE tabs (really visible, not just Prop:Hide=FALSE).
  IF SELF.bPainting
    !- set _TAB_IsVisible_ prop
    pTabCtrl.SetPropA(_TAB_IsVisible_, 1)
  END

TSheetExtBase.GetTabRect      PROCEDURE(SIGNED pTabFeq, *TRect rc)
tabCtrl                         TWnd
  CODE
  tabCtrl.Init(pTabFeq)
  rc.left = tabCtrl.GetPropA(_TAB_left)
  rc.top = tabCtrl.GetPropA(_TAB_top)
  rc.right = tabCtrl.GetPropA(_TAB_right)
  rc.bottom = tabCtrl.GetPropA(_TAB_bottom)

TSheetExtBase.SaveTabRect     PROCEDURE(SIGNED pTabFeq, *TRect rc)
tabCtrl                         TWnd
  CODE
  tabCtrl.Init(pTabFeq)
  tabCtrl.SetPropA(_TAB_left, rc.left)
  tabCtrl.SetPropA(_TAB_top, rc.top)
  tabCtrl.SetPropA(_TAB_right, rc.right)
  tabCtrl.SetPropA(_TAB_bottom, rc.bottom)

TSheetExtBase.GetTabXRect     PROCEDURE(SIGNED pTabFeq, *TRect rc)
tabCtrl                         TWnd
  CODE
  tabCtrl.Init(pTabFeq)
  IF tabCtrl.GetPropA(_X_exists)
    rc.left = tabCtrl.GetPropA(_X_left)
    rc.top = tabCtrl.GetPropA(_X_top)
    rc.right = tabCtrl.GetPropA(_X_right)
    rc.bottom = tabCtrl.GetPropA(_X_bottom)
  ELSE
    rc.Assign(0, 0, -1, -1)
  END
  
TSheetExtBase.GetTabByXPoint  PROCEDURE(POINT pt)
tabFeq                          SIGNED, AUTO
rcX                             TRect
i                               LONG, AUTO
  CODE
  LOOP i=1 TO SELF.FEQ{PROP:NumTabs}
    tabFeq = SELF.FEQ{PROP:Child, i}
    SELF.GetTabXRect(tabFeq, rcX)
    IF rcX.PtInRect(pt)
      RETURN tabFeq
    END
  END
  RETURN 0
  
TSheetExtBase.GetTabByPoint   PROCEDURE(POINT pt)
tabFeq                          SIGNED, AUTO
rc                              TRect
i                               LONG, AUTO
  CODE
  LOOP i=1 TO SELF.FEQ{PROP:NumTabs}
    tabFeq = SELF.FEQ{PROP:Child, i}
    SELF.GetTabRect(tabFeq, rc)
    IF rc.PtInRect(pt)
      RETURN tabFeq
    END
  END
  RETURN 0

TSheetExtBase.IsVisible       PROCEDURE(SIGNED pTabFeq)
tabCtrl                         TWnd
  CODE
  tabCtrl.Init(pTabFeq)
  RETURN tabCtrl.GetPropA(_TAB_IsVisible_)

TSheetExtBase.OnPaint         PROCEDURE()
dc                              TDC
res                             LONG, AUTO
i                               LONG, AUTO
tabFeq                          SIGNED, AUTO
iconWidth                       LONG, AUTO
totalWidth                      LONG, AUTO
textWidth                       LONG, AUTO
fntTab                          TLogicalFont  !- font of a TAB
dtFormat                        LONG(DT_LEFT+DT_SINGLELINE)
rcSheet                         TRect
rcText                          TRect
rcTab                           TRect
rcClose                         TRect
textSize                        LIKE(SIZE)      !- size of a text
cWidth                          LONG, AUTO      !- width of 1 ascii char
  CODE
  !- we are in OnPaint()
  !- start listen for TAB's WM_GETTEXT
  SELF.bPainting = TRUE
  
  !- clear TAB visible states
  SELF.ClearTabProps()
  
  !- Call SHEET's default proc, then custom OnPaint.
  SELF.DefSubclassProc(WM_PAINT, 0, 0)

  !- stop listen for TAB's WM_GETTEXT
  SELF.bPainting = FALSE

  dc.GetDC(SELF)
  SELF.GetClientRect(rcSheet)

  totalWidth = 0

  !- Calc each tab's rect.
  LOOP i=1 TO SELF.FEQ{PROP:NumTabs}
    tabFeq = SELF.FEQ{PROP:Child, i}
    
    IF tabFeq{PROP:Hide}
      !- skip hidden (Prop:Hide=true) tabs
      CYCLE
    END
    
!    printd('Tab %s: Visible=%b', SELF.GetOriginalTabText(tabFeq), SELF.IsVisible(tabFeq))
    IF NOT SELF.IsVisible(tabFeq)
      !- skip hidden (by scrolling) tabs
      CYCLE
    END

    iconWidth = CHOOSE(tabFeq{PROP:Icon} <> '', 18, 0)
        
    !- tab font
    fntTab.CreateFont(dc, tabFeq{PROP:FontName}, tabFeq{PROP:FontSize}, tabFeq{PROP:FontStyle}, tabFeq{PROP:FontCharset})
    fntTab.SelectObject(dc)

    !- get text rect as (0, 0, w, h)
    rcText.Assign(rcSheet)
    dc.DrawText(tabFeq{PROP:Text}, rcText, dtFormat+DT_CALCRECT, TRUE)

    !- add extra space for padding and icon
    rcText.OffsetRect(TAB::PadWidth + iconWidth, 4)
    
    !- set text rect as (x, y, w, h)
    rcText.OffsetRect(totalWidth, 0)
    
    COMPILE('New sheet properties', _C80_)
    IF SELF.FEQ{PROP:TabSheetStyle} <> TabStyle:Default
      rcText.OffsetRect(TAB::PadWidth, 0)
    END
    !'New sheet properties'
    
    !- get tab text size
    dc.GetTextExtentPoint32(tabFeq{PROP:Text}, textSize)
    textWidth = textSize.cx
    
    !- TAB rect
    rcTab.Assign(rcText)
    rcTab.left -= iconWidth+2
    rcTab.right += TAB::PadWidth
!    dc.FillSolidRect(rcTab, RANDOM(0, 0ffffffh))
    
    !- save TAB's rect
    SELF.SaveTabRect(tabFeq, rcTab)
    
    fntTab.DeleteObject()

    !- draw custom button
    SELF.OnDrawCustomButton(dc, tabFeq, rcTab)
    
    !- increase total tab rects width
    totalWidth += iconWidth + TAB::PadWidth + textWidth + TAB::BorderWidth
!    COMPILE('New sheet properties', _C80_)
!    IF SELF.FEQ{PROP:TabSheetStyle} <> TabStyle:Default
!      totalWidth += TAB::StylePad
!    END
    !'New sheet properties'
  END
  
  dc.ReleaseDC()

TSheetExtBase.OnDrawCustomButton  PROCEDURE(TDC dc, SIGNED pTabFeq, TRect pTabRect)
bNoTheme                            BOOL, AUTO
bIsSelected                         BOOL, AUTO
tabCtrl                             TWnd
fntButton                           TLogicalFont  !- font of "X" button
nFontSize                           UNSIGNED, AUTO
rc                                  TRect
rcX                                 TRect
textWidth                           LONG, AUTO
dtFormat                            LONG(DT_TOP+DT_RIGHT)
bkColor                             LONG, AUTO
textColor                           LONG, AUTO
  CODE
  COMPILE('New sheet properties', _C80_)
  bNoTheme = CHOOSE(NOT SELF.bComCtl32Enabled OR SELF.FEQ{PROP:NoTheme})
  !'New sheet properties'
  OMIT('New sheet properties', _C80_)
  bNoTheme = CHOOSE(NOT SELF.bComCtl32Enabled)
  !'New sheet properties'
      
  bIsSelected = CHOOSE(pTabFeq = SELF.FEQ{PROP:ChoiceFeq})

  !- calc font size
  IF SELF.btnInfo.FontSize > 1
    nFontSize = SELF.btnInfo.FontSize
  ELSE
    !- SELF.btnInfo.FontSize is a scale factor
    nFontSize = SELF.FEQ{PROP:FontSize} * SELF.btnInfo.FontSize
  END
  
  fntButton.CreateFont(dc, SELF.btnInfo.FontName, nFontSize, SELF.btnInfo.FontStyle, SELF.btnInfo.FontCharset)
  fntButton.SelectObject(dc)
  
  tabCtrl.Init(pTabFeq)

  COMPILE('New sheet properties', _C80_)
  IF SELF.FEQ{PROP:TabSheetStyle} <> TabStyle:Default
    rcX.OffsetRect(-TAB::StylePad, 0)
  END
  
  !- fix diagonal border issue
!  IF SELF.FEQ{PROP:TabSheetStyle} = TabStyle:BlackAndWhite
!    printd('This index %i, selected index %i', pTabFeq{PROP:ChildIndex}, (SELF.FEQ{PROP:ChoiceFeq}){PROP:ChildIndex})
!    IF NOT bIsSelected AND pTabFeq{PROP:ChildIndex} = (SELF.FEQ{PROP:ChoiceFeq}){PROP:ChildIndex} - 1
!      rcX.OffsetRect(-TAB::StylePad*2, 0)
!    END
!  END
  !'New sheet properties'
  
  !- calc background color
  IF tabCtrl.GetPropA(_X_hovered)
    !- hover color
    bkColor = SELF.btnInfo.HoverBackColor
  ELSE
    !- usual bkcolor
    bkColor = SELF.btnInfo.BackColor
    
    !- determine actual color if bkColor is COLOR:NONE
    IF bkColor = COLOR:NONE
      IF bNoTheme
        !- Common controls disabled
        bkColor = pTabFeq{PROP:Color}     !- check TAB color
        IF bkColor = COLOR:NONE
          bkColor = SELF.FEQ{PROP:Color}  !- check SHEET color
          IF bkColor = COLOR:NONE
            bkColor = COLOR:WINDOWGRAY    !- default background
            COMPILE('New sheet properties', _C80_)
            CASE SELF.FEQ{PROP:TabSheetStyle}
            OF TabStyle:BlackAndWhite OROF TabStyle:Colored
              bkColor = CHOOSE(NOT bIsSelected, COLOR:GainsboroE3, COLOR:White)    !- default background
            OF TabStyle:Boxed
              bkColor = CHOOSE(NOT bIsSelected, COLOR:WINDOWGRAY, COLOR:White)     !- default background
            END
            !'New sheet properties'
          END
        END
      ELSE
        !- Common controls enabled
        bkColor = COLOR:White           !- default background
      END
    END
  END
  
  IF bkColor <> COLOR:NONE AND BAND(bkColor, 80000000h) !- system color
    bkColor = winapi::GetSysColor(bkColor)
  END

  !- calc custom button rect
  rc.Assign(pTabRect)
  
  !- leave some space on right side
  rc.right -= TAB::PadWidth
 
  !- get button text rect
  rcX.Assign(rc)
  IF NOT SELF.btnInfo.IsUnicode
    dc.DrawText(SELF.btnInfo.ButtonText, rcX, dtFormat+DT_CALCRECT)
  ELSE
    dc.DrawTextW(CLIP(SELF.btnInfo.ButtonText), rcX, dtFormat+DT_CALCRECT)
  END
  
  textWidth = rcX.Width()
  rcX.left = rc.right - textWidth
  rcX.right = rc.right
  IF SELF.FEQ{PROP:ChoiceFEQ} = pTabFeq
    !- for selected tabs draw custom button higher a little
    rcX.OffsetRect(0, TAB::SelVertOffest)
  END

  !- fill custom button rect
  IF bkColor <> COLOR:NONE
    dc.FillSolidRect(rcX, bkColor)
  END

  !- calc text color
  IF tabCtrl.GetPropA(_X_hovered)
    textColor = SELF.btnInfo.HoverFontColor
  ELSE
    textColor = SELF.btnInfo.FontColor
  END
  IF textColor <> COLOR:NONE
    IF BAND(textColor, 80000000h) !- system color
      textColor = winapi::GetSysColor(textColor)
    END
    dc.SetTextColor(textColor)
  END
  
  !- draw custom button
  dc.SetBkMode(TRANSPARENT)
  IF NOT SELF.btnInfo.IsUnicode
    dc.DrawText(SELF.btnInfo.ButtonText, rcX, dtFormat)
  ELSE
    dc.DrawTextW(CLIP(SELF.btnInfo.ButtonText), rcX, dtFormat)
  END
  
  !- save custom button rect
  tabCtrl.SetPropA(_X_exists, TRUE)
  tabCtrl.SetPropA(_X_left, rcX.left)
  tabCtrl.SetPropA(_X_top, rcX.top)
  tabCtrl.SetPropA(_X_right, rcX.right)
  tabCtrl.SetPropA(_X_bottom, rcX.bottom)
  
  fntButton.DeleteObject()

TSheetExtBase.OnLButtonDown   PROCEDURE(UNSIGNED wParam, LONG lParam)
pt                              LIKE(POINT)
tabFeq                          SIGNED, AUTO
tabCtrl                         TWnd
rc                              TRect
  CODE
  pt.x = GET_X_LPARAM(lParam)
  pt.y = GET_Y_LPARAM(lParam)

  !- get a TAB where pt is inside custom button
  tabFeq = SELF.GetTabByXPoint(pt)
  IF tabFeq
    IF SELF.OnCustomButtonPressed(tabFeq)
      !- redraw
      SELF.RedrawWindow(RDW_INVALIDATE + RDW_UPDATENOW)
      !- don't call default handler DefSubclassProc
      RETURN FALSE
    END
  END
  !- call default handler DefSubclassProc
  RETURN TRUE
  
TSheetExtBase.OnRButtonUp     PROCEDURE(UNSIGNED wParam, LONG lParam)
pt                              LIKE(POINT)
tabFeq                          SIGNED, AUTO
tabCtrl                         TWnd
rc                              TRect
  CODE
  pt.x = GET_X_LPARAM(lParam)
  pt.y = GET_Y_LPARAM(lParam)
  
  !- get a TAB where pt is inside TAB rect
  tabFeq = SELF.GetTabByPoint(pt)
  IF tabFeq
    SELF.ClientToScreen(pt)
    SELF.OnContextMenu(tabFeq, pt)
    RETURN FALSE
  END
  !- call default handler DefSubclassProc
  RETURN TRUE
  
TSheetExtBase.OnMouseMove     PROCEDURE(UNSIGNED wParam, LONG lParam)
pt                              LIKE(POINT)
dc                              TDC
tabFeq                          SIGNED, AUTO
tabCtrl                         TWnd
rc                              TRect
rcX                             TRect
i                               LONG, AUTO
  CODE
  pt.x = GET_X_LPARAM(lParam)
  pt.y = GET_Y_LPARAM(lParam)

  dc.GetDC(SELF)
  
  LOOP i=1 TO SELF.FEQ{PROP:NumTabs}
    tabFeq = SELF.FEQ{PROP:Child, i}
    tabCtrl.Init(tabFeq)
    
    !- get tab rect
    SELF.GetTabRect(tabFeq, rc)
    !- get custom button rect
    SELF.GetTabXRect(tabFeq, rcX)
    
    IF  rcX.PtInRect(pt)
      !- if cursor hovers custom button
      IF NOT tabCtrl.GetPropA(_X_hovered)
        !- redraw custom button
        tabCtrl.SetPropA(_X_hovered, TRUE)
        SELF.OnDrawCustomButton(dc, tabFeq, rc)
      END
    ELSE
      !- if cursor is outside if custom button
      IF tabCtrl.GetPropA(_X_hovered)
        !- redraw custom button
        tabCtrl.SetPropA(_X_hovered, FALSE)
        SELF.OnDrawCustomButton(dc, tabFeq, rc)
      END
    END
  END
  
  dc.ReleaseDC()
  
TSheetExtBase.OnCustomButtonPressed   PROCEDURE(SIGNED pTabFeq)
  CODE
  RETURN TRUE
  
TSheetExtBase.OnContextMenu   PROCEDURE(SIGNED pTabFeq, POINT pPt)
  CODE
!!!endregion