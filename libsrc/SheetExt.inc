!* SHEET control extender
!* mikeduglas@yandex.ru
!* 2022

  INCLUDE('winapi.inc'), ONCE

typCustomTabButtonInfo        GROUP, TYPE
ButtonText                      STRING(6)
ExtraSpaces                     UNSIGNED    !- number of extra spaces in TAB text
FontName                        STRING(31)
FontSize                        REAL        !- > 1: font size; < 1: scale factor relative to SHEET{Prop:FontSize}
FontStyle                       UNSIGNED
FontCharset                     LONG
FontColor                       LONG
BackColor                       LONG
HoverFontColor                  LONG
HoverBackColor                  LONG
                              END

TSheetExtBase                 CLASS(TWnd), TYPE, MODULE('sheetext.clw'), LINK('sheetext.clw')
bComCtl32Enabled                BOOL, PROTECTED
btnInfo                         LIKE(typCustomTabButtonInfo), PROTECTED
bPainting                       BOOL, PROTECTED   !- WM_PAINT started

Construct                       PROCEDURE()
Destruct                        PROCEDURE(), DERIVED

Init                            PROCEDURE(SIGNED pFeq), DERIVED
SetCustomButtonInfo             PROCEDURE(<STRING pText>, <UNSIGNED pExtraSpaces>, | 
                                  <STRING pFontName>, <REAL pFontSize>, <UNSIGNED pFontStyle>, <LONG pFontCharset>, | 
                                  <LONG pFontColor>, <LONG pBackColor>, <LONG pHoverFontColor>, <LONG pHoverBackColor>)
SetCustomButtonInfo             PROCEDURE(typCustomTabButtonInfo pInfo)
SetFont                         PROCEDURE(<STRING pFontName>, <REAL pFontSize>, <UNSIGNED pFontStyle>, <LONG pFontCharset>)
SetColors                       PROCEDURE(<LONG pFontColor>, <LONG pBackColor>, <LONG pHoverFontColor>, <LONG pHoverBackColor>)
SetCustomButtonAsClose          PROCEDURE()
SetCustomButtonAsDropDown       PROCEDURE()
SetCustomButtonAsHelp           PROCEDURE()

PrepareTab                      PROCEDURE(SIGNED pTabFeq), VIRTUAL  !- call it for dynamically created TAB
GetOriginalTabText              PROCEDURE(SIGNED pTabFeq), STRING   !- returns original TAB text

ClearTabProps                   PROCEDURE(), PROTECTED
NotifyTabGetText                PROCEDURE(TWnd pTabCtrl), PROTECTED

GetTabRect                      PROCEDURE(SIGNED pTabFeq, *TRect pRect), PROTECTED
GetTabXRect                     PROCEDURE(SIGNED pTabFeq, *TRect pRect), PROTECTED
GetTabByXPoint                  PROCEDURE(POINT pt), SIGNED, PROTECTED

OnPaint                         PROCEDURE(), PROTECTED, VIRTUAL
OnLButtonDown                   PROCEDURE(UNSIGNED wParam, LONG lParam), BOOL, VIRTUAL, PROTECTED
OnMouseMove                     PROCEDURE(UNSIGNED wParam, LONG lParam), VIRTUAL, PROTECTED

OnDrawCustomButton              PROCEDURE(TDC pDc, SIGNED pTabFeq, TRect pTabRect), PROTECTED, VIRTUAL
OnCustomButtonPressed           PROCEDURE(SIGNED pTabFeq), BOOL, PROTECTED, VIRTUAL

                              END
