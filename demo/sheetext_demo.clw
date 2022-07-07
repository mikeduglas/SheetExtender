  PROGRAM

!  PRAGMA('link(sheetext_demo.exe.manifest)')

  INCLUDE('sheetext.inc'), ONCE

  MAP
    INCLUDE('printf.inc'), ONCE
  END

Window                        WINDOW('Extended Sheet control'),AT(,,395,359),CENTER,GRAY,SYSTEM, |
                                FONT('Tahoma',9)
                                SHEET,AT(18,11,357,105),USE(?shtBrowse),FONT(,12)
                                  TAB('Customers'),USE(?TAB1_1),ICON(ICON:Copy)
                                    PROMPT('Here is Customer list'),AT(41,59,201),USE(?PROMPT1),TRN,CENTER
                                  END
                                  TAB('Products'),USE(?TAB1_2),ICON(ICON:Cut)
                                    PROMPT('Here is Product list'),AT(48,59,201),USE(?PROMPT2),TRN,CENTER
                                  END
                                  TAB('Orders'),USE(?TAB1_3),ICON(ICON:Paste)
                                    PROMPT('Here is Order list'),AT(41,59,201),USE(?PROMPT3),TRN,CENTER
                                  END
                                END
                                SHEET,AT(19,135,356,106),USE(?shtWeb)
                                  TAB('Google'),USE(?TAB2_1)
                                    ENTRY(@s30),AT(34,182,163),USE(?ENTRY1),READONLY
                                    BUTTON('Search'),AT(206,181,50),USE(?BUTTON1)
                                  END
                                  TAB('StackOverflow'),USE(?TAB2_2)
                                    PROMPT('Q. Nothing works. Any suggestion?'),AT(101,182),USE(?PROMPT6),TRN, |
                                      CENTER
                                  END
                                  TAB('CodeProject'),USE(?TAB2_3)
                                    PROMPT('Yet another awesome Python library!'),AT(101,182),USE(?PROMPT7),TRN, |
                                      CENTER
                                  END
                                  TAB('ClarionHub'),USE(?TAB2_4)
                                    PROMPT('No more unread topics.'),AT(115,182),USE(?PROMPT8),TRN,CENTER
                                  END
                                  TAB('YouTube'),USE(?TAB2_5)
                                    PROMPT('This video has been blocked in your country.'),AT(101,182), |
                                      USE(?PROMPT5),TRN,CENTER
                                  END
                                END
                                SHEET,AT(19,259,354,76),USE(?shtHelp),FONT(,12)
                                  TAB('Customer options'),USE(?TAB3_1)
                                  END
                                  TAB('Product options'),USE(?TAB3_2)
                                  END
                                  TAB('Order options'),USE(?TAB3_3)
                                  END
                                END
                              END

TSheetExtExportable           CLASS(TSheetExtBase), TYPE
OnCustomButtonPressed           PROCEDURE(SIGNED pTabFeq), BOOL, PROC, PROTECTED, DERIVED
                              END
TSheetExtCloseable            CLASS(TSheetExtBase), TYPE
OnCustomButtonPressed           PROCEDURE(SIGNED pTabFeq), BOOL, PROC, PROTECTED, DERIVED
OnContextMenu                   PROCEDURE(SIGNED pTabFeq, POINT pPt), DERIVED, VIRTUAL
                              END
TSheetExtHelp                 CLASS(TSheetExtBase), TYPE
OnCustomButtonPressed           PROCEDURE(SIGNED pTabFeq), BOOL, PROC, PROTECTED, DERIVED
                              END

shCtrl1                       TSheetExtExportable
shCtrl2                       TSheetExtCloseable
shCtrl3                       TSheetExtHelp

  CODE
  OPEN(Window)

  !- set controls
  CHANGE(?ENTRY1, 'Extended SHEET control')
  SELECT(?ENTRY1, 23, 0)
  
!  ?TAB2_1{PROP:Background} = COLOR:Red
  
  
!  ?shtHelp{PROP:TabSheetStyle} = TabStyle:BlackAndWhite
!  ?shtHelp{PROP:TabSheetStyle} = TabStyle:Boxed
!  ?shtHelp{PROP:TabSheetStyle} = TabStyle:Colored
!  ?shtHelp{PROP:TabSheetStyle} = TabStyle:Squared
  
  shCtrl1.Init(?shtBrowse)
  shCtrl2.Init(?shtWeb)
  shCtrl3.Init(?shtHelp)
  
  shCtrl1.SetCustomButtonAsDropDown() !- dropdown arrow
  shCtrl2.SetCustomButtonAsClose()    !- X button
  shCtrl3.SetCustomButtonAsHelp()     !- ? button
  shCtrl3.SetCustomButtonColors(COLOR:Gray,,COLOR:White,00FFBF00h)
  
  !- Unicode test
!  shCtrl3.SetCustomButtonText('<03Dh,0D8h,0,0DEh>', TRUE) !- UTF16 character 3D D8 00 DE (smile)
!  shCtrl3.SetCustomButtonFont('Tahoma', 9, FONT:regular, CHARSET:ANSI)
!  shCtrl3.SetCustomButtonColors(COLOR:Gray,,COLOR:White,00FFBF00h)

  ACCEPT
  END
  
  
TSheetExtExportable.OnCustomButtonPressed PROCEDURE(SIGNED pTabFeq)
sTabText                                    STRING(32), AUTO
  CODE
  !- get original TAB text
  sTabText = SELF.GetOriginalTabText(pTabFeq)
  !- display context menu for this TAB
  EXECUTE POPUP(printf('Export %s to Excel|Export %s to CSV|Print %s', sTabText, sTabText, sTabText))
    MESSAGE(printf('Export %s to Excel completed!', sTabText), 'Success', ICON:Asterisk)
    MESSAGE(printf('Export %s to CSV completed!', sTabText), 'Success', ICON:Asterisk)
    MESSAGE(printf('Print %s completed!', sTabText), 'Success', ICON:Asterisk)
  END
  
  RETURN PARENT.OnCustomButtonPressed(pTabFeq)
  
  
TSheetExtCloseable.OnCustomButtonPressed  PROCEDURE(SIGNED pTabFeq)
shtFeq                                      SIGNED, AUTO
childFeq                                    SIGNED(0)
childIndex                                  UNSIGNED, AUTO
selectedTabFeq                              SIGNED, AUTO
childrenQ                                   QUEUE, PRE(childrenQ)
feq                                           SIGNED
                                            END
i                                           LONG, AUTO
  CODE
  shtFeq = pTabFeq{PROP:Parent}             !- SHEET's FEQ
  childIndex = pTabFeq{PROP:ChildIndex}     !- TAB's index
  selectedTabFeq = shtFeq{PROP:ChoiceFEQ}   !- selected TAB
  
  IF shtFeq{PROP:NumTabs} = 1
    IF MESSAGE('You''re going to delete last tab. Are you sure?', 'Question', ICON:Question, BUTTON:YES+BUTTON:NO) = BUTTON:NO
      !- do not destroy last tab
      RETURN PARENT.OnCustomButtonPressed(pTabFeq)
    END
  END
  
  !- collect all TAB's children
  LOOP
    childFeq = Window{PROP:NextField, childFeq}
    IF childFeq
      IF childFeq{PROP:Parent} = pTabFeq
        childrenQ.feq = childFeq
        ADD(childrenQ)
      END
    ELSE
      BREAK
    END
  END
  
  !- destroy all TAB's children
  LOOP i=1 TO RECORDS(childrenQ)
    GET(childrenQ, i)
    DESTROY(childrenQ.feq)
  END
  
  !- delete TAB
  DESTROY(pTabFeq)
  
  !- if this tab was selected, select next or previous tab
  IF selectedTabFeq = pTabFeq
    IF childIndex > 1
      !- prev tab
      SELECT(shtFeq, childIndex-1)
    ELSE
      SELECT(shtFeq, 1)
    END
  ELSE
      SELECT(shtFeq, selectedTabFeq{PROP:ChildIndex})
  END
  
  RETURN PARENT.OnCustomButtonPressed(pTabFeq)

TSheetExtCloseable.OnContextMenu  PROCEDURE(SIGNED pTabFeq, POINT pPt)
sTabText                            STRING(32), AUTO
  CODE
  !- get original TAB text
  sTabText = SELF.GetOriginalTabText(pTabFeq)
  EXECUTE POPUP(printf('Close "%s" tab', sTabText))
    SELF.OnCustomButtonPressed(pTabFeq)
  END
    
  
TSheetExtHelp.OnCustomButtonPressed   PROCEDURE(SIGNED pTabFeq)
sTabText                                STRING(32), AUTO
  CODE
  !- get original TAB text
  sTabText = SELF.GetOriginalTabText(pTabFeq)
  !- display help for this TAB
  MESSAGE(printf('Help for "%s" tab.', sTabText), 'Help', ICON:Question)
  
  RETURN PARENT.OnCustomButtonPressed(pTabFeq)
  