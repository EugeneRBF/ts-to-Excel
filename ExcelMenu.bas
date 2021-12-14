Attribute VB_Name = "ExcelMenu"
'
' (C) 2021 Eugene Torkhov
'
'--------------------------------------------------------------------------'
'   Menu        Adds our "OLAP menu tree to, or removes from, the Excel    '
'               main menu. The event handling code may be found in module  '
'               MenuEvent.                                                 '
'--------------------------------------------------------------------------'
Option Explicit

Dim bar        As CommandBar

'   The translated menu item strings

Private MenuTopLevel As String
Private HelpMenu   As CommandBarControl

Private MenuParseTSSheet As String
Private MenuSaveAsTSheet As String

Private gbCascadedMenu As Boolean


Public Sub AddPluginMenu()
    MenuTopLevel = "Локализация"         ' have to supply this now,
    AddPluginMenuByString "Парсинг .ts;Сохранение в .ts;"
End Sub


'--------------------------------------------------------------------------'
'  AddOLAPMenu()  Add our OLAP menu tree to the left of the Excel's "Help" '
'                 menu item                                                '
'--------------------------------------------------------------------------'

Private Sub AddPluginMenuByString(str As String)

 Dim Menu       As CommandBarControl
 Dim OurMenu    As CommandBarControl
 Dim MenuItem   As CommandBarPopup

 '  Parse the concatenated list of translated menu strings into separate
 '  variables

 ParseMenuStrings (str)

 ' Just in case its still there from last time, delete it.

 On Error Resume Next

 RemovePluginMenu
 
 '  Need to find Excel's Help menu (We have to use the published ID property
 '  because the spelling will be different for non-english locales

 Set bar = Application.CommandBars("Worksheet Menu Bar")

 Set HelpMenu = bar.FindControl(Id:=30010)
 If HelpMenu Is Nothing Then
   '  Can't find Help so add to end of menu bar
   Set OurMenu = bar.Controls.add(Type:=msoControlPopup, Temporary:=True)
 Else
   '  Insert our menu to the left of Help
   Set OurMenu = bar.Controls.add(Type:=msoControlPopup, Before:=HelpMenu.index, _
                                                      Temporary:=True)
 End If

 OurMenu.Caption = MenuTopLevel          '  Give it a caption

 '  Now add our submenus

 With OurMenu.Controls.add(Type:=msoControlButton)
   .Caption = MenuParseTSSheet
   .OnAction = "Menu_ParsingTS_OnAction"
   .Enabled = True
 End With
 
 With OurMenu.Controls.add(Type:=msoControlButton)
   .Caption = MenuSaveAsTSheet
   .OnAction = "Menu_SaveAsTS_OnAction"
   .Enabled = True
 End With
 
 End Sub


'--------------------------------------------------------------------------'
'    EnableMenu             Enable/Disable an OLAP menu item               '
'--------------------------------------------------------------------------'

Sub EnableMenu(sMenu As String, bEnable As Boolean)

Application.CommandBars("Worksheet Menu Bar").Controls(MenuTopLevel).Controls(sMenu).Enabled = bEnable

End Sub


'--------------------------------------------------------------------------'
'    RemoveOLAPMenu()      Remove the entire OLAP menu tree from the main  '
'                          Excel menu                                      '
'--------------------------------------------------------------------------'

Public Sub RemovePluginMenu()
 On Error Resume Next
        
 Application.CommandBars("Worksheet Menu Bar").Controls(MenuTopLevel).Delete
    
 '  Also remove the "OLAP Help" from Excel's Help menu

 Set bar = Application.CommandBars("Worksheet Menu Bar")

 On Error GoTo 0

End Sub


'--------------------------------------------------------------------------'
'    ParseMenuStrings()   Parses the concatenated translated menu string   '
'                         into individual menu item strings. This routine  '
'                         expects the delimited string to be in a specifc  '
'                         order!   Refer to getMenuString() in Core.java   '
'--------------------------------------------------------------------------'

Private Sub ParseMenuStrings(MenuString As String)
                                
Dim MenuItem As String
Dim Count As Integer
Dim i As Long
Dim j As Long

'  Extract the individual menu items into their own variables

i = 1
Count = 1
Do
  j = InStr(i, MenuString, ";")      '  ; is used as an item string delimiter
  If (j < 1) Then
    Exit Do
  End If
  
  MenuItem = Mid(MenuString, i, j - i)
  i = j + 1
  Select Case Count
    Case 1:  MenuParseTSSheet = MenuItem
    Case 2:  MenuSaveAsTSheet = MenuItem
  End Select
  Count = Count + 1
Loop
  
               
End Sub


