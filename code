Option Explicit
Option Compare Text

'NOTE: MAX LENGTH IS 27 FOR THE MENU LABELS
Private pHeight As Integer
Private pWidth As Integer
Private WithEvents xlApp As Application
Private pTextboxDefaultText As String
Private pTextboxColor As Long
Private ControlPressed As Boolean

Const SidebarExpandWidth As Integer = 120


'=================================================================
' INITIALIZE
'=================================================================
Private Sub UserForm_Initialize()
    
    'EVERY TIME THE USER OPENS THIS FORM, IT WILL CHECK FOR ANY UPDATES
    CheckForUpdates
    
    'CONSTRUCTORS
    Set xlApp = Application
    pHeight = Me.Height
    pWidth = Me.Width
    pTextboxDefaultText = TextBox1.Value
    pTextboxColor = TextBox1.BackColor
    
    'LOCATION\TEXTBOX START
    DisplayLocation
    Me.TextBox1.SelStart = 0
    
    'STARTUP WITH SHORTCUTKEYS, INSTEAD OF TEXTBOX
    ControlPressed = True
    hiddenButton.SetFocus
    LineNumbers True
    
    'CHECK TO SEE IF ON VC OFFER TABLES - SUGGEST BUTTONS
    If ActiveSheet.Name = "VC Pivot" Or ActiveSheet.Name = "VC Table" Then
        Label1U.Caption = "Bill Case OI"
        Label2U.Caption = "Bouncer Fill OI"
        Label3U.Caption = ""
        Label4U.Caption = ""
        Label5U.Caption = ""
        LineNumbers True
    End If
    
End Sub


'===========================================================================
' Main methods
'===========================================================================

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' THIS IS THE MAIN METHOD, THAT IS CALLED FROM ANY CLICK OF THE LABELS
' (OTHER THAN SOME OF THE SIDEBAR BUTTONS). OR ANYTHING TYPED IN
' TEXTBOX1, AND THEN ENTER IS PRESSED. OR ANY SHORTCUT KEYS PRESSED.
'
' IT WILL READ WHAT IS IN TEXTBOX1 AND RUN ANYCODE ASSOCIATED TO THAT.
' BUTTONS PRESSED & KEYBOARD SHORTCUTS, WILL FILL IN THE VALUE OF TEXTBOX1
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RunCommand()
    
    Dim cmd As String
    cmd = TextBox1.Value
    If Not cmd = "" Then Unload Me
    
    'LOG CODE RAN IN TEXTFILE... FOR DEVELOPER
    If TextBox1.Value <> "" And TextBox1.Value <> pTextboxDefaultText Then LogCode TextBox1.Value
   
    
    Select Case True
        
        Case cmd Like "TEST":
        Case cmd Like "Cic Date New":               CicDateNew.Show
        Case cmd Like "VC Offer":                   CaseAuditForm.Show
        Case cmd Like "Create List.txt":            GeneralTools.CreateList_txt
        Case cmd Like "Access Fill":                MicrosoftAccess.InputToAccess
        Case cmd Like "ASM Fill - Tribble Sheet":   GeneralTools.ASMFill
        Case cmd Like "Placement Query":            PlacementForm.Show
        Case cmd Like "Tribble Check Highlight":    Placement.WithTribble_Placement
        Case cmd Like "Placement Email":            MicrosoftOutlook.PlacementOutlookEmail
        Case cmd Like "DSD Data Sort":              DSD.DSD_Steps
        Case cmd Like "UPC From CIC":               TestingMod.TestingCreateUPCListFromCic
        Case cmd Like "Hightlight Mismatch":        TestingMod.HighlightMismatch
        Case cmd Like "Pull Pivot Data":            TableTools.CreateAllowanceOffersWorkbookFromPivotTable
        Case cmd Like "Debit Memo Fill":            runMacro ("DebitMemoFill")
        Case cmd Like "Update":                     Update.CheckForUpdates True
        Case cmd Like "Just For U Biller":          GeneralTools.j4uBox
        Case cmd Like "RemoveLines From Text File": RemoveLinesFromTextFile
        Case cmd Like "Custom Table":               customTableStyle
        Case cmd Like "TimeSheet Format":           RemoveDataFromTimeSheet
        Case cmd Like "Edeals Navigation":          Edeals.Show
        Case cmd Like "Refresh Bouncer List":       RefreshBouncerList
        Case cmd Like "Developer Log":              DeveloperLog.Show
        Case cmd Like "Goto 17/2":                  Goto_17set2
        Case cmd Like "Bill Case OI":               BillerForm.Show
        Case cmd Like "Bouncer Fill OI":            OIBouncerFill
        Case cmd Like "":                           TextboxControl
        
    End Select
    
End Sub

'=================================================================
' KEYPRESS EVENTS
'=================================================================
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' HIDDEN BUTTON IS USED FOR HAVING FOCUS NOT IN TEXTBOX.
' IT ALLOWS FOR KEYBOARD SHORTCUTS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HiddenButton_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    LineNumbers True
    
    Select Case KeyCode
        Case 97: Label1U_Click
        Case 98: Label2U_Click
        Case 99: Label3U_Click
        Case 100: Label4U_Click
        Case 101: Label5U_Click
        Case vbKey6: sidEdeals_Click
        Case vbKeyNumpad6: sidEdeals_Click
        Case vbKeyEscape: Unload Me
        Case vbKeyReturn: TextBox1.SetFocus
        Case vbKeyF1: sidBookmark_Click
        Case vbKeyF5: sidEdeals_Click
        Case vbKeyF2: butMainU_Click
        Case vbKeyF3: butPlacementU_Click
        Case vbKeyF4: butToolsU_Click
        Case vbKeyF7: Unload Me: CicAllowForAllDiv
        Case vbKeyF12: DeveloperLog.Show
        Case vbKeyM: butMainU_Click
        Case vbKeyP: butPlacementU_Click
        Case vbKeyT: butToolsU_Click
    End Select
    
End Sub
Private Sub HiddenButton_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    'SHOW LINENUMBERS FOR A HINT TO THE KEYBOARD SHORTCUTS
    LineNumbers True
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SEARCH BOX - ALLOWS USER TO SEARCH FOR CODE
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    
    menLabel1U.BackColor = blue
    menLabel1L.BackColor = blue
    menLabel1U.ForeColor = vbWhite
    
    Select Case KeyCode
        
        'SHORTCUTS
        Case vbKeyF5: sidEdeals_Click
        Case vbKeyReturn: TextBox1.Value = menLabel1U.Caption: RunCommand
        Case vbKeyTab: hiddenButton.SetFocus: ResetTextboxStyle
        
        'ESCAPE KEY
        Case vbKeyEscape:
        
            If Not TextBox1.Value = "" And Not TextBox1.Value = pTextboxDefaultText Then
                TextBox1.Value = pTextboxDefaultText
            Else
                hiddenButton.SetFocus
                LineNumbers
                ResetTextboxStyle
            End If
            
        Case Else
            
    End Select
    
End Sub

Private Sub TextBox1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    LineNumbers False
    TextboxControl
    SidebarMinimize
    SearchSuggestions
End Sub



'=========================================================================
' TEXTBOX OBJECT SECTION
'=========================================================================
Private Sub TextboxControl()
    TextBox1.BackColor = vbWhite
    
    If Len(TextBox1.Value) > Len(pTextboxDefaultText) Then
        TextBox1.Value = Replace(TextBox1.Value, pTextboxDefaultText, "")
    End If
    
    If Len(TextBox1.Value) < 1 Or TextBox1.Value = pTextboxDefaultText Then
        TextBox1.Value = pTextboxDefaultText
        TextBox1.ForeColor = veryLightGrey
        Me.TextBox1.SelStart = 0
        HidePopup
    Else
        DisplayPopup
    End If

End Sub

'-----------------------------------------------------------------------
' FUNCTIONS FOR TEXTBOX
'-----------------------------------------------------------------------
Private Sub ResetTextboxStyle()
    TextBox1.Value = pTextboxDefaultText
    TextBox1.BackColor = lightGrey
    TextBox1.ForeColor = vbWhite
    popup.Height = 0
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' TIED TO THE SEARCH BOX - SORTS LIST IN POPUP BOX
' CALLED FROM TextBox1_KeyUp
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SearchSuggestions()
    
    Dim SearchList() As Variant
    Dim L As Long
    Dim I As Integer
    
    menLabel1U.Caption = "": menLabel2U.Caption = ""
    menLabel3U.Caption = "": menLabel4U.Caption = "": menLabel5U.Caption = ""
    
    'ARRAY THAT CONTAINS THE LIST OF CODES TO DISPLAY IN THE POPUP
    ArrayPush SearchList, "Access Fill"
    ArrayPush SearchList, "ASM Fill - Tribble Sheet"
    ArrayPush SearchList, "CIC Date New"
    ArrayPush SearchList, "Create List.txt"
    ArrayPush SearchList, "Custom Table"
    ArrayPush SearchList, "Debit Memo Fill"
    ArrayPush SearchList, "Developer Log"
    ArrayPush SearchList, "DSD Data Sort"
    ArrayPush SearchList, "Edeals Navigation"
    ArrayPush SearchList, "Goto 17/2"
    ArrayPush SearchList, "Hightlight Mismatch"
    ArrayPush SearchList, "Just For U Biller"
    ArrayPush SearchList, "Placement Email"
    ArrayPush SearchList, "Placement Query"
    ArrayPush SearchList, "Pull Pivot Data"
    ArrayPush SearchList, "Refresh Bouncer List"
    ArrayPush SearchList, "RemoveLines From Text File"
    ArrayPush SearchList, "TimeSheet Format"
    ArrayPush SearchList, "Update"
    ArrayPush SearchList, "UPC From CIC"
    ArrayPush SearchList, "VC Offer"

    
    For L = LBound(SearchList, 1) To UBound(SearchList, 1)
        'Debug.Print SearchList(L)
        If VBA.Left(SearchList(L), Len(TextBox1.Value)) = Trim(TextBox1.Value) Then
            I = I + 1
            Select Case I
                Case 1: menLabel1U.Caption = SearchList(L)
                Case 2: menLabel2U.Caption = SearchList(L)
                Case 3: menLabel3U.Caption = SearchList(L)
                Case 4: menLabel4U.Caption = SearchList(L)
                Case 5: menLabel5U.Caption = SearchList(L)
            End Select
        End If
        
    Next L

'If Trim(VBA.Left(c.Value, Len(TextBox1.Value))) = Trim(TextBox1.Value) Then 'EXACT
'If Trim(c.Value) Like "*" & Trim(TextBox1.Value) & "*" Then ' CONTAINS

End Sub

'-----------------------------------------------------------------------
' EVENTS FOR TEXTBOX
'-----------------------------------------------------------------------
Private Sub TextBox1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    TextboxControl
    SidebarMinimize
End Sub
Private Sub TextBox1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    TextboxControl
    LineNumbers False
    TextBox1.BackColor = vbWhite
    If Me.Height < 41 Then Maximize
End Sub

Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetTextboxStyle
    SidebarMinimize
    Me.TextBox1.SelStart = 0
    Me.hiddenButton.SetFocus
    LineNumbers True
End Sub



'=========================================================================
' POPUP OBJECT SECTION
'=========================================================================
Private Sub DisplayPopup()
    
    Dim L As Long
    
    If Me.Height < 41 Then Maximize
    If popup.Height > 0 Then Exit Sub
    TextBox1.BackColor = vbWhite
    TextBox1.ForeColor = vbBlack
    
    popup.Height = 0
    popup.Top = 307
    
    For L = 0 To 17 '24.5 is for close to the top
        popup.Height = popup.Height + L
        popup.Top = popup.Top - L
        Me.Repaint
        Sleep 1
        
    Next L
    
End Sub

Private Sub HidePopup()

    popup.Height = 0
    popup.Top = 307

End Sub


'=================================================================
' SIDEBAR OBJECT SECTION
'=================================================================
Private Sub sidebarToggle()
    
    'FIND DIRECTION
    If Sidebar.Width = 30 Then
        SidebarMaximize
    Else
        SidebarMinimize
    End If
     
End Sub
Private Sub SidebarMinimize()
    
    If Not Sidebar.Width > 31 Then Exit Sub
    Dim I As Integer
    For I = Sidebar.Width To 30 Step -1
        Sidebar.Width = I
        Me.Repaint
        Sleep 0.01
    Next I
    
End Sub
Private Sub SidebarMaximize()
          
    If Not Sidebar.Width < SidebarExpandWidth Then Exit Sub
    popup.ZOrder (1)
    Dim I As Integer
    For I = Sidebar.Width To SidebarExpandWidth
        Sidebar.Width = I
        Me.Repaint
        Sleep 0.01
    Next I
      
End Sub



'=================================================================
' CLICK EVENTS MAIN SECTION
'=================================================================
Private Sub Label1L_Click()
    TextBox1.Value = Label1U.Caption: RunCommand
End Sub
Private Sub Label1U_Click()
    TextBox1.Value = Label1U.Caption: RunCommand
End Sub
Private Sub Label2L_Click()
    TextBox1.Value = Label2U.Caption: RunCommand
End Sub
Private Sub Label2U_Click()
    TextBox1.Value = Label2U.Caption: RunCommand
End Sub
Private Sub Label3L_Click()
    TextBox1.Value = Label3U.Caption: RunCommand
End Sub
Private Sub Label3U_Click()
    TextBox1.Value = Label3U.Caption: RunCommand
End Sub
Private Sub Label4L_Click()
    TextBox1.Value = Label4U.Caption: RunCommand
End Sub
Private Sub Label4U_Click()
    TextBox1.Value = Label4U.Caption: RunCommand
End Sub
Private Sub Label5L_Click()
    TextBox1.Value = Label5U.Caption: RunCommand
End Sub
Private Sub Label5U_Click()
    TextBox1.Value = Label5U.Caption: RunCommand
End Sub

'Sidebar Buttons
Private Sub sidBookmark_Click()
    LogCode "ToolsCatalog"
    ToolsCatalog
End Sub
Private Sub sidBookmarkColor_Click()
    sidBookmark_Click
End Sub
Private Sub sidBookmarkLabel_Click()
    sidBookmark_Click
End Sub
Private Sub sidSettings_Click()
    On Error GoTo ErrorCatch
    CreateFilePath SettingsFile
    OpenAnyFile SettingsFile
    
ErrorCatch:
End Sub
Private Sub sidSettingsColor_Click()
    sidSettings_Click
End Sub
Private Sub sidSettingsLabel_Click()
    sidSettings_Click
End Sub
Private Sub sidHome_Click()
    Header.Caption = "How can I help?" & vbNewLine & "................"
End Sub
Private Sub sidHomeColor_Click()
    sidHome_Click
End Sub
Private Sub sidHomeLabel_Click()
    sidHome_Click
End Sub
Private Sub sidEdeals_Click()
    LogCode "Edeals Lookup"
    Unload Me
    Edeals.Show
End Sub
Private Sub sidEdealsColor_Click()
   sidEdeals_Click
End Sub
Private Sub sidEdealsLabel_Click()
    sidEdeals_Click
End Sub
Private Sub sidMenu_Click()
    sidebarToggle
End Sub
Private Sub sidMenuColor_Click()
    sidMenu_Click
End Sub
Private Sub sidMenuLabel_Click()
    sidMenu_Click
End Sub


'MENU BUTTONS
Private Sub butMainU_Click()
    Label1U.Caption = "VC Offer"
    Label2U.Caption = "CIC Date New"
    Label3U.Caption = "Create List.txt"
    Label4U.Caption = "ASM Fill - Tribble Sheet"
    Label5U.Caption = "Access Fill"
    LineNumbers True
End Sub
Private Sub butMainL_Click()
    butMainU_Click
End Sub
Private Sub butPlacementU_Click()
    Label1U.Caption = "Placement Query"
    Label2U.Caption = "Tribble Check Highlight"
    Label3U.Caption = "Placement Email"
    Label4U.Caption = "Cic Date New"
    Label5U.Caption = ""
    LineNumbers True
End Sub
Private Sub butPlacementl_Click()
    butPlacementU_Click
End Sub
Private Sub butToolsU_Click()
    Label1U.Caption = "DSD Data Sort"
    Label2U.Caption = "UPC From CIC"
    Label3U.Caption = "Hightlight Mismatch"
    Label4U.Caption = "Pull Pivot Data"
    Label5U.Caption = "Debit Memo Fill"
    LineNumbers True
End Sub
Private Sub butToolsl_Click()
    butToolsU_Click
End Sub


'POPUP BUTTONS
Private Sub menLabel1L_Click()
    menLabel1U_Click
End Sub
Private Sub menLabel1U_Click()
    TextBox1.Value = menLabel1U.Caption: RunCommand
End Sub
Private Sub menLabel2L_Click()
    menLabel2U_Click
End Sub
Private Sub menLabel2U_Click()
    TextBox1.Value = menLabel2U.Caption: RunCommand
End Sub
Private Sub menLabel3L_Click()
    menLabel3U_Click
End Sub
Private Sub menLabel3U_Click()
    TextBox1.Value = menLabel3U.Caption: RunCommand
End Sub
Private Sub menLabel4L_Click()
    menLabel4U_Click
End Sub
Private Sub menLabel4U_Click()
    TextBox1.Value = menLabel4U.Caption: RunCommand
End Sub
Private Sub menLabel5L_Click()
    menLabel5U_Click
End Sub
Private Sub menLabel5U_Click()
    TextBox1.Value = menLabel5U.Caption: RunCommand
End Sub


'FORM CLICKS
Private Sub Header_Click()
    SidebarMinimize
End Sub
Private Sub Sidebar_Click()
    hiddenButton.SetFocus
End Sub


'ADD OR REMOVE LINE NUMBERS
Private Function LineNumbers(Optional Show As Boolean = True) As Boolean
    
    If Label1U.Caption <> "" Then Number1.Visible = Show Else: Number1.Visible = False
    If Label2U.Caption <> "" Then Number2.Visible = Show Else: Number2.Visible = False
    If Label3U.Caption <> "" Then Number3.Visible = Show Else: Number3.Visible = False
    If Label4U.Caption <> "" Then Number4.Visible = Show Else: Number4.Visible = False
    If Label5U.Caption <> "" Then Number5.Visible = Show Else: Number5.Visible = False
    
End Function



'=========================================================================
' RESIZING SECTION FOR THE USERFORM\SIDEBAR. IE MINIMIZE AND MAXIMIZE
'=========================================================================
Private Sub sidExpand_Click()
    ToggleSize
End Sub
Private Sub sidExpandColor_Click()
    ToggleSize
End Sub
Private Sub sidExpandLabel_Click()
    ToggleSize
End Sub

Private Sub ToggleSize()
    
    Dim I As Integer: I = 50
    SidebarMinimize
    
    If Me.Height > 41 Then
        minimize
    Else
        Maximize
    End If
    
End Sub

Public Sub Maximize()
    
    Dim L As Long
    
    Header.Visible = True
    SidMenu.Visible = True
    sidMenuColor.Visible = True
    TextBox1.Top = 307.5
    sidExpandColor.Top = 307
    sidExpand.Top = 307.5
    
    On Error Resume Next
    For L = Me.Height To pHeight Step 5
        DoEvents
        Me.Top = Me.Top - 5
        Me.Height = L
        Me.Repaint
        Sleep 0.01
        Me.Repaint
    Next L

End Sub

Public Sub minimize()
    
    ClearAll
    Me.Height = 40

    DisplayLocation
    
    Header.Visible = False
    SidMenu.Visible = False
    sidMenuColor.Visible = False
    
    ResetTextboxStyle
    TextBox1.Top = 0
    popup.Height = 0
    popup.Top = 307
    sidExpand.Top = 0
    sidExpandColor.Top = 0.5
    
End Sub

'IF THE WINDOW IS RESIZED, THIS WILL PLACE THE USERFORM IN THE LOWER RIGHT CORNER
Private Sub xlApp_WindowResize(ByVal Wb As Workbook, ByVal Wn As Window)
    DisplayLocation
End Sub
Private Sub DisplayLocation()
    
     With Me
        .StartUpPosition = 0
        .Left = Application.Left + ((Application.Width) - (.Width + 35))
        .Top = Application.Top + (Application.Height) - (.Height + 55)
    End With

End Sub



'==========================================================================
' FORM APPEARANCE SECTION - MOSTLY HIGHLIGHTING BUTTONS
'==========================================================================
'BUTTON (LABELS) HIGHLIGHT
Private Sub ButtonHighlight(Item As MSForms.Label, Optional item2 As MSForms.Label)

    ClearAll
    On Error Resume Next
    
    'sidbar
    If Left(Item.Name, 3) = "sid" Then
        If Item.BackColor = blue Or Item.BackColor = bluehover Then
            Item.BackColor = bluehover
        Else
            Item.BackColor = basicGreyHover
            item2.BackColor = basicGreyHover
        End If
    'other buttons
    Else
        Item.BackColor = bluehover
        item2.BackColor = bluehover
    End If
    
    
    Item.ForeColor = vbWhite
    item2.ForeColor = vbWhite
    
End Sub
Private Sub ClearAll()

    Dim ctl

    For Each ctl In Me.Controls
    
        'LABELS
        If TypeName(ctl) = "Label" Then
            
            'SIDEBAR DEFAULT
            If Left(ctl.Name, 3) = "sid" Then
                If ctl.BackColor = blue Or ctl.BackColor = bluehover Then
                    ctl.BackColor = blue
                Else
                    ctl.BackColor = darkGrey
                End If
                ctl.ForeColor = vbWhite
            End If
            
            'BUTTONS DEFAULT
            If Left(ctl.Name, 3) = "but" Then
                ctl.BackColor = lightGrey
                ctl.ForeColor = vbWhite
            End If
            
            'POPUP BUTTONS
            If Left(ctl.Name, 3) = "men" Then
                If ctl.Name = "menLabel1L" Or ctl.Name = "menLabel1U" Then
                    ctl.BackColor = blue
                    ctl.ForeColor = vbWhite
                Else
                    ctl.BackColor = vbWhite
                    ctl.ForeColor = vbBlack
                End If
            End If
            
            'LABEL BUTTONS
            If Left(ctl.Name, 3) = "lab" Then
                ctl.BackColor = blue
            End If
        End If
    Next ctl
    
End Sub


'SIDEBAR
Private Sub sidMenu_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonHighlight sidMenuColor
End Sub
Private Sub sidMenuColor_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonHighlight sidMenuColor
End Sub
Private Sub sidMenuLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonHighlight sidMenuColor
End Sub
Private Sub sidHome_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonHighlight sidHomeColor
End Sub
Private Sub sidHomeColor_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonHighlight sidHomeColor
End Sub
Private Sub sidHomeLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonHighlight sidHomeColor
End Sub
Private Sub sidBookmark_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonHighlight sidBookmarkColor
End Sub
Private Sub sidBookmarkColor_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonHighlight sidBookmarkColor
End Sub
Private Sub sidBookmarkLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonHighlight sidBookmarkColor
End Sub
Private Sub sidSettings_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonHighlight sidSettingsColor
End Sub
Private Sub sidSettingsColor_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonHighlight sidSettingsColor
End Sub
Private Sub sidSettingsLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonHighlight sidSettingsColor
End Sub
Private Sub sidEdeals_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonHighlight sidEdealsColor
End Sub
Private Sub sidEdealsColor_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonHighlight sidEdealsColor
End Sub
Private Sub sidEdealsLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonHighlight sidEdealsColor
End Sub



Private Sub sidExpand_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonHighlight sidExpandColor
End Sub
Private Sub sidExpandColor_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonHighlight sidExpandColor
End Sub
Private Sub sidExpandLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonHighlight sidExpandColor
End Sub


'BUTTONS
Private Sub butToolsL_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonHighlight butToolsL, butToolsU
End Sub
Private Sub butToolsU_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonHighlight butToolsL, butToolsU
End Sub
Private Sub butMainL_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonHighlight butMainL, butMainU
End Sub
Private Sub butMainU_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonHighlight butMainL, butMainU
End Sub
Private Sub butPlacementL_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonHighlight butPlacementL, butPlacementU
End Sub
Private Sub butPlacementU_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonHighlight butPlacementL, butPlacementU
End Sub

'Label Buttons
Private Sub Label1U_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonHighlight Label1U, Label1L
End Sub
Private Sub Label1l_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonHighlight Label1U, Label1L
End Sub
Private Sub Label2U_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonHighlight Label2U, Label2L
End Sub
Private Sub Label2l_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonHighlight Label2U, Label2L
End Sub
Private Sub Label3U_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonHighlight Label3U, Label3L
End Sub
Private Sub Label3l_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonHighlight Label3U, Label3L
End Sub
Private Sub Label4U_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonHighlight Label4U, Label4L
End Sub
Private Sub Label4l_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonHighlight Label4U, Label4L
End Sub
Private Sub Label5U_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonHighlight Label5U, Label5L
End Sub
Private Sub Label5l_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonHighlight Label5U, Label5L
End Sub



'Popup buttons
Private Sub menLabel1U_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonHighlight menLabel1U, menLabel1L
End Sub
Private Sub menLabel1L_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonHighlight menLabel1U, menLabel1L
End Sub

Private Sub menLabel2U_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonHighlight menLabel2U, menLabel2L
End Sub
Private Sub menLabel2L_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonHighlight menLabel2U, menLabel2L
End Sub

Private Sub menLabel3U_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonHighlight menLabel3U, menLabel3L
End Sub
Private Sub menLabel3L_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonHighlight menLabel3U, menLabel3L
End Sub

Private Sub menLabel4U_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonHighlight menLabel4U, menLabel4L
End Sub
Private Sub menLabel4L_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonHighlight menLabel4U, menLabel4L
End Sub

Private Sub menLabel5U_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonHighlight menLabel5U, menLabel5L
End Sub
Private Sub menLabel5L_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonHighlight menLabel5U, menLabel5L
End Sub


'CLEAR ALLS
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ClearAll
End Sub
Private Sub Sidebar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ClearAll
End Sub
Private Sub TextBox1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ClearAll
End Sub
Private Sub popup_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ClearAll
End Sub
