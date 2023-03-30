VERSION 4.00
Begin VB.MDIForm mdiMain 
   BackColor       =   &H80000001&
   Caption         =   "Generic Staff Roster"
   ClientHeight    =   6660
   ClientLeft      =   3210
   ClientTop       =   2475
   ClientWidth     =   9390
   Height          =   7350
   Icon            =   "MDIMAIN.frx":0000
   Left            =   3150
   LinkTopic       =   "mdiMain"
   LockControls    =   -1  'True
   Top             =   1845
   WhatsThisHelp   =   -1  'True
   Width           =   9510
   Begin Threed.SSPanel PanelToolBar 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   7
      Top             =   0
      Width           =   9390
      _version        =   65536
      _extentx        =   16563
      _extenty        =   820
      _stockprops     =   15
      forecolor       =   -2147483641
      backcolor       =   -2147483644
      BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   7.5
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      bevelwidth      =   2
      borderwidth     =   1
      bevelouter      =   1
      roundedcorners  =   0   'False
      floodtype       =   1
      floodcolor      =   -2147483646
      floodshowpct    =   0   'False
      alignment       =   0
      mouseicon       =   "MDIMAIN.frx":0CFA
      Begin Threed.SSCommand cmdStaffReport 
         Height          =   360
         Left            =   2760
         TabIndex        =   6
         Top             =   45
         Width           =   360
         _version        =   65536
         _extentx        =   635
         _extenty        =   635
         _stockprops     =   78
         forecolor       =   -2147483630
         autosize        =   2
         picture         =   "MDIMAIN.frx":15D4
      End
      Begin Threed.SSCommand cmdExceptionReport 
         Height          =   360
         Left            =   2370
         TabIndex        =   5
         Top             =   45
         Width           =   360
         _version        =   65536
         _extentx        =   635
         _extenty        =   635
         _stockprops     =   78
         forecolor       =   -2147483630
         autosize        =   2
         picture         =   "MDIMAIN.frx":1B26
      End
      Begin Threed.SSCommand cmdFonts 
         Height          =   360
         Left            =   1980
         TabIndex        =   4
         Top             =   45
         Width           =   360
         _version        =   65536
         _extentx        =   635
         _extenty        =   635
         _stockprops     =   78
         forecolor       =   -2147483630
         autosize        =   2
         picture         =   "MDIMAIN.frx":1E78
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   360
         Left            =   1590
         TabIndex        =   3
         Top             =   45
         Width           =   360
         _version        =   65536
         _extentx        =   635
         _extenty        =   635
         _stockprops     =   78
         forecolor       =   -2147483630
         autosize        =   2
         picture         =   "MDIMAIN.frx":1F8A
      End
      Begin Threed.SSRibbon cmdControlWindow 
         Height          =   360
         Left            =   30
         TabIndex        =   0
         Top             =   45
         Width           =   360
         _version        =   65536
         _extentx        =   635
         _extenty        =   635
         _stockprops     =   65
         backcolor       =   12632256
         value           =   -1  'True
         groupnumber     =   3
         pictureup       =   "MDIMAIN.frx":209C
      End
      Begin Threed.SSRibbon cmdRosterWindow 
         Height          =   360
         Left            =   810
         TabIndex        =   2
         Top             =   45
         Width           =   360
         _version        =   65536
         _extentx        =   635
         _extenty        =   635
         _stockprops     =   65
         backcolor       =   12632256
         value           =   -1  'True
         groupnumber     =   2
         pictureup       =   "MDIMAIN.frx":23EE
      End
      Begin Threed.SSRibbon cmdStaffWindow 
         Height          =   360
         Left            =   420
         TabIndex        =   1
         Top             =   45
         Width           =   360
         _version        =   65536
         _extentx        =   635
         _extenty        =   635
         _stockprops     =   65
         backcolor       =   12632256
         value           =   -1  'True
         pictureup       =   "MDIMAIN.frx":2740
      End
   End
   Begin Threed.SSPanel panelStatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   6405
      Width           =   9390
      _version        =   65536
      _extentx        =   16563
      _extenty        =   450
      _stockprops     =   15
      caption         =   "Status Bar Message Area"
      forecolor       =   -2147483639
      backcolor       =   -2147483646
      borderwidth     =   2
      bevelouter      =   1
      roundedcorners  =   0   'False
      floodshowpct    =   0   'False
      alignment       =   8
      autosize        =   2
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   90
      Top             =   570
      _version        =   65536
      _extentx        =   847
      _extenty        =   847
      _stockprops     =   0
      fontname        =   "Arial"
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLoadFromFile 
         Caption         =   "Load Roster from File"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuSavetoFile 
         Caption         =   "Save Roster to File"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSpacer2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintSetup 
         Caption         =   "Printer Set&up"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuFileSpacer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Quit 
         Caption         =   "&Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuToolbarState 
         Caption         =   "Hide &Toolbar"
      End
      Begin VB.Menu mnuStatusBarState 
         Caption         =   "Hide &Status Bar"
      End
      Begin VB.Menu mnuLockRoster 
         Caption         =   "Lock &Roster Columns"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu mnuReportS 
      Caption         =   "&Reports"
      Begin VB.Menu mnuException 
         Caption         =   "&Exception"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuStaffRpt 
         Caption         =   "S&taff"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuSingleTimesheet 
         Caption         =   "S&ingle Timesheet"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuAllTimesheets 
         Caption         =   "All Ti&mesheets"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuReportSpacer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintGrid 
         Caption         =   "&Print Current Grid"
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuMaximise 
         Caption         =   "&Maximise"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mnuCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowSpacer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowAll 
         Caption         =   "&Show All"
      End
      Begin VB.Menu mnuHideAll 
         Caption         =   "&Hide All"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp_About 
         Caption         =   "&About"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuRegister 
         Caption         =   "&Register"
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit
'[----------------------------------------------]
'[mdiMain           Main Parent Form            ]
'[----------------------------------------------]
'[GSR                                           ]
'[generic staff roster system for Windows       ]
'[(C) ---------- David Gilbert, 1997            ]
'[----------------------------------------------]
'[version                           2.2         ]
'[initial date                      21/07/1997  ]
'[----------------------------------------------]







Private Sub cmdControlWindow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    Select Case cmdControlWindow.Value
    Case True
        StatusBar "Click here to hide the control form."
    Case False
        StatusBar "Click here to show the control form."
    Case Else
    End Select
    '[---------------------------------------------------------------------------------]
    
End Sub

Private Sub cmdExceptionReport_Click()

    '[CALL SUBROUTINE WHICH PROCESSES THE EXCEPTION REPORT]
    Call procExceptionReport
    
End Sub


Private Sub cmdControlWindow_Click(Value As Integer)
    
    '[SHOW CONTROL WINDOW AND MOVE TO FRONT OF ZORDER]
    If Value = True Then
        frmControl.Show
        frmControl.ZOrder
        frmControl.WindowState = 0
    Else
        frmControl.Hide
    End If

End Sub

Private Sub cmdExceptionReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Click this button to initiate the exception report which will list any errors found within your roster design."
    '[---------------------------------------------------------------------------------]

End Sub

Private Sub cmdFonts_Click()

    '[ERROR HANDLER]
    On Error Resume Next
    
    '[GET CURRENT FONTS FROM SELECTED FORM]
    mdiMain.CommonDialog.FontName = ActiveForm.ActiveControl.Font.Name
    mdiMain.CommonDialog.FontSize = ActiveForm.ActiveControl.Font.Size
    mdiMain.CommonDialog.FontBold = ActiveForm.ActiveControl.Font.Bold
    mdiMain.CommonDialog.FontItalic = ActiveForm.ActiveControl.Font.Italic
    
    '[SET FONT DIALOG TO CURRENT OPTIONS]
    
    '[SHOW FONT DIALOG]
    mdiMain.CommonDialog.Flags = cdlCFPrinterFonts
    mdiMain.CommonDialog.ShowFont
    
    '[APPLY FONT TO CURRENT FORM FOR GRID RESIZING]
    ActiveForm.Font.Name = mdiMain.CommonDialog.FontName
    ActiveForm.Font.Size = mdiMain.CommonDialog.FontSize
    ActiveForm.Font.Bold = mdiMain.CommonDialog.FontBold
    ActiveForm.Font.Italic = mdiMain.CommonDialog.FontItalic
    
    '[APPLY TO ALL OBJECTS ON THE FORM]
    ActiveForm.ActiveControl.Font.Name = mdiMain.CommonDialog.FontName
    ActiveForm.ActiveControl.Font.Size = mdiMain.CommonDialog.FontSize
    ActiveForm.ActiveControl.Font.Bold = mdiMain.CommonDialog.FontBold
    ActiveForm.ActiveControl.Font.Italic = mdiMain.CommonDialog.FontItalic
    
End Sub

Private Sub cmdFonts_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Click this button to change the font for the currently displayed form."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdGotoControl_Click()
    
    '[MOVE CONTROL FORM TO THE FRONT]
    frmControl.ZOrder
    '[CHECK ROSTER WINDOW STATE]
    If frmRoster.WindowState = 2 Then frmControl.WindowState = 2
        
End Sub

Private Sub cmdGotoControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Clicking this button will move the control form to the front."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdGotoRoster_Click()
    
    '[MOVE ROSTER FORM TO THE FRONT]
    If frmRoster.WindowState = 1 Then frmRoster.WindowState = 0
    frmRoster.ZOrder

End Sub


Private Sub cmdGotoRoster_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Clicking this button will move the roster form to the front."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdGotoStaff_Click()

    '[MOVE STAFF FORM TO THE FRONT]
    frmStaff.ZOrder
    '[CHECK ROSTER WINDOW STATE]
    If frmRoster.WindowState = 2 Then frmStaff.WindowState = 2

End Sub



Private Sub cmdGotoStaff_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Clicking this button will move the staff details form to the front."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdPrint_Click()

    '[CALL PROCEEDURE TO PRINT CURRENT GRID]
    Call procCurrentGrid
    
End Sub

Private Sub cmdPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    '[ERROR CHECK FOR IF NO FORMS ARE LOADED]
    On Error Resume Next
    
    If mdiMain.ActiveForm.Name = "frmRoster" Then
        StatusBar "Click this button to print the displayed roster."
    ElseIf mdiMain.ActiveForm.Name = "frmReport" Then
        StatusBar "Click this button to print the displayed report grid."
    Else
        StatusBar "This button will print rosters/processed reports when the appropriate forms are highlighted."
    End If
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdRosterWindow_Click(Value As Integer)
    
    '[SHOW ROSTER WINDOW AND MOVE TO FRONT OF ZORDER]
    If Value = True Then
        frmRoster.Show
        frmRoster.ZOrder
    Else
        frmRoster.Hide
    End If

End Sub

Private Sub cmdRosterWindow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    Select Case cmdRosterWindow.Value
    Case True
        StatusBar "Click here to hide the roster form."
    Case False
        StatusBar "Click here to show the roster form."
    Case Else
    End Select
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdStaffReport_Click()

    '[CALL SUBROUTINE WHICH PROCESSES THE STAFF REPORT]
    Call procStaffReport
    
End Sub

Private Sub cmdStaffReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Click this button to run the staff report, detailing staff cost breakdowns for each active roster."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdStaffWindow_Click(Value As Integer)

    '[SHOW STAFF WINDOW AND MOVE TO FRONT OF ZORDER]
    If Value = True Then
        frmStaff.Show
        frmStaff.ZOrder
        frmStaff.WindowState = 0
    Else
        frmStaff.Hide
    End If
    
End Sub


Private Sub cmdStaffWindow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    Select Case cmdStaffWindow.Value
    Case True
        StatusBar "Click here to hide the staff form."
    Case False
        StatusBar "Click here to show the staff form."
    Case Else
    End Select
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub MDIForm_Load()
    
    '[*******************************]
    flagBeta = False
    '[TRUE      = BETA VERSION       ]
    '[FALSE     = FULL VERSION       ]
    '[*******************************]

    '[load splash form and display, wait for control to return]
    frmSplash.Show 1
    
    '[center form on screen]
    Top = (Screen.Height / 2) - (Height / 2)
    Left = (Screen.Width / 2) - (Width / 2)
   
    
    '[initialise dynasets and other variables]
    Call Initialise
    
    '[SET TOOLBAR STATE]
    If DsDefault("ToolBarState") = 0 Then
        '[HIDE TOOLBAR]
        Call mnuToolbarState_Click
    End If
    
    '[SET STATUSBAR STATE]
    If DsDefault("StatusBarState") = 0 Then
        '[HIDE STATUSBAR]
        Call mnuStatusBarState_Click
    End If
    
    '[SET ROSTER COLUMN STATE]
    If DsDefault("RosterLocked") = 0 Then
        '[UNLOCK ROSTER]
        Call mnuLockRoster_Click
    End If
    
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Information on various controls and the general use of GSR can be found here !"
    '[---------------------------------------------------------------------------------]
    
End Sub


Private Sub mnuAllTimesheets_Click()
    
    '[CALL ROUTINE TO PROCESS ALL STAFF RECORDS AND PRINT TIME SHEETS]
    Call procAllStaffRosters

End Sub

Private Sub mnuCascade_Click()

    '[ARRANGE VISIBLE WINDOWS IN CASCADE FORMAT]
    mdiMain.Arrange vbCascade

End Sub


Private Sub mnuException_Click()
    
    '[CALL SUBROUTINE WHICH PROCESSES THE EXCEPTION REPORT]
    Call procExceptionReport

End Sub

Private Sub mnuFile_Quit_Click()

    '[CALL TERMINATE SUBROUTINE]
    Call Terminate
    
End Sub


Private Sub mnuHelp_About_Click()

    '[load fmrAbout and set labelRegUser to the Registered User Value in the database]
    Load frmAbout
    frmAbout.labelRegUser.Caption = "This copy of Generic Staff Roster is licensed to :" & Chr$(vbKeyReturn) & DsDefault("RegUser")
    frmAbout.LabelEmail.Caption = "Email Address : " & constEmail
    frmAbout.labelInfo = "GSR - Generic Staff Roster.  (C) 1997, David Gilbert" & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "Version " & App.Major & "." & App.Minor & "." & App.Revision
    '[BETA FLAG]
    If flagBeta = True Then frmAbout.labelInfo = frmAbout.labelInfo & " beta"
    
    frmAbout.labelInfo = frmAbout.labelInfo & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "You have been using Generic Staff Roster for " & intDaysUsed & " days."

    frmAbout.Show 1
    mdiMain.Show
    
End Sub







Private Sub mnuHideAll_Click()
    
    '[MAKE ALL STANDARD FORMS VISIBLE]
    mdiMain.cmdControlWindow.Value = False
    mdiMain.cmdStaffWindow.Value = False
    mdiMain.cmdRosterWindow.Value = False

End Sub


Private Sub mnuLoadFromFile_Click()

    '[SET CALL TO ERRORHANDLING ROUTINE]
    On Error GoTo ErrorHandler
    Dim intFileHandle           As Integer
    
    '[SET DEFAULTS FOR FILE DIALOG]
    FileSetFilter
    
    '[DIALOG BOX TO OPEN A .FPL FIELD PLAN FILE]
    mdiMain.CommonDialog.Flags = cdlOFNFileMustExist Or cdlOFNPathMustExist
    mdiMain.CommonDialog.DialogTitle = "Open A Roster"
    mdiMain.CommonDialog.CancelError = True
    mdiMain.CommonDialog.ShowOpen

    Select Case mdiMain.CommonDialog.FileName
        Case "" '[NO FILE SELECTED]
            '[NO CHANGES TO STATUS QUO]
        Case Else '[OPEN SELECTED FILE]
            FileRead (mdiMain.CommonDialog.FileName)
            frmRoster.GridRoster.Row = 1
            frmRoster.GridRoster.Col = 0
    End Select

ErrorHandler:
    If Err.Number > 0 Then
        '[CANCEL WAS PRESSED ON THE SAVE FORM - NO PROCESSING REQUIRED]
        If Err.Number = cdlCancel Then Debug.Print "Cancel Pressed in mnuLoadFromFile"
        '[DEBUG]
        Debug.Print Err.Number & " - Error in mnuLoadFromFile Module"
    End If


End Sub

Private Sub mnuLockRoster_Click()

    '[LOCK/UNLOCK THE FIRST TWO ROSTER COLUMNS]
    Select Case mnuLockRoster.Caption
    Case "Lock &Roster Columns"
        frmRoster.GridRoster.FixedCols = 2
        mnuLockRoster.Caption = "Unlock &Roster Columns"
    Case "Unlock &Roster Columns"
        frmRoster.GridRoster.FixedCols = 0
        mnuLockRoster.Caption = "Lock &Roster Columns"
    End Select

End Sub

Private Sub mnuMaximise_Click()

    '[MAXIMISE CURRENT WINDOW]
    On Error Resume Next
    ActiveForm.WindowState = 2

End Sub

Private Sub mnuPrintGrid_Click()

    '[CALL PROCEEDURE TO PRINT CURRENTLY HIGHLIGHTED GRID]
    Call procCurrentGrid

End Sub

Private Sub mnuPrintSetup_Click()

    '[LOAD PRINTER DIALOG]
    mdiMain.CommonDialog.Flags = cdlPDPrintSetup    '[JUST DISPLAY THE PRINT SETUP BOX]
    mdiMain.CommonDialog.ShowPrinter

End Sub

Private Sub mnuRegister_Click()

    '[DECLARE VARIABLES]
    Dim Msg
    Dim Style
    Dim Title
    Dim Response

    '[SHOW REGISTER FORM AND VALIDATE REGISTRATION CODE]
    Load frmRegister
    '[PLACE VALUES IN FORM TEXT BOXES]
    frmRegister.txtRegUser = DsDefault("RegUser")
    
    '[SHOW FORM]
    frmRegister.Show 1
    
    '[PROCESS RESULT OF FORM]
    If frmRegister.CheckResult = vbChecked Then
        If Validate(frmRegister.txtRegUser) = frmRegister.txtRegCode And frmRegister.txtRegCode > "" Then
            '[CODE MATCHES, PLACE NEW NAME AND CODE INTO THE DATABASE]
            DsDefault.Edit
                DsDefault("RegUser") = frmRegister.txtRegUser
                DsDefault("RegCode") = frmRegister.txtRegCode
            DsDefault.Update
            
            '[SHOW MESSAGE FORM]
            Msg = "Thank you for registering Generic Staff Roster." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "Name : " & DsDefault("RegUser") & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "Code : " & DsDefault("RegCode")
            Style = vbOKOnly                     ' Define buttons.
            Title = "Registration Confirmed"
            Response = gsrMsg(Msg, Style, Title)
        Else
            '[CODE DOESN'T MATCH, PLACE UNREGISTERED DETAILS INTO DATABASE]
            DsDefault.Edit
                DsDefault("RegUser") = "Unregistered Version"
                DsDefault("RegCode") = ""
            DsDefault.Update
            
            '[SHOW MESSAGE FORM]
            Msg = "Your registration validation code does not match your registered user name.  Please check and re-enter." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "Name : " & frmRegister.txtRegUser & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "Code : " & frmRegister.txtRegCode
            Style = vbOKOnly                     ' Define buttons.
            Title = "Registration Incorrect"
            Response = gsrMsg(Msg, Style, Title)
        End If
    End If
    
    '[REMOVE FORM]
    Unload frmRegister

End Sub

Private Sub mnuRestore_Click()

    '[RESTORE CURRENT WINDOW]
    On Error Resume Next
    ActiveForm.WindowState = 0

End Sub




Private Sub mnuSavetoFile_Click()

    '[SAVE THIS ROSTER TO A FILE]
    '[FILENAME IS PREALLOCATED BUT MAY BE CHANGED]
    Dim strSaveFile             As String
    Dim intClassCounter         As Integer
    Dim strBookmark             As String
    Dim intSlashFound           As Integer

    '[CREATE FILENAME FOR FILE SAVE]
    strSaveFile = LCase(Trim(DsDefault("StartDate")) & "." & Trim(DsClass("Code")))
    
    '[REMOVE SLASHES FROM FILENAME]
    intSlashFound = InStr(strSaveFile, "/")
    Do While intSlashFound > 0
        '[REPLACE WITH HYPHEN]
        Mid$(strSaveFile, intSlashFound, 1) = "-"
        intSlashFound = InStr(strSaveFile, "/")
    Loop
    
    '[REMOVE SLASHES FROM FILENAME]
    intSlashFound = InStr(strSaveFile, "\")
    Do While intSlashFound > 0
        '[REPLACE WITH HYPHEN]
        Mid$(strSaveFile, intSlashFound, 1) = "-"
        intSlashFound = InStr(strSaveFile, "\")
    Loop
    
    '[SET DEFAULTS FOR FILE DIALOG]
    FileSetFilter
    
    '[CALL SAVE FILE ROUTINE]
    FileSaveAs (strSaveFile)
    
End Sub

Private Sub mnuShowAll_Click()

    '[MAKE ALL STANDARD FORMS VISIBLE]
    mdiMain.cmdControlWindow.Value = True
    mdiMain.cmdStaffWindow.Value = True
    mdiMain.cmdRosterWindow.Value = True
    
End Sub



Private Sub mnuSingleTimesheet_Click()

    '[CALL ROUTINE TO PROCESS SINGLE STAFF RECORD AND PRINT TIME SHEET]
    Call procSelectedStaffRoster

End Sub

Private Sub mnuStaffRpt_Click()

    '[CALL SUBROUTINE WHICH PROCESSES THE STAFF REPORT]
    Call procStaffReport

End Sub


Private Sub mnuStatusBarState_Click()

    '[SHOW OR HIDE THE STATUS BAR]
    Select Case mnuStatusBarState.Caption
    Case "Show Status Bar"
        mdiMain.panelStatusBar.Visible = True
        mnuStatusBarState.Caption = "Hide Status Bar"
    Case "Hide Status Bar"
        mdiMain.panelStatusBar.Visible = False
        mnuStatusBarState.Caption = "Show Status Bar"
    End Select

End Sub

Private Sub mnuToolbarState_Click()

    Select Case mnuToolbarState.Caption
    Case "Show Toolbar"
        mdiMain.PanelToolBar.Visible = True
        mnuToolbarState.Caption = "Hide Toolbar"
    Case "Hide Toolbar"
        mdiMain.PanelToolBar.Visible = False
        mnuToolbarState.Caption = "Show Toolbar"
    End Select
    
End Sub

Private Sub panelStatusBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Select 'Hide Status Bar' from the 'Options' menu to remove this bar from the screen."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub PanelToolBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Select 'Hide Toolbar' from the 'Options' menu to remove this bar from the screen."
    '[---------------------------------------------------------------------------------]

End Sub


