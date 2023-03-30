VERSION 4.00
Begin VB.Form frmControl 
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control"
   ClientHeight    =   3450
   ClientLeft      =   4155
   ClientTop       =   5160
   ClientWidth     =   7440
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   Height          =   3855
   Icon            =   "FRMCONTR.frx":0000
   Left            =   4095
   LinkTopic       =   "frmControl"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   Top             =   4815
   Width           =   7560
   Begin Threed.SSFrame FrameControl 
      Height          =   3405
      Left            =   0
      TabIndex        =   12
      Top             =   30
      Width           =   7425
      _version        =   65536
      _extentx        =   13097
      _extenty        =   6006
      _stockprops     =   14
      shadowstyle     =   1
      Begin VB.Frame FrameRosterSetup 
         Caption         =   "Roster Setup"
         Height          =   3225
         Left            =   3090
         TabIndex        =   13
         Top             =   90
         Width           =   4245
         Begin VB.CheckBox CheckDelete 
            Alignment       =   1  'Right Justify
            Caption         =   "Delete Confirmation"
            BeginProperty Font 
               name            =   "Arial"
               charset         =   1
               weight          =   400
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   90
            TabIndex        =   10
            Top             =   2250
            Value           =   1  'Checked
            Width           =   2325
         End
         Begin VB.ComboBox ComboMinute 
            Height          =   315
            ItemData        =   "FRMCONTR.frx":000C
            Left            =   1380
            List            =   "FRMCONTR.frx":001C
            TabIndex        =   4
            Text            =   "ComboMinute"
            Top             =   1080
            Width           =   1035
         End
         Begin VB.ComboBox ComboHour 
            Height          =   315
            ItemData        =   "FRMCONTR.frx":0030
            Left            =   90
            List            =   "FRMCONTR.frx":008A
            TabIndex        =   3
            Text            =   "ComboHour"
            Top             =   1080
            Width           =   1245
         End
         Begin VB.ComboBox ComboStartDay 
            BeginProperty Font 
               name            =   "Arial"
               charset         =   1
               weight          =   400
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "FRMCONTR.frx":00EE
            Left            =   90
            List            =   "FRMCONTR.frx":0107
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   480
            Width           =   2325
         End
         Begin VB.ComboBox ComboIncrement 
            BeginProperty Font 
               name            =   "Arial"
               charset         =   1
               weight          =   400
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "FRMCONTR.frx":014B
            Left            =   2520
            List            =   "FRMCONTR.frx":017E
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   450
            Width           =   1605
         End
         Begin VB.OptionButton OptionTIme 
            Caption         =   "End Time"
            Height          =   315
            Index           =   1
            Left            =   2520
            TabIndex        =   6
            Top             =   1440
            Width           =   1005
         End
         Begin VB.OptionButton OptionTIme 
            Caption         =   "Start Time"
            Height          =   315
            Index           =   0
            Left            =   2520
            TabIndex        =   5
            Top             =   1080
            Value           =   -1  'True
            Width           =   1005
         End
         Begin Threed.SSCommand cmdNukeCover 
            Height          =   600
            Left            =   3600
            TabIndex        =   11
            Top             =   2580
            Width           =   600
            _version        =   65536
            _extentx        =   1058
            _extenty        =   1058
            _stockprops     =   78
            autosize        =   1
            picture         =   "FRMCONTR.frx":0200
         End
         Begin VB.Label Label_std 
            BackStyle       =   0  'Transparent
            Caption         =   "Starting Day of Week"
            BeginProperty Font 
               name            =   "Arial"
               charset         =   1
               weight          =   400
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Index           =   0
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label Label_std 
            BackStyle       =   0  'Transparent
            Caption         =   "(Today)"
            BeginProperty Font 
               name            =   "Arial"
               charset         =   1
               weight          =   400
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   4
            Left            =   2940
            TabIndex        =   20
            Top             =   1920
            Width           =   570
         End
         Begin VB.Label Label_std 
            BackStyle       =   0  'Transparent
            Caption         =   "Start Date:"
            BeginProperty Font 
               name            =   "Arial"
               charset         =   1
               weight          =   400
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   3
            Left            =   120
            TabIndex        =   19
            Top             =   1920
            Width           =   2295
         End
         Begin MSMask.MaskEdBox MaskDate 
            DataField       =   "DateHired"
            Height          =   315
            Left            =   1140
            TabIndex        =   8
            Top             =   1860
            Width           =   1275
            _version        =   65536
            _extentx        =   2249
            _extenty        =   556
            _stockprops     =   109
            forecolor       =   0
            backcolor       =   16777215
            BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
               name            =   "Arial"
               charset         =   1
               weight          =   400
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            borderstyle     =   1
            autotab         =   -1  'True
            promptinclude   =   0   'False
            clipmode        =   1
            format          =   "Short Date"
         End
         Begin Threed.SSCommand cmdToday 
            Height          =   360
            Left            =   2520
            TabIndex        =   9
            Top             =   1830
            Width           =   360
            _version        =   65536
            _extentx        =   635
            _extenty        =   635
            _stockprops     =   78
            autosize        =   2
            picture         =   "FRMCONTR.frx":0ADA
         End
         Begin VB.Label Label_std 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Delete All Rosters"
            BeginProperty Font 
               name            =   "Arial"
               charset         =   1
               weight          =   400
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   5
            Left            =   1950
            TabIndex        =   17
            Top             =   2910
            Width           =   1530
            WordWrap        =   -1  'True
         End
         Begin MSMask.MaskEdBox MaskDayLength 
            Height          =   315
            Left            =   1140
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   1470
            Width           =   1275
            _version        =   65536
            _extentx        =   2249
            _extenty        =   556
            _stockprops     =   109
            forecolor       =   -2147483640
            backcolor       =   -2147483643
            borderstyle     =   1
            enabled         =   0   'False
            maxlength       =   5
            format          =   "hh:mm"
         End
         Begin VB.Label Label_std 
            BackStyle       =   0  'Transparent
            Caption         =   "Day Length :"
            BeginProperty Font 
               name            =   "Arial"
               charset         =   1
               weight          =   400
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   120
            TabIndex        =   16
            Top             =   1530
            Width           =   2295
         End
         Begin VB.Label Label_std 
            BackStyle       =   0  'Transparent
            Caption         =   "Roster Increment"
            BeginProperty Font 
               name            =   "Arial"
               charset         =   1
               weight          =   400
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Index           =   1
            Left            =   2550
            TabIndex        =   14
            Top             =   240
            Width           =   1575
         End
         Begin Threed.SSCommand cmdNuke 
            Height          =   600
            Left            =   3600
            TabIndex        =   18
            Top             =   2580
            Width           =   600
            _version        =   65536
            _extentx        =   1058
            _extenty        =   1058
            _stockprops     =   78
            autosize        =   1
            picture         =   "FRMCONTR.frx":0BEC
         End
      End
      Begin VB.Image ImageSwitch 
         Height          =   240
         Index           =   0
         Left            =   2340
         Picture         =   "FRMCONTR.frx":0F06
         Top             =   2880
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image ImageSwitch 
         Height          =   240
         Index           =   1
         Left            =   2580
         Picture         =   "FRMCONTR.frx":1008
         Top             =   2880
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image ImageSwitch 
         Height          =   240
         Index           =   2
         Left            =   2790
         Picture         =   "FRMCONTR.frx":110A
         Top             =   2880
         Visible         =   0   'False
         Width           =   240
      End
      Begin MSGrid.Grid GridClass 
         Height          =   3135
         Left            =   90
         TabIndex        =   0
         Top             =   180
         Width           =   2955
         _version        =   65536
         _extentx        =   5212
         _extenty        =   5530
         _stockprops     =   77
         forecolor       =   -2147483640
         backcolor       =   -2147483643
         BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
            name            =   "Arial"
            charset         =   1
            weight          =   400
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         cols            =   3
         fixedcols       =   0
         scrollbars      =   0
         mouseicon       =   "FRMCONTR.frx":120C
      End
   End
End
Attribute VB_Name = "frmControl"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit
'[----------------------------------------------]
'[frmControl        Control Panel form          ]
'[----------------------------------------------]
'[GSR                                           ]
'[generic staff roster system for Windows       ]
'[(C) ---------- David Gilbert, 1997            ]
'[----------------------------------------------]
'[version                           2.1         ]
'[initial date                      21/07/1997  ]
'[----------------------------------------------]


Private Sub CheckDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Clicking this check box to enable/disable the delete confirmation dialog which accompanies staff and roster deletions."

End Sub

Private Sub cmdNuke_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Clicking this button will permanently delete ALL roster records.  Be careful using this one."

End Sub

Private Sub cmdNukeCover_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Click this button to reveal the 'NUKE' command button."

End Sub

Private Sub cmdToday_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Click this button to place todays date (" & Format(Now, "Short Date") & ") in the Start Date box."

End Sub

Private Sub FrameRosterSetup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "The roster setup area is where you designate most of the parameters used in creating new rosters."

End Sub


Private Sub GridClass_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
    
    Case vbKeyReturn
        '[CAPTURE ENTER KEY PRESSED]
        Call GridClass_DblClick
    Case Else
    
    End Select
    
End Sub

Private Sub GridClass_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    If X > frmControl.GridClass.ColPos(2) Then
        StatusBar "Double-click the icon to activate/deactivate the selected roster.  Only active roster can be modified and processed."
    ElseIf X > frmControl.GridClass.ColPos(1) Then
        StatusBar "Double-click to modify the class description, a 20 character identifier for the selected roster."
    Else
        StatusBar "Double-click to modify the class identifier, a three character short form for the roster name."
    End If
    '[---------------------------------------------------------------------------------]

End Sub

Private Sub Label_std_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    Select Case Index
    Case 0
        StatusBar "Click on the drop down box to selected the starting day for your roster period."
    Case 1
        StatusBar "Click on the drop down box to selected the appropriate roster increment for autobuilding rosters."
    Case 2
        StatusBar "This box will display the day length calculated from the start time and end time."
    Case 3
        StatusBar "Enter the starting date for your rosters here.  This will be used for timesheets, roster printouts and reports."
    Case 4
        StatusBar "Click this button to place todays date (" & Format(Now, "Short Date") & ") in the Start Date box."
    Case 5
        StatusBar "Clicking this button twice (the first click is a safety) will permanently delete ALL roster records."
    Case Else
    End Select
    
End Sub


Private Sub MaskDate_Change()

    '[IF VALID DATE, SAVE TO DSCLASS DYNASET]
    
    If IsDate(frmControl.MaskDate.Text) Then
    
        If CDate(frmControl.MaskDate.Text) = DsDefault("StartDate") Then Exit Sub
    
        DsDefault.Edit
            DsDefault("StartDate") = frmControl.MaskDate.Text
        DsDefault.Update
    
    End If

End Sub


Private Sub cmdToday_Click()

    '[SET MSK DATE TO TODAYS DATE]
    frmControl.MaskDate.Text = Date

End Sub


Private Sub cmdNuke_Click()

    '[***********************************************************************]
    '[WARNING - DELETE ALL ROSTERS CURRENTLY STORED IN THE DATABASE - WARNING]
    '[***********************************************************************]
    '[THIS COMMAND IS NON-REVERSABLE, POPUP YES/NO DIALOG]
    Dim Msg As String
    Dim Style
    Dim Response
    Dim Title
    
    Msg = "Caution - This action will permanently delete and reset all rosters." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "This action is not reversible and should be used with extreme care." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "Do you wish to continue and delete all rosters ?"
    Style = vbYesNo                         ' Define buttons.
    Title = "Confirm Roster Deletion"     ' Define title.
    Response = gsrMsg(Msg, Style, Title)
    If Response = vbYes Then    ' User chose Yes.
        '[DELETE ROSTERS HERE]
        '[REBUILD ROSTER DYNASET WITH ALL RECORDS]
        Dim SQLStmt         As String
        
        '[SELECT APPROPRIATE RECORDS]
        SQLStmt = "SELECT * FROM Roster"
        Set DsRoster = DBMain.OpenRecordset(SQLStmt, dbOpenDynaset)
                
        '[CLEAR CURRENT DYNASET AND PREPARE]
        If DsRoster.EOF And DsRoster.BOF Then
            '[RESIZE NUKECOVER BUTTON]
            frmControl.cmdNukeCover.Visible = True
            frmControl.cmdNuke.Visible = False
            Exit Sub
        End If
        
        DsRoster.MoveLast
        Do While DsRoster.RecordCount > 0
            DsRoster.MoveFirst
            DsRoster.Delete
        Loop
        
        '[MOVE TO FIRST ITEM IN LIST]
        frmRoster.ComboClass.ListIndex = 0
        
    End If
    
    '[RESIZE NUKECOVER BUTTON]
    frmControl.cmdNukeCover.Visible = True
    frmControl.cmdNuke.Visible = False
    '[***********************************************************************]
    '[WARNING - DELETE ALL ROSTERS CURRENTLY STORED IN THE DATABASE - WARNING]
    '[***********************************************************************]
    
End Sub
Private Sub cmdNukeCover_Click()

    '[HIDE NUKE COVER]
    frmControl.cmdNukeCover.Visible = False
    frmControl.cmdNuke.Visible = True
    
End Sub



Private Sub CheckDelete_Click()

    '[CHANGE PUBLIC VARIABLE DELETE CONFIRMATION]
    If frmControl.CheckDelete.Value = 1 Then
        flagDeleteConfirm = True
    Else
        flagDeleteConfirm = False
    End If

End Sub



Private Sub ComboHour_Click()

    ProcessDayLength
    
End Sub


Private Sub ComboIncrement_Click()

    '[TAKE CHANGES MADE BY USER TO INCREMENT AND APPLY TO DEFAULT DYNASET]
    If ComboIncrement.ListIndex <> DsDefault("Increment") - 1 And ComboIncrement.ListIndex > -1 Then
        DsDefault.Edit
            DsDefault("Increment") = ComboIncrement.ListIndex + 1
        DsDefault.Update
    End If

End Sub


Private Sub ComboMinute_Click()

    ProcessDayLength
    
End Sub

Private Sub ComboStartDay_Click()

    '[TAKE CHANGES MADE BY USER TO START DAY AND APPLY TO DEFAULT DYNASET]
    If ComboStartDay.ListIndex <> DsDefault("StartDay") - 1 And ComboStartDay.ListIndex > -1 Then
        DsDefault.Edit
            DsDefault("StartDay") = ComboStartDay.ListIndex + 1
        DsDefault.Update
    End If
    
    '[SET GRID TITLES ON ROSTER GRID]
    SetGridTitles
    
    '[APPLY DAY LABELS]
    SetDayLabels
    
End Sub


Private Sub Form_Load()

    '[Resize controls and grids to match]
    frmControl.GridClass.ColWidth(2) = frmControl.ImageSwitch(0).Width
    frmControl.GridClass.ColWidth(0) = (GridClass.Width - frmControl.GridClass.ColWidth(2)) * 0.25
    frmControl.GridClass.ColWidth(1) = (GridClass.Width - frmControl.GridClass.ColWidth(2)) * 0.75

    
    '[CALL ROUTINE TO FILL GRID]
    FillClassGrid
    
    '[PLACE DEFAULT VALUES IN COMBOBOXES AND MASK BOXES]
    ComboStartDay.ListIndex = DsDefault("StartDay") - 1
    ComboIncrement.ListIndex = DsDefault("Increment") - 1
    
    '[SET START HOUR AND MINUTE]
    ComboHour = Format(Hour(DsDefault("StartTime")), "0#")
    ComboMinute = Format(Minute(DsDefault("StartTime")), "0#")
    ProcessDayLength
       
    '[SET FORM DATE]
    If IsNull(DsDefault("StartDate")) Then
        frmControl.MaskDate = Format(Now, "Short Date")
    Else
        frmControl.MaskDate = DsDefault("StartDate")
    End If
       
   
End Sub


Private Sub GridClass_DblClick()

    '[USER DOUBLE CLICKED ON CLASS GRID, POPUP INPUT BOX FOR NEW VALUE]
    Dim strTemp             As String   '[Temporary Variable for Cell Content]
    Dim strMessage          As String
    Dim strTitle            As String
    Dim strOldClass         As String
    Dim SQLStmt             As String
    Dim intCounter          As Integer
    Dim intCol, intRow      As Integer
    
    '[SAVE CURRENT LOCATION]
    intCol = GridClass.Col
    intRow = GridClass.Row
    
    '[EXIT IF TITLE ROW CLICKED]
    If intRow = 0 Then Exit Sub
    
    Select Case GridClass.Col
    Case 0
        strMessage = "The class code is used as a short form to identify which classes an employee belong to.  You are allowed to use up to three alpha-numeric characters to define this class."
    Case 1
        strMessage = "The class description is used on reports to provide a longer identifier for each class.  You may use up to 20 alpha-numeric characters in this field."
    End Select
    
    strTitle = "Class Definitions"
    strOldClass = GridClass.Text
    strTemp = GridClass.Text
    
    If GridClass.Col = 2 Then
        '[ENABLED/DISABLED]
        If GridClass.Text = vbChecked Then
            GridClass.Text = vbUnchecked
            frmControl.GridClass.Picture = frmControl.ImageSwitch(constCritical).Picture
        Else
            GridClass.Text = vbChecked
            frmControl.GridClass.Picture = frmControl.ImageSwitch(constWarning).Picture
        End If
    Else
        '[GET USER INPUT]
        strTemp = InputBox(strMessage, strTitle, strTemp)
        '[PROCESS USER INPUT]
        If Not strTemp = "" Or strTemp = GridClass.Text Then
            If GridClass.Col = 0 Then
            '[CHECK FOR EXISTANCE OF NEW CODE]
                SQLStmt = "Code = '" & strTemp & "'"
                DsClass.FindFirst SQLStmt
                If Not DsClass.NoMatch Then
                    '[RESTORE LOCATION IN GRID]
                    GridClass.Col = intCol
                    GridClass.Row = intRow
                    Exit Sub
                End If
                If Len(strTemp) > 3 Then GridClass.Text = Left$(UCase$(strTemp), 3) Else GridClass.Text = UCase$(strTemp)
            Else
                '[CHECK FOR EXISTANCE OF NEW CODE]
                SQLStmt = "Description = '" & strTemp & "'"
                DsClass.FindFirst SQLStmt
                If Not DsClass.NoMatch Then
                    '[RESTORE LOCATION IN GRID]
                    GridClass.Col = intCol
                    GridClass.Row = intRow
                    Exit Sub
                End If
                If Len(strTemp) > 20 Then GridClass.Text = Left$(strTemp, 3) Else GridClass.Text = strTemp
            End If
        Else
            '[RESTORE LOCATION IN GRID]
            GridClass.Col = intCol
            GridClass.Row = intRow
            Exit Sub
        End If
    End If
    
    '[PLACE VALUES IN DYNASET]
    DsClass.MoveFirst
    For intCounter = 1 To 10
        frmControl.GridClass.Row = intCounter
        frmControl.GridClass.Col = 0
        DsClass.Edit
            DsClass("Code") = frmControl.GridClass.Text
            frmControl.GridClass.Col = 1
            DsClass("Description") = frmControl.GridClass.Text
            frmControl.GridClass.Col = 2
            DsClass("Active") = frmControl.GridClass.Text
        DsClass.Update
        DsClass.MoveNext
    Next intCounter
    
    '[SET CLASS LABELS ON STAFF FORM]
    SetClassLabels

    '[RESTORE LOCATION IN GRID]
    GridClass.Col = intCol
    GridClass.Row = intRow
    
End Sub




Private Sub OptionTIme_Click(Index As Integer)

    Select Case Index
    Case 0
        ComboHour = Format(Hour(DsDefault("StartTime")), "0#")
        ComboMinute = Format(Minute(DsDefault("StartTime")), "0#")
    Case 1
        ComboHour = Format(Hour(DsDefault("EndTime")), "0#")
        ComboMinute = Format(Minute(DsDefault("EndTime")), "0#")
    End Select

End Sub


Private Sub OptionTIme_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    Select Case Index
    Case 0
        StatusBar "Click here to switch the drop down boxes to the left to display/modify the roster START time."
    Case 1
        StatusBar "Click here to switch the drop down boxes to the left to display/modify the roster START time."
    Case Else
    End Select

End Sub


