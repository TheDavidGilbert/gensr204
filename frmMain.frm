VERSION 4.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control"
   ClientHeight    =   3810
   ClientLeft      =   3840
   ClientTop       =   2040
   ClientWidth     =   7200
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   Height          =   4215
   Icon            =   "FRMMAIN.frx":0000
   Left            =   3780
   LinkTopic       =   "frmClass"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   Top             =   1695
   Width           =   7320
   Begin VB.Frame FrameRosterSetup 
      Caption         =   "Roster Setup"
      Height          =   1875
      Left            =   2670
      TabIndex        =   3
      Top             =   0
      Width           =   4485
      Begin VB.ComboBox ComboMinute 
         Height          =   315
         ItemData        =   "FRMMAIN.frx":000C
         Left            =   1380
         List            =   "FRMMAIN.frx":001C
         TabIndex        =   11
         Text            =   "ComboMinute"
         Top             =   1080
         Width           =   1035
      End
      Begin VB.ComboBox ComboHour 
         Height          =   315
         ItemData        =   "FRMMAIN.frx":0030
         Left            =   90
         List            =   "FRMMAIN.frx":008A
         TabIndex        =   10
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
         ItemData        =   "FRMMAIN.frx":00EE
         Left            =   90
         List            =   "FRMMAIN.frx":0107
         Style           =   2  'Dropdown List
         TabIndex        =   7
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
         ItemData        =   "FRMMAIN.frx":014B
         Left            =   2550
         List            =   "FRMMAIN.frx":017E
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   480
         Width           =   1605
      End
      Begin VB.OptionButton OptionTIme 
         Caption         =   "End Time"
         Height          =   315
         Index           =   1
         Left            =   2700
         TabIndex        =   5
         Top             =   1440
         Width           =   1005
      End
      Begin VB.OptionButton OptionTIme 
         Caption         =   "Start Time"
         Height          =   315
         Index           =   0
         Left            =   2700
         TabIndex        =   4
         Top             =   1080
         Value           =   -1  'True
         Width           =   1005
      End
      Begin MSMask.MaskEdBox MaskDayLength 
         Height          =   315
         Left            =   1140
         TabIndex        =   13
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
         AutoSize        =   -1  'True
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
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   1530
         Width           =   915
      End
      Begin VB.Label Label_std 
         AutoSize        =   -1  'True
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
         Height          =   210
         Index           =   15
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1530
      End
      Begin VB.Label Label_std 
         AutoSize        =   -1  'True
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
         Height          =   210
         Index           =   0
         Left            =   2550
         TabIndex        =   8
         Top             =   240
         Width           =   1230
      End
   End
   Begin VB.TextBox TextPath 
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
      Left            =   2730
      TabIndex        =   0
      Top             =   3390
      Width           =   4425
   End
   Begin MSGrid.Grid GridClass 
      Height          =   3675
      Left            =   90
      TabIndex        =   2
      Top             =   60
      Width           =   2475
      _version        =   65536
      _extentx        =   4366
      _extenty        =   6482
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
      fixedcols       =   0
      scrollbars      =   0
      mouseicon       =   "FRMMAIN.frx":0200
   End
   Begin VB.Label Label_std 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Program Path"
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
      Left            =   2760
      TabIndex        =   1
      Top             =   3090
      Width           =   960
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit


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
    
End Sub


Private Sub Form_Load()

    '[Resize controls and grids to match]
    GridClass.ColWidth(0) = GridClass.Width * 0.25
    GridClass.ColWidth(1) = GridClass.Width * 0.75
    
    '[CALL ROUTINE TO FILL GRID]
    FillClassGrid
    
    '[PLACE DEFAULT VALUES IN COMBOBOXES AND MASK BOXES]
    TextPath.Text = DsDefault("Path")
    ComboStartDay.ListIndex = DsDefault("StartDay") - 1
    ComboIncrement.ListIndex = DsDefault("Increment") - 1
    
    '[SET START HOUR AND MINUTE]
    ComboHour = Format(Hour(DsDefault("StartTime")), "0#")
    ComboMinute = Format(Minute(DsDefault("StartTime")), "0#")
    ProcessDayLength
       
   
End Sub


Private Sub GridClass_DblClick()

    '[USER DOUBLE CLICKED ON CLASS GRID, POPUP INPUT BOX FOR NEW VALUE]
    Dim strTemp             As String   '[Temporary Variable for Cell Content]
    Dim strMessage          As String
    Dim strTitle            As String
    Dim strOldClass         As String
    Dim SQLStmt             As String
    Dim intcounter          As Integer
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
    
    '[GET USER INPUT]
    strTemp = InputBox(strMessage, strTitle, strTemp)
    
    '[PROCESS USER INPUT]
    If Not strTemp = "" Or strTemp = GridClass.Text Then
        
        If GridClass.Col = 0 Then
        '[CHECK FOR EXISTANCE OF NEW CODE]
            SQLStmt = "Code = '" & strTemp & "'"
            DsClass.FindFirst SQLStmt
            If Not DsClass.NoMatch Then Exit Sub
            If Len(strTemp) > 3 Then GridClass.Text = Left$(UCase$(strTemp), 3) Else GridClass.Text = UCase$(strTemp)
        Else
            If Len(strTemp) > 20 Then GridClass.Text = Left$(strTemp, 3) Else GridClass.Text = strTemp
        End If
        
    
        '[PLACE VALUES IN DYNASET]
        DsClass.MoveFirst
        For intcounter = 1 To 10
            frmMain.GridClass.Row = intcounter
            frmMain.GridClass.Col = 0
            DsClass.Edit
                DsClass("Code") = frmMain.GridClass.Text
                frmMain.GridClass.Col = 1
                DsClass("Description") = frmMain.GridClass.Text
            DsClass.Update
            DsClass.MoveNext
        Next intcounter
    
        '[SET CLASS LABELS ON STAFF FORM]
        SetClassLabels
    
    End If

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

Private Sub TextPath_Change()
    
    '[TAKE CHANGES MADE BY USER AND APPLY TO THE DEFAULT DYNASET]
    If TextPath.Text <> DsDefault("Path") Then
        DsDefault.Edit
            DsDefault("Path") = TextPath.Text
        DsDefault.Update
    End If

End Sub


