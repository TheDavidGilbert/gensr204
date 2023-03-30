VERSION 4.00
Begin VB.Form frmReport 
   Caption         =   "Report"
   ClientHeight    =   5235
   ClientLeft      =   2910
   ClientTop       =   3225
   ClientWidth     =   8400
   ClipControls    =   0   'False
   BeginProperty Font 
      name            =   "Arial"
      charset         =   1
      weight          =   400
      size            =   8.25
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   Height          =   5640
   Icon            =   "FRMREPOR.frx":0000
   Left            =   2850
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   8400
   Top             =   2880
   Width           =   8520
   Begin Threed.SSPanel PanelToolBar 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   8400
      _version        =   65536
      _extentx        =   14817
      _extenty        =   661
      _stockprops     =   15
      forecolor       =   -2147483641
      backcolor       =   -2147483644
      bevelouter      =   0
      floodtype       =   1
      floodcolor      =   -2147483646
      floodshowpct    =   0   'False
      alignment       =   0
      autosize        =   2
      mouseicon       =   "FRMREPOR.frx":08CA
      Begin VB.Label LabelTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report Title"
         BeginProperty Font 
            name            =   "Arial"
            charset         =   1
            weight          =   700
            size            =   12
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   810
         TabIndex        =   4
         Top             =   60
         Width           =   1320
      End
      Begin Threed.SSCommand cmdInfo 
         Height          =   360
         Left            =   0
         TabIndex        =   3
         Top             =   15
         Width           =   360
         _version        =   65536
         _extentx        =   635
         _extenty        =   635
         _stockprops     =   78
         autosize        =   2
         picture         =   "FRMREPOR.frx":11A4
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   360
         Left            =   390
         TabIndex        =   2
         Top             =   15
         Width           =   360
         _version        =   65536
         _extentx        =   635
         _extenty        =   635
         _stockprops     =   78
         autosize        =   2
         picture         =   "FRMREPOR.frx":12B6
      End
      Begin VB.Image ImageWarning 
         Height          =   240
         Index           =   2
         Left            =   3510
         Picture         =   "FRMREPOR.frx":13C8
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image ImageWarning 
         Height          =   240
         Index           =   1
         Left            =   3300
         Picture         =   "FRMREPOR.frx":14CA
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image ImageWarning 
         Height          =   240
         Index           =   0
         Left            =   3060
         Picture         =   "FRMREPOR.frx":15CC
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin MSGrid.Grid GridReport 
      Height          =   4815
      Left            =   30
      TabIndex        =   1
      Top             =   390
      Width           =   8355
      _version        =   65536
      _extentx        =   14737
      _extenty        =   8493
      _stockprops     =   77
      forecolor       =   0
      backcolor       =   16777215
      cols            =   6
      fixedcols       =   0
      mouseicon       =   "FRMREPOR.frx":16CE
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit
'[----------------------------------------------]
'[frmReport         Exception report form       ]
'[----------------------------------------------]
'[GSR                                           ]
'[generic staff roster system for Windows       ]
'[(C) -------- & David Gilbert, 1997            ]
'[----------------------------------------------]
'[version                           2.1         ]
'[initial date                      21/07/1997  ]
'[----------------------------------------------]

Private Sub cmdDelete_Click()

    '[DELETE SELECTED ROW FROM REPORT GRID]
    Dim intRow          As Integer
    
    If frmReport.GridReport.Row = 0 Or frmReport.GridReport.Rows = 2 Then Exit Sub
    
    '[ASSIGN ROW]
    intRow = frmReport.GridReport.Row
    frmReport.GridReport.RemoveItem intRow
    
    '[SET NEW ROW]
    intRow = intRow - 1
    
    If intRow = 0 Then intRow = 1
    If intRow > (frmReport.GridReport.Rows - 1) Then intRow = (frmReport.GridReport.Rows - 1)

End Sub


Private Sub cmdDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Click here to delete the currently selected row from the report grid."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdInfo_Click()

    '[SHOW HELP INFORMATION FOR THE SELECTED WARNING]
    Dim intLevel                As Integer
    Dim strClassDescription     As String
    Dim strDayName              As String
    Dim intStartDay             As Integer
    Dim intCounter              As Integer
    Dim strFullname             As String
    Dim strTime                 As String
    '[POPUP YES/NO DIALOG TO CONFIRM DELETION]
    Dim Msg As String
    Dim Style
    Dim Response
    Dim Title
    
    '[EXIT THE ROUTINE IF THE USER CLICKED ON THE FIRST ROW]
    If frmReport.GridReport.Text = "" Or frmReport.GridReport.Row = 0 Then Exit Sub
    '[  0    ROSTER     DAY     TIME     STAFF     WARNING]
    '[  0    1          2       3        4         5      ]
    '[MOVE TO THE FIRST COLUMN AND GET THE WARNING LEVEL]
    frmReport.GridReport.Col = 0
    intLevel = Val(frmReport.GridReport.Text)
    If frmReport.LabelTitle = "GSR Staff Report" Then intLevel = 4
        
    '[PROCESS WARNING LEVEL]
    Select Case intLevel
    Case 0  '[HOURS+/HOURS- FOR A PARTICULAR STAFF MEMBER]
        Msg = "This exception has been reported because the hours allocated for this staff member are not within the limits set on the staff form." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "You may have to reduce the hours allocated on the rosters or alternatively adjust the hour limits on the staff form." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "Double-click the entry to move to the appropriate staff record."
        Style = vbOKOnly                    ' Define buttons.
        Title = "About Required Hours"      ' Define title.
        
    Case 1  '[STAFF MEMBER NOT IN CLASS]
        Msg = "This exception has been reported for this staff member because they do not belong to this class." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "You can replace the staff member in the roster or alternatively you may make this staff member available for this roster by clicking the appropriate check box on the staff form." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "Double-click the entry to move to the appropriate roster."
        Style = vbOKOnly                    ' Define buttons.
        Title = "About Staff Not In Class"      ' Define title.
    
    Case 2  '[UNAVAILABLE STAFF MEMBER IN ROSTER]
        Msg = "This exception has been reported because the staff member you have allocated to this roster is marked as being unavailable on this day." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "You may have to remove this staff member from the roster or change the availability of the staff member on the staff form." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "Double-click the entry to move to the appropriate roster."
        Style = vbOKOnly                    ' Define buttons.
        Title = "About Unavailable Staff Member"      ' Define title.
    
    Case 3  '[CONFLICT IN ROSTER]
        Msg = "This exception has been reported because this staff member has been allocated to two or more rosters during the same roster period." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "You will have to remove this staff member from one of the conflicting rosters." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "Double-click the entry to move to the appropriate roster."
        Style = vbOKOnly                    ' Define buttons.
        Title = "About Conflict In Roster"  ' Define title.
    
    Case 4 '[STAFF REPORT]
        Msg = "The staff report lists hours and currency amounts for all staff members who appear in active rosters." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "It then summarizes and totals these values for all rosters." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "If a staff member has not been assigned an hourly rate (on the staff form), only hours assigned will be reported."
        Style = vbOKOnly                    ' Define buttons.
        Title = "About Staff Report"      ' Define title.
        
    Case 5  '[STAFF MEMBER IN ROSTER NOT FOUND IN STAFF LIST]
        Msg = "This exception has been reported because a staff member who no longer appears in the staff list has been allocated to the roster." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "You may have to remove this staff member from the roster." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "Double-click the entry to move to the appropriate roster."
        Style = vbOKOnly                    ' Define buttons.
        Title = "About Name Not Found"      ' Define title.
    
    Case Else
    End Select
     
    '[CALL DISPLAY WINDOW]
    If intLevel <= 5 Then Response = gsrMsg(Msg, Style, Title)

End Sub

Private Sub cmdInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Click here to see more information about the highlighted report item."
    '[---------------------------------------------------------------------------------]


End Sub


Private Sub Form_Resize()

    '[CALL SUBROUTINE TO RESIZE FORM]
    Call resReportForm
    
End Sub


Private Sub GridReport_DblClick()

    '[GOTO THE SOURCE OF THE PROBLEM - MAY REQUIRE REPOSITIONING OF SOME DYNASETS AND LISTS]
    Dim intLevel                As Integer
    Dim strClassDescription     As String
    Dim strDayName              As String
    Dim intStartDay             As Integer
    Dim intCounter              As Integer
    Dim strFullname             As String
    Dim strTime                 As String
    Dim intCol                  As Integer
    Dim intRow                  As Integer
    
    '[MOVE TO THE FIRST COLUMN AND GET THE WARNING LEVEL]
    frmReport.GridReport.Col = 0
    '[EXIT THE ROUTINE IF THE USER CLICKED ON THE FIRST ROW]
    If frmReport.GridReport.Text = "" Or frmReport.GridReport.Row = 0 Or Not IsNumeric(frmReport.GridReport.Text) Then Exit Sub
    
    intLevel = frmReport.GridReport.Text
    
    '[  0    ROSTER     DAY     TIME     STAFF     WARNING]
    '[  0    1          2       3        4         5      ]
        
    '[PROCESS WARNING LEVEL]
    Select Case intLevel
    Case 0  '[HOURS+/HOURS- FOR A PARTICULAR STAFF MEMBER]
        frmReport.GridReport.Col = 4
        strFullname = frmReport.GridReport.Text
        '[RELOCATE STAFF NAME]
        For intCounter = 0 To (frmStaff.ListStaff.ListCount - 1)
            If frmStaff.ListStaff.List(intCounter) = strFullname Then frmStaff.ListStaff.ListIndex = intCounter
        Next intCounter
        '[MOVE TO STAFF FORM]
        If frmStaff.Visible = False Then frmStaff.Visible = True
        frmStaff.ZOrder
        '[IF DAY CHECK FRAME IS NOT VISIBLE THEN MAKE IT VISIBLE]
        If frmStaff.FrameDays.Visible = False Then
            frmStaff.FrameDays.Visible = True
            frmStaff.FrameClass.Visible = False
        End If
        
    Case 1, 2, 3, 5 '[INNAPROPRIATE STAFF MEMBER IN FORM] OR [UNAVAILABLE STAFF MEMBER IN ROSTER] OR [ROSTER CONFLICT] OR [NAME NOT IN STAFF LIST]
        frmReport.GridReport.Col = 1
        strClassDescription = frmReport.GridReport.Text
        frmReport.GridReport.Col = 2
        strDayName = frmReport.GridReport.Text
        frmReport.GridReport.Col = 3
        strTime = frmReport.GridReport.Text

        '[FIND APPROPRIATE ROSTER]
        If frmRoster.Visible = False Then frmRoster.Visible = True
        frmRoster.ComboClass.Text = strClassDescription
        
        '[FIND APPROPRIATE COLUMN]
        frmRoster.GridRoster.Row = 0
        For intCounter = 2 To 8
            frmRoster.GridRoster.Col = intCounter
            If frmRoster.GridRoster.Text = strDayName Then intCol = intCounter
        Next intCounter
        
        '[FIND APPROPRIATE ROW]
        frmRoster.GridRoster.Col = 0
        For intCounter = 1 To (frmRoster.GridRoster.Rows - 1)
            frmRoster.GridRoster.Row = intCounter
            If frmRoster.GridRoster.Text = strTime Then intRow = intCounter
        Next intCounter
        
        '[SELECT/SHOW CURRENT ROW AND COL]
        frmRoster.GridRoster.TopRow = 1

        frmRoster.GridRoster.Row = intRow
        frmRoster.GridRoster.Col = intCol
        frmRoster.GridRoster.LeftCol = 0
        frmRoster.GridRoster.TopRow = 1
        Do While Not frmRoster.GridRoster.RowIsVisible(intRow)
            frmRoster.GridRoster.TopRow = frmRoster.GridRoster.TopRow + 1
        Loop
        Do While Not frmRoster.GridRoster.ColIsVisible(intCol)
            frmRoster.GridRoster.LeftCol = frmRoster.GridRoster.LeftCol + 1
        Loop
        
        '[SELECT CELL]
        frmRoster.GridRoster.SelStartRow = intRow
        frmRoster.GridRoster.SelStartCol = intCol
        frmRoster.GridRoster.SelEndRow = intRow
        frmRoster.GridRoster.SelEndCol = intCol
        
        '[MOVE TO FORM]
        
        frmRoster.ZOrder
        
    Case Else
    End Select
    
End Sub

Private Sub GridReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Click the printer icon on the main toolbar to print this report."
    '[---------------------------------------------------------------------------------]

End Sub


