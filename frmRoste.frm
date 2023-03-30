VERSION 4.00
Begin VB.Form frmRoster 
   BackColor       =   &H80000004&
   Caption         =   "Roster"
   ClientHeight    =   4935
   ClientLeft      =   1965
   ClientTop       =   6105
   ClientWidth     =   8400
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      name            =   "Arial"
      charset         =   1
      weight          =   400
      size            =   8.25
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   Height          =   5340
   Icon            =   "FRMROSTE.frx":0000
   Left            =   1905
   LinkTopic       =   "frmRoster"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4935
   ScaleWidth      =   8400
   Top             =   5760
   Width           =   8520
   Begin VB.ComboBox ComboClass 
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   30
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   420
      Width           =   2295
   End
   Begin VB.ListBox ListStaff 
      DragIcon        =   "FRMROSTE.frx":08CA
      Height          =   4050
      Left            =   30
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   810
      Width           =   2295
   End
   Begin Threed.SSPanel PanelToolBar 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   12
      Top             =   0
      Width           =   8400
      _version        =   65536
      _extentx        =   14817
      _extenty        =   661
      _stockprops     =   15
      forecolor       =   -2147483641
      bevelouter      =   0
      floodtype       =   1
      floodcolor      =   -2147483646
      floodshowpct    =   0   'False
      alignment       =   0
      autosize        =   2
      mouseicon       =   "FRMROSTE.frx":1194
      Begin GaugeLib.Gauge GaugeProgress 
         Height          =   300
         Left            =   5625
         TabIndex        =   13
         Top             =   60
         Width           =   2670
         _version        =   65536
         _extentx        =   4710
         _extenty        =   529
         _stockprops     =   73
         forecolor       =   255
         backcolor       =   -2147483633
         innertop        =   0
         innerleft       =   0
         innerright      =   0
         innerbottom     =   0
         value           =   50
         needlewidth     =   2
      End
      Begin Threed.SSCommand cmdNote 
         Height          =   360
         Left            =   2340
         TabIndex        =   15
         Top             =   15
         Width           =   360
         _version        =   65536
         _extentx        =   635
         _extenty        =   635
         _stockprops     =   78
         autosize        =   2
         picture         =   "FRMROSTE.frx":1A6E
      End
      Begin Threed.SSCommand cmdPaste 
         Height          =   360
         Left            =   1950
         TabIndex        =   8
         Top             =   15
         Width           =   360
         _version        =   65536
         _extentx        =   635
         _extenty        =   635
         _stockprops     =   78
         autosize        =   2
         picture         =   "FRMROSTE.frx":1B80
      End
      Begin Threed.SSCommand cmdCopy 
         Height          =   360
         Left            =   1560
         TabIndex        =   7
         Top             =   15
         Width           =   360
         _version        =   65536
         _extentx        =   635
         _extenty        =   635
         _stockprops     =   78
         autosize        =   2
         picture         =   "FRMROSTE.frx":1C92
      End
      Begin VB.Label LabelInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Height          =   300
         Left            =   3930
         TabIndex        =   14
         Top             =   60
         Width           =   1665
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   360
         Left            =   780
         TabIndex        =   5
         Top             =   15
         Width           =   360
         _version        =   65536
         _extentx        =   635
         _extenty        =   635
         _stockprops     =   78
         autosize        =   2
         picture         =   "FRMROSTE.frx":1DA4
      End
      Begin Threed.SSCommand cmdRemove 
         Height          =   360
         Left            =   0
         TabIndex        =   3
         Top             =   15
         Width           =   360
         _version        =   65536
         _extentx        =   635
         _extenty        =   635
         _stockprops     =   78
         BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
            name            =   "Arial"
            charset         =   1
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         font3d          =   2
         autosize        =   2
         picture         =   "FRMROSTE.frx":1EB6
      End
      Begin Threed.SSCommand cmdTransfer 
         Height          =   360
         Left            =   390
         TabIndex        =   4
         Top             =   15
         Width           =   360
         _version        =   65536
         _extentx        =   635
         _extenty        =   635
         _stockprops     =   78
         BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
            name            =   "Arial"
            charset         =   1
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         font3d          =   2
         autosize        =   2
         picture         =   "FRMROSTE.frx":1FC8
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   360
         Left            =   1170
         TabIndex        =   6
         Top             =   15
         Visible         =   0   'False
         Width           =   360
         _version        =   65536
         _extentx        =   635
         _extenty        =   635
         _stockprops     =   78
         BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   400
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         autosize        =   2
         picture         =   "FRMROSTE.frx":20DA
      End
      Begin Threed.SSCommand cmdRebuild 
         Height          =   360
         Left            =   2730
         TabIndex        =   9
         Top             =   15
         Width           =   360
         _version        =   65536
         _extentx        =   635
         _extenty        =   635
         _stockprops     =   78
         autosize        =   2
         picture         =   "FRMROSTE.frx":21EC
      End
   End
   Begin Threed.SSCommand cmdRemoveRow 
      Height          =   255
      Left            =   2340
      TabIndex        =   11
      Top             =   1080
      Width           =   255
      _version        =   65536
      _extentx        =   450
      _extenty        =   450
      _stockprops     =   78
      forecolor       =   16777215
      bevelwidth      =   0
      autosize        =   2
      picture         =   "FRMROSTE.frx":22FE
   End
   Begin Threed.SSCommand cmdInsertRow 
      Height          =   255
      Left            =   2340
      TabIndex        =   10
      Top             =   810
      Width           =   255
      _version        =   65536
      _extentx        =   450
      _extenty        =   450
      _stockprops     =   78
      forecolor       =   16777215
      bevelwidth      =   0
      autosize        =   2
      picture         =   "FRMROSTE.frx":23F8
   End
   Begin MSGrid.Grid GridRoster 
      Height          =   4455
      Left            =   2640
      TabIndex        =   2
      Top             =   420
      Width           =   5685
      _version        =   65536
      _extentx        =   10028
      _extenty        =   7858
      _stockprops     =   77
      backcolor       =   16777215
      rows            =   24
      cols            =   9
      fixedcols       =   0
      mouseicon       =   "FRMROSTE.frx":24F2
   End
End
Attribute VB_Name = "frmRoster"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit
'[----------------------------------------------]
'[frmRoster         Roster Creation form        ]
'[----------------------------------------------]
'[GSR                                           ]
'[generic staff roster system for Windows       ]
'[(C) ---------- David Gilbert, 1997            ]
'[----------------------------------------------]
'[version                           2.1         ]
'[initial date                      21/07/1997  ]
'[----------------------------------------------]

Private Sub cmdCopy_Click()

    '[ADD CLIP AREA OF ROSTER TO strCLIP]
    strClip = frmRoster.GridRoster.Clip

End Sub


Private Sub cmdCopy_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Click here to copy selected roster cells to the clipboard."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdDelete_Click()

    '[CLEAR SELECTED CELLS IN ROSTER]
    Dim intCounter          As Integer
    Dim intCol              As Integer
    Dim intRow              As Integer

    '[POPUP YES/NO DIALOG TO CONFIRM DELETION]
    Dim Msg As String
    Dim Style
    Dim Response
    Dim Title
    
    If frmRoster.GridRoster.Col <= 1 Or frmRoster.GridRoster.Row = 0 Then Exit Sub
    
    '[CHECK DELETION FLAG]
    If flagDeleteConfirm Then
        Msg = "This action will delete the contents of the selected cell/cells." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "Do you wish to continue ?"
        Style = vbYesNo                     ' Define buttons.
        Title = "Confirmation Required"     ' Define title.
        Response = gsrMsg(Msg, Style, Title)
    Else
        Response = vbYes
    End If
    
    If Response = vbYes Then    ' User chose Yes.

        '[DETERMINE WHETHER IT IS A MULTISELECT OR SINGLE CELL FILL]
        If frmRoster.GridRoster.SelStartCol = -1 Or frmRoster.GridRoster.SelEndCol = -1 Then
            '[SINGLE CELL FILL]
            '[CLEAR CELL CONTENTS]
            frmRoster.GridRoster.Text = ""
        Else
            '[MULTI CELL FILL]
            For intCol = frmRoster.GridRoster.SelStartCol To frmRoster.GridRoster.SelEndCol
                frmRoster.GridRoster.Col = intCol
                For intRow = frmRoster.GridRoster.SelStartRow To frmRoster.GridRoster.SelEndRow
                    frmRoster.GridRoster.Row = intRow
                    '[CLEAR CELL CONTENTS]
                    If frmRoster.GridRoster.Col <= 1 Or frmRoster.GridRoster.Row = 0 Then Exit For
                    frmRoster.GridRoster.Text = ""
                Next intRow
            Next intCol
        End If
        
        '[MAKE SAVE BUTTON VISIBLE]
        frmRoster.cmdSave.Visible = True
    
    End If
    
End Sub


Private Sub cmdDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Click here to delete the selected cell contents from the roster."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdInsertRow_Click()

    '[INSERT ROW INTO ROSTER]
    Dim intRow As Integer
    If frmRoster.GridRoster.Row = 0 Then
        '[SET ROW TO FIRST ROW]
        intRow = 1
    Else
        intRow = frmRoster.GridRoster.Row + 1
    End If
    
    '[ADD NEW ROW ITEM TO THE GRID]
    frmRoster.GridRoster.AddItem Format(DsDefault("StartTime"), "Medium Time") & Chr$(vbKeyTab) & Format(DsDefault("EndTime"), "Medium Time"), intRow
    '[OLD TIME FORMAT] -> "hh:mm AMPM"
    
    '[MOVE TO THIS NEW ROW]
    frmRoster.GridRoster.Row = intRow
    
    '[MAKE SAVE BUTTON VISIBLE]
    frmRoster.cmdSave.Visible = True

End Sub

Private Sub cmdInsertRow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Click here to insert a new row into the roster grid at the highlighted position."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdNote_Click()

    '[PLACE ROSTER TEXT IN THE TEXT BOX ON THE MESSAGE FORM AND ALLOW CHANGES]
    Dim flagResult      As Boolean
    
    flagResult = gsrMsg(DsClass("Note"), vbQuestion, "Enter " & frmRoster.ComboClass.Text & " Roster Notes")
    Select Case flagResult
    Case vbOK
        DsClass.Edit
            DsClass("Note") = Trim(gsrNote & " ")
        DsClass.Update
    Case Else
    End Select
    
End Sub

Private Sub cmdNote_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Click here to edit notes for this roster."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdPaste_Click()
    
    '[PUT strClip IN ROSTER]
    If Len(strClip) = 0 Then Exit Sub
    frmRoster.GridRoster.Clip = strClip
    '[MAKE SAVE BUTTON VISIBLE]
    frmRoster.cmdSave.Visible = True
    
End Sub

Private Sub cmdPaste_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Click here to paste the contents of the clipboard into the selected cells in the roster."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdRebuild_Click()

    '[IF CHANGES HAVE BEEN MADE AND NOT SAVED, CHECK WITH USER]
    '[CHECK FOR CHANGED DATA AND NOTIFY]
    Dim Msg As String
    Dim Style
    Dim Response
    Dim Title

    If frmRoster.cmdSave.Visible = True Then
        '[ROSTER IS UNSAVED, POPUP YES/NO DIALOG]
        Msg = "You have made changes to this roster (" & frmRoster.ComboClass.Text & ") but have not saved these changes." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "If you choose not to save now, any changes you have made since your last save will be lost." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "Do you wish to save these changes before you rebuild this roster ?"
        Style = vbYesNoCancel ' Define buttons.
        Title = "Roster Changes Not Saved"  ' Define title.
        Response = gsrMsg(Msg, Style, Title)
        If Response = vbYes Then    ' User chose Yes.
            SaveRosterGrid
            '[MAKE SAVE BUTTON VISIBLE]
            frmRoster.cmdSave.Visible = True
        ElseIf Response = vbCancel Then
            Exit Sub
        End If
    Else
        '[REBUILDING ROSTER, POPUP YES/NO DIALOG]
        Msg = "Caution - This action will rebuild this roster using the settings on the Control Form - start time, end time and increment." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "This will result in this roster being cleared of all data currently entered." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "Do you wish to continue and rebuild this roster ?"
        Style = vbYesNo             ' Define buttons.
        Title = "Rebuilding Roster"  ' Define title.
        Response = gsrMsg(Msg, Style, Title)
        If Response = vbNo Then    ' User chose No.
            Exit Sub
        End If
    End If

    '[REBUILD ROSTER]
    RebuildRoster
    
    '[MOVE TO ROSTER FORM IF IT IS VISIBLE]
    If frmRoster.WindowState = 1 Then frmRoster.WindowState = 0
    frmRoster.ZOrder
    
    '[MAKE SAVE BUTTON VISIBLE]
    frmRoster.cmdSave.Visible = True

End Sub



Private Sub cmdRebuild_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Click here to rebuild this roster using the defaults specified on the control form."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdRemove_Click()

    '[REMOVE ALL SELECTED STAFF MEMBERS FROM THE SELECTED CELL/CELLS IN THE ROSTER GRID]
    Dim intCounter          As Integer
    Dim intTrCount          As Integer
    Dim intSelCount         As Integer
    
    '[SELECTION COUNT - ITEMS SELECTED]
    intSelCount = frmRoster.ListStaff.SelCount
    If intSelCount = 0 Then Exit Sub
    intTrCount = 1
    
    '[ROSTER LABEL INFO]
    Call RosterInfo("Removing from Roster", 0)
        
    For intCounter = 0 To (frmRoster.ListStaff.ListCount - 1)
        If frmRoster.ListStaff.Selected(intCounter) Then
            '[ITEM IS SELECTED SO TRANSFER TO ROSTER]
            RemoveFromRoster (frmRoster.ListStaff.List(intCounter))
            '[TURN SAVE BUTTON ON]
            frmRoster.cmdSave.Visible = True
            '[PROGRESS BAR]
            ProgressBar ((intTrCount / intSelCount) * 100)
            intTrCount = intTrCount + 1
        End If
    Next intCounter

    '[ZERO PROGRESS BAR]
    ProgressBar (0)

    '[ROSTER LABEL INFO]
    Call RosterInfo("", 0)

End Sub



Private Sub cmdRemove_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Click here to remove the selected staff member/members from the selected cell/cells in the roster."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdRemoveRow_Click()

    '[REMOVE SELECTED ROW FROM GRID]
    Dim intCounter          As Integer
    Dim intCol              As Integer
    Dim intRow              As Integer

    '[POPUP YES/NO DIALOG TO CONFIRM DELETION]
    Dim Msg As String
    Dim Style
    Dim Response
    Dim Title
    
    If frmRoster.GridRoster.Row = 0 Then Exit Sub
    intCol = frmRoster.GridRoster.Col
    
    '[CHECK DELETION FLAG]
    If flagDeleteConfirm Then
        Msg = "This action will remove this time slot from the roster." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "Do you wish to continue ?"
        Style = vbYesNo                     ' Define buttons.
        Title = "Confirmation Required"     ' Define title.
        Response = gsrMsg(Msg, Style, Title)
    Else
        Response = vbYes
    End If
    
    If Response = vbYes Then    ' User chose Yes.
        '[REMOVE ROW FROM ROSTER]
        intRow = frmRoster.GridRoster.Row
        If frmRoster.GridRoster.Rows = 2 Then
            For intCounter = 0 To 8
                frmRoster.GridRoster.Col = intCounter
                frmRoster.GridRoster.Text = ""
            Next intCounter
        Else
            frmRoster.GridRoster.RemoveItem intRow
        End If
        '[MOVE TO A NEW ROW]
        If intRow = frmRoster.GridRoster.Rows Then frmRoster.GridRoster.Row = intRow - 1
        '[MAKE SAVE BUTTON VISIBLE]
        frmRoster.cmdSave.Visible = True
    End If

    frmRoster.GridRoster.Col = intCol


End Sub

Private Sub cmdRemoveRow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Click here to remove the selected row from the roster grid."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdSave_Click()

    '[SAVE DATA IN ROSTER GRID TO DYNASET]
    SaveRosterGrid
    
    '[HIDE SAVE BUTTON]
    frmRoster.cmdSave.Visible = False

End Sub


Private Sub cmdSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Click here to save the changes you have made to this roster."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdTransfer_Click()

    '[TRANSFER ALL SELECTED STAFF MEMBERS TO THE SELECTED CELL/CELLS IN THE ROSTER GRID]
    Dim intCounter          As Integer
    Dim intTrCount          As Integer
    Dim intSelCount         As Integer
    
    '[SELECTION COUNT - ITEMS SELECTED]
    intSelCount = frmRoster.ListStaff.SelCount
    If intSelCount = 0 Then Exit Sub
    intTrCount = 1
    
    '[ROSTER LABEL INFO]
    Call RosterInfo("Transferring to Roster", 0)
    
    For intCounter = 0 To (frmRoster.ListStaff.ListCount - 1)
        If frmRoster.ListStaff.Selected(intCounter) Then
            '[ITEM IS SELECTED SO TRANSFER TO ROSTER]
            TransferToRoster (frmRoster.ListStaff.List(intCounter))
            '[TURN SAVE BUTTON ON]
            frmRoster.cmdSave.Visible = True
            '[PROGRESS BAR]
            intTrCount = intTrCount + 1
            ProgressBar ((intTrCount / intSelCount) * 100)
        End If
    Next intCounter
    
    '[ZERO PROGRESS BAR]
    ProgressBar (0)
    
    Call RosterInfo("", 0)
    
End Sub


Private Sub cmdTransfer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Click here to transfer the selected staff member/members to the selected cell/cells in the roster."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub comboClass_Click()
    
    '[CHECK FOR CHANGED DATA AND NOTIFY]
    If frmRoster.cmdSave.Visible = True Then
        '[ROSTER IS UNSAVED, POPUP YES/NO DIALOG]
        Dim Msg As String
        Dim Style
        Dim Response
        Dim Title
        
        Msg = "You have made changes to this roster (" & frmRoster.ComboClass.Text & ") but have not saved these changes." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "If you choose not to save now, any changes you have made since your last save will be lost." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "Do you wish to save these changes before you move to another roster ?"
        Style = vbYesNoCancel ' Define buttons.
        Title = "Roster Changes Not Saved"  ' Define title.
        Response = gsrMsg(Msg, Style, Title)
        If Response = vbYes Then    ' User chose Yes.
            SaveRosterGrid
        ElseIf Response = vbCancel Then
            Exit Sub
        End If
    End If
    
    '[HIDE SAVE BUTTON]
    frmRoster.cmdSave.Visible = False
    
    '[LOCATE CLASS IN DSCLASS DYNASET]
    LocateClass (frmRoster.ComboClass.Text)
    
    '[APPLY THE NEW CLASS ID TO THE PUBLIC VARIABLE]
    intRosterClass = frmRoster.ComboClass.ItemData(frmRoster.ComboClass.ListIndex)
    
    '[CHANGE CAPTION ON FORM TO CLASS NAME]
    frmRoster.Caption = "Roster - " & frmRoster.ComboClass.Text
    
    '[FILL STAFF ROSTER LIST WITH THOSE STAFF THAT MATCH]
    FillStaffRosterList
    
    '[REBUILD DSROSTER DYNASET WITH DATA FROM DATABASE]
    BuildRosterDynaset (intRosterClass)
    
    '[FILL GRID WITH ROSTER FROM DSROSTER ARRAY]
    FillRosterGrid

End Sub


Private Sub ComboClass_GotFocus()
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Click on this box to drop down a list of available rosters (those checked as active on the control form)."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub Form_Load()

    '[RESIZE GRID AND COLS]
    Dim intCounter      As Integer
    frmRoster.GridRoster.ColWidth(0) = frmRoster.TextWidth("800:008AM")
    frmRoster.GridRoster.ColWidth(1) = frmRoster.TextWidth("800:008AM")

    For intCounter = 2 To 8
        frmRoster.GridRoster.ColWidth(intCounter) = (frmRoster.GridRoster.Width * 0.9) / 9
    Next intCounter
    '[SET INDEX TO FIRST ITEM IN LIST (IF LIST HAS ITEMS)]
    If frmRoster.ComboClass.ListCount > 0 Then frmRoster.ComboClass.ListIndex = 0
    
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "The roster form is used to create and modify your list of rosters."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub Form_Resize()

    '[TEMP VARIABLES SO WE CAN CATCH ILLEGAL WIDTHS]
    Dim sinWidth        As Single
    Dim sinHeight       As Single
    
    '[IF FORM IS MINIMISED THEN EXIT THIS ROUTINE]
    If frmRoster.WindowState = 1 Then Exit Sub
    '[RESIZE GRID AND LIST BOXES AND ARRANGE CONTROLS ON FORM]
    
    '[STAFF LIST]
    sinHeight = frmRoster.Height - frmRoster.ListStaff.Top - PanelToolBar.Height - 180
    If sinHeight > 0 Then frmRoster.ListStaff.Height = sinHeight
    
    '[ROSTER GRID]
    frmRoster.GridRoster.Top = frmRoster.ComboClass.Top
    sinWidth = frmRoster.Width - (frmRoster.ListStaff.Width + frmRoster.cmdInsertRow + 640)
    If sinWidth > 0 Then frmRoster.GridRoster.Width = sinWidth
    sinHeight = frmRoster.ListStaff.Height + frmRoster.ComboClass.Height + 90
    If sinHeight > 0 Then frmRoster.GridRoster.Height = sinHeight
        
    '[PROGRESS BAR]
    sinWidth = frmRoster.Width - frmRoster.GaugeProgress.Left - 280
    If sinWidth > 0 Then frmRoster.GaugeProgress.Width = sinWidth
    
End Sub


Private Sub GridRoster_DblClick()
    
    '[ALLOW USER TO CHANGE THE SHIFT START/END TIME]
    Dim Msg As String
    Dim Style
    Dim Response
    Dim Title

    If frmRoster.GridRoster.Row = 0 Then Exit Sub
    
    Select Case frmRoster.GridRoster.Col
    Case 0, 1     '[SHIFT START TIME/END TIME]
        Load frmTime
        
        '[SET TIME ON FORM]
        If frmRoster.GridRoster.Text = "" Then
            If frmRoster.GridRoster.Col = 0 Then
                '[SHIFT START TIME]
                frmTime.Caption = "Select Start Time"
                frmTime.ComboHour = Format(Hour(DsDefault("StartTime")), "0#")
                frmTime.ComboMinute = Format(Minute(DsDefault("StartTime")), "0#")
                frmTime.Refresh
            Else
                '[SHIFT END TIME]
                frmTime.Caption = "Select End Time"
                frmTime.ComboHour = Format(Hour(DsDefault("EndTime")), "0#")
                frmTime.ComboMinute = Format(Minute(DsDefault("EndTime")), "0#")
                frmTime.Refresh
            End If
        Else
            If IsDate(frmRoster.GridRoster.Text) Then
                frmTime.ComboHour = Format(Hour(CDate(frmRoster.GridRoster.Text)), "0#")
                frmTime.ComboMinute = Format(Minute(CDate(frmRoster.GridRoster.Text)), "0#")
            Else
                frmTime.ComboHour = "00"
                frmTime.ComboMinute = "00"
            End If
        End If
        frmTime.Show 1
        '[PROCESS RESULT OF FORM]
        If frmTime.CheckResult = vbChecked Then
            frmRoster.GridRoster.Text = Format(frmTime.ComboHour.Text & ":" & frmTime.ComboMinute.Text, "Medium Time")
            '[OLD FORMAT COMMAND] -> "hh:mm AMPM"
            '[CHECK FOR CELL SIZE AND RESIZE IF NECESSARY]
            '[ALLOW 10% MARGIN FOR TEXT ADJUSTMENT]
            If frmRoster.GridRoster.ColWidth(frmRoster.GridRoster.Col) < (TextWidth(frmRoster.GridRoster.Text) * 1.1) Then frmRoster.GridRoster.ColWidth(frmRoster.GridRoster.Col) = (TextWidth(frmRoster.GridRoster.Text) * 1.1)
            '[MAKE SAVE BUTTON VISIBLE]
            frmRoster.cmdSave.Visible = True
        End If
        '[REMOVE FORM]
        'Unload frmTime

    Case Else   '[SHOW FULL CELL CONTENTS]
        If Len(Trim(frmRoster.GridRoster.Text)) = 0 Then Exit Sub
        Msg = frmRoster.GridRoster.Text
        Style = vbOKOnly                        ' Define buttons.
        Title = "Cell Contents"                 ' Define title.
        Response = gsrMsg(Msg, Style, Title)
    
    End Select

End Sub

Private Sub GridRoster_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
    Case vbKeyReturn
        '[CAPTURE ENTER KEY]
        Call GridRoster_DblClick
    Case vbKeyDelete
        '[CAPTURE DELETE KEY]
        Call cmdDelete_Click
    Case Else
    End Select

End Sub


Private Sub GridRoster_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    If X > frmControl.GridClass.ColPos(2) Then
        StatusBar "Double-click to display the names which appear in this cell."
    Else
        StatusBar "Double-click to modify the shift start/finish time."
    End If
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub ListStaff_Click()

    '[CHECK NUMBER OF SELECTED ITEMS AND CHANGE ICON ON TRANSFER BUTTON]
    If frmRoster.ListStaff.SelCount = 0 Then
        frmRoster.cmdTransfer.Enabled = False
    Else
        frmRoster.cmdTransfer.Enabled = True
    End If

End Sub


Private Sub ListStaff_DblClick()

    '[TRANSFER STAFF MEMBER ON DOUBLE CLICK -IF- NAME ISN'T IN CELL]
    If InStr(frmRoster.GridRoster.Text, frmRoster.ListStaff.List(frmRoster.ListStaff.ListIndex)) > 0 Then
        '[NAME IS IN CELL, REMOVE NAME]
        cmdRemove_Click
    Else
        '[NAME ISN'T IN CELL, ADD NAME]
        cmdTransfer_Click
    End If

End Sub


Private Sub ListStaff_GotFocus()
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Double click a staff name to transfer it to, or remove it from, the current roster."
    '[---------------------------------------------------------------------------------]

End Sub


