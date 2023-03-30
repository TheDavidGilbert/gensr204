VERSION 4.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GSR - Main"
   ClientHeight    =   6390
   ClientLeft      =   1920
   ClientTop       =   2010
   ClientWidth     =   7515
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   Height          =   6750
   Icon            =   "FRMCLASS.frx":0000
   Left            =   1860
   LinkTopic       =   "frmClass"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   Top             =   1710
   Width           =   7635
   Begin VB.ComboBox ComboIncrement 
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   700
         size            =   9
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "FRMCLASS.frx":000C
      Left            =   4740
      List            =   "FRMCLASS.frx":003F
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   240
      Width           =   1935
   End
   Begin VB.ComboBox ComboStartDay 
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   700
         size            =   9
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "FRMCLASS.frx":00C1
      Left            =   2760
      List            =   "FRMCLASS.frx":00DA
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   240
      Width           =   1935
   End
   Begin VB.Frame FrameClass 
      Caption         =   "Class Definitions"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3495
      Left            =   60
      TabIndex        =   0
      Top             =   240
      Width           =   2595
      Begin VB.Data DataClass 
         Caption         =   "DataClass"
         Connect         =   "Access"
         DatabaseName    =   "C:\CONTRACT\GSR\GSR.MDB"
         Exclusive       =   0   'False
         Height          =   315
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Class"
         Top             =   3720
         Visible         =   0   'False
         Width           =   4035
      End
      Begin MSDBGrid.DBGrid GridClass 
         Bindings        =   "FRMCLASS.frx":011E
         Height          =   2835
         Left            =   120
         OleObjectBlob   =   "FRMCLASS.frx":0130
         TabIndex        =   1
         Top             =   300
         Width           =   2355
      End
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
      Left            =   4740
      TabIndex        =   5
      Top             =   0
      Width           =   1230
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
      Left            =   2760
      TabIndex        =   3
      Top             =   0
      Width           =   1530
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Private Sub ComboIncrement_Click()

    '[Set Roster Increment]
    If ComboIncrement.ListIndex <> mdiMain.DataDefaults.Recordset("Increment") - 1 And ComboIncrement.ListIndex > -1 Then
    
        mdiMain.DataDefaults.Recordset.Edit
                mdiMain.DataDefaults.Recordset("Increment") = ComboIncrement.ListIndex + 1
        mdiMain.DataDefaults.Recordset.Update
        
    End If

End Sub


Private Sub ComboStartDay_Click()

    '[Set Starting Day of Week]
    If ComboStartDay.ListIndex <> mdiMain.DataDefaults.Recordset("StartDay") - 1 And ComboStartDay.ListIndex > -1 Then
    
        mdiMain.DataDefaults.Recordset.Edit
                mdiMain.DataDefaults.Recordset("StartDay") = ComboStartDay.ListIndex + 1
        mdiMain.DataDefaults.Recordset.Update
        
    End If

End Sub


Private Sub Form_Load()

    '[Resize controls and grids to match]
    GridClass.Height = GridClass.RowHeight * 11
    FrameClass.Height = GridClass.Height * 1.15
    
    '[Set Starting Day of Week]
    ComboStartDay.ListIndex = mdiMain.DataDefaults.Recordset("StartDay") - 1
    
    '[Set Roster Increment]
    ComboIncrement.ListIndex = mdiMain.DataDefaults.Recordset("Increment") - 1

End Sub

Private Sub Form_Terminate()

    mdiMain.mnuView_Class.Checked = False
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

    mdiMain.mnuView_Class.Checked = False

End Sub


