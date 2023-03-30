VERSION 4.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5280
   ClientLeft      =   6270
   ClientTop       =   2490
   ClientWidth     =   5685
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00808080&
   Height          =   5685
   Icon            =   "FRMABOUT.frx":0000
   Left            =   6210
   LinkTopic       =   "frmAbout"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   Top             =   2145
   Width           =   5805
   Begin VB.Image ImageInfo 
      Height          =   480
      Index           =   2
      Left            =   60
      Picture         =   "FRMABOUT.frx":000C
      Top             =   4110
      Width           =   480
   End
   Begin Threed.SSCommand cmdReturn 
      Height          =   360
      Left            =   5280
      TabIndex        =   0
      Top             =   4860
      Width           =   360
      _version        =   65536
      _extentx        =   635
      _extenty        =   635
      _stockprops     =   78
      forecolor       =   -2147483630
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
      picture         =   "FRMABOUT.frx":08D6
   End
   Begin VB.Image ImageInfo 
      Height          =   480
      Index           =   1
      Left            =   60
      Picture         =   "FRMABOUT.frx":09E8
      Top             =   3450
      Width           =   480
   End
   Begin VB.Image ImageInfo 
      Height          =   480
      Index           =   0
      Left            =   90
      Picture         =   "FRMABOUT.frx":12B2
      Top             =   2790
      Width           =   480
   End
   Begin VB.Label LabelInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Generic Staff Roster,  (C) 1997 David Gilbert."
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   1875
      Left            =   705
      TabIndex        =   3
      Top             =   2700
      Width           =   4845
      WordWrap        =   -1  'True
   End
   Begin VB.Image ImageEmail 
      Height          =   480
      Left            =   90
      Picture         =   "FRMABOUT.frx":1B7C
      Top             =   4710
      Width           =   480
   End
   Begin VB.Image ImageLicence 
      Height          =   480
      Left            =   60
      Picture         =   "FRMABOUT.frx":2446
      Top             =   2100
      Width           =   480
   End
   Begin VB.Image ImageLogo 
      Height          =   2070
      Left            =   30
      Picture         =   "FRMABOUT.frx":2D10
      Top             =   0
      Width           =   5640
   End
   Begin VB.Label LabelEmail 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Email Address :"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Left            =   700
      TabIndex        =   1
      Top             =   4890
      Width           =   4065
   End
   Begin VB.Label labelRegUser 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "This copy of GSR is licensed to ______________________________"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   700
      TabIndex        =   2
      Top             =   2130
      Width           =   4850
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit
'[----------------------------------------------]
'[frmAbout          About program and author    ]
'[----------------------------------------------]
'[GSR                                           ]
'[generic staff roster system for Windows       ]
'[(C) ---------- David Gilbert, 1997            ]
'[----------------------------------------------]
'[version                           2.1         ]
'[initial date                      21/07/1997  ]
'[----------------------------------------------]



Private Sub cmdReturn_Click()

    '[CLOSE FORM]
    Unload frmAbout

End Sub

Private Sub cmdReturn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Click here to return to Generic Staff Roster."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub Form_Click()

    Unload frmAbout

End Sub

Private Sub Form_Load()

    '[center form on screen]
    Top = (Screen.Height / 2) - (Height / 2)
    Left = (Screen.Width / 2) - (Width / 2)
        
End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Click anywhere on this form to return to Generic Staff Roster."
    '[---------------------------------------------------------------------------------]

End Sub



Private Sub ImageEmail_DblClick()
    
    '[COPY ADDRESS TO CLIPBOARD]
    Clipboard.SetText constEmail, vbCFText

End Sub


Private Sub ImageEmail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Double-click here to copy the authors email address to the clipboard."
    '[---------------------------------------------------------------------------------]

End Sub

Private Sub ImageInfo_Click(Index As Integer)

    Select Case Index
    Case 0
        labelInfo.Caption = "A Generic Staff Rostering software system for small businesses."
        labelInfo.Caption = labelInfo.Caption & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "Generic Staff Roster was designed to assist small business owners/operators in allocating their available human resources easily and efficiently."
        labelInfo.Caption = labelInfo.Caption & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "For more information on buying Generic Staff Roster, see the ordering form which accompanies this program."
    Case 1
        labelInfo.Caption = "Generic Staff Roster was designed as a simple alternative to pencil and paper staff roster creation, allowing the maintainance of ten seperate weekly rosters.  GSR incorporates several roster failsafes, such as duplicate shift allocations."
        labelInfo.Caption = labelInfo.Caption & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "The concept was spawned from various comments about the difficulty of creating multiple weekly rosters on paper."
    Case 2
        labelInfo.Caption = "David Gilbert lives in Toowoomba, Australia and works as a freelance programmer.  He is " & Format(Date - CDate("25/12/1967"), "yy") & " years old and spends most of his time throwing a tennis ball for his two dogs."
        labelInfo.Caption = labelInfo.Caption & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "GSR is his first non-commissioned software release and was written during a frenzied two week period in September, 1997."
    Case Else
        labelInfo.Caption = "Generic Staff Roster"
        labelInfo.Caption = labelInfo.Caption & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "(C) 1997 David Gilbert"
    End Select

End Sub

Private Sub ImageInfo_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Click on these icons for more information ..."
    '[---------------------------------------------------------------------------------]

End Sub



Private Sub ImageLogo_Click()
    
    Unload frmAbout

End Sub



Private Sub labelInfo_Click()

    Unload frmAbout

End Sub


Private Sub labelRegUser_Click()

    Unload frmAbout

End Sub





