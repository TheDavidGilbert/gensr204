VERSION 4.00
Begin VB.Form frmMsg 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3945
   ClientLeft      =   5535
   ClientTop       =   7365
   ClientWidth     =   5250
   ControlBox      =   0   'False
   BeginProperty Font 
      name            =   "Arial"
      charset         =   1
      weight          =   400
      size            =   8.25
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   Height          =   4350
   Left            =   5475
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   Top             =   7020
   Width           =   5370
   Begin GaugeLib.Gauge GaugeProgress 
      Height          =   270
      Left            =   120
      TabIndex        =   6
      Top             =   3555
      Visible         =   0   'False
      Width           =   4995
      _version        =   65536
      _extentx        =   8811
      _extenty        =   476
      _stockprops     =   73
      forecolor       =   255
      innertop        =   0
      innerleft       =   0
      innerright      =   0
      innerbottom     =   0
      value           =   50
      needlewidth     =   2
   End
   Begin VB.TextBox TextNote 
      Height          =   2535
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   690
      Visible         =   0   'False
      Width           =   5200
   End
   Begin VB.Label LabelMessage 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label Message"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   9
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   2300
      Left            =   150
      TabIndex        =   5
      Top             =   800
      Width           =   4935
      WordWrap        =   -1  'True
   End
   Begin VB.Line LineDivider 
      BorderWidth     =   2
      X1              =   30
      X2              =   5220
      Y1              =   630
      Y2              =   630
   End
   Begin VB.Label LabelTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Title Label"
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
      Left            =   2130
      TabIndex        =   4
      Top             =   150
      Width           =   1185
   End
   Begin VB.Image ImageInfo 
      Height          =   480
      Index           =   0
      Left            =   120
      Picture         =   "FRMMSG.frx":0000
      Top             =   90
      Width           =   480
   End
   Begin Threed.SSCommand cmdOK 
      Height          =   375
      Left            =   3690
      TabIndex        =   3
      Top             =   3450
      Width           =   945
      _version        =   65536
      _extentx        =   1667
      _extenty        =   661
      _stockprops     =   78
      caption         =   "OK"
   End
   Begin Threed.SSCommand cmdCancel 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   375
      Left            =   2670
      TabIndex        =   2
      Top             =   3450
      Width           =   945
      _version        =   65536
      _extentx        =   1667
      _extenty        =   661
      _stockprops     =   78
      caption         =   "Cancel"
   End
   Begin Threed.SSCommand cmdNo 
      Height          =   375
      Left            =   1650
      TabIndex        =   1
      Top             =   3450
      Width           =   945
      _version        =   65536
      _extentx        =   1667
      _extenty        =   661
      _stockprops     =   78
      caption         =   "No"
   End
   Begin Threed.SSCommand cmdYes 
      Height          =   375
      Left            =   630
      TabIndex        =   0
      Top             =   3450
      Width           =   945
      _version        =   65536
      _extentx        =   1667
      _extenty        =   661
      _stockprops     =   78
      caption         =   "Yes"
   End
   Begin VB.Label LabelInfo 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   130
      TabIndex        =   7
      Top             =   3300
      Visible         =   0   'False
      Width           =   4995
   End
   Begin VB.Shape ShapeBackGround 
      BackColor       =   &H80000002&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H80000002&
      FillStyle       =   0  'Solid
      Height          =   2500
      Left            =   30
      Top             =   690
      Width           =   5200
   End
End
Attribute VB_Name = "frmMsg"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit
'[----------------------------------------------]
'[frmMsg            GSR Message Form            ]
'[----------------------------------------------]
'[GSR                                           ]
'[generic staff roster system for Windows       ]
'[(C) ---------- David Gilbert, 1997            ]
'[----------------------------------------------]
'[version                           2.1         ]
'[initial date                      21/07/1997  ]
'[----------------------------------------------]



Private Sub cmdCancel_Click()
    
    '[SAVE RETURN CODE AND UNLOAD FORM]
    gsrReturn = vbCancel
    Unload frmMsg

End Sub

Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Click here to CANCEL this operation ..."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdNo_Click()
    
    '[SAVE RETURN CODE AND UNLOAD FORM]
    gsrReturn = vbNo
    Unload frmMsg

End Sub

Private Sub cmdNo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Click here to answer NO to the question and continue ..."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdOK_Click()
    
    '[SAVE RETURN CODE AND UNLOAD FORM]
    gsrReturn = vbOK
    '[IF THIS IS A TEXT GET MESSAGE, SAVE THE CHANGED TEXT]
    If frmMsg.TextNote.Visible = True Then gsrNote = frmMsg.TextNote.Text
    Unload frmMsg

End Sub

Private Sub cmdOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Click here to continue once you have read the message ..."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdYes_Click()

    '[SAVE RETURN CODE AND UNLOAD FORM]
    gsrReturn = vbYes
    Unload frmMsg
    

End Sub


Private Sub cmdYes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Click here to answer YES to the question and continue ..."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub Form_Load()

    '[center form on screen]
    frmMsg.Top = (Screen.Height / 2) - (frmMsg.Height / 2)
    frmMsg.Left = (Screen.Width / 2) - (frmMsg.Width / 2)

End Sub


Private Sub GaugeProgress_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "This bar displays the progress percentage of the current operation (" & frmMsg.GaugeProgress.Value & "%)"
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub LabelMessage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "This form displays relevant messages about GSR's activites."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub TextNote_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Type any notes for the selected roster here."
    '[---------------------------------------------------------------------------------]

End Sub


