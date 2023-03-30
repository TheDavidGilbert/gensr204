VERSION 4.00
Begin VB.Form frmRegister 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1275
   ClientLeft      =   7275
   ClientTop       =   7845
   ClientWidth     =   4995
   ClipControls    =   0   'False
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
   Height          =   1680
   Left            =   7215
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   Top             =   7500
   Width           =   5115
   Begin VB.CheckBox CheckResult 
      Alignment       =   1  'Right Justify
      Caption         =   "Result"
      Height          =   315
      Left            =   60
      TabIndex        =   6
      Top             =   1380
      Value           =   1  'Checked
      Width           =   1965
   End
   Begin VB.TextBox txtRegCode 
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   870
      Width           =   3225
   End
   Begin VB.TextBox txtRegUser 
      Height          =   315
      Left            =   60
      MaxLength       =   50
      TabIndex        =   0
      Top             =   270
      Width           =   4875
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter your validation code here :"
      Height          =   210
      Index           =   1
      Left            =   90
      TabIndex        =   5
      Top             =   630
      Width           =   2340
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter your registered user name here :"
      Height          =   210
      Index           =   0
      Left            =   90
      TabIndex        =   4
      Top             =   30
      Width           =   2805
   End
   Begin Threed.SSCommand cmdOK 
      Default         =   -1  'True
      Height          =   345
      Left            =   3390
      TabIndex        =   3
      Top             =   870
      Width           =   750
      _version        =   65536
      _extentx        =   1323
      _extenty        =   609
      _stockprops     =   78
      caption         =   "OK"
   End
   Begin Threed.SSCommand cmdCancel 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   4200
      TabIndex        =   2
      Top             =   870
      Width           =   750
      _version        =   65536
      _extentx        =   1323
      _extenty        =   609
      _stockprops     =   78
      caption         =   "Cancel"
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

    '[SET RESULT OF BUTTON CHOICE]
    frmRegister.CheckResult = vbUnchecked
    '[CLOSE FORM / ONLY UNLOAD THOUGH, LET CALLING ROUTINE UNLOAD FORM]
    frmRegister.Hide

End Sub


Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Click here to cancel registration and continue."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdOK_Click()

    '[SET RESULT OF BUTTON CHOICE]
    frmRegister.CheckResult = vbChecked
    '[CLOSE FORM / ONLY UNLOAD THOUGH, LET CALLING ROUTINE UNLOAD FORM]
    frmRegister.Hide

End Sub


Private Sub cmdOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Click here to accept the displayed registration information and continue."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub Form_Load()
    
    '[center form on screen]
    Top = (Screen.Height / 2) - (Height / 2)
    Left = (Screen.Width / 2) - (Width / 2)

End Sub


Private Sub txtRegCode_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Enter your registration validation code here."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub txtRegUser_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Enter your registered user name here."
    '[---------------------------------------------------------------------------------]

End Sub


