VERSION 4.00
Begin VB.Form frmTime 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Time"
   ClientHeight    =   405
   ClientLeft      =   6825
   ClientTop       =   10485
   ClientWidth     =   3315
   ControlBox      =   0   'False
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
   Height          =   810
   Icon            =   "FRMTIME.frx":0000
   Left            =   6765
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   405
   ScaleWidth      =   3315
   Top             =   10140
   Width           =   3435
   Begin VB.CheckBox CheckResult 
      Alignment       =   1  'Right Justify
      Caption         =   "Result"
      Height          =   315
      Left            =   30
      TabIndex        =   4
      Top             =   510
      Value           =   1  'Checked
      Width           =   1965
   End
   Begin VB.ComboBox ComboHour 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   330
      ItemData        =   "FRMTIME.frx":0442
      Left            =   30
      List            =   "FRMTIME.frx":049C
      TabIndex        =   0
      Text            =   "ComboHour"
      Top             =   30
      Width           =   795
   End
   Begin VB.ComboBox ComboMinute 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   330
      ItemData        =   "FRMTIME.frx":0500
      Left            =   870
      List            =   "FRMTIME.frx":0510
      TabIndex        =   1
      Text            =   "ComboMinute"
      Top             =   30
      Width           =   795
   End
   Begin Threed.SSCommand cmdCancel 
      Cancel          =   -1  'True
      Height          =   330
      Left            =   2520
      TabIndex        =   3
      Top             =   30
      Width           =   750
      _version        =   65536
      _extentx        =   1323
      _extenty        =   582
      _stockprops     =   78
      caption         =   "Cancel"
   End
   Begin Threed.SSCommand cmdOK 
      Default         =   -1  'True
      Height          =   330
      Left            =   1710
      TabIndex        =   2
      Top             =   30
      Width           =   750
      _version        =   65536
      _extentx        =   1323
      _extenty        =   582
      _stockprops     =   78
      caption         =   "OK"
   End
End
Attribute VB_Name = "frmTime"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

    '[SET RESULT OF BUTTON CHOICE]
    frmTime.CheckResult = vbUnchecked
    '[CLOSE FORM / ONLY UNLOAD THOUGH, LET CALLING ROUTINE UNLOAD FORM]
    frmTime.Hide

End Sub

Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Click here to abandon any changes you have made and return."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdOK_Click()

    '[SET RESULT OF BUTTON CHOICE]
    frmTime.CheckResult = vbChecked
    '[CLOSE FORM / ONLY UNLOAD THOUGH, LET CALLING ROUTINE UNLOAD FORM]
    frmTime.Hide

End Sub


Private Sub cmdOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Click here to accept any changes you have made and return."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub Form_Load()

    '[CENTER FORM]
    frmTime.Left = (Screen.Width / 2) - (frmTime.Width / 2)
    frmTime.Top = (Screen.Height / 2) - (frmTime.Height / 2)

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Click on the drop down lists to select the required hour/minute."
    '[---------------------------------------------------------------------------------]

End Sub


