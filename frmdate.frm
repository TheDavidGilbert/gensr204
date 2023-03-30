VERSION 4.00
Begin VB.Form frmDate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Date"
   ClientHeight    =   390
   ClientLeft      =   7560
   ClientTop       =   10110
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
   Height          =   795
   Icon            =   "FRMDATE.frx":0000
   Left            =   7500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   390
   ScaleWidth      =   3315
   Top             =   9765
   Width           =   3435
   Begin VB.CheckBox CheckResult 
      Alignment       =   1  'Right Justify
      Caption         =   "Result"
      Height          =   315
      Left            =   1650
      TabIndex        =   2
      Top             =   420
      Value           =   1  'Checked
      Width           =   1965
   End
   Begin MSMask.MaskEdBox MaskDate 
      DataField       =   "DateHired"
      Height          =   345
      Left            =   60
      TabIndex        =   3
      Top             =   30
      Width           =   1005
      _version        =   65536
      _extentx        =   1773
      _extenty        =   609
      _stockprops     =   109
      forecolor       =   0
      backcolor       =   16777215
      borderstyle     =   1
      autotab         =   -1  'True
      promptinclude   =   0   'False
      clipmode        =   1
      format          =   "Short Date"
      appearance      =   0
   End
   Begin Threed.SSCommand cmdCancel 
      Cancel          =   -1  'True
      Height          =   330
      Left            =   2520
      TabIndex        =   1
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
      TabIndex        =   0
      Top             =   30
      Width           =   750
      _version        =   65536
      _extentx        =   1323
      _extenty        =   582
      _stockprops     =   78
      caption         =   "OK"
   End
End
Attribute VB_Name = "frmDate"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

    '[SET RESULT OF BUTTON CHOICE]
    frmDate.CheckResult = vbUnchecked
    '[CLOSE FORM / ONLY UNLOAD THOUGH, LET CALLING ROUTINE UNLOAD FORM]
    frmDate.Hide

End Sub

Private Sub cmdOK_Click()

    '[SET RESULT OF BUTTON CHOICE]
    frmDate.CheckResult = vbChecked
    '[CLOSE FORM / ONLY UNLOAD THOUGH, LET CALLING ROUTINE UNLOAD FORM]
    frmDate.Hide

End Sub


Private Sub Form_Load()

    '[CENTER FORM]
    frmTime.Left = (Screen.Width / 2) - (frmTime.Width / 2)
    frmTime.Top = (Screen.Height / 2) - (frmTime.Height / 2)

End Sub


