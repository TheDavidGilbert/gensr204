VERSION 4.00
Begin VB.Form frmValidate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Validation Code Generation"
   ClientHeight    =   1755
   ClientLeft      =   6225
   ClientTop       =   6990
   ClientWidth     =   4950
   BeginProperty Font 
      name            =   "Arial"
      charset         =   0
      weight          =   400
      size            =   8.25
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   Height          =   2160
   Left            =   6165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   Top             =   6645
   Width           =   5070
   Begin VB.TextBox txtModifier 
      Height          =   315
      Left            =   810
      TabIndex        =   0
      Text            =   "129"
      Top             =   60
      Width           =   885
   End
   Begin VB.TextBox txtRegUser 
      Height          =   315
      Left            =   30
      MaxLength       =   50
      TabIndex        =   1
      Top             =   690
      Width           =   4875
   End
   Begin VB.TextBox txtRegCode 
      Height          =   315
      Left            =   30
      TabIndex        =   2
      Top             =   1320
      Width           =   3225
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Modifier :"
      Height          =   210
      Index           =   2
      Left            =   60
      TabIndex        =   5
      Top             =   120
      Width           =   660
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter your registered user name here :"
      Height          =   210
      Index           =   0
      Left            =   60
      TabIndex        =   4
      Top             =   450
      Width           =   2805
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Validation Code :"
      Height          =   210
      Index           =   1
      Left            =   60
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "frmValidate"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Function Validate(strValidate) As String

    '[FUNCTION TO VALIDATE A STRING AND RETURN THE VALIDATION CODE]
    Dim strRegCode      As String       '[RETURNED HEX VALIDATION CODE]
    Dim sinValue        As Single       '[ACCUMULATED VALUE]
    Dim intCounter      As Integer      '[COUNTER FOR LENGTH]

    If IsNull(strValidate) Or Len(strValidate) = 0 Then
        '[NO CODE TO VALIDATE SO RETURN NULL]
        strRegCode = ""
    Else
        For intCounter = 1 To Len(strValidate)
        '[CYCLE THROUGH VALIDATION STRING AND ACCUMULATE VALUES]
            sinValue = sinValue + (Asc(Mid$(strValidate, intCounter, 1)) * Val(txtModifier))
        Next intCounter
        strRegCode = Hex(sinValue)
    End If
    
    '[RETURN CODE VALUE]
    Validate = strRegCode

End Function


Private Sub Form_Load()
    
    '[center form on screen]
    Top = (Screen.Height / 2) - (Height / 2)
    Left = (Screen.Width / 2) - (Width / 2)


End Sub

Private Sub txtRegUser_Change()

    '[SHOW VALIDATION CODE]
    txtRegCode.Text = Validate(txtRegUser.Text)

End Sub


