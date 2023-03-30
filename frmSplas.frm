VERSION 4.00
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   2085
   ClientLeft      =   5700
   ClientTop       =   5610
   ClientWidth     =   5640
   ControlBox      =   0   'False
   Height          =   2490
   Icon            =   "FRMSPLAS.frx":0000
   Left            =   5640
   LinkTopic       =   "frmSplash"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   Top             =   5265
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Width           =   5760
   Begin VB.Timer TimerStart 
      Interval        =   100
      Left            =   5190
      Top             =   30
   End
   Begin VB.Label LabelVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   1830
      Width           =   645
   End
   Begin VB.Image ImageLogo 
      Height          =   2070
      Left            =   0
      Picture         =   "FRMSPLAS.frx":000C
      Top             =   0
      Width           =   5640
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit
'[----------------------------------------------]
'[frmSplash         Splash startup screen       ]
'[----------------------------------------------]
'[GSR                                           ]
'[generic staff roster system for Windows       ]
'[(C) ---------- David Gilbert, 1997            ]
'[----------------------------------------------]
'[version                           2.1         ]
'[initial date                      21/07/1997  ]
'[----------------------------------------------]

Dim Shared flagInitialised
Dim Shared intTimerCount    'counter for timer


Private Sub Form_Load()

    '[size form]
    frmSplash.ZOrder
    Width = ImageLogo.Width
    Height = ImageLogo.Height
    '[center form on screen]
    Top = (Screen.Height / 2) - (Height / 2)
    Left = (Screen.Width / 2) - (Width / 2)
        
    frmSplash.LabelVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    '[BETA VERSION]
    If flagBeta = True Then frmSplash.LabelVersion.Caption = frmSplash.LabelVersion.Caption + " beta"
    
End Sub


Private Sub ImageLogo_Click()
    
    '[IF USER CLICKS ON THE FORM, UNLOAD]
    Unload frmSplash

End Sub

Private Sub TimerStart_Timer()

    '[after constDelay seconds, unload form and continue]
    intTimerCount = intTimerCount + 0.1                     ' add 10ths of seconds
    If intTimerCount >= constDelay Then Unload frmSplash

End Sub


