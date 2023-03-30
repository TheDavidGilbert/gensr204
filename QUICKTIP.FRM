VERSION 4.00
Begin VB.Form Quick_Tip 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Quick Tip"
   ClientHeight    =   3285
   ClientLeft      =   3540
   ClientTop       =   4785
   ClientWidth     =   6315
   BeginProperty Font 
      name            =   "MS Sans Serif"
      charset         =   1
      weight          =   700
      size            =   8.25
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Height          =   3690
   Left            =   3480
   LinkTopic       =   "Quick_Tip"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   6315
   Top             =   4440
   Width           =   6435
   Begin VB.CheckBox show_tip 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Show Tips at Startup"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Value           =   1  'Checked
      Width           =   3495
   End
   Begin VB.CommandButton Quick_Tip_ 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Help"
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton More_Tip_Button 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&More Tips..."
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Next_Tip_Button 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Next Tip"
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton Quick_Tip_OK_Button 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   360
      Picture         =   "QUICKTIP.frx":0000
      Top             =   250
      Width           =   420
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Did You Know..."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label tip_text 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Tip Text...................."
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   360
      TabIndex        =   6
      Top             =   960
      Width           =   4095
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   2535
      Left            =   240
      Top             =   240
      Width           =   4335
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   4680
      X2              =   4680
      Y1              =   2880
      Y2              =   120
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   4680
      X2              =   120
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   4680
      X2              =   120
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   2880
   End
End
Attribute VB_Name = "Quick_Tip"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Declare Function GetPrivateProfileString Lib "Kernel" (ByVal Section As String, ByVal Entry As String, ByVal DefaultValueEntry As String, ByVal ReturnedEntry As String, ByVal EntryValueSize As Integer, ByVal FileName As String) As Integer
Private Declare Function WritePrivateProfileString Lib "Kernel" (ByVal Section As String, ByVal Entry As String, ByVal EntryValue As String, ByVal FileName As String) As Integer

Dim current_tip As String
Dim quick_ini As String

Private Sub Form_Load()
    
    tip_tex = ""
    quick_ini = CurDir$ + "\QUICKTIP.INI"
    start_tip = Trim(rtv_ini_data(quick_ini, "TIP", "START"))
    
    If start_tip = "T" Then
        show_tip = 1
    Else
        show_tip = 0
    End If

    load_tip
    
End Sub

Private Sub load_tip()
    quick_ini = CurDir$ + "\QUICKTIP.INI"
    current_tip = rtv_ini_data(quick_ini, "TIP", "CURRENT")
    tip_text = rtv_ini_data(quick_ini, "TIP", current_tip)
    current_tip = CStr(Val(current_tip) + 1)
    error_check = rtv_ini_data(quick_ini, "TIP", current_tip)
    If error_check = "unknown" Then
        current_tip = "1"
    End If
    xx = upd_ini_data(quick_ini, "TIP", "CURRENT", current_tip)
End Sub

Private Sub More_Tip_Button_Click()
    MsgBox "Place more Tip help under this button", 0
End Sub

Private Sub Next_Tip_Button_Click()
    load_tip
End Sub

Private Sub Quick_Tip__Click()
    MsgBox "Place your help under this button", 0
End Sub

Private Sub Quick_Tip_OK_Button_Click()
    Dim start_tip As String
    If show_tip Then
        start_tip = "T"
    Else
        start_tip = "F"
    End If
    quick_ini = CurDir$ + "\QUICKTIP.INI"
    xx = upd_ini_data(quick_ini, "TIP", "START", start_tip)
    
    End
    'unload QUICKTIP
End Sub

Private Function rtv_ini_data$(ini_File$, ini_select$, ini_entry$)
        ' rtv_ini_data("XXX.INI","ENVIRONMENT","PATH")
        Dim temp_string As String
        
        rtv_ini_data$ = ""
        
        Section$ = ini_select$
        Entry$ = ini_entry$
        DefaultValue$ = "unknown"
        EntryValue$ = Space$(1000)
        EntryValueSize = 1000
        
        X% = GetPrivateProfileString(Section$, Entry$, DefaultValue$, EntryValue$, EntryValueSize, ini_File$)
        If X% = 0 Then
            box_msg$ = " Invalid Entry -> " + Section$ + "/" + Entry$
            MsgBox box_msg$, 16, "Loading Error"
        End If
        
        Len_string = Len(Trim$(EntryValue$)) - 1
        temp_string = Left$(EntryValue$, Len_string)
        rtv_ini_data$ = temp_string
End Function

Private Function upd_ini_data$(ini_File$, ini_select$, ini_entry$, EntryValue$)
        ' upd_ini_data("XXX.INI","ENVIRONMENT","PATH", "C:\")
        
        upd_ini_data$ = ""
        
        Section$ = ini_select$
        Entry$ = ini_entry$
        
        X% = WritePrivateProfileString(Section$, Entry$, EntryValue$, ini_File$)
        If X% = 0 Then
            box_msg$ = " Invalid Entry -> " + Section$ + "/" + Entry$
            MsgBox box_msg$, 16, "Loading Error"
        End If
        
End Function

