VERSION 4.00
Begin VB.Form frmStaff 
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Staff List"
   ClientHeight    =   4710
   ClientLeft      =   4575
   ClientTop       =   5400
   ClientWidth     =   7470
   ControlBox      =   0   'False
   Height          =   5115
   Icon            =   "FRMSTAFF.frx":0000
   Left            =   4515
   LinkTopic       =   "frmEmpList"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   Top             =   5055
   Width           =   7590
   Begin VB.ListBox ListStaff 
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   4050
      Left            =   30
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   630
      Width           =   2235
   End
   Begin Threed.SSPanel PanelToolBar 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   38
      Top             =   0
      Width           =   7470
      _version        =   65536
      _extentx        =   13176
      _extenty        =   661
      _stockprops     =   15
      forecolor       =   -2147483641
      bevelouter      =   0
      floodtype       =   1
      floodcolor      =   -2147483646
      floodshowpct    =   0   'False
      alignment       =   0
      autosize        =   2
      mouseicon       =   "FRMSTAFF.frx":08CA
      Begin Threed.SSCommand cmdSelectedStaffRoster 
         Height          =   360
         Left            =   1170
         TabIndex        =   35
         Top             =   15
         Width           =   360
         _version        =   65536
         _extentx        =   635
         _extenty        =   635
         _stockprops     =   78
         autosize        =   2
         picture         =   "FRMSTAFF.frx":11A4
      End
      Begin Threed.SSCommand cmdAllStaffRosters 
         Height          =   360
         Left            =   1560
         TabIndex        =   36
         Top             =   15
         Width           =   360
         _version        =   65536
         _extentx        =   635
         _extenty        =   635
         _stockprops     =   78
         autosize        =   2
         picture         =   "FRMSTAFF.frx":16F6
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   360
         Left            =   390
         TabIndex        =   33
         Top             =   15
         Width           =   360
         _version        =   65536
         _extentx        =   635
         _extenty        =   635
         _stockprops     =   78
         autosize        =   2
         picture         =   "FRMSTAFF.frx":1A48
      End
      Begin Threed.SSCommand cmdNew 
         Height          =   360
         Left            =   0
         TabIndex        =   32
         Top             =   15
         Width           =   360
         _version        =   65536
         _extentx        =   635
         _extenty        =   635
         _stockprops     =   78
         autosize        =   2
         picture         =   "FRMSTAFF.frx":1B5A
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   360
         Left            =   780
         TabIndex        =   34
         Top             =   15
         Visible         =   0   'False
         Width           =   360
         _version        =   65536
         _extentx        =   635
         _extenty        =   635
         _stockprops     =   78
         autosize        =   2
         picture         =   "FRMSTAFF.frx":1C6C
      End
   End
   Begin Threed.SSFrame FrameStaffDetails 
      Height          =   4395
      Left            =   2340
      TabIndex        =   39
      Top             =   300
      Width           =   5115
      _version        =   65536
      _extentx        =   9022
      _extenty        =   7752
      _stockprops     =   14
      shadowstyle     =   1
      Begin VB.TextBox TextNote 
         Appearance      =   0  'Flat
         DataField       =   "Note"
         BeginProperty Font 
            name            =   "Arial"
            charset         =   1
            weight          =   400
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   2085
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Text            =   "FRMSTAFF.frx":1D7E
         Top             =   2190
         Width           =   2685
      End
      Begin VB.TextBox TextStaffID 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "StaffID"
         BeginProperty Font 
            name            =   "Arial"
            charset         =   1
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   150
         MaxLength       =   10
         TabIndex        =   1
         Text            =   "StaffID"
         Top             =   420
         Width           =   1000
      End
      Begin VB.TextBox TextMiddleName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "MiddleName"
         BeginProperty Font 
            name            =   "Arial"
            charset         =   1
            weight          =   400
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3810
         MaxLength       =   25
         TabIndex        =   4
         Text            =   "MiddleName"
         Top             =   420
         Width           =   1200
      End
      Begin VB.TextBox TextFirstName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "FirstName"
         BeginProperty Font 
            name            =   "Arial"
            charset         =   1
            weight          =   400
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2520
         MaxLength       =   25
         TabIndex        =   3
         Text            =   "FirstName"
         Top             =   420
         Width           =   1200
      End
      Begin VB.TextBox TextLastName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "LastName"
         BeginProperty Font 
            name            =   "Arial"
            charset         =   1
            weight          =   400
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1230
         MaxLength       =   25
         TabIndex        =   2
         Text            =   "LastName"
         Top             =   420
         Width           =   1200
      End
      Begin VB.Frame FrameClass 
         Caption         =   "Staff Classification"
         Height          =   2415
         Left            =   2850
         TabIndex        =   40
         Top             =   1890
         Width           =   2205
         Begin VB.CheckBox CheckClass 
            Alignment       =   1  'Right Justify
            Caption         =   "Class"
            DataField       =   "Class_1"
            BeginProperty Font 
               name            =   "Arial"
               charset         =   1
               weight          =   400
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   90
            TabIndex        =   22
            Top             =   210
            Width           =   2000
         End
         Begin VB.CheckBox CheckClass 
            Alignment       =   1  'Right Justify
            Caption         =   "Class"
            DataField       =   "Class_2"
            BeginProperty Font 
               name            =   "Arial"
               charset         =   1
               weight          =   400
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   90
            TabIndex        =   23
            Top             =   420
            Width           =   2000
         End
         Begin VB.CheckBox CheckClass 
            Alignment       =   1  'Right Justify
            Caption         =   "Class"
            DataField       =   "Class_3"
            BeginProperty Font 
               name            =   "Arial"
               charset         =   1
               weight          =   400
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   2
            Left            =   90
            TabIndex        =   24
            Top             =   630
            Width           =   2000
         End
         Begin VB.CheckBox CheckClass 
            Alignment       =   1  'Right Justify
            Caption         =   "Class"
            DataField       =   "Class_4"
            BeginProperty Font 
               name            =   "Arial"
               charset         =   1
               weight          =   400
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   3
            Left            =   90
            TabIndex        =   25
            Top             =   840
            Width           =   2000
         End
         Begin VB.CheckBox CheckClass 
            Alignment       =   1  'Right Justify
            Caption         =   "Class"
            DataField       =   "Class_5"
            BeginProperty Font 
               name            =   "Arial"
               charset         =   1
               weight          =   400
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   4
            Left            =   90
            TabIndex        =   26
            Top             =   1050
            Width           =   2000
         End
         Begin VB.CheckBox CheckClass 
            Alignment       =   1  'Right Justify
            Caption         =   "Class"
            DataField       =   "Class_6"
            BeginProperty Font 
               name            =   "Arial"
               charset         =   1
               weight          =   400
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   5
            Left            =   90
            TabIndex        =   27
            Top             =   1260
            Width           =   2000
         End
         Begin VB.CheckBox CheckClass 
            Alignment       =   1  'Right Justify
            Caption         =   "Class"
            DataField       =   "Class_7"
            BeginProperty Font 
               name            =   "Arial"
               charset         =   1
               weight          =   400
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   6
            Left            =   90
            TabIndex        =   28
            Top             =   1470
            Width           =   2000
         End
         Begin VB.CheckBox CheckClass 
            Alignment       =   1  'Right Justify
            Caption         =   "Class"
            DataField       =   "Class_8"
            BeginProperty Font 
               name            =   "Arial"
               charset         =   1
               weight          =   400
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   7
            Left            =   90
            TabIndex        =   29
            Top             =   1680
            Width           =   2000
         End
         Begin VB.CheckBox CheckClass 
            Alignment       =   1  'Right Justify
            Caption         =   "Class"
            DataField       =   "Class_9"
            BeginProperty Font 
               name            =   "Arial"
               charset         =   1
               weight          =   400
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   8
            Left            =   90
            TabIndex        =   30
            Top             =   1890
            Width           =   2000
         End
         Begin VB.CheckBox CheckClass 
            Alignment       =   1  'Right Justify
            Caption         =   "Class"
            DataField       =   "Class_10"
            BeginProperty Font 
               name            =   "Arial"
               charset         =   1
               weight          =   400
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   9
            Left            =   90
            TabIndex        =   31
            Top             =   2100
            Width           =   2000
         End
      End
      Begin VB.Frame FrameDays 
         Caption         =   "Availability"
         BeginProperty Font 
            name            =   "Arial"
            charset         =   1
            weight          =   400
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   2850
         TabIndex        =   53
         Top             =   1890
         Visible         =   0   'False
         Width           =   2205
         Begin VB.CommandButton cmdSelectNone 
            Caption         =   "Select None"
            BeginProperty Font 
               name            =   "Arial"
               charset         =   1
               weight          =   400
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   1140
            TabIndex        =   21
            Top             =   1980
            Width           =   1000
         End
         Begin VB.CommandButton cmdSelectAll 
            Caption         =   "Select All"
            BeginProperty Font 
               name            =   "Arial"
               charset         =   1
               weight          =   400
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   60
            TabIndex        =   20
            Top             =   1980
            Width           =   1000
         End
         Begin VB.CheckBox CheckDay 
            Alignment       =   1  'Right Justify
            Caption         =   "Day"
            DataField       =   "Class_8"
            BeginProperty Font 
               name            =   "Arial"
               charset         =   1
               weight          =   400
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   6
            Left            =   90
            TabIndex        =   19
            Top             =   1470
            Width           =   2000
         End
         Begin VB.CheckBox CheckDay 
            Alignment       =   1  'Right Justify
            Caption         =   "Day"
            DataField       =   "Class_6"
            BeginProperty Font 
               name            =   "Arial"
               charset         =   1
               weight          =   400
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   5
            Left            =   90
            TabIndex        =   18
            Top             =   1260
            Width           =   2000
         End
         Begin VB.CheckBox CheckDay 
            Alignment       =   1  'Right Justify
            Caption         =   "Day"
            DataField       =   "Class_5"
            BeginProperty Font 
               name            =   "Arial"
               charset         =   1
               weight          =   400
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   4
            Left            =   90
            TabIndex        =   17
            Top             =   1050
            Width           =   2000
         End
         Begin VB.CheckBox CheckDay 
            Alignment       =   1  'Right Justify
            Caption         =   "Day"
            DataField       =   "Class_4"
            BeginProperty Font 
               name            =   "Arial"
               charset         =   1
               weight          =   400
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   3
            Left            =   90
            TabIndex        =   16
            Top             =   840
            Width           =   2000
         End
         Begin VB.CheckBox CheckDay 
            Alignment       =   1  'Right Justify
            Caption         =   "Day"
            DataField       =   "Class_3"
            BeginProperty Font 
               name            =   "Arial"
               charset         =   1
               weight          =   400
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   2
            Left            =   90
            TabIndex        =   15
            Top             =   630
            Width           =   2000
         End
         Begin VB.CheckBox CheckDay 
            Alignment       =   1  'Right Justify
            Caption         =   "Day"
            DataField       =   "Class_2"
            BeginProperty Font 
               name            =   "Arial"
               charset         =   1
               weight          =   400
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   90
            TabIndex        =   13
            Top             =   420
            Width           =   2000
         End
         Begin VB.CheckBox CheckDay 
            Alignment       =   1  'Right Justify
            Caption         =   "Day"
            DataField       =   "Class_1"
            BeginProperty Font 
               name            =   "Arial"
               charset         =   1
               weight          =   400
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   90
            TabIndex        =   14
            Top             =   210
            Width           =   2000
         End
      End
      Begin Threed.SSCommand cmdClassDayToggle 
         Height          =   345
         Left            =   4680
         TabIndex        =   12
         Top             =   1530
         Width           =   345
         _version        =   65536
         _extentx        =   609
         _extenty        =   609
         _stockprops     =   78
         autosize        =   2
         picture         =   "FRMSTAFF.frx":1D87
      End
      Begin MSMask.MaskEdBox MaskMaxHours 
         DataField       =   "HourRate"
         Height          =   300
         Left            =   4200
         TabIndex        =   11
         Top             =   1140
         Width           =   795
         _version        =   65536
         _extentx        =   1402
         _extenty        =   529
         _stockprops     =   109
         forecolor       =   0
         backcolor       =   16777215
         BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
            name            =   "Arial"
            charset         =   1
            weight          =   400
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         borderstyle     =   1
         autotab         =   -1  'True
         promptinclude   =   0   'False
         clipmode        =   1
         format          =   "#0.00"
         appearance      =   0
      End
      Begin MSMask.MaskEdBox MaskMinHours 
         DataField       =   "HourRate"
         Height          =   300
         Left            =   4200
         TabIndex        =   10
         Top             =   840
         Width           =   795
         _version        =   65536
         _extentx        =   1402
         _extenty        =   529
         _stockprops     =   109
         forecolor       =   0
         backcolor       =   16777215
         BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
            name            =   "Arial"
            charset         =   1
            weight          =   400
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         borderstyle     =   1
         autotab         =   -1  'True
         promptinclude   =   0   'False
         clipmode        =   1
         format          =   "#0.00"
         appearance      =   0
      End
      Begin VB.Label Label_std 
         BackStyle       =   0  'Transparent
         Caption         =   "Hours Max"
         BeginProperty Font 
            name            =   "Arial"
            charset         =   1
            weight          =   400
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   15
         Left            =   3360
         TabIndex        =   52
         Top             =   1170
         Width           =   1635
      End
      Begin VB.Label Label_std 
         BackStyle       =   0  'Transparent
         Caption         =   "Hours Min"
         BeginProperty Font 
            name            =   "Arial"
            charset         =   1
            weight          =   400
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   14
         Left            =   3360
         TabIndex        =   51
         Top             =   870
         Width           =   1635
      End
      Begin VB.Label Label_std 
         BackStyle       =   0  'Transparent
         Caption         =   "Hourly Rate"
         BeginProperty Font 
            name            =   "Arial"
            charset         =   1
            weight          =   400
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   12
         Left            =   120
         TabIndex        =   50
         Top             =   1530
         Width           =   2115
      End
      Begin VB.Label Label_std 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Employed"
         BeginProperty Font 
            name            =   "Arial"
            charset         =   1
            weight          =   400
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   11
         Left            =   120
         TabIndex        =   49
         Top             =   1200
         Width           =   2115
      End
      Begin VB.Label Label_std 
         BackStyle       =   0  'Transparent
         Caption         =   "Birth Date"
         BeginProperty Font 
            name            =   "Arial"
            charset         =   1
            weight          =   400
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   120
         TabIndex        =   48
         Top             =   900
         Width           =   2115
      End
      Begin VB.Label Label_std 
         BackStyle       =   0  'Transparent
         Caption         =   "Home Phone"
         BeginProperty Font 
            name            =   "Arial"
            charset         =   1
            weight          =   400
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   13
         Left            =   120
         TabIndex        =   47
         Top             =   1860
         Width           =   2475
      End
      Begin VB.Label Label_std 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Middle Name"
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
         Index           =   3
         Left            =   3840
         TabIndex        =   46
         Top             =   210
         Width           =   900
      End
      Begin VB.Label Label_std 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
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
         Index           =   2
         Left            =   2550
         TabIndex        =   45
         Top             =   210
         Width           =   765
      End
      Begin VB.Label Label_std 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name"
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
         Index           =   1
         Left            =   1260
         TabIndex        =   44
         Top             =   210
         Width           =   765
      End
      Begin VB.Label Label_std 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Staff ID"
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
         Left            =   180
         TabIndex        =   43
         Top             =   210
         Width           =   540
      End
      Begin MSMask.MaskEdBox MaskHourRate 
         DataField       =   "HourRate"
         Height          =   300
         Left            =   1230
         TabIndex        =   7
         Top             =   1500
         Width           =   1005
         _version        =   65536
         _extentx        =   1773
         _extenty        =   529
         _stockprops     =   109
         forecolor       =   0
         backcolor       =   16777215
         BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
            name            =   "Arial"
            charset         =   1
            weight          =   400
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         borderstyle     =   1
         autotab         =   -1  'True
         promptinclude   =   0   'False
         clipmode        =   1
         format          =   "Currency"
         appearance      =   0
      End
      Begin MSMask.MaskEdBox MaskDateHired 
         DataField       =   "DateHired"
         Height          =   300
         Left            =   1230
         TabIndex        =   6
         Top             =   1170
         Width           =   1000
         _version        =   65536
         _extentx        =   1773
         _extenty        =   529
         _stockprops     =   109
         forecolor       =   0
         backcolor       =   16777215
         BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
            name            =   "Arial"
            charset         =   1
            weight          =   400
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         borderstyle     =   1
         autotab         =   -1  'True
         promptinclude   =   0   'False
         clipmode        =   1
         format          =   "Short Date"
         appearance      =   0
      End
      Begin MSMask.MaskEdBox MaskBirthDate 
         DataField       =   "Birthdate"
         Height          =   300
         Left            =   1230
         TabIndex        =   5
         Top             =   840
         Width           =   1000
         _version        =   65536
         _extentx        =   1773
         _extenty        =   529
         _stockprops     =   109
         forecolor       =   0
         backcolor       =   16777215
         BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
            name            =   "Arial"
            charset         =   1
            weight          =   400
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         borderstyle     =   1
         autotab         =   -1  'True
         promptinclude   =   0   'False
         clipmode        =   1
         format          =   "Short Date"
         appearance      =   0
      End
      Begin MSMask.MaskEdBox MaskHomePhone 
         DataField       =   "HomePhone"
         Height          =   300
         Left            =   1230
         TabIndex        =   8
         Top             =   1830
         Width           =   1365
         _version        =   65536
         _extentx        =   2408
         _extenty        =   529
         _stockprops     =   109
         forecolor       =   0
         backcolor       =   16777215
         BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
            name            =   "Arial"
            charset         =   1
            weight          =   400
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         borderstyle     =   1
         autotab         =   -1  'True
         promptinclude   =   0   'False
         clipmode        =   1
         maxlength       =   13
         mask            =   "##9 #####999"
         appearance      =   0
      End
      Begin VB.Label LabelDaysEmployed 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "DaysEmp"
         BeginProperty Font 
            name            =   "Arial"
            charset         =   1
            weight          =   400
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2280
         TabIndex        =   42
         Top             =   1200
         Width           =   1005
      End
      Begin VB.Label LabelAge 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         BeginProperty Font 
            name            =   "Arial"
            charset         =   1
            weight          =   400
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2280
         TabIndex        =   41
         Top             =   900
         Width           =   1005
      End
      Begin VB.Shape ShapeBGround 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         Height          =   615
         Index           =   0
         Left            =   90
         Top             =   180
         Width           =   4965
      End
      Begin VB.Shape ShapeBGround 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         Height          =   765
         Index           =   1
         Left            =   3300
         Top             =   750
         Width           =   1755
      End
   End
   Begin VB.Label Label_std 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Staff List"
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
      Index           =   16
      Left            =   30
      TabIndex        =   37
      Top             =   420
      Width           =   660
   End
End
Attribute VB_Name = "frmStaff"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit
'[----------------------------------------------]
'[frmStaff          Staff Details form          ]
'[----------------------------------------------]
'[GSR                                           ]
'[generic staff roster system for Windows       ]
'[(C) ---------- David Gilbert, 1997            ]
'[----------------------------------------------]
'[version                           2.1         ]
'[initial date                      21/07/1997  ]
'[----------------------------------------------]

Dim Shared intStartDay As Integer       '[Start time - week day]
Dim Shared intFinishDay As Integer      '[Finish time - week day]



Private Sub CheckAvailable_Click()

    '[ENABLE SAVE BUTTON IF A CHANGE HAS BEEN MADE]
    frmStaff.cmdSave.Visible = True

End Sub

Private Sub CheckClass_Click(Index As Integer)

    '[MAKE SAVE BUTTON VISIBLE]
    frmStaff.cmdSave.Visible = True
    
End Sub


Private Sub CheckClass_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "These check boxes determine which rosters (or 'classes') this staff member is available to."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub CheckDay_Click(Index As Integer)
    
    '[MAKE SAVE BUTTON VISIBLE]
    frmStaff.cmdSave.Visible = True

End Sub

Private Sub CheckDay_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "These check boxes determine on which days the staff member is available to be placed into a roster."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdAllStaffRosters_Click()
    
    '[CALL ROUTINE TO PROCESS ALL STAFF RECORDS AND PRINT TIME SHEETS]
    Call procAllStaffRosters

End Sub

Private Sub cmdAllStaffRosters_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Click this button to print the weekly roster for all staff members in the list who have been allocated to rosters."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdClassDayToggle_Click()

    '[TOGGLE BETWEEN CLASS FRAME AND DAY AVAILABLE FRAM]
    If frmStaff.FrameDays.Visible = True Then
        frmStaff.FrameDays.Visible = False
        frmStaff.FrameClass.Visible = True
    Else
        frmStaff.FrameDays.Visible = True
        frmStaff.FrameClass.Visible = False
    End If

End Sub

Private Sub cmdClassDayToggle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Click here to switch between the 'Days Available' and the 'Rosters Available' boxes."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdDelete_Click()

    '[*********************************************************************]
    '[DON'T GET TOO COMPLICATED IN HERE, OTHERWISE OTHER EVENTS MAY TRIGGER]
    '[*********************************************************************]
    
    '[DELETE A RECORD FROM THE STAFF LIST]
    Dim strDisplayname      As String
    Dim intCounter          As Integer
    
        '[POPUP YES/NO DIALOG TO CONFIRM DELETION]
        Dim Msg As String
        Dim Style
        Dim Response
        Dim Title
        
        If frmStaff.ListStaff.ListIndex = -1 Then Exit Sub
        
        '[CHECK DELETION FLAG]
        If flagDeleteConfirm Then
            Msg = "You have chosen to delete this staff member - " & DsStaff("LastName") & ", " & DsStaff("FirstName") & "." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "This will permanently remove this staff member from the list." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "Do you wish to continue ?"
            Style = vbYesNo ' Define buttons.
            Title = "Confirmation Required"  ' Define title.
            Response = gsrMsg(Msg, Style, Title)
        Else
            Response = vbYes
        End If
        
        If Response = vbYes Then    ' User chose Yes.
            
            '[DELETE STAFF MEMBER FROM DYNASET]
            DsStaff.Delete
            
            '[SAVE DISPLAYED NAME TO TEMPORARY STRING]
            If frmStaff.ListStaff.ListCount > 1 Then
                strDisplayname = frmStaff.ListStaff.List(frmStaff.ListStaff.ListIndex - 1)
            Else
                strDisplayname = ""
            End If
            
            '[FILL STAFF LIST SO WE GET ORDER]
            FillStaffList
                
            '[RELOCATE STAFF NAME]
            For intCounter = 0 To (frmStaff.ListStaff.ListCount - 1)
                If frmStaff.ListStaff.List(intCounter) = strDisplayname Then
                    frmStaff.ListStaff.ListIndex = intCounter
                    Exit For
                End If
            Next intCounter
        
            '[CHECK FOR NO STAFF MEMBERS]
            If DsStaff.RecordCount = 0 Then
                '[CALL ROUTINE TO ADD NEW STAFF MEMBER AND REPOSITION LIST]
                AddNewStaff
            End If
        
            '[MOVE TO FIRST ITEM IF NO ITEM IS SELECTED]
            If frmStaff.ListStaff.ListIndex < 0 Then frmStaff.ListStaff.ListIndex = 0
        
        End If
    

End Sub

Private Sub cmdDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Click this button to delete the currently selected staff member (" & frmStaff.ListStaff.List(frmStaff.ListStaff.ListIndex) & ") from the list."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdNew_Click()

    Dim Msg As String
    Dim Style
    Dim Response
    Dim Title
    '[ADD A NEW STAFF MEMBER TO THE LIST]
    '[CHECK TO SEE IF WE ARE MOVING FROM AN UNSAVED RECORD]
    If frmStaff.cmdSave.Visible = True Then
        '[RECORD IS UNSAVED, POPUP YES/NO DIALOG]
        Msg = "You have made changes to this staff record (" & DsStaff("Lastname") & ", " & DsStaff("FirstName") & ") but have not saved these changes." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "If you choose not to save now, any changes you have made since your last save will be lost." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "Do you wish to save these changes before you add another staff member ?"
        Style = vbYesNoCancel              ' Define buttons.
        Title = "Staff Changes Not Saved"  ' Define title.
        Response = gsrMsg(Msg, Style, Title)
        If Response = vbYes Then    ' User chose Yes.
            '[COMMIT RECORD CHANGES TO THE DYNASET]
            SaveStaffDetails
        ElseIf Response = vbCancel Then
            Exit Sub
        End If
    End If
    
    '[CALL ROUTINE TO ADD NEW STAFF MEMBER AND REPOSITION LIST]
    AddNewStaff

End Sub

Private Sub cmdNew_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Click this button to add a new staff member to the list."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdSave_Click()

    Dim strDisplayname      As String
    Dim intCounter          As Integer

    '[COMMIT RECORD CHANGES TO THE DYNASET]
    SaveStaffDetails
    
    '[CONVERT DYNASET LASTNAME, FIRSTNAME TO DISPLAY FORMAT]
    strDisplayname = DsStaff("LastName") & ", " & DsStaff("FirstName")
    
    '[IF NAME HASN'T CHANGED, EXIT THIS SUB NOW]
    If frmStaff.ListStaff.List(frmStaff.ListStaff.ListIndex) = strDisplayname Then Exit Sub
    
    '[FILL STAFF LIST SO WE GET ORDER]
    FillStaffList
        
    '[RELOCATE STAFF NAME]
    For intCounter = 0 To (frmStaff.ListStaff.ListCount - 1)
        If frmStaff.ListStaff.List(intCounter) = strDisplayname Then frmStaff.ListStaff.ListIndex = intCounter
    Next intCounter
    
End Sub

Private Sub cmdSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Click this button to any changes made to the current staff record."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdSelectAll_Click()
    
    '[SELECT ALL AVAILABLE DAYS]
    Dim intCounter      As Integer
    For intCounter = 1 To 7
        frmStaff.CheckDay(intCounter - 1).Value = vbChecked
    Next intCounter
    
    '[MAKE SAVE BUTTON VISIBLE]
    frmStaff.cmdSave.Visible = True

End Sub

Private Sub cmdSelectAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Click here to mark all days as selected."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdSelectedStaffRoster_Click()

    '[CALL ROUTINE TO PROCESS SINGLE STAFF RECORD AND PRINT TIME SHEET]
    Call procSelectedStaffRoster

End Sub

Private Sub cmdSelectedStaffRoster_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Click this button to print the weekly roster for " & frmStaff.ListStaff.List(frmStaff.ListStaff.ListIndex) & "."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub cmdSelectNone_Click()
    
    '[SELECT NO AVAILABLE DAYS]
    Dim intCounter      As Integer
    For intCounter = 1 To 7
        frmStaff.CheckDay(intCounter - 1).Value = vbUnchecked
    Next intCounter
    
    '[MAKE SAVE BUTTON VISIBLE]
    frmStaff.cmdSave.Visible = True

End Sub


Private Sub cmdSelectNone_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Click here to mark all days as unselected."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub Form_Load()

    '[FILL STAFF LIST FROM DYNASET]
    FillStaffList
    '[APPLY CLASS LABELS]
    SetClassLabels
    '[APPLY DAY LABELS]
    SetDayLabels
    '[LOCATE FIRST STAFF MEMBER]
    If frmStaff.ListStaff.ListCount > 0 Then frmStaff.ListStaff.ListIndex = 0
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "The staff form allows you to add, delete and modify staff records."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub Form_Resize()

    '[RESIZE CONTROLS ON THIS FORM]
    
    '[TEMP VARIABLES SO WE CAN CATCH ILLEGAL WIDTHS]
    Dim sinWidth        As Single
    Dim sinHeight       As Single
    
    '[IF FORM IS MINIMISED THEN EXIT THIS ROUTINE]
    If frmStaff.WindowState = 1 Then Exit Sub
    '[RESIZE GRID AND LIST BOXES AND ARRANGE CONTROLS ON FORM]
    
    '[STAFF LIST]
    sinHeight = frmStaff.Height - frmStaff.ListStaff.Top - frmStaff.PanelToolBar.Height
    If sinHeight > 0 Then frmStaff.ListStaff.Height = sinHeight

End Sub


Private Sub Label_std_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Select Case Index
    Case 10
        '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
        StatusBar "Enter the birth date of the staff member here."
        '[---------------------------------------------------------------------------------]
    Case 11
        '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
        StatusBar "Enter the employment date of the staff member here."
        '[---------------------------------------------------------------------------------]
    Case 12
        '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
        StatusBar "Enter the hourly rate of the staff member here."
        '[---------------------------------------------------------------------------------]
    Case 13
        '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
        StatusBar "Enter the home phone number of the staff member here."
        '[---------------------------------------------------------------------------------]
    Case 14
        '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
        StatusBar "If the staff member requires a minimum number of hours a week, enter that figure here."
        '[---------------------------------------------------------------------------------]
    Case 15
        '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
        StatusBar "Enter the maximum hours a week that this staff member can/may work."
        '[---------------------------------------------------------------------------------]
    Case Else
    End Select
    
End Sub


Private Sub ListStaff_Click()

    '[LOCATE SELECTED STAFF MEMBER]
    LocateStaff

End Sub

Private Sub ListStaff_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
    Case vbKeyDelete
        '[CAPTURE DELETE KEY]
        Call cmdDelete_Click
    Case Else
    End Select
    
End Sub


Private Sub ListStaff_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Click on a name in this list to display the details for the selected staff member."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub MaskBirthDate_Change()

    '[ENABLE SAVE BUTTON IF A CHANGE HAS BEEN MADE]
    frmStaff.cmdSave.Visible = True

End Sub

Private Sub MaskDateHired_Change()

    '[ENABLE SAVE BUTTON IF A CHANGE HAS BEEN MADE]
    frmStaff.cmdSave.Visible = True

End Sub

Private Sub MaskEmerPhone_Change()


    '[ENABLE SAVE BUTTON IF A CHANGE HAS BEEN MADE]
    frmStaff.cmdSave.Visible = True

End Sub

Private Sub MaskHomePhone_Change()


    '[ENABLE SAVE BUTTON IF A CHANGE HAS BEEN MADE]
    frmStaff.cmdSave.Visible = True

End Sub

Private Sub MaskHourRate_Change()


    '[ENABLE SAVE BUTTON IF A CHANGE HAS BEEN MADE]
    frmStaff.cmdSave.Visible = True

End Sub

Private Sub MaskWorkPhone_Change()


    '[ENABLE SAVE BUTTON IF A CHANGE HAS BEEN MADE]
    frmStaff.cmdSave.Visible = True

End Sub

Private Sub TextDetail_Change(Index As Integer)


    '[ENABLE SAVE BUTTON IF A CHANGE HAS BEEN MADE]
    frmStaff.cmdSave.Visible = True

End Sub

Private Sub MaskMaxHours_Change()
    
    '[ENABLE SAVE BUTTON IF A CHANGE HAS BEEN MADE]
    frmStaff.cmdSave.Visible = True

End Sub

Private Sub MaskMinHours_Change()
    
    '[ENABLE SAVE BUTTON IF A CHANGE HAS BEEN MADE]
    frmStaff.cmdSave.Visible = True

End Sub

Private Sub PanelToolBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "The staff form allows you to add, delete and modify staff records."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub TextFirstName_Change()


    '[ENABLE SAVE BUTTON IF A CHANGE HAS BEEN MADE]
    frmStaff.cmdSave.Visible = True

End Sub

Private Sub TextFirstName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Enter the first name of the staff member here (or an initial if more space is required)."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub TextLastName_Change()

    '[ENABLE SAVE BUTTON IF A CHANGE HAS BEEN MADE]
    frmStaff.cmdSave.Visible = True

End Sub

Private Sub TextLastName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Enter the staff members last name here."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub TextMiddleName_Change()


    '[ENABLE SAVE BUTTON IF A CHANGE HAS BEEN MADE]
    frmStaff.cmdSave.Visible = True

End Sub

Private Sub TextMiddleName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Enter the middle name of the staff member here. This field is not compulsory."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub TextNote_Change()


    '[ENABLE SAVE BUTTON IF A CHANGE HAS BEEN MADE]
    frmStaff.cmdSave.Visible = True

End Sub

Private Sub TextNote_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Enter any notes for this staff member here. These notes will be printed on the individual staff roster."
    '[---------------------------------------------------------------------------------]

End Sub


Private Sub TextStaffID_Change()

    '[ENABLE SAVE BUTTON IF A CHANGE HAS BEEN MADE]
    frmStaff.cmdSave.Visible = True

End Sub


Private Sub TextStaffID_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
    StatusBar "Enter a unique staff identification code here."
    '[---------------------------------------------------------------------------------]

End Sub


