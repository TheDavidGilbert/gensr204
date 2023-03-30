VERSION 4.00
Begin VB.Form frmEmpList 
   Caption         =   "Employee List"
   ClientHeight    =   5055
   ClientLeft      =   1785
   ClientTop       =   4560
   ClientWidth     =   9840
   Height          =   5460
   Icon            =   "FRMEMPLI.frx":0000
   Left            =   1725
   LinkTopic       =   "frmEmpList"
   MDIChild        =   -1  'True
   ScaleHeight     =   5055
   ScaleWidth      =   9840
   Top             =   4215
   Width           =   9960
   Begin VB.Data DataStaffList 
      Caption         =   "DataStaffList"
      Connect         =   "Access"
      DatabaseName    =   "C:\CONTRACT\GSR\GSR.MDB"
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Employees"
      Top             =   4560
      Visible         =   0   'False
      Width           =   2745
   End
   Begin MSDBCtls.DBList DBStaffList 
      Bindings        =   "FRMEMPLI.frx":0442
      DataSource      =   "DataStaffList"
      Height          =   4350
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   2745
      _version        =   65536
      _extentx        =   4842
      _extenty        =   7673
      _stockprops     =   77
      forecolor       =   -2147483640
      backcolor       =   -2147483643
      listfield       =   "LastName"
      boundcolumn     =   "LastName"
   End
End
Attribute VB_Name = "frmEmpList"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Private Sub DBListEmpList_Click()

    '[LOCATE RECORD IN RECORDSET]
'    DataEmpList.Recordset.FindFirst "StaffID = " & DBListEmpList.BoundText

    
End Sub


