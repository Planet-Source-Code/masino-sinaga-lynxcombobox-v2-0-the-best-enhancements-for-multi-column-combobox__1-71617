VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWait 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loading Data ..."
   ClientHeight    =   1290
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5115
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   5115
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar prgBar1 
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblWait 
      Caption         =   "Please wait while program loading data from database ..."
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   4455
   End
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
