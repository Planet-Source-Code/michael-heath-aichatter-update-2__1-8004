VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPgress 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Importanting file...."
   ClientHeight    =   780
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5700
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   780
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   225
      Left            =   30
      TabIndex        =   0
      Top             =   450
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Importing file... please wait."
      Height          =   225
      Left            =   30
      TabIndex        =   1
      Top             =   120
      Width           =   5505
   End
End
Attribute VB_Name = "frmPgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
frmCenter Me
End Sub
