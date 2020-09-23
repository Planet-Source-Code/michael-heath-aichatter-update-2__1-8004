VERSION 5.00
Begin VB.Form frmDebug 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sentence Debug - Future use"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   2715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2985
      ItemData        =   "Form1.frx":0000
      Left            =   30
      List            =   "Form1.frx":0002
      TabIndex        =   0
      Top             =   0
      Width           =   2625
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Left = 0
Me.Top = 0

End Sub
