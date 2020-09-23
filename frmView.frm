VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmView 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   8610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox txtThoughts 
      Height          =   5805
      Left            =   0
      TabIndex        =   0
      Top             =   90
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   10239
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmView.frx":0000
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Print #1, strUser & " has viewed their chatlog. " & Now
Close #100
Open App.Path & "\logs\chat" & strUser & ".log" For Input As #99
    txtThoughts.Text = Input$(LOF(99), 99)
Close #99
Open App.Path & "\logs\chat" & strUser & ".log" For Append As #100

End Sub
