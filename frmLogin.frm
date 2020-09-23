VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Please Identify Yourself."
   ClientHeight    =   615
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   615
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLogin 
      Caption         =   "&Login"
      Height          =   345
      Left            =   3330
      TabIndex        =   1
      Top             =   120
      Width           =   1245
   End
   Begin VB.TextBox txtLogin 
      Height          =   345
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   3105
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLogin_Click()
' Login as a user
If txtLogin.Text = "" Then
    MsgBox "You must create a username so I know what to call you.", vbOKOnly + vbCritical, "Login Error"
    Exit Sub
End If
strUser = txtLogin.Text
Main
Unload Me
End Sub

Private Sub Form_Load()
' center the form
frmCenter Me
End Sub


Private Sub txtLogin_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
    cmdLogin_Click
End If
End Sub
