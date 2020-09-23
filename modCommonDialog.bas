Attribute VB_Name = "modCommonDialog"
Public CurrentFileName As String     ' Holds filename
Public pos As Double                 ' Counts position in file
Public Char As Byte                  ' Binary Access String
Public Char2 As String
Public vCount(1 To 65) As String 'Holds size of each file
Public FileString(1 To 65) As String 'Holds name of each file
Public Position(1 To 65) As String 'Holds the starting position(in bytes) of each file
Public NewPosition As Long
Public CurrentCount As Double 'Last position in appended file
Public vSize As Long 'Size of file
Public lngOffset As Long 'Magic Offset used to extract from the appended file
Public Sub OpenFile(vForm As Form)
' Handle errors
    On Error GoTo OpenProblem
    vForm.CommonDialog1.InitDir = App.Path
    vForm.CommonDialog1.Filter = "ques.txt|ques.txt|All Text Files|*.txt"
    vForm.CommonDialog1.FilterIndex = 1
    ' Display an Open dialog box.
    vForm.CommonDialog1.Action = 1
    
    ' vForm.Caption = strAppName & " - " & vForm.CommonDialog1.FileName
    ' vForm.Caption = strAppName & " - " & vForm.CommonDialog1.FileName
    CurrentFileName = vForm.CommonDialog1.filename
    ImportFile

    Exit Sub
OpenProblem:
    ' Cancel button clicked
    Exit Sub

End Sub

Public Sub SaveFile(vForm As Form)
On Error GoTo SaveERR
    Dim FileNum As Integer
    ' Set Initial Directory to open and FileTypes
        vForm.CommonDialog1.InitDir = App.Path & "\Save"
        vForm.CommonDialog1.Filter = "ALL Files | *.*"
        
        'CurrentFileName = 'FileString(frmMain.lstFiles.ListIndex)
        If CurrentFileName = "" Then
            'vForm.CommonDialog1.FileName = FileString(frmMain.lstFiles.ListIndex)
        Else
            vForm.CommonDialog1.filename = CurrentFileName
        End If
        
            vForm.CommonDialog1.ShowSave
            CurrentFileName = vForm.CommonDialog1.filename
Exit Sub
SaveERR:
        ' I don't know what the error was, but I want to let you know and then
        ' Exit the sub.
End Sub

Public Sub ImportFile()
Dim i As Long
Dim g As Long
Dim j As Long
Dim x As Long
'frmPgress.Show
i = FileLen(CurrentFileName)
frmMain.pb1.Min = 1
frmMain.pb1.Max = 100

Open CurrentFileName For Input As #99
Open App.Path & "\data\general.dat" For Append As #98
    Do While Not EOF(99)
    g = g + 1
    If g >= 99 Then
        g = 98
    End If
    frmMain.pb1.Value = g
    frmMain.Refresh
    Input #99, strQuestion, strAnswer
    Write #98, strQuestion, strAnswer
    DoEvents
Loop
Close #98
Close #99
frmMain.pb1.Value = frmMain.pb1.Max
MsgBox "Import complete...", vbOKOnly, "Importing"
End Sub
Public Function PutFileInString(sFileName As String) As String
    'sFileName must include Path and file na
    '     me
    'eg "c:\Windows\notepad.exe"
    Dim iFree As Integer, sizeOfFile As Long
    Dim sFileString As String, sTemp As String
    iFree = FreeFile
    Open sFileName For Binary Access Read As iFree
    sizeOfFile = LOF(iFree)
    sFileString = Space$(sizeOfFile)
    Get iFree, , sFileString
    Close #iFree
    PutFileInString = sFileString
    
    ' Thanks to Robert Carter for this function
End Function
