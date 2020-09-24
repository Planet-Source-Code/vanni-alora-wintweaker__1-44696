Attribute VB_Name = "modFunctions"
Option Explicit

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function SetComputerName Lib "kernel32" Alias "SetComputerNameA" (ByVal lpComputerName As String) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim hWnd As Long
Dim conSwNormal As Long

Public strTemp As String

Public Function ColoredValue(ByVal chkBox As CheckBox)
    On Error Resume Next
    If chkBox.Value = 1 Then
        chkBox.ForeColor = vbRed
    Else
        chkBox.ForeColor = vbBlack
    End If
End Function

Public Function Description(ByVal object As Object, ByVal label As label, ByVal myDescription As String)
    label.Caption = myDescription
End Function

Public Sub GetMachineName(ByVal txtBox As TextBox)
    ' Current Machine Name
    strTemp = String(255, Chr(0))
    GetComputerName strTemp, 255
    strTemp = Replace(strTemp, Chr(0), "")
    txtBox.Text = UCase(strTemp)
End Sub

Public Sub SetMachineName(ByVal txtBox As TextBox)
    ' change machine name...
    If txtBox.Text = "" Then
        MsgBox "Please enter a valid Machine Name.", vbCritical, "Error"
        Exit Sub
    ElseIf frmMain.chkNetworkChangeName.Value = 0 Then
        Exit Sub
    Else
        SetComputerName UCase(txtBox.Text)
    End If
End Sub

' go to website...
Public Sub GoToUrl(ByVal strUrl As String)
    ShellExecute hWnd, "open", strUrl, vbNullString, vbNullString, conSwNormal
End Sub

' send email...
Public Sub SendEmail(strTo As String, strSubject As String)
    Dim ret As Integer
    Dim vbnormailfocus As Long
    'Open the default email client with the strings passed to it.
    'The second string can be ommited
    ret = ShellExecute(0, vbNullString, "mailto:" & strTo & "?subject=" & strSubject, vbNullString, vbNullString, vbnormailfocus)
End Sub

