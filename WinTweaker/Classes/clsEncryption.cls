VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEncryption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function Cryption(strSource As String, strPassword As String, EnDeCrypt As Boolean) As String
    Dim intPassword As Long
    Dim intCrypt As Long

    For X = 1 To Len(strPassword)
        intPassword = intPassword + Asc(Mid$(strPassword, X, 1))
    Next X

    For X = 1 To Len(strSource)

        If EnDeCrypt = True Then
            intCrypt = Asc(Mid$(strSource, X, 1)) + intPassword + X

            Do Until intCrypt <= 255
                intCrypt = intCrypt - 255
            Loop
        Else
            intCrypt = Asc(Mid$(strSource, X, 1)) - intPassword - X

            Do Until intCrypt > 0
                intCrypt = intCrypt + 255
            Loop
        End If
        Cryption = Cryption & Chr(intCrypt)
    Next X
End Function

