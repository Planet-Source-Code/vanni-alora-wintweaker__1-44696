Attribute VB_Name = "modWExplorer"
Option Explicit

Public Sub ApplyWExplorer()
    On Error Resume Next
    ' apply Windows Explorer settings...
    ' NoViewContextMenu...
    If frmMain.chkWExplorer(0).Value = 1 Then
        modRegistry1.SetBinaryValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoViewContextMenu", Chr$(&H1) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0)
    Else
        modRegistry2.DeleteSetting "", "Explorer", "NoViewContextMenu", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' NoFileMenu...
    If frmMain.chkWExplorer(1).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFileMenu", 1
    Else
        modRegistry2.DeleteSetting "", "Explorer", "NoFileMenu", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' SmallIcons...
    If frmMain.chkWExplorer(2).Value = 1 Then
        modRegistry1.SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\SmallIcons", "SmallIcons", "YES"
    Else
        modRegistry1.SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\SmallIcons", "SmallIcons", "NO"
    End If
    
    ' NoBandCustomize...
    If frmMain.chkWExplorer(3).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoBandCustomize", 1
    Else
        modRegistry2.DeleteSetting "", "Explorer", "NoBandCustomize", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' NoDrives...
    If frmMain.chkWExplorer(4).Value = 1 Then
        modRegistry1.SetBinaryValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDrives", Chr$(&HFF) + Chr$(&HFF) + Chr$(&HFF) + Chr$(&H3)
    Else
        modRegistry2.DeleteSetting "", "Explorer", "NoDrives", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
End Sub

Public Sub ReadWExplorer()
    On Error Resume Next
    ' read Windows Explorer settings...
    ' NoViewContextMenu
    If modRegistry1.GetBinaryValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoViewContextMenu") = Chr$(&H1) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0) Then
        frmMain.chkWExplorer(0).Value = 1
    Else
        frmMain.chkWExplorer(0).Value = 0
    End If
    ' NoFileMenu
    frmMain.chkWExplorer(1).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFileMenu")
    ' SmallIcons
    If modRegistry1.GetStringValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\SmallIcons", "SmallIcons") = "YES" Then
        frmMain.chkWExplorer(2).Value = 1
    Else
        frmMain.chkWExplorer(2).Value = 0
    End If
    ' NoBandCustomize
    frmMain.chkWExplorer(3).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoBandCustomize")
    ' NoDrives
    If modRegistry1.GetBinaryValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDrives") = Chr$(&HFF) + Chr$(&HFF) + Chr$(&HFF) + Chr$(&H3) Then
        frmMain.chkWExplorer(4).Value = 1
    Else
        frmMain.chkWExplorer(4).Value = 0
    End If
End Sub
