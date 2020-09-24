Attribute VB_Name = "modDesktop"
Option Explicit

Public Sub ApplyDesktop()
    On Error Resume Next
    ' apply Desktop settings...
    ' NoSaveSettings...
    If frmMain.chkDesktop(0).Value = 1 Then
        modRegistry1.SetBinaryValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSaveSettings", Chr$(&H1) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0)
    Else
        modRegistry2.DeleteSetting "", "Explorer", "NoSaveSettings", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' NoDesktop...
    If frmMain.chkDesktop(1).Value = 1 Then
        modRegistry1.SetBinaryValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDesktop", Chr$(&H1) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0)
    Else
        modRegistry2.DeleteSetting "", "Explorer", "NoDesktop", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' NoActiveDesktop...
    If frmMain.chkDesktop(2).Value = 1 Then
        modRegistry1.SetBinaryValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoActiveDesktop", Chr$(&H1) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0)
    Else
        modRegistry2.DeleteSetting "", "Explorer", "NoActiveDesktop", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' NoActiveDesktopChanges...
    If frmMain.chkDesktop(3).Value = 1 Then
        modRegistry1.SetBinaryValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoActiveDesktopChanges", Chr$(&H1) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0)
    Else
        modRegistry2.DeleteSetting "", "Explorer", "NoActiveDesktopChanges", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' NoAddingComponents...
    If frmMain.chkDesktop(4).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop", "NoAddingComponents", 1
    Else
        modRegistry2.DeleteSetting "", "ActiveDesktop", "NoAddingComponents", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' NoDeletingComponents...
    If frmMain.chkDesktop(5).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop", "NoDeletingComponents", 1
    Else
        modRegistry2.DeleteSetting "", "ActiveDesktop", "NoDeletingComponents", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' NoEditingComponents...
    If frmMain.chkDesktop(6).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop", "NoEditingComponents", 1
    Else
        modRegistry2.DeleteSetting "", "ActiveDesktop", "NoEditingComponents", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    ' NoClosingComponents...
    If frmMain.chkDesktop(7).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop", "NoClosingComponents", 1
    Else
        modRegistry2.DeleteSetting "", "ActiveDesktop", "NoClosingComponents", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' NoHTMLWallPaper...
    If frmMain.chkDesktop(8).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop", "NoHTMLWallPaper", 1
    Else
        modRegistry2.DeleteSetting "", "ActiveDesktop", "NoHTMLWallPaper", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' NoChangingWallPaper...
    If frmMain.chkDesktop(9).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop", "NoChangingWallPaper", 1
    Else
        modRegistry2.DeleteSetting "", "ActiveDesktop", "NoChangingWallPaper", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' NoCloseDragDropBands...
    If frmMain.chkDesktop(10).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop", "NoCloseDragDropBands", 1
    Else
        modRegistry2.DeleteSetting "", "ActiveDesktop", "NoCloseDragDropBands", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' NoMovingBands...
    If frmMain.chkDesktop(11).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop", "NoMovingBands", 1
    Else
        modRegistry2.DeleteSetting "", "ActiveDesktop", "NoMovingBands", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' NoStartBanner...
    If frmMain.chkDesktop(12).Value = 1 Then
        modRegistry1.SetBinaryValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoStartBanner", Chr$(&H1) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0)
    Else
        modRegistry2.DeleteSetting "", "Explorer", "NoStartBanner", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' NoTrayContextMenu...
    If frmMain.chkDesktop(13).Value = 1 Then
        modRegistry1.SetBinaryValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoTrayContextMenu", Chr$(&H1) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0)
    Else
        modRegistry2.DeleteSetting "", "Explorer", "NoTrayContextMenu", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' PaintDesktopVersion...
    If frmMain.chkDesktop(14).Value = 1 Then
        modRegistry1.SetStringValue "HKEY_CURRENT_USER\Control Panel\Desktop", "PaintDesktopVersion", "1"
    Else
        modRegistry2.DeleteSetting "", "Desktop", "PaintDesktopVersion", HKEY_CURRENT_USER, "Control Panel"
    End If
    
    ' Attributes (Recycle Bin)...
    If frmMain.chkDesktop(15).Value = 1 Then
        modRegistry1.SetBinaryValue "HKEY_CLASSES_ROOT\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\ShellFolder", "Attributes", Chr$(&H70) + Chr$(&H1) + Chr$(&H0) + Chr$(&H20)
    Else
        modRegistry1.SetBinaryValue "HKEY_CLASSES_ROOT\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\ShellFolder", "Attributes", Chr$(&H40) + Chr$(&H1) + Chr$(&H0) + Chr$(&H20)
    End If
    
    ' NoNetHood...
    If frmMain.chkDesktop(16).Value = 1 Then
        modRegistry1.SetBinaryValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoNetHood", Chr$(&H1) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0)
    Else
        modRegistry2.DeleteSetting "", "Explorer", "NoNetHood", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' NoInternetIcon...
    If frmMain.chkDesktop(17).Value = 1 Then
        modRegistry1.SetBinaryValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoInternetIcon", Chr$(&H1) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0)
    Else
        modRegistry2.DeleteSetting "", "Explorer", "NoInternetIcon", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
End Sub

Public Sub ReadDesktop()
    On Error Resume Next
    ' read Desktop settings...
    ' NoSaveSettings
    If modRegistry1.GetBinaryValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSaveSettings") = Chr$(&H1) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0) Then
        frmMain.chkDesktop(0).Value = 1
    Else
        frmMain.chkDesktop(0).Value = 0
    End If
    ' NoDesktop
    If modRegistry1.GetBinaryValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDesktop") = Chr$(&H1) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0) Then
        frmMain.chkDesktop(1).Value = 1
    Else
        frmMain.chkDesktop(1).Value = 0
    End If
    ' NoActiveDesktop
    If modRegistry1.GetBinaryValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoActiveDesktop") = Chr$(&H1) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0) Then
        frmMain.chkDesktop(2).Value = 1
    Else
        frmMain.chkDesktop(2).Value = 0
    End If
    ' NoActiveDesktopChanges
    If modRegistry1.GetBinaryValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoActiveDesktopChanges") = Chr$(&H1) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0) Then
        frmMain.chkDesktop(3).Value = 1
    Else
        frmMain.chkDesktop(3).Value = 0
    End If
    ' NoAddingComponents
    frmMain.chkDesktop(4).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop", "NoAddingComponents")
    ' NoDeletingComponents
    frmMain.chkDesktop(5).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop", "NoDeletingComponents")
    ' NoEditingComponents
    frmMain.chkDesktop(6).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop", "NoEditingComponents")
    ' NoClosingComponents
    frmMain.chkDesktop(7).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop", "NoClosingComponents")
    ' NoHTMLWallPaper
    frmMain.chkDesktop(8).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop", "NoHTMLWallPaper")
    ' NoChangingWallPaper
    frmMain.chkDesktop(9).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop", "NoChangingWallPaper")
    ' NoCloseDragDropBands
    frmMain.chkDesktop(10).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop", "NoCloseDragDropBands")
    ' NoMovingBands
    frmMain.chkDesktop(11).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop", "NoMovingBands")
    ' NoStartBanner
    If modRegistry1.GetBinaryValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoStartBanner") = Chr$(&H1) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0) Then
        frmMain.chkDesktop(12).Value = 1
    Else
        frmMain.chkDesktop(12).Value = 0
    End If
    ' NoTrayContextMenu
    If modRegistry1.GetBinaryValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoTrayContextMenu") = Chr$(&H1) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0) Then
        frmMain.chkDesktop(13).Value = 1
    Else
        frmMain.chkDesktop(13).Value = 0
    End If
    frmMain.chkDesktop(14).Value = modRegistry1.GetStringValue("HKEY_CURRENT_USER\Control Panel\Desktop", "PaintDesktopVersion")
    ' Attribute (Recycle Bin)
    If modRegistry1.GetBinaryValue("HKEY_CLASSES_ROOT\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\ShellFolder", "Attributes") = Chr$(&H70) + Chr$(&H1) + Chr$(&H0) + Chr$(&H20) Then
        frmMain.chkDesktop(15).Value = 1
    ElseIf modRegistry1.GetBinaryValue("HKEY_CLASSES_ROOT\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\ShellFolder", "Attributes") = Chr$(&H40) + Chr$(&H1) + Chr$(&H0) + Chr$(&H20) Then
        frmMain.chkDesktop(15).Value = 0
    Else
        frmMain.chkDesktop(15).Value = 2
    End If
    ' NoNetHood
    If modRegistry1.GetBinaryValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoNetHood") = Chr$(&H1) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0) Then
        frmMain.chkDesktop(16).Value = 1
    Else
        frmMain.chkDesktop(16).Value = 0
    End If
    ' NoInternetIcon
    If modRegistry1.GetBinaryValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoInternetIcon") = Chr$(&H1) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0) Then
        frmMain.chkDesktop(17).Value = 1
    Else
        frmMain.chkDesktop(17).Value = 0
    End If
End Sub
