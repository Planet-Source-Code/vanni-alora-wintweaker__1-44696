Attribute VB_Name = "modIExplorer"
Option Explicit

Public Sub ApplyIExplorer()
    On Error Resume Next
    ' apply Internet Explorer settings...
    ' NoFileOpen...
    If frmMain.chkIExplorer(0).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoFileOpen", 1
    Else
        modRegistry2.DeleteSetting "", "Restrictions", "NoFileOpen", HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer"
    End If
    
    ' NoFileNew...
    If frmMain.chkIExplorer(1).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoFileNew", 1
    Else
        modRegistry2.DeleteSetting "", "Restrictions", "NoFileNew", HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer"
    End If
    
    ' NoBrowserSaveAs...
    If frmMain.chkIExplorer(2).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoBrowserSaveAs", 1
    Else
        modRegistry2.DeleteSetting "", "Restrictions", "NoBrowserSaveAs", HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer"
    End If
    
    ' NoBrowserSaveWebComplete...
    If frmMain.chkIExplorer(3).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoBrowserSaveWebComplete", 1
    Else
        modRegistry2.DeleteSetting "", "Restrictions", "NoBrowserSaveWebComplete", HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer"
    End If
    
    ' NoBrowserOptions...
    If frmMain.chkIExplorer(4).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoBrowserOptions", 1
    Else
        modRegistry2.DeleteSetting "", "Restrictions", "NoBrowserOptions", HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer"
    End If
    
    ' NoWindowsUpdate...
    If frmMain.chkIExplorer(5).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoWindowsUpdate", 1
    Else
        modRegistry2.DeleteSetting "", "Restrictions", "NoWindowsUpdate", HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer"
    End If
    
    ' NoTheaterMode...
    If frmMain.chkIExplorer(6).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoTheaterMode", 1
    Else
        modRegistry2.DeleteSetting "", "Restrictions", "NoTheaterMode", HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer"
    End If
    
    ' NoFavorites...
    If frmMain.chkIExplorer(7).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Policies\Microsoft\Internet Explorer\Restrictions", "Nofavorites", 1
    Else
        modRegistry2.DeleteSetting "", "Restrictions", "NoFavorites", HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer"
    End If
    
    ' NoSelectDownloadDir...
    If frmMain.chkIExplorer(8).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoSelectDownloadDir", 1
    Else
        modRegistry2.DeleteSetting "", "Restrictions", "NoSelectDownloadDir", HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer"
    End If
    
    ' NoBrowserContextMenu...
    If frmMain.chkIExplorer(9).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoBrowserContextMenu", 1
    Else
        modRegistry2.DeleteSetting "", "Restrictions", "NoBrowserContextMenu", HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer"
    End If
    
    ' NoBrowserClose...
    If frmMain.chkIExplorer(10).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoBrowserClose", 1
    Else
        modRegistry2.DeleteSetting "", "Restrictions", "NoBrowserClose", HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer"
    End If
    
    ' NoFindFiles...
    If frmMain.chkIExplorer(11).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoFindFiles", 1
    Else
        modRegistry2.DeleteSetting "", "Restrictions", "NoFindFiles", HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer"
    End If
    
    ' LockIconSize...
    If frmMain.chkIExplorer(12).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Policies\Microsoft\Internet Explorer\Restrictions", "LockIconSize", 1
    Else
        modRegistry2.DeleteSetting "", "Restrictions", "LockIconSize", HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer"
    End If
    
    ' NoChannelUI...
    If frmMain.chkIExplorer(13).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoChannelUI", 1
    Else
        modRegistry2.DeleteSetting "", "Restrictions", "NoChannelUI", HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer"
    End If
    
    ' NoFileUrl
    If frmMain.chkIExplorer(14).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoFileUrl", 1
    Else
        modRegistry2.DeleteSetting "", "Restrictions", "NoFileUrl", HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer"
    End If
    
    ' NoSplash...
    If frmMain.chkIExplorer(15).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoSplash", 1
    Else
        modRegistry2.DeleteSetting "", "Restrictions", "NoSplash", HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer"
    End If
    
    ' NoToolbarOptions...
    If frmMain.chkIExplorer(16).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Policies\Microsoft\Internet Explorer\Toolbars\Restrictions", "NoToolbarOptions", 1
    Else
        modRegistry2.DeleteSetting "", "Restrictions", "NoToolbarOptions", HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Toolbars"
    End If
    
    ' NoToolbarCustomize...
    If frmMain.chkIExplorer(17).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoToolbarCustomize", 1
    Else
        modRegistry2.DeleteSetting "", "Explorer", "NoToolbarCustomize", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' ShowFonts...
    If frmMain.chkIExplorer(18).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Toolbar", "ShowFonts", 1
    Else
        modRegistry2.DeleteSetting "", "Toolbar", "ShowFonts", HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer"
    End If
    
    ' Disable Hotmail...
    If frmMain.chkIExplorer(19).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Outlook Express", "Disable Hotmail", 1
    Else
        modRegistry2.DeleteSetting "", "Outlook Express", "Disable Hotmail", HKEY_CURRENT_USER, "Software\Microsoft"
    End If
End Sub

Public Sub ReadIExplorer()
    ' NoFileOpen
    frmMain.chkIExplorer(0).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoFileOpen")
    ' NoFileNew
    frmMain.chkIExplorer(1).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoFileNew")
    ' NoBrowserSaveAs
    frmMain.chkIExplorer(2).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoBrowserSaveAs")
    ' NoBrowserSaveWebComplete
    frmMain.chkIExplorer(3).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoBrowserSaveWebComplete")
    ' NoBrowserOptions
    frmMain.chkIExplorer(4).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoBrowserOptions")
    ' NoWindowsUpdate
    frmMain.chkIExplorer(5).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoWindowsUpdate")
    ' NoTheaterMode
    frmMain.chkIExplorer(6).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoTheaterMode")
    ' NoFavorites
    frmMain.chkIExplorer(7).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoFavorites")
    ' NoSelectDownloadDir
    frmMain.chkIExplorer(8).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoSelectDownloadDir")
    ' NoBrowserContextMenu
    frmMain.chkIExplorer(9).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoBrowserContextMenu")
    ' NoBrowserClose
    frmMain.chkIExplorer(10).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoBrowserClose")
    ' NoFindFiles
    frmMain.chkIExplorer(11).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoFindFiles")
    ' LockIconSize
    frmMain.chkIExplorer(12).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Policies\Microsoft\Internet Explorer\Restrictions", "LockIconSize")
    ' NoChannelUI
    frmMain.chkIExplorer(13).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoChannelUI")
    ' NoFileUrl
    frmMain.chkIExplorer(14).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoFileUrl")
    ' NoSplash
    frmMain.chkIExplorer(15).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Policies\Microsoft\Internet Explorer\Restrictions", "NoSplash")
    ' NoToolbarOptions
    frmMain.chkIExplorer(16).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Policies\Microsoft\Internet Explorer\Toolbars\Restrictions", "NoToolbarOptions")
    ' NoToolbarCustomize
    frmMain.chkIExplorer(17).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoToolbarCustomize")
    ' ShowFonts
    frmMain.chkIExplorer(18).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Toolbar", "ShowFonts")
    ' Disable Hotmail
    frmMain.chkIExplorer(19).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Outlook Express", "Disable Hotmail")
End Sub
