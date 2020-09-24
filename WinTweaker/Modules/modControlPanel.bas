Attribute VB_Name = "modControlPanel"
Option Explicit

Public Sub ApplyControlPanel()
    On Error Resume Next
    ' apply Control Panel settings...
    ' NoPrinters...
    If frmMain.chkControlPanel(0).Value = 1 Then
        modRegistry1.SetBinaryValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoPrinters", Chr$(&H1) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0)
    Else
        modRegistry2.DeleteSetting "", "Explorer", "NoPrinters", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' NoAddPrinter...
    If frmMain.chkControlPanel(1).Value = 1 Then
        modRegistry1.SetBinaryValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoAddPrinter", Chr$(&H1) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0)
    Else
        modRegistry2.DeleteSetting "", "Explorer", "NoAddPrinter", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' NoDeletePrinter...
    If frmMain.chkControlPanel(2).Value = 1 Then
        modRegistry1.SetBinaryValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDeletePrinter", Chr$(&H1) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0)
    Else
        modRegistry2.DeleteSetting "", "Explorer", "NoDeletePrinter", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' NoPrinterTabs...
    If frmMain.chkControlPanel(3).Value = 1 Then
        modRegistry1.SetBinaryValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoPrinterTabs", Chr$(&H1) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0)
    Else
        modRegistry2.DeleteSetting "", "Explorer", "NoPrinterTabs", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' NoDispCPL...
    If frmMain.chkControlPanel(4).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDispCPL", 1
    Else
        modRegistry2.DeleteSetting "", "System", "NoDispCPL", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' NoDispBackgroundPage...
    If frmMain.chkControlPanel(5).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDispBackgroundPage", 1
    Else
        modRegistry2.DeleteSetting "", "System", "NoDispBackgroundPage", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' NoDispScrSavPage...
    If frmMain.chkControlPanel(6).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDispScrSavPage", 1
    Else
        modRegistry2.DeleteSetting "", "System", "NoDispScrSavPage", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' NoDispAppearancePage...
    If frmMain.chkControlPanel(7).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDispAppearancePage", 1
    Else
        modRegistry2.DeleteSetting "", "System", "NoDispAppearancePage", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' NoDispSettingsPage...
    If frmMain.chkControlPanel(8).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDispSettingsPage", 1
    Else
        modRegistry2.DeleteSetting "", "System", "NoDispSettingsPage", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' NoSecCPL...
    If frmMain.chkControlPanel(9).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoSecCPL", 1
    Else
        modRegistry2.DeleteSetting "", "System", "NoSecCPL", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' NoPwdPage...
    If frmMain.chkControlPanel(10).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoPwdPage", 1
    Else
        modRegistry2.DeleteSetting "", "System", "NoPwdPage", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' NoAdminPage...
    If frmMain.chkControlPanel(11).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoAdminPage", 1
    Else
        modRegistry2.DeleteSetting "", "System", "NoAdminPage", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' NoProfilePage...
    If frmMain.chkControlPanel(12).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoProfilePage", 1
    Else
        modRegistry2.DeleteSetting "", "System", "NoProfilePage", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' NoConfigPage...
    If frmMain.chkControlPanel(13).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoConfigPage", 1
    Else
        modRegistry2.DeleteSetting "", "System", "NoConfigPage", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' NoDevMgrPage...
    If frmMain.chkControlPanel(14).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDevMgrPage", 1
    Else
        modRegistry2.DeleteSetting "", "System", "NoDevMgrPage", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' NoFileSysPage...
    If frmMain.chkControlPanel(15).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoFileSysPage", 1
    Else
        modRegistry2.DeleteSetting "", "System", "NoFileSysPage", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' NoVirtMemPage...
    If frmMain.chkControlPanel(16).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoVirtMemPage", 1
    Else
        modRegistry2.DeleteSetting "", "System", "NoVirtMemPage", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
End Sub

Public Sub ReadControlPanel()
    On Error Resume Next
    ' read Control Panel settings...
    ' NoPrinters
    If modRegistry1.GetBinaryValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoPrinters") = Chr$(&H1) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0) Then
        frmMain.chkControlPanel(0).Value = 1
    Else
        frmMain.chkControlPanel(0).Value = 0
    End If
    ' NoAddPrinter
    If modRegistry1.GetBinaryValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoAddPrinter") = Chr$(&H1) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0) Then
        frmMain.chkControlPanel(1).Value = 1
    Else
        frmMain.chkControlPanel(1).Value = 0
    End If
    ' NoDeletePrinter
    If modRegistry1.GetBinaryValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDeletePrinter") = Chr$(&H1) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0) Then
        frmMain.chkControlPanel(2).Value = 1
    Else
        frmMain.chkControlPanel(2).Value = 0
    End If
    ' NoPrinterTabs
    If modRegistry1.GetBinaryValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoPrinterTabs") = Chr$(&H1) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0) Then
        frmMain.chkControlPanel(3).Value = 1
    Else
        frmMain.chkControlPanel(3).Value = 0
    End If
    ' NoDispCPL
    frmMain.chkControlPanel(4).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDispCPL")
    ' NoDispBackgroundPage
    frmMain.chkControlPanel(5).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDispBackgroundPage")
    ' NoDispScrSavPage
    frmMain.chkControlPanel(6).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDispScrSavPage")
    ' NoDispAppearancePage
    frmMain.chkControlPanel(7).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDispAppearancePage")
    ' NoDispSettingsPage
    frmMain.chkControlPanel(8).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDispSettingsPage")
    ' NoSecCPL
    frmMain.chkControlPanel(9).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoSecCPL")
    ' NoPwdPage
    frmMain.chkControlPanel(10).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoPwdPage")
    ' NoAdminPage
    frmMain.chkControlPanel(11).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoAdminPage")
    ' NoProfilePage
    frmMain.chkControlPanel(12).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoProfilePage")
    ' NoConfigPage
    frmMain.chkControlPanel(13).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoConfigPage")
    ' NoDevMgrPage
    frmMain.chkControlPanel(14).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDevMgrPage")
    ' NoFileSysPage
    frmMain.chkControlPanel(15).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoFileSysPage")
    ' NoVirtMemPage
    frmMain.chkControlPanel(16).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoVirtMemPage")
End Sub
