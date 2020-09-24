Attribute VB_Name = "modSystem"
Option Explicit

Public Sub ApplySystem()
    On Error Resume Next
    ' apply System settings...
    ' RegisteredOwner...
    modRegistry1.SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "RegisteredOwner", frmMain.txtSystemOwner.Text
    
    ' RegisteredOrganization...
    modRegistry1.SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "RegisteredOrganization", frmMain.txtSystemOrganization.Text
    
    ' SourcePath...
    modRegistry1.SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Setup", "SourcePath", frmMain.txtSystemSourcePath.Text
    
    ' DisableRegistryTools...
    If frmMain.chkSystem(0).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", 1
    Else
        modRegistry2.DeleteSetting "", "System", "DisableRegistryTools", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' NoDevMgrUpdate...
    If frmMain.chkSystem(1).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDevMgrUpdate", 1
    Else
        modRegistry2.DeleteSetting "", "Explorer", "NoDevMgrUpdate", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' NoDriveTypeAutoRun...
    If frmMain.chkSystem(2).Value = 1 Then
        modRegistry1.SetBinaryValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDriveTypeAutoRun", Chr$(&HB5) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0)
    Else
        modRegistry1.SetBinaryValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDriveTypeAutoRun", Chr$(&H95) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0)
    End If
    
    ' EnforceShellExtensionSecurity...
    If frmMain.chkSystem(3).Value = 1 Then
        modRegistry1.SetBinaryValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "EnforceShellExtensionSecurity", Chr$(&H1) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0)
    Else
        modRegistry2.DeleteSetting "", "Explorer", "EnforceShellExtensionSecurity", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' NoRealMode...
    If frmMain.chkSystem(4).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\WinOldApp", "NoRealMode", 1
    Else
        modRegistry2.DeleteSetting "", "WinOldApp", "NoRealMode", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' Disabled...
    If frmMain.chkSystem(5).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\WinOldApp", "Disabled", 1
    Else
        modRegistry2.DeleteSetting "", "WinOldApp", "Disabled", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' RegDone...
    If frmMain.chkSystem(6).Value = 1 Then
        modRegistry1.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion", "RegDone", "1"
        modRegistry1.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Welcome\RegWiz", "@", "1"
    Else
        modRegistry1.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion", "RegDone", ""
        modRegistry2.DeleteSetting "", "RegWiz", "@", HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Welcome"
    End If
    
    ' Shell Icon Size...
    If frmMain.chkSystem(7).Value = 1 Then
        modRegistry1.SetStringValue "HKEY_CURRENT_USER\Control Panel\Desktop\WindowMetrics", "Shell Icon Size", "48"
    Else
        modRegistry1.SetStringValue "HKEY_CURRENT_USER\Control Panel\Desktop\WindowMetrics", "Shell Icon Size", "32"
    End If
    
    ' Shell Icon BPP...
    If frmMain.chkSystem(8).Value = 1 Then
        modRegistry1.SetStringValue "HKEY_CURRENT_USER\Control Panel\Desktop\WindowMetrics", "Shell Icon BPP", "16"
    Else
        modRegistry1.SetStringValue "HKEY_CURRENT_USER\Control Panel\Desktop\WindowMetrics", "Shell Icon BPP", "4"
    End If
    
    ' MinAnimate...
    If frmMain.chkSystem(9).Value = 1 Then
        modRegistry1.SetStringValue "HKEY_CURRENT_USER\Control Panel\Desktop\WindowMetrics", "MinAnimate", "1"
    Else
        modRegistry1.SetStringValue "HKEY_CURRENT_USER\Control Panel\Desktop\WindowMetrics", "MinAnimate", "0"
    End If
    
    ' FontSmoothing...
    If frmMain.chkSystem(10).Value = 1 Then
        modRegistry1.SetStringValue "HKEY_CURRENT_USER\Control Panel\Desktop", "FontSmoothing", "1"
    Else
        modRegistry1.SetStringValue "HKEY_CURRENT_USER\Control Panel\Desktop", "FontSmoothing", "0"
    End If
    
    ' DragFullWindows...
    If frmMain.chkSystem(11).Value = 1 Then
        modRegistry1.SetStringValue "HKEY_CURRENT_USER\Control Panel\Desktop", "DragFullWindows", "1"
    Else
        modRegistry1.SetStringValue "HKEY_CURRENT_USER\Control Panel\Desktop", "DragFullWindows", "0"
    End If
    
    ' ClassicShell...
    If frmMain.chkSystem(12).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "ClassicShell", 1
    Else
        modRegistry2.DeleteSetting "", "Explorer", "ClassicShell", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' IsShortcut...
    If frmMain.chkSystem(13).Value = 1 Then
        modRegistry1.SetStringValue "HKEY_CLASSES_ROOT\lnkfile", "IsShortcut", "1"
        modRegistry1.SetStringValue "HKEY_CLASSES_ROOT\piffile", "IsShortcut", "1"
    Else
        modRegistry2.DeleteSetting "", "lnkfile", "IsShortcut", HKEY_CLASSES_ROOT, ""
        modRegistry2.DeleteSetting "", "piffile", "IsShortcut", HKEY_CLASSES_ROOT, ""
        'modRegistry1.SetStringValue "HKEY_CLASSES_ROOT\lnkfile", "IsShortcut", ""
        'modRegistry1.SetStringValue "HKEY_CLASSES_ROOT\piffile", "IsShortcut", ""
    End If
End Sub

Public Sub ReadSystem()
    On Error Resume Next
    ' read System settings...
    ' RegisteredOwner
    frmMain.txtSystemOwner.Text = modRegistry2.GetSetting("", "CurrentVersion", "RegisteredOwner", "", HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows")
    ' RegisteredOrganization
    frmMain.txtSystemOrganization.Text = modRegistry2.GetSetting("", "CurrentVersion", "RegisteredOrganization", "", HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows")
    ' SourcePath
    frmMain.txtSystemSourcePath.Text = modRegistry2.GetSetting("", "Setup", "SourcePath", "", HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion")
    ' DisableRegistryTools
    frmMain.chkSystem(0).Value = modRegistry2.GetSetting("", "System", "DisableRegistryTools", "", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies")
    ' NoDevMgrUpdate
    frmMain.chkSystem(1).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDevMgrUpdate")
    ' NoDriveTypeAutoRun
    If modRegistry1.GetBinaryValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDriveTypeAutoRun") = Chr$(&HB5) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0) Then
        frmMain.chkSystem(2).Value = 1
    ElseIf modRegistry1.GetBinaryValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDriveTypeAutoRun") = Chr$(&H95) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0) Then
        frmMain.chkSystem(2).Value = 0
    Else
        frmMain.chkSystem(2).Value = 2
    End If
    ' EnforceShellExtensionSecurity
    If modRegistry1.GetBinaryValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "EnforceShellExtensionSecurity") = Chr$(&H1) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0) Then
        frmMain.chkSystem(3).Value = 1
    Else
        frmMain.chkSystem(3).Value = 0
    End If
    ' NoRealMode
    frmMain.chkSystem(4).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\WinOldApp", "NoRealMode")
    ' Disabled
    frmMain.chkSystem(5).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\WinOldApp", "Disabled")
    ' RegDone @
    frmMain.chkSystem(6).Value = modRegistry1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion", "RegDone")
    ' Shell Icon Size
    If modRegistry1.GetStringValue("HKEY_CURRENT_USER\Control Panel\Desktop\WindowMetrics", "Shell Icon Size") = 48 Then
        frmMain.chkSystem(7).Value = 1
    Else
        frmMain.chkSystem(7).Value = 0
    End If
    ' Shell Icon BPP
    If modRegistry1.GetStringValue("HKEY_CURRENT_USER\Control Panel\Desktop\WindowMetrics", "Shell Icon BPP") = 16 Then
        frmMain.chkSystem(8).Value = 1
    Else
        frmMain.chkSystem(8).Value = 0
    End If
    ' MinAnimate
    frmMain.chkSystem(9).Value = modRegistry1.GetStringValue("HKEY_CURRENT_USER\Control Panel\Desktop\WindowMetrics", "MinAnimate")
    ' FontSmoothing
    frmMain.chkSystem(10).Value = modRegistry1.GetStringValue("HKEY_CURRENT_USER\Control Panel\Desktop", "FontSmoothing")
    ' DragFullWindows
    frmMain.chkSystem(11).Value = modRegistry1.GetStringValue("HKEY_CURRENT_USER\Control Panel\Desktop", "DragFullWindows")
    ' ClassicShell
    frmMain.chkSystem(12).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "ClassicShell")
    ' IsShortcut
    frmMain.chkSystem(13).Value = modRegistry1.GetStringValue("HKEY_CLASSES_ROOT\lnkfile", "IsShortcut")
End Sub
