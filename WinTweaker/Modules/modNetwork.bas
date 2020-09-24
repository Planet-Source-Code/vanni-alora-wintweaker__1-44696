Attribute VB_Name = "modNetwork"
Option Explicit

Public Sub ApplyNetwork()
    On Error Resume Next
    ' apply Network settings...
    ' NoNetSetup...
    If frmMain.chkNetwork(0).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Network", "NoNetSetup", 1
    Else
        modRegistry2.DeleteSetting "", "Network", "NoNetSetup", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' NoNetSetupIDPage...
    If frmMain.chkNetwork(1).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Network", "NoNetSetupIDPage", 1
    Else
        modRegistry2.DeleteSetting "", "Network", "NoNetSetupIDPage", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' NoNetSetupSecurityPage...
    If frmMain.chkNetwork(2).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Network", "NoNetSetupSecurityPage", 1
    Else
        modRegistry2.DeleteSetting "", "Network", "NoNetSetupSecurityPage", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' NoEntireNetwork...
    If frmMain.chkNetwork(3).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Network", "NoEntireNetwork", 1
    Else
        modRegistry2.DeleteSetting "", "Network", "NoEntireNetwork", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' NoWorkgroupContents...
    If frmMain.chkNetwork(4).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Network", "NoWorkgroupContents", 1
    Else
        modRegistry2.DeleteSetting "", "Network", "NoWorkgroupContents", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' NoNetConnectDisconnect...
    If frmMain.chkNetwork(5).Value = 1 Then
        modRegistry1.SetBinaryValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoNetConnectDisconnect", Chr$(&H1) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0)
    Else
        modRegistry2.DeleteSetting "", "Explorer", "NoNetConnectDisconnect", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' NoFileSharingControl...
    If frmMain.chkNetwork(6).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Network", "NoFileSharingControl", 1
    Else
        modRegistry2.DeleteSetting "", "Network", "NoFileSharingControl", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' NoPrintSharingControl...
    If frmMain.chkNetwork(7).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Network", "NoPrintSharingControl", 1
    Else
        modRegistry2.DeleteSetting "", "Network", "NoPrintSharingControl", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' MustBeValidated...
    If frmMain.chkNetwork(8).Value = 1 Then
        modRegistry1.SetDWORDValue "HKEY_LOCAL_MACHINE\Network\Logon", "MustBeValidated", 1
    Else
        modRegistry2.DeleteSetting "", "Logon", "MustBeValidated", HKEY_LOCAL_MACHINE, "Network"
    End If
    
    ' MinPwdLen...
    If frmMain.txtNetworkPwdLength.Text <> "" Then
        modRegistry1.SetBinaryValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Network", "MinPwdLen", frmMain.txtNetworkPwdLength.Text
    Else
        modRegistry2.DeleteSetting "", "Network", "MinPwdLen", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
    End If
    
    ' DefaultTTL...
    If frmMain.txtNetworkTTLTcp.Text <> "" Then
        modRegistry1.SetStringValue "HKEY_LOCAL_MACHINE\System\CurrentControlSet\Services\VxD\MSTCP", "DefaultTTL", frmMain.txtNetworkTTLTcp.Text
    Else
        modRegistry2.DeleteSetting "", "MSTCP", "DefaultTTL", HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\VxD"
    End If
End Sub

Public Sub ReadNetwork()
    On Error Resume Next
    ' read Network settings...
    ' NoNetSetup
    frmMain.chkNetwork(0).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Network", "NoNetSetup")
    ' NoNetSetupIDPage
    frmMain.chkNetwork(1).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Network", "NoNetSetupIDPage")
    ' NoNetSetupSecurityPage
    frmMain.chkNetwork(2).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Network", "NoNetSetupSecurityPage")
    ' NoEntireNetwork
    frmMain.chkNetwork(3).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Network", "NoEntireNetwork")
    ' NoWorkgroupContents
    frmMain.chkNetwork(4).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Network", "NoWorkgroupContents")
    ' NoNetConnectDisconnect
    If modRegistry1.GetBinaryValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoNetConnectDisconnect") = Chr$(&H1) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0) Then
        frmMain.chkNetwork(5).Value = 1
    Else
        frmMain.chkNetwork(5).Value = 0
    End If
    ' NoFileSharingControl
    frmMain.chkNetwork(6).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Network", "NoFileSharingControl")
    ' NoPrintSharingControl
    frmMain.chkNetwork(7).Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Network", "NoPrintSharingControl")
    ' MustBeValidated
    frmMain.chkNetwork(8).Value = modRegistry1.GetDWORDValue("HKEY_LOCAL_MACHINE\Network\Logon", "MustBeValidated")
    ' MinPwdLen
    frmMain.txtNetworkPwdLength.Text = modRegistry2.GetSetting("", "Network", "MinPwdLen", "", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies")
    'frmMain.txtNetworkPwdLength.Text = modRegistry1.GetBinaryValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Network", "MinPwdLen")
    ' DefaultTTL
    frmMain.txtNetworkTTLTcp.Text = modRegistry2.GetSetting("", "MSTCP", "DefaultTTL", "", HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\VxD")
    'frmMain.txtNetworkTTLTcp.Text = modRegistry1.GetStringValue("HKEY_LOCAL_MACHINE\System\CurrentControlSet\Services\VxD\MSTCP", "DefaultTTL")
End Sub
