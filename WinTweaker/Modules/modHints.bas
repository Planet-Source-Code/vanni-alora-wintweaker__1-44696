Attribute VB_Name = "modHints"
Option Explicit

Public Sub ApplyHints()
    On Error Resume Next
    ' apply Hints settings...
    ' Recycle Bin...
    'modRegistry1.SetStringValue "HKEY_CLASSES_ROOT\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}", "InfoTips", frmMain.txtHintsRecycleBin.Text
    modRegistry2.SaveSetting "", "{645FF040-5081-101B-9F08-00AA002F954E}", "InfoTip", frmMain.txtHintsRecycleBin, HKEY_CLASSES_ROOT, "CLSID"
    ' My Computer...
    modRegistry2.SaveSetting "", "{20D04FE0-3AEA-1069-A2D8-08002B30309D}", "InfoTip", frmMain.txtHintsMyComputer, HKEY_CLASSES_ROOT, "CLSID"
    ' My Documents...
    modRegistry2.SaveSetting "", "{450D8FBA-AD25-11D0-98A8-0800361B1103}", "InfoTip", frmMain.txtHintsMyDocuments.Text, HKEY_CLASSES_ROOT, "CLSID"
    ' Network Neighborhood...
    modRegistry2.SaveSetting "", "{208D2C60-3AEA-1069-A2D7-08002B30309D}", "InfoTip", frmMain.txtHintsNetworkNeighborhood.Text, HKEY_CLASSES_ROOT, "CLSID"
    ' Internet...
    modRegistry2.SaveSetting "", "{3DC7A020-0ACD-11CF-A9BB-00AA004AE837}", "InfoTip", frmMain.txtHintsInternet.Text, HKEY_CLASSES_ROOT, "CLSID"
    ' Internet Explorer...
    modRegistry2.SaveSetting "", "{871C5380-42A0-1069-A2EA-08002B30309D}", "InfoTip", frmMain.txtHintsInternetExplorer.Text, HKEY_CLASSES_ROOT, "CLSID"
    ' Printers...
    modRegistry2.SaveSetting "", "{2227A280-3AEA-1069-A2DE-08002B30309D}", "InfoTip", frmMain.txtHintsPrinters.Text, HKEY_CLASSES_ROOT, "CLSID"
    ' Control Panel...
    modRegistry2.SaveSetting "", "{21EC2020-3AEA-1069-A2DD-08002B30309D}", "InfoTip", frmMain.txtHintsControlPanel, HKEY_CLASSES_ROOT, "CLSID"
    ' Dial-Up Networking...
    modRegistry2.SaveSetting "", "{992CFFA0-F557-101A-88EC-00DD010CCC48}", "InfoTip", frmMain.txtHintsDialUp.Text, HKEY_CLASSES_ROOT, "CLSID"
    ' Briefcase...
    modRegistry2.SaveSetting "", "{85BBD920-42A0-1069-A2E4-08002B30309D}", "InfoTip", frmMain.txtHintsBriefcase.Text, HKEY_CLASSES_ROOT, "CLSID"
    ' Tasks...
    modRegistry2.SaveSetting "", "{D6277990-4C6A-11CF-8D87-00AA0060F5BF}", "InfoTip", frmMain.txtHintsTasks.Text, HKEY_CLASSES_ROOT, "CLSID"
    ' Web Folders...
    modRegistry2.SaveSetting "", "{BDEADF00-C265-11d0-BCED-00A0C90AB50F}", "InfoTip", frmMain.txtHintsWebFolders.Text, HKEY_CLASSES_ROOT, "CLSID"
End Sub

Public Sub ReadHints()
    On Error Resume Next
    ' read Hints settings...
    ' Recycle Bin
    'frmMain.txtHintsRecycleBin.Text = modRegistry1.GetStringValue("HKEY_CLASSES_ROOT\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}", "InfoTips")
    frmMain.txtHintsRecycleBin.Text = modRegistry2.GetSetting("", "{645FF040-5081-101B-9F08-00AA002F954E}", "InfoTip", "", HKEY_CLASSES_ROOT, "CLSID")
    ' My Computer
    frmMain.txtHintsMyComputer.Text = modRegistry2.GetSetting("", "{20D04FE0-3AEA-1069-A2D8-08002B30309D}", "InfoTip", "", HKEY_CLASSES_ROOT, "CLSID")
    ' My Documents
    frmMain.txtHintsMyDocuments.Text = modRegistry2.GetSetting("", "{450D8FBA-AD25-11D0-98A8-0800361B1103}", "InfoTip", "", HKEY_CLASSES_ROOT, "CLSID")
    ' Network Neighborhood
    frmMain.txtHintsNetworkNeighborhood.Text = modRegistry2.GetSetting("", "{208D2C60-3AEA-1069-A2D7-08002B30309D}", "InfoTip", "", HKEY_CLASSES_ROOT, "CLSID")
    ' Internet
    frmMain.txtHintsInternet.Text = modRegistry2.GetSetting("", "{3DC7A020-0ACD-11CF-A9BB-00AA004AE837}", "InfoTip", "", HKEY_CLASSES_ROOT, "CLSID")
    ' Internet Explorer
    frmMain.txtHintsInternetExplorer.Text = modRegistry2.GetSetting("", "{871C5380-42A0-1069-A2EA-08002B30309D}", "InfoTip", "", HKEY_CLASSES_ROOT, "CLSID")
    ' Printers
    frmMain.txtHintsPrinters.Text = modRegistry2.GetSetting("", "{2227A280-3AEA-1069-A2DE-08002B30309D}", "InfoTip", "", HKEY_CLASSES_ROOT, "CLSID")
    ' Control Panel
    frmMain.txtHintsControlPanel.Text = modRegistry2.GetSetting("", "{21EC2020-3AEA-1069-A2DD-08002B30309D}", "InfoTip", "", HKEY_CLASSES_ROOT, "CLSID")
    ' Dial-Up Networking
    frmMain.txtHintsDialUp.Text = modRegistry2.GetSetting("", "{992CFFA0-F557-101A-88EC-00DD010CCC48}", "InfoTip", "", HKEY_CLASSES_ROOT, "CLSID")
    ' Briefcase
    frmMain.txtHintsBriefcase.Text = modRegistry2.GetSetting("", "{85BBD920-42A0-1069-A2E4-08002B30309D}", "InfoTip", "", HKEY_CLASSES_ROOT, "CLSID")
    ' Tasks
    frmMain.txtHintsTasks.Text = modRegistry2.GetSetting("", "{D6277990-4C6A-11CF-8D87-00AA0060F5BF}", "InfoTip", "", HKEY_CLASSES_ROOT, "CLSID")
    ' Web Folders
    frmMain.txtHintsWebFolders.Text = modRegistry2.GetSetting("", "{BDEADF00-C265-11d0-BCED-00A0C90AB50F}", "InfoTip", "", HKEY_CLASSES_ROOT, "CLSID")
End Sub
