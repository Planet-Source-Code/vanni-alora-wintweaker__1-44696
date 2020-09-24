VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "WinTweak 1.0"
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4830
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3030
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrTime 
      Enabled         =   0   'False
      Left            =   120
      Top             =   1800
   End
   Begin VB.Label lblCompany 
      BackStyle       =   0  'Transparent
      Caption         =   "Company: VJ™ Software"
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   2040
      Width           =   3975
   End
   Begin VB.Label lblGPL 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSplash.frx":0992
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Width           =   4335
   End
   Begin VB.Label lblURL 
      BackStyle       =   0  'Transparent
      Caption         =   "WebSite: http://www.vanjo.com.ph"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   1800
      Width           =   3975
   End
   Begin VB.Label lblEmail 
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail: vanjo08@msn.com"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Label lblAuthor 
      BackStyle       =   0  'Transparent
      Caption         =   "Author: Vanni Alora"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   1320
      Width           =   3975
   End
   Begin VB.Image imgSplash 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   3030
      Left            =   0
      Picture         =   "frmSplash.frx":0A2A
      Top             =   0
      Width           =   4830
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const timeEnd As Integer = 3
Dim tCount As Integer

Private Sub Form_Load()
    On Error Resume Next
    tmrTime.Interval = 1000: tmrTime.Enabled = True
    tCount = 0
    
    ' create key at startup...
    modRegistry1.CreateKey "HKEY_LOCAL_MACHINE\Software\VJ Software\WinTweaker"
    modRegistry1.SetDWORDValue "HKEY_LOCAL_MACHINE\Software\VJ Software\WinTweaker", "FirstRun", 0
    modRegistry1.SetStringValue "HKEY_LOCAL_MACHINE\Software\VJ Software\WinTweaker", "Author", "Vanni Alora"
    modRegistry1.SetStringValue "HKEY_LOCAL_MACHINE\Software\VJ Software\WinTweaker", "Email", "vanjo08@msn.com"
    modRegistry1.SetStringValue "HKEY_LOCAL_MACHINE\Software\VJ Software\WinTweaker", "Company", "VJ™ Software"
    modRegistry1.SetStringValue "HKEY_LOCAL_MACHINE\Software\VJ Software\WinTweaker", "Url", "http://www.vanjo.com.ph"
    modRegistry1.SetStringValue "HKEY_LOCAL_MACHINE\Software\VJ Software\WinTweaker", "Product", App.Title
    modRegistry1.SetStringValue "HKEY_LOCAL_MACHINE\Software\VJ Software\WinTweaker", "Version", App.Major & "." & App.Minor & "." & App.Revision
    
    modRegistry1.CreateKey "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    modRegistry1.CreateKey "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System"
    modRegistry1.CreateKey "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop"
    modRegistry1.CreateKey "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Network"
    modRegistry1.CreateKey "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\WinOldApp"
    modRegistry1.CreateKey "HKEY_CURRENT_USER\Software\Policies\Microsoft\Internet Explorer\Restrictions"
    modRegistry1.CreateKey "HKEY_CURRENT_USER\Software\Policies\Microsoft\Internet Explorer\Toolbars\Restrictions"
End Sub

Private Sub tmrTime_Timer()
    On Error Resume Next
    tCount = tCount + 1
    
    If tCount = timeEnd Then
        ' Unload Me: frmMain.Show
        If modRegistry1.GetDWORDValue("HKEY_LOCAL_MACHINE\Software\VJ Software\WinTweaker", "EnablePassword") = 1 Then
            Unload Me: frmLogin.Show
        Else
            Unload Me: frmMain.Show
        End If
    End If
End Sub
