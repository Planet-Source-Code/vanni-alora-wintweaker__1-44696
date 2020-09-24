VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   0  'None
   Caption         =   "Set Password"
   ClientHeight    =   2775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3975
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPassword.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2775
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picBackground 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   0
      ScaleHeight     =   2505
      ScaleWidth      =   3945
      TabIndex        =   1
      Top             =   240
      Width           =   3975
      Begin WinTweaker.Button2K cmdCancel 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   2640
         TabIndex        =   9
         Top             =   2040
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "Cancel"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPassword.frx":0442
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin WinTweaker.Button2K cmdOK 
         Default         =   -1  'True
         Height          =   375
         Left            =   1440
         TabIndex        =   8
         Top             =   2040
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "OK"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPassword.frx":045E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Frame fraSetPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   3495
         Begin VB.TextBox txtPassword2 
            BeginProperty Font 
               Name            =   "Wingdings"
               Size            =   8.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   240
            PasswordChar    =   "l"
            TabIndex        =   7
            Top             =   1080
            Width           =   3015
         End
         Begin VB.TextBox txtPassword1 
            BeginProperty Font 
               Name            =   "Wingdings"
               Size            =   8.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   240
            PasswordChar    =   "l"
            TabIndex        =   6
            Top             =   480
            Width           =   3015
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Re-Type Password:"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "New Password:"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   4
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.CheckBox chkPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Enable Password Protection."
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.Label lblSetPassword 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Set Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim encrypt As New clsEncryption
Dim setX As Integer, setY As Integer

Private Sub chkPassword_Click()
    On Error Resume Next
    If chkPassword.Value = 1 Then
        fraSetPassword.Enabled = True
        txtPassword1.Text = encrypt.Cryption(modRegistry2.GetSetting("", "WinTweaker", "Password", "", HKEY_LOCAL_MACHINE, "Software\VJ Software"), "WinTweaker", False)
        txtPassword2.Text = encrypt.Cryption(modRegistry2.GetSetting("", "WinTweaker", "Password", "", HKEY_LOCAL_MACHINE, "Software\VJ Software"), "WinTweaker", False)
        txtPassword1.SetFocus
    Else
        fraSetPassword.Enabled = False
        txtPassword1.Text = ""
        txtPassword2.Text = ""
        cmdOK.SetFocus
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error Resume Next
    If chkPassword.Value = 1 Then
        If txtPassword1.Text <> txtPassword2.Text Then
            MsgBox "Invalid password combination... please try again...", vbCritical, App.Title
            txtPassword1.Text = "": txtPassword2.Text = ""
            txtPassword1.SetFocus
            Exit Sub
        ElseIf Len(txtPassword1.Text) < 6 Then
            MsgBox "A valid password must be 6 characters long...", vbCritical, App.Title
            txtPassword1.Text = "": txtPassword2.Text = ""
            txtPassword1.SetFocus
            Exit Sub
        Else
            modRegistry1.SetDWORDValue "HKEY_LOCAL_MACHINE\Software\VJ Software\WinTweaker", "EnablePassword", chkPassword.Value
            modRegistry1.SetStringValue "HKEY_LOCAL_MACHINE\Software\VJ Software\WinTweaker", "Password", encrypt.Cryption(txtPassword1.Text, "WinTweaker", True)
            MsgBox "Administrator password successfully applied...", vbInformation, App.Title
            Unload Me
        End If
    Else
        modRegistry1.SetDWORDValue "HKEY_LOCAL_MACHINE\Software\VJ Software\WinTweaker", "EnablePassword", chkPassword.Value
        modRegistry1.SetStringValue "HKEY_LOCAL_MACHINE\Software\VJ Software\WinTweaker", "Password", ""
        MsgBox "Administrator password successfully removed...", vbInformation, App.Title
        Unload Me
    End If
    'modRegistry1.SetStringValue "HKEY_LOCAL_MACHINE\Software\VJ Software\WinTweaker", "Password", Encrypt.Cryption(txtPassword1.Text, "WinTweaker", True)
End Sub

Private Sub Form_Load()
    On Error Resume Next
    chkPassword.Value = modRegistry1.GetDWORDValue("HKEY_LOCAL_MACHINE\Software\VJ Software\WinTweaker", "EnablePassword")
    'txtPassword1.Text = Encrypt.Cryption(modRegistry2.GetSetting("", "WinTweaker", "Password", "", HKEY_LOCAL_MACHINE, "Software\VJ Software"), "WinTweaker", False)
    'txtPassword1.Text = modRegistry2.GetSetting("", "WinTweaker", "Password", "", HKEY_LOCAL_MACHINE, "Software\VJ Software")
End Sub

Private Sub lblSetPassword_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    setX = X
    setY = Y
End Sub

Private Sub lblSetPassword_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    ' move the form...
    If Button = 1 Then
        Me.Left = Me.Left + (X - setX)
        Me.Top = Me.Top + (Y - setY)
    End If
End Sub
