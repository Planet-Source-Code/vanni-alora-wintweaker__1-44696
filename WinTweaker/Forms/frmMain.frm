VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "WinTweaker 1.0 for Win 9x/ME"
   ClientHeight    =   7335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8535
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   8535
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picBackground 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7095
      Left            =   0
      ScaleHeight     =   7065
      ScaleWidth      =   8505
      TabIndex        =   1
      Top             =   240
      Width           =   8530
      Begin VB.PictureBox picArrow 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   9
         Left            =   2160
         Picture         =   "frmMain.frx":0442
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   34
         Top             =   4530
         Width           =   300
      End
      Begin VB.PictureBox picArrow 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   8
         Left            =   2160
         Picture         =   "frmMain.frx":04CA
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   33
         Top             =   4050
         Width           =   300
      End
      Begin VB.PictureBox picArrow 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   7
         Left            =   2160
         Picture         =   "frmMain.frx":0552
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   32
         Top             =   3570
         Width           =   300
      End
      Begin VB.PictureBox picArrow 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   6
         Left            =   2160
         Picture         =   "frmMain.frx":05DA
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   31
         Top             =   3090
         Width           =   300
      End
      Begin VB.PictureBox picArrow 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   5
         Left            =   2160
         Picture         =   "frmMain.frx":0662
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   30
         Top             =   2610
         Width           =   300
      End
      Begin VB.PictureBox picArrow 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   4
         Left            =   2160
         Picture         =   "frmMain.frx":06EA
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   29
         Top             =   2130
         Width           =   300
      End
      Begin VB.PictureBox picArrow 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   3
         Left            =   2160
         Picture         =   "frmMain.frx":0772
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   28
         Top             =   1650
         Width           =   300
      End
      Begin VB.PictureBox picArrow 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   2
         Left            =   2160
         Picture         =   "frmMain.frx":07FA
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   27
         Top             =   1170
         Width           =   300
      End
      Begin VB.PictureBox picArrow 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   1
         Left            =   2160
         Picture         =   "frmMain.frx":0882
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   26
         Top             =   690
         Width           =   300
      End
      Begin VB.Frame fraDescription 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Description"
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   360
         TabIndex        =   9
         Top             =   5880
         Width           =   7815
         Begin VB.Label lblDescription 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "WinTweaker 1.0 (Windows Personal Security). Copyright © 2001 - 2003 VJ™ Software"
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   7575
         End
      End
      Begin VB.PictureBox picArrow 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   0
         Left            =   2160
         Picture         =   "frmMain.frx":090A
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   8
         Top             =   210
         Width           =   300
      End
      Begin WinTweaker.Button2K cmdAbout 
         Height          =   375
         Left            =   6840
         TabIndex        =   11
         Top             =   5400
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "About"
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
         BCOL            =   16744576
         BCOLO           =   16744576
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMain.frx":0992
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin WinTweaker.Button2K cmdSetPassword 
         Height          =   375
         Left            =   3960
         TabIndex        =   12
         Top             =   5400
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "Set Password"
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
         BCOL            =   16744576
         BCOLO           =   16744576
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMain.frx":09AE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin WinTweaker.Button2K cmdTurnOff 
         Height          =   375
         Left            =   5400
         TabIndex        =   13
         Top             =   5400
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "Shutdown"
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
         BCOL            =   16744576
         BCOLO           =   16744576
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMain.frx":09CA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin WinTweaker.Button2K cmdApply 
         Height          =   375
         Left            =   2520
         TabIndex        =   14
         Top             =   5400
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "Apply"
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
         BCOL            =   16744576
         BCOLO           =   16744576
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMain.frx":09E6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin WinTweaker.Button2K cmdWExplorer 
         Height          =   375
         Left            =   360
         TabIndex        =   15
         Top             =   3960
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "Window Explorer"
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
         BCOL            =   16744576
         BCOLO           =   16744576
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMain.frx":0A02
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin WinTweaker.Button2K cmdIExplorer 
         Height          =   375
         Left            =   360
         TabIndex        =   16
         Top             =   3480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "Internet Explorer"
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
         BCOL            =   16744576
         BCOLO           =   16744576
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMain.frx":0A1E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin WinTweaker.Button2K cmdHints 
         Height          =   375
         Left            =   360
         TabIndex        =   17
         Top             =   3000
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "Hints"
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
         BCOL            =   16744576
         BCOLO           =   16744576
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMain.frx":0A3A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin WinTweaker.Button2K cmdSystem 
         Height          =   375
         Left            =   360
         TabIndex        =   18
         Top             =   2520
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "System"
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
         BCOL            =   16744576
         BCOLO           =   16744576
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMain.frx":0A56
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin WinTweaker.Button2K cmdNetwork 
         Height          =   375
         Left            =   360
         TabIndex        =   19
         Top             =   2040
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "Network"
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
         BCOL            =   16744576
         BCOLO           =   16744576
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMain.frx":0A72
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin WinTweaker.Button2K cmdDesktop 
         Height          =   375
         Left            =   360
         TabIndex        =   20
         Top             =   1560
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "Desktop"
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
         BCOL            =   16744576
         BCOLO           =   16744576
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMain.frx":0A8E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin WinTweaker.Button2K cmdControlPanel 
         Height          =   375
         Left            =   360
         TabIndex        =   21
         Top             =   1080
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "Control Panel"
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
         BCOL            =   16744576
         BCOLO           =   16744576
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMain.frx":0AAA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin WinTweaker.Button2K cmdStartMenu 
         Height          =   375
         Left            =   360
         TabIndex        =   22
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "Start Menu"
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
         BCOL            =   16744576
         BCOLO           =   16744576
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMain.frx":0AC6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin WinTweaker.Button2K cmdAddRemove 
         Height          =   375
         Left            =   360
         TabIndex        =   23
         Top             =   120
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "Add/Remove"
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
         BCOL            =   16744576
         BCOLO           =   16744576
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMain.frx":0AE2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin WinTweaker.Button2K cmdRestrictRun 
         Height          =   375
         Left            =   360
         TabIndex        =   73
         Top             =   4440
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "Restrict Run"
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
         BCOL            =   16744576
         BCOLO           =   16744576
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMain.frx":0AFE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Frame fraWExplorer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5295
         Left            =   2520
         TabIndex        =   187
         Top             =   0
         Width           =   5655
         Begin VB.CheckBox chkWExplorer 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Hide all Disk Drives."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   192
            Top             =   1200
            Width           =   5055
         End
         Begin VB.CheckBox chkWExplorer 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable toolbar customization in Windows Explorer."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   191
            Top             =   960
            Width           =   5055
         End
         Begin VB.CheckBox chkWExplorer 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Small toolbar icons in Windows Explorer and Internet Explorer."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   190
            Top             =   720
            Width           =   5055
         End
         Begin VB.CheckBox chkWExplorer 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable ""File"" menu in Windows Explorer and Internet Explorer."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   189
            Top             =   480
            Width           =   5055
         End
         Begin VB.CheckBox chkWExplorer 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable context menu of Windows Explorer and Desktop."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   188
            Top             =   240
            Width           =   5055
         End
      End
      Begin VB.Frame fraAbout 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5295
         Left            =   2520
         TabIndex        =   7
         Top             =   0
         Width           =   5655
         Begin VB.Label lblAbout 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   $"frmMain.frx":0B1A
            ForeColor       =   &H80000008&
            Height          =   1215
            Index           =   9
            Left            =   480
            TabIndex        =   165
            Top             =   3960
            Width           =   4695
         End
         Begin VB.Label lblAbout 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WebSite: http://www.vanjo.com.ph"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   8
            Left            =   480
            MouseIcon       =   "frmMain.frx":0C7E
            MousePointer    =   99  'Custom
            TabIndex        =   164
            ToolTipText     =   "Visit us online..."
            Top             =   3600
            Width           =   2655
         End
         Begin VB.Label lblAbout 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "E-Mail: vanjo08@msn.com"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   7
            Left            =   480
            MouseIcon       =   "frmMain.frx":0F88
            MousePointer    =   99  'Custom
            TabIndex        =   163
            ToolTipText     =   "Send E-mail for any feedback..."
            Top             =   3360
            Width           =   1935
         End
         Begin VB.Label lblAbout 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Author: Vanni Alora"
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   6
            Left            =   480
            TabIndex        =   162
            Top             =   3120
            Width           =   4575
         End
         Begin VB.Label lblAbout 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "All Rights Reserved"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   480
            TabIndex        =   161
            Top             =   2760
            Width           =   4575
         End
         Begin VB.Label lblAbout 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "VJ™ Software"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   480
            TabIndex        =   160
            Top             =   2520
            Width           =   4575
         End
         Begin VB.Label lblAbout 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Copyright © 2001 - 2003"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   480
            TabIndex        =   159
            Top             =   2280
            Width           =   4575
         End
         Begin VB.Label lblAbout 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "For Windows 98/ME (Millenium Edition)"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   480
            TabIndex        =   158
            Top             =   2040
            Width           =   4575
         End
         Begin VB.Label lblAbout 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Version"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   157
            Top             =   1800
            Width           =   4575
         End
         Begin VB.Label lblAbout 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WinTweaker (Windows Personal Security)"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   156
            Top             =   1560
            Width           =   4575
         End
         Begin VB.Image imgAbout 
            Appearance      =   0  'Flat
            Height          =   5100
            Left            =   120
            Picture         =   "frmMain.frx":1292
            Top             =   150
            Width           =   5400
         End
      End
      Begin VB.Frame fraAddRemove 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5295
         Left            =   2520
         TabIndex        =   2
         Top             =   0
         Width           =   5655
         Begin VB.TextBox txtUninstallString 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   53
            Top             =   4440
            Width           =   5415
         End
         Begin VB.ListBox lstPrograms 
            Appearance      =   0  'Flat
            Height          =   3930
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   3
            Top             =   240
            Width           =   5415
         End
         Begin WinTweaker.Button2K cmdRefresh 
            Height          =   375
            Left            =   600
            TabIndex        =   4
            Top             =   4800
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            BTYPE           =   5
            TX              =   "Refresh"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   0   'False
            BCOL            =   14737632
            BCOLO           =   14737632
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmMain.frx":4549
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin WinTweaker.Button2K cmdRemove 
            Height          =   375
            Left            =   3720
            TabIndex        =   5
            Top             =   4800
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            BTYPE           =   5
            TX              =   "Remove"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   0   'False
            BCOL            =   14737632
            BCOLO           =   14737632
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmMain.frx":4565
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin WinTweaker.Button2K cmdUninstall 
            Height          =   375
            Left            =   2160
            TabIndex        =   6
            Top             =   4800
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            BTYPE           =   5
            TX              =   "Uninstall"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   0   'False
            BCOL            =   14737632
            BCOLO           =   14737632
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmMain.frx":4581
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Command:"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   4200
            Width           =   855
         End
      End
      Begin VB.Frame fraStartMenu 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5295
         Left            =   2520
         TabIndex        =   24
         Top             =   0
         Width           =   5655
         Begin VB.TextBox txtStartMenuDelay 
            Height          =   285
            Left            =   2040
            TabIndex        =   52
            Text            =   "400"
            Top             =   4400
            Width           =   2175
         End
         Begin VB.CheckBox chkStartMenu 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable the Start Menu Edition."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   16
            Left            =   240
            TabIndex        =   50
            Top             =   4080
            Width           =   5175
         End
         Begin VB.CheckBox chkStartMenu 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable context menu of the Start Menu."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   15
            Left            =   240
            TabIndex        =   49
            Top             =   3840
            Width           =   5175
         End
         Begin VB.CheckBox chkStartMenu 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Hide the Start Menu subfolders."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   14
            Left            =   240
            TabIndex        =   48
            Top             =   3600
            Width           =   5175
         End
         Begin VB.CheckBox chkStartMenu 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Do not remember recent documents."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   13
            Left            =   240
            TabIndex        =   47
            Top             =   3360
            Width           =   5175
         End
         Begin VB.CheckBox chkStartMenu 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Clear ""Documents"" on exit."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   12
            Left            =   240
            TabIndex        =   46
            Top             =   3120
            Width           =   5175
         End
         Begin VB.CheckBox chkStartMenu 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable ""Settings"" -> ""Windows Update"" option."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   11
            Left            =   240
            TabIndex        =   45
            Top             =   2880
            Width           =   5175
         End
         Begin VB.CheckBox chkStartMenu 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable ""Settings"" -> ""Active Desktop"" option."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   10
            Left            =   240
            TabIndex        =   44
            Top             =   2640
            Width           =   5175
         End
         Begin VB.CheckBox chkStartMenu 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable ""Settings"" -> ""Folder Options"" option."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   9
            Left            =   240
            TabIndex        =   43
            Top             =   2400
            Width           =   5175
         End
         Begin VB.CheckBox chkStartMenu 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable ""Settings"" -> ""TaskBar"" and ""Start Menu"" option."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   8
            Left            =   240
            TabIndex        =   42
            Top             =   2160
            Width           =   5175
         End
         Begin VB.CheckBox chkStartMenu 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable ""Settings"" -> ""Control Panel"" and ""Printers"" option."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   41
            Top             =   1920
            Width           =   5175
         End
         Begin VB.CheckBox chkStartMenu 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable ""Documents"" option."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   40
            Top             =   1680
            Width           =   5175
         End
         Begin VB.CheckBox chkStartMenu 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable ""Favorites"" option."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   39
            Top             =   1440
            Width           =   5175
         End
         Begin VB.CheckBox chkStartMenu 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable ""Shut Down"" option."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   38
            Top             =   1200
            Width           =   5175
         End
         Begin VB.CheckBox chkStartMenu 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable ""Log Off"" option."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   37
            Top             =   960
            Width           =   5175
         End
         Begin VB.CheckBox chkStartMenu 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable ""Help"" option."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   36
            Top             =   720
            Width           =   5175
         End
         Begin VB.CheckBox chkStartMenu 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable ""Find"" option."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   35
            Top             =   480
            Width           =   5175
         End
         Begin VB.CheckBox chkStartMenu 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable ""Run"" option."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   25
            Top             =   240
            Width           =   5175
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Submenu show delay:"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   360
            TabIndex        =   51
            Top             =   4440
            Width           =   1575
         End
      End
      Begin VB.Frame fraControlPanel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5295
         Left            =   2520
         TabIndex        =   55
         Top             =   0
         Width           =   5655
         Begin VB.CheckBox chkControlPanel 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable ""System"" -> ""Virtual Memory"" page."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   16
            Left            =   240
            TabIndex        =   56
            Top             =   4080
            Width           =   5175
         End
         Begin VB.CheckBox chkControlPanel 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable ""System"" -> ""File System..."" page."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   15
            Left            =   240
            TabIndex        =   57
            Top             =   3840
            Width           =   5175
         End
         Begin VB.CheckBox chkControlPanel 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable ""System"" -> ""Device Manager"" page."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   14
            Left            =   240
            TabIndex        =   58
            Top             =   3600
            Width           =   5175
         End
         Begin VB.CheckBox chkControlPanel 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable ""System"" -> ""Hardware Profiles"" page."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   13
            Left            =   240
            TabIndex        =   59
            Top             =   3360
            Width           =   5175
         End
         Begin VB.CheckBox chkControlPanel 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable ""Passwords"" -> ""User Profiles"" page."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   12
            Left            =   240
            TabIndex        =   60
            Top             =   3120
            Width           =   5175
         End
         Begin VB.CheckBox chkControlPanel 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable ""Passwords"" -> ""Remote Administration"" page."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   11
            Left            =   240
            TabIndex        =   61
            Top             =   2880
            Width           =   5175
         End
         Begin VB.CheckBox chkControlPanel 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable ""Passwords"" -> ""Change Passwords"" page."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   10
            Left            =   240
            TabIndex        =   62
            Top             =   2640
            Width           =   5175
         End
         Begin VB.CheckBox chkControlPanel 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable ""Control Panel"" -> ""Passwords"" item."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   9
            Left            =   240
            TabIndex        =   63
            Top             =   2400
            Width           =   5175
         End
         Begin VB.CheckBox chkControlPanel 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable ""Display"" -> ""Settings"" page."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   8
            Left            =   240
            TabIndex        =   64
            Top             =   2160
            Width           =   5175
         End
         Begin VB.CheckBox chkControlPanel 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable ""Display"" -> ""Apperance"" page."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   65
            Top             =   1920
            Width           =   5175
         End
         Begin VB.CheckBox chkControlPanel 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable ""Display"" -> ""Screen Saver"" page."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   66
            Top             =   1680
            Width           =   5175
         End
         Begin VB.CheckBox chkControlPanel 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable ""Display"" -> ""Background"" page."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   67
            Top             =   1440
            Width           =   5175
         End
         Begin VB.CheckBox chkControlPanel 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable ""Control Panel"" -> ""Display"" item."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   68
            Top             =   1200
            Width           =   5175
         End
         Begin VB.CheckBox chkControlPanel 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable ""Printer"" -> ""General"" and ""Details"" page."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   69
            Top             =   960
            Width           =   5175
         End
         Begin VB.CheckBox chkControlPanel 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable deletion of Printers."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   70
            Top             =   720
            Width           =   5175
         End
         Begin VB.CheckBox chkControlPanel 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable addition of Printers."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   71
            Top             =   480
            Width           =   5175
         End
         Begin VB.CheckBox chkControlPanel 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable ""Control Panel"" -> Printer item."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   72
            Top             =   240
            Width           =   5175
         End
      End
      Begin VB.Frame fraDesktop 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5295
         Left            =   2520
         TabIndex        =   74
         Top             =   0
         Width           =   5655
         Begin VB.CheckBox chkDesktop 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Hide ""Internet Explorer"" icon from Desktop."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   17
            Left            =   240
            TabIndex        =   92
            Top             =   4320
            Width           =   5055
         End
         Begin VB.CheckBox chkDesktop 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Hide ""Network Neighborhood"" icon from Desktop."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   16
            Left            =   240
            TabIndex        =   91
            Top             =   4080
            Width           =   5055
         End
         Begin VB.CheckBox chkDesktop 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Allow changing name and removing ""Recycle Bin"" from Desktop."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   15
            Left            =   240
            TabIndex        =   90
            Top             =   3840
            Width           =   5055
         End
         Begin VB.CheckBox chkDesktop 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Show Windows Version on Desktop."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   14
            Left            =   240
            TabIndex        =   89
            Top             =   3600
            Width           =   5055
         End
         Begin VB.CheckBox chkDesktop 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable context menu of Taskbar."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   13
            Left            =   240
            TabIndex        =   88
            Top             =   3360
            Width           =   5055
         End
         Begin VB.CheckBox chkDesktop 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable ""Click here to begin"" on Taskbar."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   12
            Left            =   240
            TabIndex        =   87
            Top             =   3120
            Width           =   5055
         End
         Begin VB.CheckBox chkDesktop 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable resizing all toolbars on Taskbar."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   11
            Left            =   240
            TabIndex        =   86
            Top             =   2880
            Width           =   5055
         End
         Begin VB.CheckBox chkDesktop 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable dragging, dropping and closing all toolbars on Taskbar."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   10
            Left            =   240
            TabIndex        =   85
            Top             =   2640
            Width           =   5055
         End
         Begin VB.CheckBox chkDesktop 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable changing wallpaper."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   9
            Left            =   240
            TabIndex        =   84
            Top             =   2400
            Width           =   5055
         End
         Begin VB.CheckBox chkDesktop 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable choosing HTML page as a wallpaper."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   8
            Left            =   240
            TabIndex        =   83
            Top             =   2160
            Width           =   5055
         End
         Begin VB.CheckBox chkDesktop 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable closing components of Active Desktop."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   82
            Top             =   1920
            Width           =   5055
         End
         Begin VB.CheckBox chkDesktop 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable editing components of Active Desktop."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   81
            Top             =   1680
            Width           =   5055
         End
         Begin VB.CheckBox chkDesktop 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable deleting components from Active Desktop."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   80
            Top             =   1440
            Width           =   5055
         End
         Begin VB.CheckBox chkDesktop 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable adding components to Active Desktop."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   79
            Top             =   1200
            Width           =   5055
         End
         Begin VB.CheckBox chkDesktop 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable Active Desktop changing."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   78
            Top             =   960
            Width           =   5055
         End
         Begin VB.CheckBox chkDesktop 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable Active Desktop."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   77
            Top             =   720
            Width           =   5055
         End
         Begin VB.CheckBox chkDesktop 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable Desktop."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   76
            Top             =   480
            Width           =   5055
         End
         Begin VB.CheckBox chkDesktop 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Do not save Desktop setting on exit."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   75
            Top             =   240
            Width           =   5055
         End
      End
      Begin VB.Frame fraNetwork 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5295
         Left            =   2520
         TabIndex        =   93
         Top             =   0
         Width           =   5655
         Begin VB.CheckBox chkNetworkChangeName 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Change Machine Name?"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2640
            TabIndex        =   109
            Top             =   3680
            Width           =   2295
         End
         Begin VB.TextBox txtNetworkCurrentName 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2640
            TabIndex        =   108
            Top             =   3360
            Width           =   2655
         End
         Begin VB.TextBox txtNetworkTTLTcp 
            Height          =   285
            Left            =   2640
            TabIndex        =   106
            Top             =   2880
            Width           =   2655
         End
         Begin VB.TextBox txtNetworkPwdLength 
            Height          =   285
            Left            =   2640
            TabIndex        =   105
            Top             =   2520
            Width           =   2655
         End
         Begin VB.CheckBox chkNetwork 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable close button on logon dialog box."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   8
            Left            =   240
            TabIndex        =   102
            Top             =   2160
            Width           =   4935
         End
         Begin VB.CheckBox chkNetwork 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Hide printer sharing."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   101
            Top             =   1920
            Width           =   4935
         End
         Begin VB.CheckBox chkNetwork 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Hide file sharing."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   100
            Top             =   1680
            Width           =   4935
         End
         Begin VB.CheckBox chkNetwork 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable mapping and disconnecting network drive."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   99
            Top             =   1440
            Width           =   4935
         End
         Begin VB.CheckBox chkNetwork 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Hide Workgroups from ""Network Neighborhood""."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   98
            Top             =   1200
            Width           =   4935
         End
         Begin VB.CheckBox chkNetwork 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Hide ""Entire Network"" from ""Network Neighborhood""."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   97
            Top             =   960
            Width           =   4935
         End
         Begin VB.CheckBox chkNetwork 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable ""Network"" -> ""Access Control"" page."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   96
            Top             =   720
            Width           =   4935
         End
         Begin VB.CheckBox chkNetwork 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable ""Network"" -> ""Identification"" page."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   95
            Top             =   480
            Width           =   4935
         End
         Begin VB.CheckBox chkNetwork 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable ""Control Panel"" -> ""Network"" item."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   94
            Top             =   240
            Width           =   4935
         End
         Begin VB.Label lblMachineName 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Current Machine Name:"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   360
            TabIndex        =   107
            Top             =   3400
            Width           =   1815
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Default Time To Live for TCP:"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   360
            TabIndex        =   104
            Top             =   2925
            Width           =   2175
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Minimum Password Length:"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   360
            TabIndex        =   103
            Top             =   2565
            Width           =   2175
         End
      End
      Begin VB.Frame fraSystem 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5295
         Left            =   2520
         TabIndex        =   110
         Top             =   0
         Width           =   5655
         Begin VB.CheckBox chkSystem 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Show arrow on "".lnk"" shortcuts."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   13
            Left            =   240
            TabIndex        =   130
            Top             =   4440
            Width           =   4935
         End
         Begin VB.CheckBox chkSystem 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable new Shell."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   12
            Left            =   240
            TabIndex        =   129
            Top             =   4200
            Width           =   4935
         End
         Begin VB.CheckBox chkSystem 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Show window contents while dragging."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   11
            Left            =   240
            TabIndex        =   128
            Top             =   3960
            Width           =   4935
         End
         Begin VB.CheckBox chkSystem 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Smooth edges of screen fonts."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   10
            Left            =   240
            TabIndex        =   127
            Top             =   3720
            Width           =   4935
         End
         Begin VB.CheckBox chkSystem 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Animate windows, menus and lists."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   9
            Left            =   240
            TabIndex        =   126
            Top             =   3480
            Width           =   4935
         End
         Begin VB.CheckBox chkSystem 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Show icons using all possible colors."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   8
            Left            =   240
            TabIndex        =   125
            Top             =   3240
            Width           =   4935
         End
         Begin VB.CheckBox chkSystem 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Use large icons."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   124
            Top             =   3000
            Width           =   4935
         End
         Begin VB.CheckBox chkSystem 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Set registration on Windows Update."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   123
            Top             =   2760
            Width           =   4935
         End
         Begin VB.CheckBox chkSystem 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable MS-DOS mode."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   122
            Top             =   2520
            Width           =   4935
         End
         Begin VB.CheckBox chkSystem 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable restarting in MS-DOS mode."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   121
            Top             =   2280
            Width           =   4935
         End
         Begin VB.CheckBox chkSystem 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Only allow approved Shell extensions."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   120
            Top             =   2040
            Width           =   4935
         End
         Begin VB.CheckBox chkSystem 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable AutoRun for CD-ROM's."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   119
            Top             =   1800
            Width           =   4935
         End
         Begin VB.CheckBox chkSystem 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable Windows Update during driver searching."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   118
            Top             =   1560
            Width           =   4935
         End
         Begin VB.CheckBox chkSystem 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable Registry Editor."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   117
            Top             =   1320
            Width           =   4935
         End
         Begin VB.TextBox txtSystemSourcePath 
            Height          =   285
            Left            =   2160
            TabIndex        =   116
            Top             =   960
            Width           =   3135
         End
         Begin VB.TextBox txtSystemOrganization 
            Height          =   285
            Left            =   2160
            TabIndex        =   115
            Top             =   600
            Width           =   3135
         End
         Begin VB.TextBox txtSystemOwner 
            Height          =   285
            Left            =   2160
            TabIndex        =   114
            Top             =   240
            Width           =   3135
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Setup Source Path:"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   113
            Top             =   1000
            Width           =   1815
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Registered Organization:"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   112
            Top             =   640
            Width           =   1815
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Registered Owner:"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   111
            Top             =   280
            Width           =   1815
         End
      End
      Begin VB.Frame fraHints 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5295
         Left            =   2520
         TabIndex        =   131
         Top             =   0
         Width           =   5655
         Begin VB.TextBox txtHintsWebFolders 
            Height          =   285
            Left            =   2160
            TabIndex        =   155
            Top             =   4200
            Width           =   3135
         End
         Begin VB.TextBox txtHintsTasks 
            Height          =   285
            Left            =   2160
            TabIndex        =   154
            Top             =   3840
            Width           =   3135
         End
         Begin VB.TextBox txtHintsBriefcase 
            Height          =   285
            Left            =   2160
            TabIndex        =   153
            Top             =   3480
            Width           =   3135
         End
         Begin VB.TextBox txtHintsDialUp 
            Height          =   285
            Left            =   2160
            TabIndex        =   152
            Top             =   3120
            Width           =   3135
         End
         Begin VB.TextBox txtHintsControlPanel 
            Height          =   285
            Left            =   2160
            TabIndex        =   151
            Top             =   2760
            Width           =   3135
         End
         Begin VB.TextBox txtHintsPrinters 
            Height          =   285
            Left            =   2160
            TabIndex        =   150
            Top             =   2400
            Width           =   3135
         End
         Begin VB.TextBox txtHintsInternetExplorer 
            Height          =   285
            Left            =   2160
            TabIndex        =   149
            Top             =   2040
            Width           =   3135
         End
         Begin VB.TextBox txtHintsInternet 
            Height          =   285
            Left            =   2160
            TabIndex        =   148
            Top             =   1680
            Width           =   3135
         End
         Begin VB.TextBox txtHintsNetworkNeighborhood 
            Height          =   285
            Left            =   2160
            TabIndex        =   147
            Top             =   1320
            Width           =   3135
         End
         Begin VB.TextBox txtHintsMyDocuments 
            Height          =   285
            Left            =   2160
            TabIndex        =   146
            Top             =   960
            Width           =   3135
         End
         Begin VB.TextBox txtHintsMyComputer 
            Height          =   285
            Left            =   2160
            TabIndex        =   145
            Top             =   600
            Width           =   3135
         End
         Begin VB.TextBox txtHintsRecycleBin 
            Height          =   285
            Left            =   2160
            TabIndex        =   144
            Top             =   240
            Width           =   3135
         End
         Begin VB.Label Label19 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Web Folders"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   143
            Top             =   4240
            Width           =   1815
         End
         Begin VB.Label Label18 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Tasks"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   142
            Top             =   3880
            Width           =   1815
         End
         Begin VB.Label Label17 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Briefcase"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   141
            Top             =   3520
            Width           =   1815
         End
         Begin VB.Label Label16 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Dial-Up Networking"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   140
            Top             =   3160
            Width           =   1815
         End
         Begin VB.Label Label15 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Control Panel"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   139
            Top             =   2800
            Width           =   1815
         End
         Begin VB.Label Label14 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Printers"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   138
            Top             =   2440
            Width           =   1815
         End
         Begin VB.Label Label13 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Internet Explorer"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   137
            Top             =   2080
            Width           =   1815
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Internet"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   136
            Top             =   1720
            Width           =   1815
         End
         Begin VB.Label Label11 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Network Neighborhood"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   135
            Top             =   1360
            Width           =   1815
         End
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "My Documents"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   134
            Top             =   1000
            Width           =   1815
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "My Computer"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   133
            Top             =   640
            Width           =   1815
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Recycle Bin"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   132
            Top             =   280
            Width           =   1815
         End
      End
      Begin VB.Frame fraIExplorer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5295
         Left            =   2520
         TabIndex        =   166
         Top             =   0
         Width           =   5655
         Begin VB.CheckBox chkIExplorer 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable Hotmail box in Outlook Express."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   19
            Left            =   240
            TabIndex        =   186
            Top             =   4800
            Width           =   4815
         End
         Begin VB.CheckBox chkIExplorer 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Show Font button in Internet Explorer."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   18
            Left            =   240
            TabIndex        =   185
            Top             =   4560
            Width           =   4815
         End
         Begin VB.CheckBox chkIExplorer 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable toolbar buttons customization."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   17
            Left            =   240
            TabIndex        =   184
            Top             =   4320
            Width           =   4815
         End
         Begin VB.CheckBox chkIExplorer 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable adding, removing and moving toolbars."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   16
            Left            =   240
            TabIndex        =   183
            Top             =   4080
            Width           =   4815
         End
         Begin VB.CheckBox chkIExplorer 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable Splash Screen."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   15
            Left            =   240
            TabIndex        =   182
            Top             =   3840
            Width           =   4815
         End
         Begin VB.CheckBox chkIExplorer 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable browsing local files."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   14
            Left            =   240
            TabIndex        =   181
            Top             =   3600
            Width           =   4815
         End
         Begin VB.CheckBox chkIExplorer 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable Channel user interface."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   13
            Left            =   240
            TabIndex        =   180
            Top             =   3360
            Width           =   4815
         End
         Begin VB.CheckBox chkIExplorer 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable icon size changing."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   12
            Left            =   240
            TabIndex        =   179
            Top             =   3120
            Width           =   4815
         End
         Begin VB.CheckBox chkIExplorer 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable Find Files command."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   11
            Left            =   240
            TabIndex        =   178
            Top             =   2880
            Width           =   4815
         End
         Begin VB.CheckBox chkIExplorer 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable closing browser."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   10
            Left            =   240
            TabIndex        =   177
            Top             =   2640
            Width           =   4815
         End
         Begin VB.CheckBox chkIExplorer 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable HTML context menu."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   9
            Left            =   240
            TabIndex        =   176
            Top             =   2400
            Width           =   4815
         End
         Begin VB.CheckBox chkIExplorer 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable selecting download directory."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   8
            Left            =   240
            TabIndex        =   175
            Top             =   2160
            Width           =   4815
         End
         Begin VB.CheckBox chkIExplorer 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable ""Favorites"" menu."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   174
            Top             =   1920
            Width           =   4815
         End
         Begin VB.CheckBox chkIExplorer 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable ""View"" -> ""Full Screen"" option."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   173
            Top             =   1680
            Width           =   4815
         End
         Begin VB.CheckBox chkIExplorer 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable ""Tools"" -> ""Windows Update"" option."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   172
            Top             =   1440
            Width           =   4815
         End
         Begin VB.CheckBox chkIExplorer 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable ""Tools"" -> ""Internet Options""."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   171
            Top             =   1200
            Width           =   4815
         End
         Begin VB.CheckBox chkIExplorer 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable ""Save As Web (Complete)"" option."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   170
            Top             =   960
            Width           =   4815
         End
         Begin VB.CheckBox chkIExplorer 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable ""Save"" and ""Save As""."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   169
            Top             =   720
            Width           =   4815
         End
         Begin VB.CheckBox chkIExplorer 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable ""File"" -> ""New"" option."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   168
            Top             =   480
            Width           =   4815
         End
         Begin VB.CheckBox chkIExplorer 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable ""File"" -> ""Open"" option."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   167
            Top             =   240
            Width           =   4815
         End
      End
   End
   Begin VB.Image imgClose 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   8190
      Picture         =   "frmMain.frx":459D
      ToolTipText     =   "Terminate the program..."
      Top             =   0
      Width           =   330
   End
   Begin VB.Label lblTitleBar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "WinTweaker 1.0 for Win9x/ME"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8535
   End
   Begin VB.Image imgTitleBar 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   0
      Picture         =   "frmMain.frx":4698
      Top             =   0
      Width           =   8550
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private setX As Integer, setY As Integer

Private Sub chkControlPanel_Click(Index As Integer)
    On Error Resume Next
    ' call ColoredValue function
    ColoredValue chkControlPanel(Index)
End Sub

Private Sub chkControlPanel_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    ' show description...
    Select Case Index
        Case 0
            Description chkControlPanel(0), lblDescription, "Disable Printer item."
        Case 1
            Description chkControlPanel(1), lblDescription, "Disable addition of printers."
        Case 2
            Description chkControlPanel(2), lblDescription, "Disable deletion of printers."
        Case 3
            Description chkControlPanel(3), lblDescription, "Disable Printer's ""General"" and ""Details"" page."
        Case 4
            Description chkControlPanel(4), lblDescription, "Disable ""Display Properties"" neither from Control Panel nor from the context menu (right click) of Desktop."
        Case 5
            Description chkControlPanel(5), lblDescription, "Disable Display's ""Background"" page."
        Case 6
            Description chkControlPanel(6), lblDescription, "Disable Display's ""Screen Saver"" page."
        Case 7
            Description chkControlPanel(7), lblDescription, "Disable Display's ""Appearance"" page."
        Case 8
            Description chkControlPanel(8), lblDescription, "Disable Display's ""Settings"" page."
        Case 9
            Description chkControlPanel(9), lblDescription, "Disable ""Passwords Properties"" item from Control Panel."
        Case 10
            Description chkControlPanel(10), lblDescription, "Disable Password's ""Change Passwords"" page."
        Case 11
            Description chkControlPanel(11), lblDescription, "Disable Password's ""Remote Administration"" page."
        Case 12
            Description chkControlPanel(12), lblDescription, "Disable Password's ""User Profiles"" page."
        Case 13
            Description chkControlPanel(13), lblDescription, "Disable System's ""Hardware Profiles"" page."
        Case 14
            Description chkControlPanel(14), lblDescription, "Disable System's ""Device Manager"" page."
        Case 15
            Description chkControlPanel(15), lblDescription, "Disable System's ""File System..."" page."
        Case 16
            Description chkControlPanel(16), lblDescription, "Disable System's ""Virtual Memory"" page."
    End Select
End Sub

Private Sub chkDesktop_Click(Index As Integer)
    On Error Resume Next
    ' call ColoredValue function...
    ColoredValue chkDesktop(Index)
End Sub

Private Sub chkDesktop_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    ' show description...
    Select Case Index
        Case 0
            Description chkDesktop(0), lblDescription, "No changes on Desktop will be stored."
        Case 1
            Description chkDesktop(1), lblDescription, "All icons from Desktop will be removed and you cannot invoke context menu (right click) on it. You can run ""Display Properties"" from Control Panel."
        Case 2
            Description chkDesktop(2), lblDescription, "You cannot enable this time-consuming Active Desktop."
        Case 3
            Description chkDesktop(3), lblDescription, "You cannot change anything on Active Desktop."
        Case 4
            Description chkDesktop(4), lblDescription, "You cannot add components to Active Desktop."
        Case 5
            Description chkDesktop(5), lblDescription, "You cannot delete components to Active Desktop."
        Case 6
            Description chkDesktop(6), lblDescription, "You cannot edit components to Active Desktop."
        Case 7
            Description chkDesktop(7), lblDescription, "You cannot close components to Active Desktop."
        Case 8
            Description chkDesktop(8), lblDescription, "You can no longer choose your favorite HTML page as a wallpaper."
        Case 9
            Description chkDesktop(9), lblDescription, "You cannot change Desktop wallpaper."
        Case 10
            Description chkDesktop(10), lblDescription, "You cannot drag, drop or close any toolbars on Taskbar."
        Case 11
            Description chkDesktop(11), lblDescription, "However you can double click on the bar on the left side of the toolbar to resize it."
        Case 12
            Description chkDesktop(12), lblDescription, "Hides an arrow and a string ""Click here to begin..."", which appear after Windows start."
        Case 13
            Description chkDesktop(13), lblDescription, "You cannot show context menu (right click) for Taskbar together with ""Start"" button, user toolbars and the clock."
        Case 14
            Description chkDesktop(14), lblDescription, "Show name and version of the Operating System in the right-bottom corner of Desktop."
        Case 15
            Description chkDesktop(15), lblDescription, "Two new options appear in the context menu (right click) of ""Recycle Bin"": ""Rename"" and ""Delete""."
        Case 16
            Description chkDesktop(16), lblDescription, "Use this option with care. They say it maybe dangerous when you use local network."
        Case 17
            Description chkDesktop(17), lblDescription, "Hide the ""Internet Explorer"" icon on Desktop. Useful to clean up your Desktop."
    End Select
End Sub

Private Sub chkIExplorer_Click(Index As Integer)
    On Error Resume Next
    ' call ColoredValue function...
    ColoredValue chkIExplorer(Index)
End Sub

Private Sub chkIExplorer_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    ' show description...
    Select Case Index
        Case 0
            Description chkIExplorer(Index), lblDescription, "Also ""CTRL + O"" and ""CTRL + L""."
        Case 1
            Description chkIExplorer(Index), lblDescription, "Also ""CTRL + N""."
        Case 2
            Description chkIExplorer(Index), lblDescription, "Disable ""Save"" options and ""CTRL + S"" key."
        Case 3
            Description chkIExplorer(Index), lblDescription, ""
        Case 4
            Description chkIExplorer(Index), lblDescription, "Disable changing browser settings."
        Case 5
            Description chkIExplorer(Index), lblDescription, ""
        Case 6
            Description chkIExplorer(Index), lblDescription, "Disable ""Full Screen"" and ""F11"" key."
        Case 7
            Description chkIExplorer(Index), lblDescription, "Disable ""Favorites"" menu also adding to favorites or organize favorites."
        Case 8
            Description chkIExplorer(Index), lblDescription, "Prevents user from being able to select download folder by not displaying the ""Save As"" dialog box when a file is downloaded."
        Case 9
            Description chkIExplorer(Index), lblDescription, "Which you normally access by right click."
        Case 10
            Description chkIExplorer(Index), lblDescription, "Also disable the ""ALT + F4"" key."
        Case 11
            Description chkIExplorer(Index), lblDescription, "Disables the ""F3"" key. Pressing ""F3"" in Internet Explorer normally causes the Windows file-find dialog to pop up."
        Case 12
            Description chkIExplorer(Index), lblDescription, "Forces the icons to remain at the size they were when last configured."
        Case 13
            Description chkIExplorer(Index), lblDescription, "Hides the ""Channel"" user interface which came with Internet Explorer 4.0."
        Case 14
            Description chkIExplorer(Index), lblDescription, "Removes the ability to browse the local file system through ""file://"" URLs. This option works differently depending on IE version. (See Q179221 on Microsoft's Web Page)"
        Case 15
            Description chkIExplorer(Index), lblDescription, ""
        Case 16
            Description chkIExplorer(Index), lblDescription, ""
        Case 17
            Description chkIExplorer(Index), lblDescription, ""
        Case 18
            Description chkIExplorer(Index), lblDescription, ""
        Case 19
            Description chkIExplorer(Index), lblDescription, ""
    End Select
End Sub

Private Sub chkNetwork_Click(Index As Integer)
    On Error Resume Next
    ' call ColoredValue function...
    ColoredValue chkNetwork(Index)
End Sub

Private Sub chkNetwork_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    ' show description...
    Select Case Index
        Case 0
            Description chkNetwork(0), lblDescription, "You cannot run ""Network Properties"" from Control Panel."
        Case 1
            Description chkNetwork(0), lblDescription, "Hide ""Indentification"" page from ""Network Properties""."
        Case 2
            Description chkNetwork(0), lblDescription, "Hide ""Access Control"" page from ""Network Properties""."
        Case 3
            Description chkNetwork(0), lblDescription, "Hide ""Entire Network"" icon from ""Network Neighborhood""."
        Case 4
            Description chkNetwork(0), lblDescription, "Hide ""Workgroups"" icon from ""Network Neighborhood""."
        Case 5
            Description chkNetwork(0), lblDescription, "Hides ""Map Network Drive"" and ""Disconnect Network Drive"" buttons of Explorer toolbar and from context menu (right click) of ""My Computer"" and ""Tools"" menu of Explorer."
        Case 6
            Description chkNetwork(0), lblDescription, "You can no longer control file sharing."
        Case 7
            Description chkNetwork(0), lblDescription, "You can no longer control printer sharing."
        Case 8
            Description chkNetwork(0), lblDescription, "User will have to provide a valid login and password to gain access to the local machine. Your machine must be part of a Windows Domain for this tweak to work, as the user must be authenticated by the network."
    End Select
End Sub

Private Sub chkNetworkChangeName_Click()
    On Error Resume Next
    If chkNetworkChangeName.Value = 1 Then
        txtNetworkCurrentName.Enabled = True
        txtNetworkCurrentName.SetFocus
    Else
        GetMachineName txtNetworkCurrentName
        txtNetworkCurrentName.Enabled = False
    End If
End Sub

Private Sub chkStartMenu_Click(Index As Integer)
    On Error Resume Next
    ' call ColoredValue function...
    ColoredValue chkStartMenu(Index)
End Sub

Private Sub chkStartMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    ' show description...
    Select Case Index
        Case 0
            Description chkStartMenu(0), lblDescription, "Hides ""Run"" option from the Start Menu, however you can use ""Win + R"" combination."
        Case 1
            Description chkStartMenu(1), lblDescription, "Hides ""Find"" option from the Start Menu, however you can use combination of ""Win + F"" or just press ""F3"" key."
        Case 2
            Description chkStartMenu(2), lblDescription, "Hides ""Help"" option from the Start Menu. Attention!!! This option might not work on your computer."
        Case 3
            Description chkStartMenu(3), lblDescription, "Hides ""Log Off"" option from the Start Menu."
        Case 4
            Description chkStartMenu(4), lblDescription, "Hides ""Shut Down..."" option from the Start Menu. You cannot close system neither by ""Alt + F4"", nor by ""Ctrl + Alt + Del"". However, you can restart your computer by pressing these three keys twice."
        Case 5
            Description chkStartMenu(5), lblDescription, "Hides ""Favorites"" option from the Start Menu."
        Case 6
            Description chkStartMenu(6), lblDescription, "Hides ""Documents"" option from the Start Menu."
        Case 7
            Description chkStartMenu(7), lblDescription, "Hides ""Control Panel"" and ""Printers"" from the ""Settings"" submenu and from Explorer. You can run Control Panel modules by running appropriate file with CPL extension. Additionally you cannot run Explorer by ""Win + E""."
        Case 8
            Description chkStartMenu(8), lblDescription, "Hides ""Taskbar"" and ""Start Menu"" from the ""Settings"" submenu. You cannot run Taskbar properties from its context menu (right click)."
        Case 9
            Description chkStartMenu(9), lblDescription, "Hides ""Folder Options"" from the ""Settings"" submenu and from the main menu of Explorer."
        Case 10
            Description chkStartMenu(10), lblDescription, "Hides ""Active Desktop"" option from ""Settings"" submenu."
        Case 11
            Description chkStartMenu(11), lblDescription, "Hides ""Windows Update"" option from the ""Settings"" submenu and you cannot execute it from Internet Explorer."
        Case 12
            Description chkStartMenu(12), lblDescription, "Everytime you close Windows... all shortcuts to documents stored in ""Documents"" submenu will be erased. So, next user will not know what you have opened."
        Case 13
            Description chkStartMenu(13), lblDescription, "No shortcuts to documents will be stored in ""Documents"" submenu. However this menu is not hidden."
        Case 14
            Description chkStartMenu(14), lblDescription, "Hides user programs folders from the Start Menu."
        Case 15
            Description chkStartMenu(15), lblDescription, "Disables context menu (right click) and moving elements of the Start Menu."
        Case 16
            Description chkStartMenu(16), lblDescription, "On this page you can set the appearance and options for the Start Menu."
    End Select
End Sub

Private Sub chkSystem_Click(Index As Integer)
    On Error Resume Next
    ' call ColoredValue function...
    ColoredValue chkSystem(Index)
End Sub

Private Sub chkSystem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    ' show description...
    Select Case Index
        Case 0
            Description chkSystem(Index), lblDescription, "You cannot run ""Registry Editor"" and some other registry editing applications."
        Case 1
            Description chkSystem(Index), lblDescription, "During searching for new drivers an option to search in Microsoft Windows Update will be disabled."
        Case 2
            Description chkSystem(Index), lblDescription, "Prevents autorun of CD-ROM's."
        Case 3
            Description chkSystem(Index), lblDescription, "You can run only files with extensions specified by System Administrator in the registry."
        Case 4
            Description chkSystem(Index), lblDescription, "Hides ""Restart in MS-DOS Mode"" option from the Shutdown window."
        Case 5
            Description chkSystem(Index), lblDescription, "You cannot run any MS-DOS programs."
        Case 6
            Description chkSystem(Index), lblDescription, "If you don't want the system to ask you about registration on Windows Update... check this."
        Case 7
            Description chkSystem(Index), lblDescription, "Sets icon size to 48 pixels instead of 32."
        Case 8
            Description chkSystem(Index), lblDescription, "Show icons using 16 bit (instead of 4) color palette."
        Case 9
            Description chkSystem(Index), lblDescription, "When you maximize and minimize your windows, there will be some animate effects. The same applies to menus etc."
        Case 10
            Description chkSystem(Index), lblDescription, "All fonts on a screen are smoother but Desktop is repainted more slowly."
        Case 11
            Description chkSystem(Index), lblDescription, "Not only a frame, but a whole window is showed while dragging it. Not for slower computers."
        Case 12
            Description chkSystem(Index), lblDescription, "Specifies whether Active Desktop, Web View and Thumbnails Views are disabled. This entry also specifies whether users can configure their system to open items by single-clicking."
        Case 13
            Description chkSystem(Index), lblDescription, "Hides the little arrow on "".lnk"" file shortcuts."
    End Select
End Sub

Private Sub chkWExplorer_Click(Index As Integer)
    On Error Resume Next
    ' call ColoredValue function...
    ColoredValue chkWExplorer(Index)
End Sub

Private Sub chkWExplorer_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    ' show descriptions...
    Select Case Index
        Case 0
            Description chkWExplorer(Index), lblDescription, "You cannot show context menu (right click) for Desktop and for right panel of Explorer."
        Case 1
            Description chkWExplorer(Index), lblDescription, "Disable the ""File"" menu in Windows Explorer and Internet Explorer."
        Case 2
            Description chkWExplorer(Index), lblDescription, ""
        Case 3
            Description chkWExplorer(Index), lblDescription, ""
        Case 4
            Description chkWExplorer(Index), lblDescription, "Hide all Disk Drives in ""My Computer"" and in ""Windows Explorer""."
    End Select
End Sub

Private Sub cmdAbout_Click()
    On Error Resume Next
    ' show the About frame box...
    fraAbout.Visible = True
    picArrow(0).Visible = False: fraAddRemove.Visible = False
    picArrow(1).Visible = False: fraStartMenu.Visible = False
    picArrow(2).Visible = False: fraControlPanel.Visible = False
    picArrow(3).Visible = False: fraDesktop.Visible = False
    picArrow(4).Visible = False: fraNetwork.Visible = False
    picArrow(5).Visible = False: fraSystem.Visible = False
    picArrow(6).Visible = False: fraHints.Visible = False
    picArrow(7).Visible = False: fraIExplorer.Visible = False
    picArrow(8).Visible = False: fraWExplorer.Visible = False
    picArrow(9).Visible = False
End Sub

Private Sub cmdAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    ' show description...
    lblDescription.Caption = "Show the About Box Dialog..."
End Sub

Private Sub cmdAddRemove_Click()
    On Error Resume Next
    ' show the Add/Remove Window and the little arrow...
    fraAbout.Visible = False
    picArrow(0).Visible = True: fraAddRemove.Visible = True: Call cmdRefresh_Click
    picArrow(1).Visible = False: fraStartMenu.Visible = False
    picArrow(2).Visible = False: fraControlPanel.Visible = False
    picArrow(3).Visible = False: fraDesktop.Visible = False
    picArrow(4).Visible = False: fraNetwork.Visible = False
    picArrow(5).Visible = False: fraSystem.Visible = False
    picArrow(6).Visible = False: fraHints.Visible = False
    picArrow(7).Visible = False: fraIExplorer.Visible = False
    picArrow(8).Visible = False: fraWExplorer.Visible = False
    picArrow(9).Visible = False
End Sub

Private Sub cmdApply_Click()
    On Error Resume Next
    ' apply all changes...
    Call SetMachineName(txtNetworkCurrentName)
    
    Call ApplyStartMenu
    Call ApplyControlPanel
    Call ApplyDesktop
    Call ApplyNetwork
    Call ApplySystem
    Call ApplyHints
    Call ApplyIExplorer
    Call ApplyWExplorer
End Sub

Private Sub cmdApply_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    ' show descriptions...
    lblDescription.Caption = "This will apply all the changes you've done to the system. In order for changes to take effect, you must restart the system..."
End Sub

Private Sub cmdControlPanel_Click()
    On Error Resume Next
    ' show the ControlPanel Window and the little arrow...
    fraAbout.Visible = False
    picArrow(0).Visible = False: fraAddRemove.Visible = False
    picArrow(1).Visible = False: fraStartMenu.Visible = False
    picArrow(2).Visible = True: fraControlPanel.Visible = True
    picArrow(3).Visible = False: fraDesktop.Visible = False
    picArrow(4).Visible = False: fraNetwork.Visible = False
    picArrow(5).Visible = False: fraSystem.Visible = False
    picArrow(6).Visible = False: fraHints.Visible = False
    picArrow(7).Visible = False: fraIExplorer.Visible = False
    picArrow(8).Visible = False: fraWExplorer.Visible = False
    picArrow(9).Visible = False
End Sub

Private Sub cmdDesktop_Click()
    On Error Resume Next
    ' show the Desktop Window and the little arrow...
    fraAbout.Visible = False
    picArrow(0).Visible = False: fraAddRemove.Visible = False
    picArrow(1).Visible = False: fraStartMenu.Visible = False
    picArrow(2).Visible = False: fraControlPanel.Visible = False
    picArrow(3).Visible = True: fraDesktop.Visible = True
    picArrow(4).Visible = False: fraNetwork.Visible = False
    picArrow(5).Visible = False: fraSystem.Visible = False
    picArrow(6).Visible = False: fraHints.Visible = False
    picArrow(7).Visible = False: fraIExplorer.Visible = False
    picArrow(8).Visible = False: fraWExplorer.Visible = False
    picArrow(9).Visible = False
End Sub

Private Sub cmdHints_Click()
    On Error Resume Next
    ' show the Desktop Window and the little arrow...
    fraAbout.Visible = False
    picArrow(0).Visible = False: fraAddRemove.Visible = False
    picArrow(1).Visible = False: fraStartMenu.Visible = False
    picArrow(2).Visible = False: fraControlPanel.Visible = False
    picArrow(3).Visible = False: fraDesktop.Visible = False
    picArrow(4).Visible = False: fraNetwork.Visible = False
    picArrow(5).Visible = False: fraSystem.Visible = False
    picArrow(6).Visible = True: fraHints.Visible = True
    picArrow(7).Visible = False: fraIExplorer.Visible = False
    picArrow(8).Visible = False: fraWExplorer.Visible = False
    picArrow(9).Visible = False
End Sub

Private Sub cmdIExplorer_Click()
    On Error Resume Next
    ' show the Desktop Window and the little arrow...
    fraAbout.Visible = False
    picArrow(0).Visible = False: fraAddRemove.Visible = False
    picArrow(1).Visible = False: fraStartMenu.Visible = False
    picArrow(2).Visible = False: fraControlPanel.Visible = False
    picArrow(3).Visible = False: fraDesktop.Visible = False
    picArrow(4).Visible = False: fraNetwork.Visible = False
    picArrow(5).Visible = False: fraSystem.Visible = False
    picArrow(6).Visible = False: fraHints.Visible = False
    picArrow(7).Visible = True: fraIExplorer.Visible = True
    picArrow(8).Visible = False: fraWExplorer.Visible = False
    picArrow(9).Visible = False
End Sub

Private Sub cmdNetwork_Click()
    On Error Resume Next
    ' show the Desktop Window and the little arrow...
    fraAbout.Visible = False
    picArrow(0).Visible = False: fraAddRemove.Visible = False
    picArrow(1).Visible = False: fraStartMenu.Visible = False
    picArrow(2).Visible = False: fraControlPanel.Visible = False
    picArrow(3).Visible = False: fraDesktop.Visible = False
    picArrow(4).Visible = True: fraNetwork.Visible = True
    picArrow(5).Visible = False: fraSystem.Visible = False
    picArrow(6).Visible = False: fraHints.Visible = False
    picArrow(7).Visible = False: fraIExplorer.Visible = False
    picArrow(8).Visible = False: fraWExplorer.Visible = False
    picArrow(9).Visible = False
End Sub

Private Sub cmdRefresh_Click()
    On Error Resume Next
    ' refresh the list of Registered Programs...
    Dim Count As Integer
    Dim returnName As Collection
    Dim returnSubs As Collection
    Dim displayName As String
    Dim uninstallString As String
    Dim Version As String
    
    cmdUninstall.Enabled = False
    cmdRemove.Enabled = False
    txtUninstallString.Text = ""
    
    lstPrograms.Clear
    Call EnumRegKeys(returnName, returnSubs, "HKEY_LOCAL_MACHINE", "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
    If returnName.Count > 0 Then
        For Count = 1 To returnName.Count
            displayName = GetSetting("", returnName(Count), "DisplayName", "", HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
            uninstallString = GetSetting("", returnName(Count), "UninstallString", "", HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
            Version = GetSetting("", returnName(Count), "Version", "", HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
            If displayName <> "" And uninstallString <> "" Then 'And Version <> "") Then
                Call lstPrograms.AddItem(displayName) ' & " - " & UninstallString)
            End If
        Next Count
    End If
End Sub

Private Sub cmdRefresh_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    ' show description...
    lblDescription.Caption = "Refresh the list of Add/Remove box..."
End Sub

Private Sub cmdRemove_Click()
    On Error Resume Next
    ' remove the Uninstall Entry from the registry...
    Dim Count As Integer
    Dim Count2 As Integer
    Dim returnName As Collection
    Dim returnSubs As Collection
    Dim displayName As String

    If lstPrograms.SelCount > 0 Then
        If MsgBox("Are you sure you want to delete this entry from the registry?", vbYesNo + vbQuestion, "Warning...") = vbYes Then
            For Count = 0 To lstPrograms.ListCount - 1
                If lstPrograms.Selected(Count) = True Then
                    Call EnumRegKeys(returnName, returnSubs, "HKEY_LOCAL_MACHINE", "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
                    If returnName.Count > 0 Then
                        For Count2 = 1 To returnName.Count
                            displayName = GetSetting("", returnName(Count2), "DisplayName", "", HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
                            If displayName = lstPrograms.List(Count) Then
                                Call DeleteSetting("", returnName(Count2), "", HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
                            End If
                        Next Count2
                    End If
                End If
            Next Count
        Else
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    
    Call cmdRefresh_Click
End Sub

Private Sub cmdRemove_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    ' show description...
    lblDescription.Caption = "Removes program's entry from registry... use this option with caution!!!"
End Sub

Private Sub cmdRestrictRun_Click()
    On Error Resume Next
    MsgBox "Sorry, but this option will be available in the next version of this software.", vbInformation + vbOKOnly, App.Title & " - RestrictRun"
End Sub

Private Sub cmdSetPassword_Click()
    On Error Resume Next
    ' show Set Password dialog box...
    frmPassword.Show vbModal, Me
End Sub

Private Sub cmdStartMenu_Click()
    On Error Resume Next
    ' show the Start Menu frame...
    fraAbout.Visible = False
    picArrow(0).Visible = False: fraAddRemove.Visible = False
    picArrow(1).Visible = True: fraStartMenu.Visible = True
    picArrow(2).Visible = False: fraControlPanel.Visible = False
    picArrow(3).Visible = False: fraDesktop.Visible = False
    picArrow(4).Visible = False: fraNetwork.Visible = False
    picArrow(5).Visible = False: fraSystem.Visible = False
    picArrow(6).Visible = False: fraHints.Visible = False
    picArrow(7).Visible = False: fraIExplorer.Visible = False
    picArrow(8).Visible = False: fraWExplorer.Visible = False
    picArrow(9).Visible = False
End Sub

Private Sub cmdSystem_Click()
    On Error Resume Next
    ' show the Start Menu frame...
    fraAbout.Visible = False
    picArrow(0).Visible = False: fraAddRemove.Visible = False
    picArrow(1).Visible = False: fraStartMenu.Visible = False
    picArrow(2).Visible = False: fraControlPanel.Visible = False
    picArrow(3).Visible = False: fraDesktop.Visible = False
    picArrow(4).Visible = False: fraNetwork.Visible = False
    picArrow(5).Visible = True: fraSystem.Visible = True
    picArrow(6).Visible = False: fraHints.Visible = False
    picArrow(7).Visible = False: fraIExplorer.Visible = False
    picArrow(8).Visible = False: fraWExplorer.Visible = False
    picArrow(9).Visible = False
End Sub

Private Sub cmdTurnOff_Click()
    On Error Resume Next
    ' show shutdown dialog box...
    frmShutdown.Show vbModal, Me
End Sub

Private Sub cmdUninstall_Click()
    On Error Resume Next
    ' run the Uninstall Command...
    Dim Count As Integer
    Dim Count2 As Integer
    Dim returnName As Collection
    Dim returnSubs As Collection
    Dim displayName As String
    Dim uninstallString As String

    If lstPrograms.SelCount > 0 Then
        For Count = 0 To lstPrograms.ListCount - 1
            If lstPrograms.Selected(Count) = True Then
                Call EnumRegKeys(returnName, returnSubs, "HKEY_LOCAL_MACHINE", "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
                If returnName.Count > 0 Then
                    For Count2 = 1 To returnName.Count
                        displayName = GetSetting("", returnName(Count2), "DisplayName", "", HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
                        uninstallString = GetSetting("", returnName(Count2), "UninstallString", "", HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall")
                        If displayName = lstPrograms.List(Count) Then
                            Call Shell(uninstallString, vbNormalFocus)
                        End If
                    Next Count2
                End If
            End If
        Next Count
    Else
        Exit Sub
    End If
End Sub

Private Sub cmdUninstall_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    ' show description...
    lblDescription.Caption = "Uninstall the selected program..."
End Sub

Private Sub cmdWExplorer_Click()
    On Error Resume Next
    ' show the Desktop Window and the little arrow...
    fraAbout.Visible = False
    picArrow(0).Visible = False: fraAddRemove.Visible = False
    picArrow(1).Visible = False: fraStartMenu.Visible = False
    picArrow(2).Visible = False: fraControlPanel.Visible = False
    picArrow(3).Visible = False: fraDesktop.Visible = False
    picArrow(4).Visible = False: fraNetwork.Visible = False
    picArrow(5).Visible = False: fraSystem.Visible = False
    picArrow(6).Visible = False: fraHints.Visible = False
    picArrow(7).Visible = False: fraIExplorer.Visible = False
    picArrow(8).Visible = True: fraWExplorer.Visible = True
    picArrow(9).Visible = False
End Sub

Private Sub Form_Load()
    On Error Resume Next
    ' default intro...
    lblAbout(1).Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    
    Call cmdAbout_Click
    
    Call GetMachineName(txtNetworkCurrentName)
    
    Call ReadStartMenu
    Call ReadControlPanel
    Call ReadDesktop
    Call ReadNetwork
    Call ReadSystem
    Call ReadHints
    Call ReadIExplorer
    Call ReadWExplorer
End Sub

Private Sub fraHints_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    ' show description...
    lblDescription.Caption = "On this page you can set hints which are displayed after pointing on a system icon or in Web mode of Explorer."
End Sub

Private Sub imgClose_Click()
    On Error Resume Next
    ' close the form and terminate the program...
    Unload Me: End
End Sub

Private Sub imgTitleBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    ' set the X and Y of form...
    setX = X
    setY = Y
End Sub

Private Sub imgTitleBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    ' move the form...
    If Button = 1 Then
        Me.Left = Me.Left + (X - setX)
        Me.Top = Me.Top + (Y - setY)
    End If
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    ' show description...
    lblDescription.Caption = "Sets time in milliseconds how fast a submenu is expanded in the Start Menu."
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    ' show description...
    lblDescription.Caption = "Please enter two-digit hexadecimal value."
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    ' show description...
    lblDescription.Caption = "Specifies the default Time To Live (TTL) value in the header of outgoing IP packets. The TTL determines how long an IP packet that has not reached its destination can remain on the network before it is discarded."
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    ' show description...
    lblDescription.Caption = "Change the name of registered owner which is displayed on System Properties."
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    ' show description...
    lblDescription.Caption = "Change the name of registered organization which is displayed on System Properties."
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    ' show description...
    lblDescription.Caption = "Change the default path, from which the system was installed."
End Sub

Private Sub lblAbout_Click(Index As Integer)
    On Error Resume Next
    ' send email...
    If Index = 7 Then
        modFunctions.SendEmail "vanjo@msn.com", "WinTweaker Feedback"
    ElseIf Index = 8 Then
        modFunctions.GoToUrl "http://www.vanjo.com.ph"
    Else
        Exit Sub
    End If
End Sub

Private Sub lblDescription_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    ' show description...
    lblDescription.Caption = "You can see a detailed description of almost every item here..."
End Sub

Private Sub lblTitleBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    ' set the X and Y for form...
    setX = X
    setY = Y
End Sub

Private Sub lblTitleBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    ' move the form...
    If Button = 1 Then
        Me.Left = Me.Left + (X - setX)
        Me.Top = Me.Top + (Y - setY)
    End If
End Sub

Private Sub lstPrograms_Click()
    On Error Resume Next
    ' display the Uninstall String...
    If lstPrograms.SelCount > 0 Then
        cmdRemove.Enabled = True
        cmdUninstall.Enabled = True
    Else
        cmdRemove.Enabled = False
        cmdUninstall.Enabled = False
    End If
    
    ' run the Uninstall Command...
    Dim Count As Integer
    Dim Count2 As Integer
    Dim returnName As Collection
    Dim returnSubs As Collection
    Dim displayName As String
    Dim uninstallString As String

    If lstPrograms.SelCount > 0 Then
        For Count = 0 To lstPrograms.ListCount - 1
            If lstPrograms.Selected(Count) = True Then
                Call EnumRegKeys(returnName, returnSubs, "HKEY_LOCAL_MACHINE", "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
                If returnName.Count > 0 Then
                    For Count2 = 1 To returnName.Count
                        displayName = GetSetting("", returnName(Count2), "DisplayName", "", HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
                        uninstallString = GetSetting("", returnName(Count2), "UninstallString", "", HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall")
                        If displayName = lstPrograms.List(Count) Then
                            txtUninstallString.Text = uninstallString
                        End If
                    Next Count2
                End If
            End If
        Next Count
    Else
        Exit Sub
    End If

End Sub

Private Sub lstPrograms_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    ' show description...
    lblDescription.Caption = "You can uninstall programs or delete their entries from registry on this page..."
End Sub

Private Sub picBackground_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    ' show description...
    lblDescription.Caption = "WinTweaker 1.0 (Windows Personal Security). Copyright © 2001 - 2003 VJ™ Software"
End Sub

Private Sub txtNetworkPwdLength_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    ' show description...
    lblDescription.Caption = "Please enter two-digit hexadecimal value."
End Sub

Private Sub txtNetworkTTLTcp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    ' show description...
    lblDescription.Caption = "Specifies the default Time To Live (TTL) value in the header of outgoing IP packets. The TTL determines how long an IP packet that has not reached its destination can remain on the network before it is discarded."
End Sub

Private Sub txtStartMenuDelay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    ' show description...
    lblDescription.Caption = "Sets time in milliseconds how fast a submenu is expanded in the Start Menu."
End Sub

Private Sub txtSystemOrganization_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    ' show description...
    lblDescription.Caption = "Change the name of registered organization which is displayed on System Properties."
End Sub

Private Sub txtSystemOwner_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    ' show description...
    lblDescription.Caption = "Change the name of registered owner which is displayed on System Properties."
End Sub

Private Sub txtSystemSourcePath_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    ' show description...
    lblDescription.Caption = "Change the default path, from which the system was installed."
End Sub

Private Sub txtUninstallString_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    lblDescription.Caption = "Display the ""UninstallString"" command that registered on the registry."
End Sub
