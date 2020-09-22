VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmMain"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7425
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
   MaxButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   7425
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cmDlg 
      Left            =   1680
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Browse..."
   End
   Begin VB.FileListBox filTemp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      ReadOnly        =   0   'False
      TabIndex        =   55
      Top             =   5040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "cmdHelp"
      Height          =   345
      Left            =   120
      TabIndex        =   23
      Top             =   4870
      Width           =   1125
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "cmdCancel"
      Height          =   345
      Left            =   6120
      TabIndex        =   3
      Top             =   4870
      Width           =   1125
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "cmdNext"
      Default         =   -1  'True
      Height          =   345
      Left            =   4800
      TabIndex        =   2
      Top             =   4870
      Width           =   1125
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "cmdBack"
      Height          =   345
      Left            =   3600
      TabIndex        =   1
      Top             =   4870
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.PictureBox picBottom 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      ScaleHeight     =   255
      ScaleWidth      =   7215
      TabIndex        =   4
      Top             =   4570
      Visible         =   0   'False
      Width           =   7215
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CoolTemplate - Templates are Cool"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000011&
         Height          =   195
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   2655
      End
      Begin VB.Line lneBottom 
         BorderColor     =   &H80000011&
         Index           =   0
         X1              =   2640
         X2              =   7155
         Y1              =   105
         Y2              =   105
      End
      Begin VB.Line lneBottom 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   2640
         X2              =   7155
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Label Label2 
         Caption         =   "CoolTemplate - Templates are Cool"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   15
         TabIndex        =   6
         Top             =   15
         Width           =   2640
      End
   End
   Begin VB.PictureBox picTop 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   0
      ScaleHeight     =   885
      ScaleWidth      =   7455
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   7455
      Begin VB.Line lneTop 
         BorderColor     =   &H0099A8AC&
         X1              =   0
         X2              =   7440
         Y1              =   850
         Y2              =   850
      End
      Begin VB.Label lblStep 
         BackColor       =   &H00FFFFFF&
         Caption         =   "lblStep"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   120
         Width           =   4515
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "lblInfo"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   360
         Width           =   6015
      End
      Begin VB.Image imgTopIcon 
         Height          =   825
         Left            =   6480
         Picture         =   "frmMain.frx":57E2
         Top             =   0
         Width           =   825
      End
   End
   Begin VB.PictureBox picWizard 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4705
      Index           =   2
      Left            =   0
      ScaleHeight     =   4710
      ScaleWidth      =   7425
      TabIndex        =   29
      Top             =   0
      Width           =   7430
      Begin VB.TextBox txtRHeight 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5220
         TabIndex        =   52
         Text            =   "600"
         Top             =   4200
         Width           =   735
      End
      Begin VB.TextBox txtRWidth 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         TabIndex        =   51
         Text            =   "800"
         Top             =   4200
         Width           =   735
      End
      Begin VB.CheckBox chkStretch 
         Caption         =   "chkStretch"
         Height          =   255
         Left            =   240
         TabIndex        =   48
         Top             =   3960
         Value           =   1  'Checked
         Width           =   3135
      End
      Begin VB.TextBox txtURL 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   39
         Top             =   3240
         Width           =   5775
      End
      Begin VB.TextBox txtDescription 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1095
         TabIndex        =   36
         Top             =   2520
         Width           =   5775
      End
      Begin VB.TextBox txtAuthor 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   33
         Top             =   1800
         Width           =   5775
      End
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   30
         Top             =   1080
         Width           =   5775
      End
      Begin VB.Label lblRHeight 
         Caption         =   "lblRHeight"
         Enabled         =   0   'False
         Height          =   195
         Left            =   3720
         TabIndex        =   50
         Top             =   4245
         Width           =   1560
      End
      Begin VB.Label lblRWidth 
         Caption         =   "lblRWidth"
         Enabled         =   0   'False
         Height          =   195
         Left            =   960
         TabIndex        =   49
         Top             =   4245
         Width           =   1515
      End
      Begin VB.Label lblURL 
         AutoSize        =   -1  'True
         Caption         =   "URL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1800
         MouseIcon       =   "frmMain.frx":59E4
         MousePointer    =   99  'Custom
         TabIndex        =   42
         Top             =   3600
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Label lblCheck 
         Caption         =   "lblCheck"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1080
         TabIndex        =   41
         Top             =   3600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblTURL 
         Caption         =   "lblTURL"
         Height          =   195
         Left            =   600
         TabIndex        =   40
         Top             =   3270
         Width           =   465
      End
      Begin VB.Label lblTDescription 
         Caption         =   "lblTDescription"
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   2550
         Width           =   975
      End
      Begin VB.Label lblDescriptionEx 
         Caption         =   "lblDescriptionEx"
         Height          =   195
         Left            =   3720
         TabIndex        =   37
         Top             =   2880
         Width           =   3195
      End
      Begin VB.Label lblAuthorEx 
         Caption         =   "lblAuthorEx"
         Height          =   195
         Left            =   3720
         TabIndex        =   35
         Top             =   2160
         Width           =   3195
      End
      Begin VB.Label lblTAuthor 
         Caption         =   "lblTAuthor"
         Height          =   195
         Left            =   360
         TabIndex        =   34
         Top             =   1830
         Width           =   675
      End
      Begin VB.Label lblTNameEx 
         Caption         =   "lblTNameEx"
         Height          =   195
         Left            =   3720
         TabIndex        =   32
         Top             =   1440
         Width           =   3135
      End
      Begin VB.Label lblTName 
         Caption         =   "lblTName"
         Height          =   195
         Left            =   480
         TabIndex        =   31
         Top             =   1110
         Width           =   585
      End
   End
   Begin VB.PictureBox picWizard 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4705
      Index           =   3
      Left            =   0
      ScaleHeight     =   4710
      ScaleWidth      =   7425
      TabIndex        =   43
      Top             =   0
      Width           =   7430
      Begin ComctlLib.Toolbar tbrMenuItem 
         Height          =   390
         Left            =   6840
         TabIndex        =   47
         Top             =   920
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   688
         ButtonWidth     =   635
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         ImageList       =   "imlToolbar"
         _Version        =   327682
         BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
            NumButtons      =   1
            BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "MenuItem"
               Object.Tag             =   ""
               ImageIndex      =   1
            EndProperty
         EndProperty
      End
      Begin RichTextLib.RichTextBox txtCode 
         Height          =   3135
         Left            =   240
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   1320
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   5530
         _Version        =   393217
         HideSelection   =   0   'False
         ScrollBars      =   3
         DisableNoScroll =   -1  'True
         Appearance      =   0
         RightMargin     =   1.00000e5
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmMain.frx":5CEE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.ComboBox cmbComments 
         Height          =   315
         ItemData        =   "frmMain.frx":5D6A
         Left            =   840
         List            =   "frmMain.frx":5D6C
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label lblComment 
         AutoSize        =   -1  'True
         Caption         =   "lblComment"
         Height          =   195
         Left            =   3000
         TabIndex        =   46
         Top             =   1035
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label lblAdd 
         Caption         =   "lblAdd"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   44
         Top             =   1020
         Width           =   615
      End
   End
   Begin VB.PictureBox picWizard 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4705
      Index           =   1
      Left            =   0
      ScaleHeight     =   4710
      ScaleWidth      =   7425
      TabIndex        =   10
      Top             =   0
      Width           =   7430
      Begin VB.Frame Frame1 
         Height          =   855
         Left            =   600
         TabIndex        =   24
         Top             =   3120
         Width           =   6255
         Begin VB.TextBox txtTemplate 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1200
            TabIndex        =   27
            Top             =   360
            Width           =   3615
         End
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4920
            ScaleHeight     =   375
            ScaleWidth      =   1215
            TabIndex        =   25
            Top             =   320
            Width           =   1215
            Begin VB.CommandButton cmdBrowseTemp 
               Caption         =   "cmdBrowseTemp"
               Height          =   345
               Left            =   0
               TabIndex        =   26
               Top             =   0
               Width           =   1125
            End
         End
         Begin VB.Label lblTemplateFile 
            Caption         =   "lblTemplateFile"
            Height          =   195
            Left            =   150
            TabIndex        =   28
            Top             =   390
            Width           =   1050
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1800
         Left            =   -240
         Picture         =   "frmMain.frx":5D6E
         ScaleHeight     =   1800
         ScaleWidth      =   1920
         TabIndex        =   53
         Top             =   1080
         Width           =   1920
      End
      Begin VB.Label lblTemplateInfo 
         Caption         =   "lblTemplateInfo"
         Height          =   615
         Left            =   1920
         TabIndex        =   54
         Top             =   1560
         Width           =   5295
      End
   End
   Begin VB.PictureBox picWizard 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4705
      Index           =   0
      Left            =   0
      ScaleHeight     =   4710
      ScaleWidth      =   7425
      TabIndex        =   17
      Top             =   0
      Width           =   7430
      Begin VB.PictureBox picStart 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4685
         Left            =   0
         Picture         =   "frmMain.frx":7453
         ScaleHeight     =   4680
         ScaleWidth      =   2415
         TabIndex        =   18
         Top             =   0
         Width           =   2415
      End
      Begin VB.Line Line3 
         BorderColor     =   &H0099A8AC&
         X1              =   -120
         X2              =   7440
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Label lblWelcome 
         BackColor       =   &H00FFFFFF&
         Caption         =   "lblWelcome(0)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   2640
         TabIndex        =   22
         Top             =   240
         Width           =   4530
      End
      Begin VB.Label lblWelcome 
         BackColor       =   &H00FFFFFF&
         Caption         =   "lblWelcome(1)"
         Height          =   795
         Index           =   1
         Left            =   2640
         TabIndex        =   21
         Top             =   960
         Width           =   4665
      End
      Begin VB.Label lblWelcome 
         BackColor       =   &H00FFFFFF&
         Caption         =   "lblWelcome(2)"
         Height          =   855
         Index           =   2
         Left            =   2640
         TabIndex        =   20
         Top             =   1800
         Width           =   4575
      End
      Begin VB.Label lblWelcome 
         BackColor       =   &H00FFFFFF&
         Caption         =   "lblWelcome(3)"
         Height          =   195
         Index           =   3
         Left            =   2640
         TabIndex        =   19
         Top             =   2760
         Width           =   4485
      End
   End
   Begin VB.PictureBox picWizard 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4705
      Index           =   4
      Left            =   0
      ScaleHeight     =   4710
      ScaleWidth      =   7425
      TabIndex        =   11
      Top             =   0
      Width           =   7430
      Begin VB.CheckBox chkRun 
         BackColor       =   &H00FFFFFF&
         Caption         =   "chkRun"
         Height          =   195
         Left            =   2640
         TabIndex        =   13
         Top             =   2400
         Value           =   1  'Checked
         Width           =   4455
      End
      Begin VB.PictureBox picFinish 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4685
         Left            =   0
         ScaleHeight     =   4680
         ScaleWidth      =   2415
         TabIndex        =   12
         Top             =   0
         Width           =   2415
      End
      Begin VB.Line Line5 
         BorderColor     =   &H0099A8AC&
         X1              =   -120
         X2              =   7440
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Label lblFinish 
         BackColor       =   &H00FFFFFF&
         Caption         =   "lblFinish(0)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   2640
         TabIndex        =   16
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label lblFinish 
         BackColor       =   &H00FFFFFF&
         Caption         =   "lblFinish(1)"
         Height          =   915
         Index           =   1
         Left            =   2640
         TabIndex        =   15
         Top             =   840
         Width           =   4665
      End
      Begin VB.Label lblFinish 
         BackColor       =   &H00FFFFFF&
         Caption         =   "lblFinish(2)"
         Height          =   195
         Index           =   2
         Left            =   2640
         TabIndex        =   14
         Top             =   1800
         Width           =   4365
      End
   End
   Begin ComctlLib.ImageList imlToolbar 
      Left            =   2280
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":8583
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "HelpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuOptions 
         Caption         =   "mnuOptions"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "mnuAbout"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function PathIsRelative Lib "shlwapi.dll" Alias "PathIsRelativeA" (ByVal pszPath As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
                                                                                                      ByVal lpKeyName As String, _
                                                                                                      ByVal lpString As Any, _
                                                                                                      ByVal lpFileName As String) As Long
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, _
                                                                    ByVal lpNewFileName As String, _
                                                                    ByVal bFailIfExists As Long) As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, _
                                                    ByVal hWndInsertAfter As Long, _
                                                    ByVal X As Long, _
                                                    ByVal Y As Long, _
                                                    ByVal cx As Long, _
                                                    ByVal cy As Long, _
                                                    ByVal wFlags As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const GWL_STYLE As Long = (-16)
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const hWnd_NOTOPMOST As Integer = -2
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_FRAMECHANGED As Long = &H20

Private sMenuItem As String
Private sMainCode As String
Private iStep As Integer
Private Sub Form_Load()
    cTrans.SetTranslation Me
    Me.Height = 5865
    picWizard(0).ZOrder
End Sub
Private Sub Form_Resize()
    If Not Me.WindowState = 1 Then
        If iStep = 3 Or iStep = 4 Then
            With Me
                cmdHelp.Top = .Height - 990
                cmdCancel.Top = .Height - 990
                cmdCancel.Left = .Width - 1400
                cmdNext.Top = .Height - 990
                cmdNext.Left = .Width - 2720
                cmdBack.Top = .Height - 990
                cmdBack.Left = .Width - 3920
                picTop.Width = .Width
                lneTop.X2 = .Width
                imgTopIcon.Left = .Width - 1040
                picBottom.Top = .Height - 1290
                picBottom.Width = .Width
                lneBottom(0).X2 = .Width - 360
                lneBottom(1).X2 = .Width - 360
                picWizard(3).Width = .Width - 150
                picWizard(3).Height = .Height - 1200
                txtCode.Height = .Height - 2730
                txtCode.Width = .Width - 540
                tbrMenuItem.Left = .Width - 670
            End With
        End If
    End If
End Sub
Private Sub cmdNext_Click()
Dim sPath As String
    Select Case iStep
    Case 0
        cmdBack.Visible = True
        picTop.Visible = True
        picBottom.Visible = True
        lblStep.Caption = cTrans.GetString(1)
        lblInfo.Caption = cTrans.GetString(2)
        picWizard(1).ZOrder
        picBottom.ZOrder
        picTop.ZOrder
        txtTemplate.SetFocus
        iStep = iStep + 1
    Case 1
        If Not FileExists(txtTemplate.Text) Or LenB(txtTemplate.Text) = 0 Or CBool(PathIsRelative(txtTemplate.Text)) Or Right$(txtTemplate.Text, 1) = "\" Then
            MessageBox cTrans.GetString(502), cTrans.GetString(501)
            txtTemplate.SetFocus
        Else
            lblStep.Caption = cTrans.GetString(3)
            lblInfo.Caption = cTrans.GetString(4)
            picWizard(2).ZOrder
            picBottom.ZOrder
            picTop.ZOrder
            txtName.SetFocus
            iStep = iStep + 1
        End If
    Case 2
        If LenB(txtName.Text) = 0 Then
            MessageBox cTrans.GetString(504), cTrans.GetString(503)
            txtName.SetFocus
        ElseIf Dir(GetSetting("") & "\Templates\" & Replace$(LCase$(txtName.Text), " ", "_"), vbDirectory) <> "" Then
            MessageBox cTrans.GetString(506), cTrans.GetString(505)
            txtName.SetFocus
        Else
            EnableMaxButton
            lblStep.Caption = cTrans.GetString(5)
            lblInfo.Caption = cTrans.GetString(6)
            With cmbComments
                .Clear
                .AddItem "<!--Title-->"
                .AddItem "<!--Author-->"
                .AddItem "<!--PageName-->"
                .AddItem "<!--PageContent-->"
                .AddItem "<!--MenuItems-->"
                .AddItem "<!--Year-->"
                .AddItem "<!--Date-->"
            End With
            txtCode.LoadFile txtTemplate.Text, rtfText
            picWizard(3).ZOrder
            picBottom.ZOrder
            picTop.ZOrder
            txtCode.SetFocus
            iStep = iStep + 1
        End If
    Case 3
        sMainCode = txtCode.Text
        lblStep.Caption = cTrans.GetString(7)
        lblInfo.Caption = cTrans.GetString(8)
        lblComment.Caption = vbNullString
        cmdNext.Caption = cTrans.GetString(103)
        cmdNext.ToolTipText = cTrans.GetString(104)
        With cmbComments
            .Clear
            .AddItem "<!--ItemName-->"
            .AddItem "<!--ItemPath-->"
        End With
        tbrMenuItem.Visible = False
        txtCode.Text = sMenuItem
        txtCode.SetFocus
        iStep = iStep + 1
    Case 4
        If LenB(txtCode.Text) = 0 Then
            MessageBox cTrans.GetString(508), cTrans.GetString(507)
            txtCode.SetFocus
        Else
            frmWait.Show
            sPath = Replace$(LCase$(txtName.Text), " ", "_")
            sPath = GetSetting("") & "\Templates\" & sPath & "\files\"
            CreateFolder sPath
            sMenuItem = txtCode.Text
            Me.Hide
            DoEvents
            Build
            DisableMaxButton
            picFinish.Picture = picStart.Picture
            cmdBack.Enabled = False
            cmdNext.Caption = cTrans.GetString(105)
            cmdNext.ToolTipText = cTrans.GetString(106)
            picTop.Visible = False
            picBottom.Visible = False
            If FileExists(GetSetting("") & "\CoolWeb.exe") = False Then
                chkRun.Value = 0
                chkRun.Enabled = False
            End If
            picWizard(4).ZOrder
            Unload frmWait
            Me.Show
            iStep = iStep + 1
        End If
    Case 5
        If chkRun.Value = 1 Then
            ShellExecute 0, vbNullString, GetSetting("") & "\CoolWeb.exe", vbNullString, "", 10
        End If
        Unload Me
    End Select
End Sub
Private Sub cmdBack_Click()
    Select Case iStep
    Case 1
        cmdBack.Visible = False
        picTop.Visible = False
        picBottom.Visible = False
        picWizard(0).ZOrder
        iStep = iStep - 1
    Case 2
        lblStep.Caption = cTrans.GetString(1)
        lblInfo.Caption = cTrans.GetString(2)
        picWizard(1).ZOrder
        picBottom.ZOrder
        picTop.ZOrder
        txtTemplate.SetFocus
        iStep = iStep - 1
    Case 3
        DisableMaxButton
        lblStep.Caption = cTrans.GetString(3)
        lblInfo.Caption = cTrans.GetString(4)
        picWizard(2).ZOrder
        picBottom.ZOrder
        picTop.ZOrder
        txtName.SetFocus
        iStep = iStep - 1
    Case 4
        lblStep.Caption = cTrans.GetString(5)
        lblInfo.Caption = cTrans.GetString(6)
        lblComment.Caption = ""
        cmdNext.Caption = cTrans.GetString(101)
        cmdNext.ToolTipText = cTrans.GetString(102)
        With cmbComments
            .Clear
            .AddItem "<!--Title-->"
            .AddItem "<!--Author-->"
            .AddItem "<!--PageName-->"
            .AddItem "<!--PageContent-->"
            .AddItem "<!--MenuItems-->"
            .AddItem "<!--Year-->"
            .AddItem "<!--Date-->"
        End With
        tbrMenuItem.Visible = True
        txtCode.Text = sMainCode
        txtCode.SetFocus
        iStep = iStep - 1
    End Select
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdHelp_Click()
    PopupMenu mnuHelp, , cmdHelp.Left, cmdHelp.Top + 350
End Sub
Private Sub mnuOptions_Click()
    frmOptions.Show vbModal, Me
End Sub
Private Sub mnuAbout_Click()
    frmAbout.Show vbModal, Me
End Sub
Private Sub cmdBrowseTemp_Click()
    On Error GoTo ErrHandler
    With cmDlg
        .Filter = "HTML Files(*.htm,*.html)|*.htm;*.html"
        .ShowOpen
        txtTemplate.Text = .FileName
    End With
    txtTemplate.SetFocus
ErrHandler:
    If Err.Number = 32755 Then
        txtTemplate.SetFocus
    End If
End Sub
Private Sub cmbComments_Click()
    lblComment.Visible = True
    txtCode.SelText = cmbComments.List(cmbComments.ListIndex)
    With cTrans
        Select Case cmbComments.List(cmbComments.ListIndex)
        Case "<!--Title-->"
            lblComment.Caption = .GetString(107)
        Case "<!--Author-->"
            lblComment.Caption = .GetString(108)
        Case "<!--PageName-->"
            lblComment.Caption = .GetString(109)
        Case "<!--PageContent-->"
            lblComment.Caption = .GetString(110)
        Case "<!--MenuItems-->"
            lblComment.Caption = .GetString(111)
        Case "<!--Year-->"
            lblComment.Caption = .GetString(112)
        Case "<!--Date-->"
            lblComment.Caption = .GetString(113)
        Case "<!--ItemName-->"
            lblComment.Caption = .GetString(114)
        Case "<!--ItemPath-->"
            lblComment.Caption = .GetString(115)
        End Select
    End With
    txtCode.SetFocus
End Sub
Private Sub chkStretch_Click()
    txtRWidth.Enabled = Not (chkStretch.Value = 1)
    txtRHeight.Enabled = Not (chkStretch.Value = 1)
    lblRWidth.Enabled = Not (chkStretch.Value = 1)
    lblRHeight.Enabled = Not (chkStretch.Value = 1)
    If chkStretch.Value = 0 Then
        txtRWidth.SelStart = 0
        txtRWidth.SelLength = Len(txtRWidth.Text)
        txtRWidth.SetFocus
    End If
End Sub
Private Sub txtURL_Change()
    If LenB(txtURL.Text) = 0 Then
        lblURL.Visible = False
        lblCheck.Visible = False
    Else
        lblURL.Visible = True
        lblCheck.Visible = True
        lblURL.Caption = txtURL.Text
    End If
End Sub
Private Sub lblURL_Click()
    ShellExecute 0, "open", lblURL.Caption, "", "", 10
End Sub
Private Sub tbrMenuItem_ButtonClick(ByVal Button As ComctlLib.Button)
    sMenuItem = txtCode.SelText
End Sub
Private Sub txtAuthor_GotFocus()
    txtAuthor.SelStart = 0
    txtAuthor.SelLength = Len(txtAuthor.Text)
End Sub
Private Sub txtDescription_GotFocus()
    txtDescription.SelStart = 0
    txtDescription.SelLength = Len(txtDescription.Text)
End Sub
Private Sub txtName_GotFocus()
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName.Text)
End Sub
Private Sub txtTemplate_GotFocus()
    txtTemplate.SelStart = 0
    txtTemplate.SelLength = Len(txtTemplate.Text)
End Sub
Private Sub txtRHeight_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr$(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0
End Sub
Private Sub txtRWidth_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr$(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0
End Sub
Private Sub txtCode_KeyDown(KeyCode As Integer, _
                            Shift As Integer)
    If KeyCode = vbKeyTab Then
        txtCode.SelText = vbTab
        KeyCode = 0
    End If
End Sub
Private Sub EnableMaxButton()
    With Me
        SetWindowLong .hWnd, GWL_STYLE, GetWindowLong(.hWnd, GWL_STYLE) Or WS_MAXIMIZEBOX
        SetWindowPos .hWnd, hWnd_NOTOPMOST, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_FRAMECHANGED
    End With
End Sub
Private Sub DisableMaxButton()
    With Me
        SetWindowLong .hWnd, GWL_STYLE, GetWindowLong(.hWnd, GWL_STYLE) And Not WS_MAXIMIZEBOX
        SetWindowPos .hWnd, hWnd_NOTOPMOST, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_FRAMECHANGED
        .WindowState = 0
    End With
End Sub
'Done by Xpert
'This does the magic of creating a folder even if some intermediates
'subfolders of the specified path neither exist
Private Function CreateFolder(ByVal sPath As String) As Boolean
Dim carps() As String
Dim i       As Long
    On Error GoTo errHandle
    If Dir$(sPath, vbDirectory) = "" Then
        carps = Split(sPath, "\")
        For i = 1 To UBound(carps)
            carps(i) = carps(i - 1) & "\" & carps(i)
        Next i
        For i = 0 To UBound(carps)
            If LenB(Dir$(carps(i), vbDirectory)) = 0 Then
                MkDir (carps(i))
            End If
        Next i
        CreateFolder = True
    End If
errHandle:
End Function
Private Function WriteINI(Section As String, _
                          KeyName As String, _
                          NewString As String, _
                          FileName As String) As Integer
    WritePrivateProfileString Section, KeyName, NewString, FileName
End Function
Private Sub DoFiles(ByVal sDir As String, _
                    ByVal sTo As String, _
                    Optional JustDelete As Boolean = False)
Dim sCurPath As String
Dim sName    As String
Dim sLastDir As String
    sCurPath = sDir
    sName = vbNullString
    sName = Dir(sCurPath, vbDirectory)
    Do While sName <> vbNullString
        If (GetAttr(sCurPath & sName) And vbDirectory) = vbDirectory Then
            If sName <> "." Then
                If sName <> ".." Then
                    sLastDir = sName
                    DoFiles sCurPath & sName & "\", ""
                    sName = Dir(sCurPath, vbDirectory)
                    Do While sName <> sLastDir
                        sName = Dir
                    Loop
                End If
            End If
        Else
            If JustDelete Then
                DeleteFile sCurPath & sName
            Else
                CopyFile sCurPath & sName, sTo & sName, 0
            End If
        End If
        sName = Dir
    Loop
End Sub
Private Function ChangeHTML(sString As String, _
                            sFind As String) As String
Dim X    As Long
Dim tmp1 As String
Dim tmp2 As String
Dim tmp3 As String
    ChangeHTML = sString
    Do While InStr(X + 1, sString, sFind)
        X = InStr(X + 1, sString, sFind)
        tmp1 = InStr(X, sString, sFind) + Len(sFind) - 1
        tmp2 = InStr(tmp1 + 1, sString, """") - 1
        tmp3 = Mid$(sString, tmp1 + 1, tmp2 - tmp1)
        tmp1 = tmp3
        If Left$(tmp3, 1) <> "#" Then
            If InStr(1, tmp3, "/") = 0 And InStr(1, tmp3, "\") = 0 Then
                tmp1 = "files/" & tmp3
            ElseIf InStr(1, tmp3, "\") = 0 And Left$(tmp3, 7) <> "http://" Then
                tmp1 = "files/" & Right$(tmp3, Len(tmp3) - InStrRev(tmp3, "/"))
            End If
        End If
        ChangeHTML = Replace$(ChangeHTML, sFind & tmp3 & """", sFind & tmp1 & """")
    Loop
End Function
Private Function ChangeCSS(sString As String) As String
Dim X    As Long
Dim tmp1 As String
Dim tmp2 As String
Dim tmp3 As String
    ChangeCSS = sString
    Do While InStr(X + 1, sString, "url(")
        X = InStr(X + 1, sString, "url(")
        tmp1 = InStr(X, sString, "url(") + 4
        tmp2 = InStr(tmp1 + 1, sString, ")") - 1
        tmp3 = Mid$(sString, tmp1, tmp2 - tmp1 + 1)
        tmp1 = tmp3
        If InStr(1, tmp3, "\") = 0 Then
            If InStr(1, tmp3, "/") <> 0 Then
                If Left$(tmp3, 1) = "'" Then
                    tmp1 = "'" & Right$(tmp3, Len(tmp3) - InStrRev(tmp3, "/"))
                Else
                    tmp1 = Right$(tmp3, Len(tmp3) - InStrRev(tmp3, "/"))
                End If
            End If
        End If
        ChangeCSS = Replace$(ChangeCSS, "url(" & tmp3 & ")", "url(" & tmp1 & ")")
    Loop
End Function
Private Sub MakeShots(TemplateDir As String)
Dim sTemp1 As String
Dim sTemp2 As String
Dim sMain  As String
    Open TemplateDir & "\main.txt" For Input As #1
    sMain = Input$(LOF(1), 1)
    Close #1
    Open TemplateDir & "\menuitem.txt" For Input As #1
    sTemp2 = Input$(LOF(1), 1)
    Close #1
    sTemp1 = sTemp2
    sTemp2 = Replace$(sTemp2, "<!--ItemName-->", "Home") & sTemp1
    sTemp2 = Replace$(sTemp2, "<!--ItemName-->", "Downloads") & sTemp1
    sTemp2 = Replace$(sTemp2, "<!--ItemName-->", "Games") & sTemp1
    sTemp2 = Replace$(sTemp2, "<!--ItemName-->", "Contact") & sTemp1
    sTemp2 = Replace$(sTemp2, "<!--ItemName-->", "About") & sTemp1
    sTemp2 = Replace$(sTemp2, "<!--ItemName-->", "Links")
    sMain = Replace$(sMain, "<!--Title-->", txtName.Text)
    sMain = Replace$(sMain, "<!--Author-->", txtAuthor.Text)
    sMain = Replace$(sMain, "<!--MenuItems-->", sTemp2)
    sMain = Replace$(sMain, "<!--PageName-->", "Home")
    sMain = Replace$(sMain, "<!--PageContent-->", "<p>This is to test the sample template, there is nothing special in this. Anything can be here, since it is going to be in the screenshot. There are two screenshots, one of 80x80 and the other of 190x190. This content is fake. This content is waste. This content is scrap.</p><p>Now a new paragraph! I have no idea what I am writing, so it is useless to read it. If you have enough time to read this junk, then you gotta be real boring. How about a list. Here is a list coming:</p><ul><li>Useless Item 1</li><li>Junk Item 2</li><li>Meaningless Item 3</li><li>Unpredictable Item 4</li></ul><p>Are you still reading it? Stop! Stop! Stop!. Instead of reading this, read knowledgeable books and be the second Newton.</p>")
    sMain = Replace$(sMain, "<!--Date-->", Date)
    DoFiles TemplateDir & "\files\", App.Path & "\Sample\Files\"
    Open App.Path & "\Sample\temp.htm" For Output As #1
    Print #1, sMain
    Close #1
    sTemp1 = GetSetting("WebShot")
    If chkStretch.Value = 1 Then
        ShellExecute Me.hWnd, "open", sTemp1, "/url """ & App.Path & "\Sample\temp.htm"" /out """ & TemplateDir & "\screen_full.jpg"" /width 190 /height 190", "", 0
    Else
        ShellExecute Me.hWnd, "open", sTemp1, "/url """ & App.Path & "\Sample\temp.htm"" /out """ & TemplateDir & "\screen_full.jpg"" /width 190 /height 190 /bwidth " & txtRWidth.Text & " /bheight " & txtRHeight.Text, "", 0
    End If
    Sleep 2000
    If chkStretch.Value = 1 Then
        ShellExecute Me.hWnd, "open", sTemp1, "/url """ & App.Path & "\Sample\temp.htm"" /out """ & TemplateDir & "\screen_thumb.jpg"" /width 80 /height 80", "", 0
    Else
        ShellExecute Me.hWnd, "open", sTemp1, "/url """ & App.Path & "\Sample\temp.htm"" /out """ & TemplateDir & "\screen_thumb.jpg"" /width 80 /height 80 /bwidth " & txtRWidth.Text & " /bheight " & txtRHeight.Text, "", 0
    End If
    Sleep 3000
    DoFiles App.Path & "\Sample\", "", True
    DoFiles App.Path & "\Sample\Files\", "", True
End Sub
Private Sub Build()
Dim sTemp As String
Dim i     As Integer
    sTemp = Replace$(LCase$(txtName.Text), " ", "_")
    sTemp = GetSetting("") & "\Templates\" & sTemp
    sMainCode = ChangeHTML(sMainCode, "src=""")
    sMainCode = ChangeHTML(sMainCode, "href=""")
    Open sTemp & "\main.txt" For Output As #1
    Print #1, sMainCode
    Close #1
    Open sTemp & "\menuitem.txt" For Output As #1
    Print #1, sMenuItem
    Close #1
    DoFiles Left$(txtTemplate.Text, InStrRev(txtTemplate.Text, "\")), sTemp & "\files\"
    DeleteFile sTemp & "\files\" & Right$(txtTemplate.Text, Len(txtTemplate.Text) - InStrRev(txtTemplate.Text, "\"))
    With filTemp
        .Pattern = "*.css"
        .Path = sTemp & "\files\"
        For i = 0 To .ListCount - 1
            sMainCode = vbNullString
            Open sTemp & "\files\" & .List(i) For Input As #1
            sMainCode = Input$(LOF(1), 1)
            Close #1
            sMainCode = ChangeCSS(sMainCode)
            Open sTemp & "\files\" & .List(i) For Output As #1
            Print #1, sMainCode
            Close #1
        Next i
    End With
    WriteINI "Data", "Name", txtName.Text, sTemp & "\data.ini"
    WriteINI "Data", "Author", txtAuthor.Text, sTemp & "\data.ini"
    WriteINI "Data", "Description", txtDescription.Text, sTemp & "\data.ini"
    WriteINI "Data", "URL", txtURL.Text, sTemp & "\data.ini"
    If GetSetting("WebShot") <> "Not Installed" Then
        MakeShots sTemp
    End If
End Sub
