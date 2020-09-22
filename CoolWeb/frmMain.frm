VERSION 5.00
Object = "{683364A1-B37D-11D1-ADC5-006008A5848C}#1.0#0"; "dhtmled.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmMain"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   540
   ClientWidth     =   8220
   Icon            =   "frmMain.frx":0000
   MaxButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   8220
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdHelp 
      Caption         =   "cmdHelp"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   51
      Top             =   5280
      Width           =   1125
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "cmdBack"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4440
      TabIndex        =   3
      Top             =   5280
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "cmdNext"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5640
      TabIndex        =   0
      Top             =   5280
      Width           =   1125
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "cmdCancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6960
      TabIndex        =   2
      Top             =   5280
      Width           =   1125
   End
   Begin MSComDlg.CommonDialog cmDlg 
      Left            =   2040
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Browse..."
      Filter          =   "HTML Files(*.htm,*.html)|*.htm;*.html"
      Flags           =   4
      InitDir         =   "C:"
   End
   Begin VB.PictureBox picBottom 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   8295
      TabIndex        =   5
      Top             =   4920
      Visible         =   0   'False
      Width           =   8295
      Begin VB.Line lneBottom 
         BorderColor     =   &H80000011&
         Index           =   0
         X1              =   2040
         X2              =   8025
         Y1              =   105
         Y2              =   105
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "CoolWeb - Truly a Wizard"
         ForeColor       =   &H80000011&
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "CoolWeb - Truly a Wizard"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   135
         TabIndex        =   7
         Top             =   15
         Width           =   1815
      End
      Begin VB.Line lneBottom 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   2040
         X2              =   8025
         Y1              =   120
         Y2              =   120
      End
   End
   Begin VB.FileListBox filTemp 
      Height          =   285
      Left            =   2040
      TabIndex        =   68
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picTop 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   0
      ScaleHeight     =   885
      ScaleWidth      =   8295
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   8295
      Begin VB.Image imgTopIcon 
         Height          =   720
         Left            =   7200
         Picture         =   "frmMain.frx":57E2
         Top             =   75
         Width           =   720
      End
      Begin VB.Label lblStepInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "lblStepInfo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   360
         Width           =   6375
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
         TabIndex        =   14
         Top             =   120
         Width           =   4875
      End
      Begin VB.Line lneTop 
         BorderColor     =   &H0099A8AC&
         X1              =   0
         X2              =   8280
         Y1              =   855
         Y2              =   855
      End
   End
   Begin VB.PictureBox picWizard 
      BorderStyle     =   0  'None
      Height          =   4935
      Index           =   1
      Left            =   0
      ScaleHeight     =   4935
      ScaleWidth      =   8265
      TabIndex        =   52
      Top             =   0
      Width           =   8265
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   1350
         Left            =   6120
         Picture         =   "frmMain.frx":6227
         ScaleHeight     =   1350
         ScaleWidth      =   1350
         TabIndex        =   65
         Top             =   1080
         Width           =   1350
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   6720
         ScaleHeight     =   1695
         ScaleWidth      =   1215
         TabIndex        =   61
         Top             =   2620
         Width           =   1215
         Begin VB.CommandButton cmdBrowsePro 
            Caption         =   "cmdBrowsePro(0)"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   0
            TabIndex        =   63
            Top             =   120
            Width           =   1125
         End
         Begin VB.CommandButton cmdBrowsePro 
            Caption         =   "cmdBrowsePro(1)"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   1
            Left            =   0
            TabIndex        =   62
            Top             =   1300
            Width           =   1125
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   3135
         Left            =   360
         ScaleHeight     =   3135
         ScaleWidth      =   6135
         TabIndex        =   53
         Top             =   1440
         Width           =   6135
         Begin VB.OptionButton optProject 
            Caption         =   "optProject(1)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   58
            Top             =   960
            Width           =   5535
         End
         Begin VB.OptionButton optProject 
            Caption         =   "optProject(0)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   57
            Top             =   240
            Value           =   -1  'True
            Width           =   5415
         End
         Begin VB.OptionButton optProject 
            Caption         =   "optProject(2)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   56
            Top             =   2160
            Width           =   5595
         End
         Begin VB.TextBox txtProject 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   1080
            TabIndex        =   55
            Top             =   1320
            Width           =   5055
         End
         Begin VB.TextBox txtProject 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   1080
            TabIndex        =   54
            Top             =   2520
            Width           =   5055
         End
         Begin VB.Label lblSaveTo 
            AutoSize        =   -1  'True
            Caption         =   "lblSaveTo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   60
            Top             =   1350
            Width           =   690
         End
         Begin VB.Label lblLoadFrom 
            AutoSize        =   -1  'True
            Caption         =   "lblLoadFrom"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   59
            Top             =   2550
            Width           =   855
         End
      End
   End
   Begin VB.PictureBox picWizard 
      BorderStyle     =   0  'None
      Height          =   4935
      Index           =   5
      Left            =   0
      ScaleHeight     =   4935
      ScaleWidth      =   8265
      TabIndex        =   38
      Top             =   0
      Width           =   8265
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   1920
         Left            =   0
         Picture         =   "frmMain.frx":6F1F
         ScaleHeight     =   1920
         ScaleWidth      =   1920
         TabIndex        =   64
         Top             =   960
         Width           =   1920
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   480
         TabIndex        =   39
         Top             =   3360
         Width           =   7215
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   5880
            ScaleHeight     =   375
            ScaleWidth      =   1215
            TabIndex        =   41
            Top             =   320
            Width           =   1215
            Begin VB.CommandButton cmdBrowseDes 
               Caption         =   "cmdBrowseDes"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   0
               TabIndex        =   42
               Top             =   0
               Width           =   1125
            End
         End
         Begin VB.TextBox txtDestination 
            Height          =   285
            Left            =   1200
            TabIndex        =   40
            Top             =   360
            Width           =   4455
         End
         Begin VB.Label lblDestination 
            AutoSize        =   -1  'True
            Caption         =   "lblDestination"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   43
            Top             =   390
            Width           =   960
         End
      End
      Begin VB.Label lblDesInfo 
         Caption         =   "lblDesInfo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   2160
         TabIndex        =   44
         Top             =   1800
         Width           =   5820
      End
   End
   Begin VB.PictureBox picWizard 
      BorderStyle     =   0  'None
      Height          =   4935
      Index           =   3
      Left            =   0
      ScaleHeight     =   4935
      ScaleWidth      =   8265
      TabIndex        =   25
      Top             =   0
      Width           =   8265
      Begin VB.CheckBox chkUnicode 
         Caption         =   "chkUnicode"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   67
         Top             =   4200
         Width           =   5655
      End
      Begin VB.TextBox txtAuthor 
         Height          =   285
         Left            =   2040
         TabIndex        =   30
         Top             =   3120
         Width           =   5775
      End
      Begin VB.TextBox txtTitle 
         Height          =   285
         Left            =   2040
         TabIndex        =   26
         Top             =   1800
         Width           =   5775
      End
      Begin VB.Image Image1 
         Height          =   2040
         Left            =   360
         Picture         =   "frmMain.frx":8879
         Top             =   1920
         Width           =   1440
      End
      Begin VB.Label lblWAuthorEx 
         Caption         =   "lblWAuthorEx"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3960
         TabIndex        =   33
         Top             =   3720
         Width           =   3795
      End
      Begin VB.Label lblWAuthorInfo 
         Caption         =   "lblWAuthorInfo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3960
         TabIndex        =   32
         Top             =   3480
         Width           =   3780
      End
      Begin VB.Label lblWAuthor 
         Caption         =   "lblWAuthor"
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
         Left            =   2040
         TabIndex        =   31
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label lblWTitleEx 
         Caption         =   "lblWTitleEx"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3720
         TabIndex        =   29
         Top             =   2400
         Width           =   4035
      End
      Begin VB.Label lblWTitleInfo 
         Caption         =   "lblWTitleInfo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3720
         TabIndex        =   28
         Top             =   2160
         Width           =   4020
      End
      Begin VB.Label lblWTitle 
         Caption         =   "lblWTitle"
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
         Left            =   2040
         TabIndex        =   27
         Top             =   1560
         Width           =   1425
      End
   End
   Begin VB.PictureBox picWizard 
      BorderStyle     =   0  'None
      Height          =   4935
      Index           =   2
      Left            =   0
      ScaleHeight     =   4935
      ScaleWidth      =   8265
      TabIndex        =   8
      Top             =   0
      Width           =   8265
      Begin VB.PictureBox picSelTemp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2885
         Left            =   3840
         ScaleHeight     =   190
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   190
         TabIndex        =   18
         Top             =   960
         Width           =   2885
      End
      Begin VB.DirListBox dirTemplates 
         Height          =   1665
         Left            =   360
         TabIndex        =   17
         Top             =   2640
         Visible         =   0   'False
         Width           =   1815
      End
      Begin ComctlLib.ListView lstTemplates 
         Height          =   3855
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   6800
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDragMode     =   1
         _Version        =   327682
         Icons           =   "imlTemplates"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDragMode     =   1
         NumItems        =   0
      End
      Begin VB.Image imgNoPreview 
         Height          =   1200
         Index           =   1
         Left            =   6480
         Picture         =   "frmMain.frx":9A49
         Top             =   2520
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Image imgNoPreview 
         Height          =   2850
         Index           =   0
         Left            =   6360
         Picture         =   "frmMain.frx":A3BA
         Top             =   1080
         Visible         =   0   'False
         Width           =   2850
      End
      Begin ComctlLib.ImageList imlTemplates 
         Left            =   2520
         Top             =   2280
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   80
         ImageHeight     =   80
         MaskColor       =   12632256
         UseMaskColor    =   0   'False
         _Version        =   327682
      End
      Begin VB.Label lblDescription 
         AutoSize        =   -1  'True
         Caption         =   "lblDescription"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3960
         TabIndex        =   24
         Top             =   4560
         Width           =   945
      End
      Begin VB.Label lblWebsite 
         AutoSize        =   -1  'True
         Caption         =   "lblWebsite"
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
         Left            =   3960
         MouseIcon       =   "frmMain.frx":C555
         MousePointer    =   99  'Custom
         TabIndex        =   23
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label lblAuthor 
         AutoSize        =   -1  'True
         Caption         =   "lblAuthor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3960
         TabIndex        =   22
         Top             =   4080
         Width           =   645
      End
      Begin VB.Label lblTWebsite 
         AutoSize        =   -1  'True
         Caption         =   "lblTWebsite"
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
         Left            =   2640
         TabIndex        =   21
         Top             =   4320
         Width           =   990
      End
      Begin VB.Label lblTDescription 
         AutoSize        =   -1  'True
         Caption         =   "lblTDescription"
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
         Left            =   2640
         TabIndex        =   20
         Top             =   4560
         Width           =   1260
      End
      Begin VB.Label lblTAuthor 
         AutoSize        =   -1  'True
         Caption         =   "lblTAuthor"
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
         Left            =   2640
         TabIndex        =   19
         Top             =   4080
         Width           =   885
      End
   End
   Begin VB.PictureBox picWizard 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5055
      Index           =   6
      Left            =   0
      ScaleHeight     =   5055
      ScaleWidth      =   8265
      TabIndex        =   45
      Top             =   0
      Width           =   8265
      Begin VB.PictureBox picFinish 
         BorderStyle     =   0  'None
         Height          =   5030
         Left            =   0
         ScaleHeight     =   5025
         ScaleWidth      =   2655
         TabIndex        =   47
         Top             =   0
         Width           =   2655
      End
      Begin VB.CheckBox chkView 
         BackColor       =   &H00FFFFFF&
         Caption         =   "chkView"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2880
         TabIndex        =   46
         Top             =   2640
         Value           =   1  'Checked
         Width           =   4935
      End
      Begin VB.Label lblFinish 
         BackColor       =   &H00FFFFFF&
         Caption         =   "lblFinish(2)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   2880
         TabIndex        =   50
         Top             =   2040
         Width           =   4725
      End
      Begin VB.Label lblFinish 
         BackColor       =   &H00FFFFFF&
         Caption         =   "lblFinish(1)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   1
         Left            =   2880
         TabIndex        =   49
         Top             =   1080
         Width           =   4905
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
         Left            =   2880
         TabIndex        =   48
         Top             =   480
         Width           =   4815
      End
      Begin VB.Line Line5 
         BorderColor     =   &H0099A8AC&
         X1              =   -120
         X2              =   8280
         Y1              =   5030
         Y2              =   5030
      End
   End
   Begin VB.PictureBox picWizard 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5055
      Index           =   0
      Left            =   0
      ScaleHeight     =   5055
      ScaleWidth      =   8265
      TabIndex        =   4
      Top             =   0
      Width           =   8265
      Begin VB.PictureBox picStart 
         BorderStyle     =   0  'None
         Height          =   5030
         Left            =   0
         Picture         =   "frmMain.frx":C85F
         ScaleHeight     =   5025
         ScaleWidth      =   2655
         TabIndex        =   66
         Top             =   0
         Width           =   2655
      End
      Begin VB.Label lblWelcome 
         BackColor       =   &H00FFFFFF&
         Caption         =   "lblWelcome(3)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   2880
         TabIndex        =   12
         Top             =   2880
         Width           =   5205
      End
      Begin VB.Label lblWelcome 
         BackColor       =   &H00FFFFFF&
         Caption         =   "lblWelcome(2)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   2
         Left            =   2880
         TabIndex        =   11
         Top             =   1800
         Width           =   5295
      End
      Begin VB.Label lblWelcome 
         BackColor       =   &H00FFFFFF&
         Caption         =   "lblWelcome(1)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   1
         Left            =   2880
         TabIndex        =   10
         Top             =   1200
         Width           =   5265
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
         Height          =   300
         Index           =   0
         Left            =   2880
         TabIndex        =   9
         Top             =   480
         Width           =   5100
      End
      Begin VB.Line Line3 
         BorderColor     =   &H0099A8AC&
         X1              =   -120
         X2              =   8280
         Y1              =   5020
         Y2              =   5020
      End
   End
   Begin VB.PictureBox picWizard 
      BorderStyle     =   0  'None
      Height          =   4935
      Index           =   4
      Left            =   0
      ScaleHeight     =   4935
      ScaleWidth      =   8265
      TabIndex        =   34
      Top             =   0
      Width           =   8265
      Begin ComctlLib.Toolbar tbrEdit 
         Height          =   390
         Left            =   2040
         TabIndex        =   37
         Top             =   915
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   688
         ButtonWidth     =   635
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         ImageList       =   "imlToolbars"
         _Version        =   327682
         BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
            NumButtons      =   20
            BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Bold"
               Object.Tag             =   ""
               ImageIndex      =   3
               Style           =   1
            EndProperty
            BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Italic"
               Object.Tag             =   ""
               ImageIndex      =   4
               Style           =   1
            EndProperty
            BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Underline"
               Object.Tag             =   ""
               ImageIndex      =   5
               Style           =   1
            EndProperty
            BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Left"
               Object.Tag             =   ""
               ImageIndex      =   6
               Style           =   2
               Value           =   1
            EndProperty
            BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Center"
               Object.Tag             =   ""
               ImageIndex      =   7
               Style           =   2
            EndProperty
            BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Right"
               Object.Tag             =   ""
               ImageIndex      =   8
               Style           =   2
            EndProperty
            BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Indent"
               Object.Tag             =   ""
               ImageIndex      =   9
            EndProperty
            BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Outdent"
               Object.Tag             =   ""
               ImageIndex      =   10
            EndProperty
            BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Ordered"
               Object.Tag             =   ""
               ImageIndex      =   11
               Style           =   1
            EndProperty
            BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Unordered"
               Object.Tag             =   ""
               ImageIndex      =   12
               Style           =   1
            EndProperty
            BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Hyperlink"
               Object.Tag             =   ""
               ImageIndex      =   13
            EndProperty
            BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Image"
               Object.Tag             =   ""
               ImageIndex      =   14
            EndProperty
            BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Font"
               Object.Tag             =   ""
               ImageIndex      =   15
            EndProperty
            BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button19 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "DelFormat"
               Object.Tag             =   ""
               ImageIndex      =   16
            EndProperty
            BeginProperty Button20 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "CharMap"
               Object.Tag             =   ""
               ImageIndex      =   17
            EndProperty
         EndProperty
      End
      Begin ComctlLib.Toolbar tbrMenu 
         Height          =   390
         Left            =   120
         TabIndex        =   36
         Top             =   915
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   688
         ButtonWidth     =   635
         ButtonHeight    =   582
         ImageList       =   "imlToolbars"
         _Version        =   327682
         BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
            NumButtons      =   2
            BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Add"
               Object.Tag             =   ""
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Delete"
               Object.Tag             =   ""
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
      Begin DHTMLEDLibCtl.DHTMLEdit dhePage 
         Height          =   3465
         Index           =   0
         Left            =   2040
         TabIndex        =   1
         Top             =   1335
         Width           =   6000
         ActivateApplets =   0   'False
         ActivateActiveXControls=   0   'False
         ActivateDTCs    =   -1  'True
         ShowDetails     =   0   'False
         ShowBorders     =   0   'False
         Appearance      =   0
         Scrollbars      =   -1  'True
         ScrollbarAppearance=   1
         SourceCodePreservation=   -1  'True
         AbsoluteDropMode=   0   'False
         SnapToGrid      =   0   'False
         SnapToGridX     =   50
         SnapToGridY     =   50
         BrowseMode      =   0   'False
         UseDivOnCarriageReturn=   0   'False
      End
      Begin VB.ListBox lstMenu 
         Height          =   3495
         IntegralHeight  =   0   'False
         ItemData        =   "frmMain.frx":1895C
         Left            =   120
         List            =   "frmMain.frx":1895E
         TabIndex        =   35
         Top             =   1320
         Width           =   1800
      End
      Begin ComctlLib.ImageList imlToolbars 
         Left            =   1200
         Top             =   840
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16711935
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   17
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":18960
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":18CB2
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":19004
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":19356
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":196A8
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":199FA
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":19D4C
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":1A09E
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":1A3F0
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":1A742
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":1AA94
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":1ADE6
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":1B138
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":1B48A
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":1B7DC
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":1BB2E
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":1BE80
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Shape shpEditBorder 
         BorderColor     =   &H00B99D7F&
         Height          =   3495
         Left            =   2025
         Top             =   1320
         Width           =   6030
      End
   End
   Begin VB.Menu mnuPage 
      Caption         =   "PageMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuUndo 
         Caption         =   "mnuUndo"
      End
      Begin VB.Menu mnuRedo 
         Caption         =   "mnuRedo"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "mnuCut"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "mnuCopy"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "mnuPaste"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "mnuDelete"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "mnuSelectAll"
      End
   End
   Begin VB.Menu mnuItem 
      Caption         =   "ItemMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuAddPage 
         Caption         =   "mnuAddPage"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeletePage 
         Caption         =   "mnuDeletePage"
      End
      Begin VB.Menu mnuRenamePage 
         Caption         =   "mnuRenamePage"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "HelpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuOptions 
         Caption         =   "mnuOptions"
      End
      Begin VB.Menu sep4 
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

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, _
                                                                              nSize As Long) As Long
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
                                                                                                  ByVal lpKeyName As String, _
                                                                                                  ByVal lpDefault As String, _
                                                                                                  ByVal lpReturnedString As String, _
                                                                                                  ByVal nSize As Long, _
                                                                                                  ByVal lpFileName As String) As Long
Private Declare Function PathIsRelative Lib "shlwapi.dll" Alias "PathIsRelativeA" (ByVal pszPath As String) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, _
                                                                                        ByVal nSize As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                        ByVal wMsg As Long, _
                                                                        ByVal wParam As Long, _
                                                                        ByVal lParam As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal hWndInsertAfter As Long, _
                                                    ByVal x As Long, _
                                                    ByVal y As Long, _
                                                    ByVal cx As Long, _
                                                    ByVal cy As Long, _
                                                    ByVal wFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long, _
                                                                            ByVal dwNewLong As Long) As Long
Private Declare Function GetLongPathName Lib "kernel32.dll" Alias "GetLongPathNameA" (ByVal lpszShortPath As String, _
                                                                                      ByVal lpszLongPath As String, _
                                                                                      ByVal cchBuffer As Long) As Long
                                                                                      
Private Const FO_COPY As Long = &H2
Private Const FOF_NOCONFIRMATION As Long = &H10
Private Const LB_ITEMFROMPOINT As Long = &H1A9
Private Const HKEY_CURRENT_USER As Long = &H80000001
Private Const GWL_STYLE As Long = (-16)
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const hWnd_NOTOPMOST As Integer = -2
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_FRAMECHANGED  As Long = &H20

Private Type SHFILEOPSTRUCT
    lHwnd As Long
    lFunc As Long
    sFrom As String
    sTo As String
    iFlags As Integer
    lAnyOperationsAborted As Long
    lNameMappings As Long
    sProgressTitle As String
End Type

Private iStep As Integer
Private iLoaded As Integer

Private Sub Form_Load()
Dim sTemp As String
Dim lRet  As Long
    cTrans.SetTranslation Me
    LoadTemplates
    Me.Height = 6285
    picWizard(0).ZOrder
    lstMenu.AddItem cTrans.GetString(113)
    iStep = 0
    If Command = "" Then
        sTemp = String$(100, vbNullChar)
        GetUserName sTemp, 100
        txtAuthor.Text = Left$(sTemp, InStr(sTemp, vbNullChar) - 1)
        chkUnicode.Value = GetSetting("Unicode")
        If GetSetting("DestDir") = 1 Then
            txtDestination.Text = GetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Desktop") & "\My Website"
            txtProject(0).Text = txtDestination.Text & ".cwp"
        Else
            sTemp = GetSetting("CustDes")
            txtDestination.Text = sTemp
            If Len(sTemp) = 3 Then
                txtProject(0).Text = sTemp & "My Website.cwp"
            ElseIf Right$(sTemp, 1) = "\" Then
                txtProject(0).Text = Left$(sTemp, Len(sTemp) - 1) & ".cwp"
            Else
                txtProject(0).Text = sTemp & ".cwp"
            End If
        End If
        iLoaded = 0
    Else
        LoadProject Command
        sTemp = Space$(255)
        lRet = GetLongPathName(Command, sTemp, 255)
        SaveSetting "LastProject", Left$(sTemp, lRet)
        sTemp = txtDestination.Text
        If Len(sTemp) = 3 Then
            txtProject(0).Text = sTemp & "My Website.cwp"
        ElseIf Right$(sTemp, 1) = "\" Then
            txtProject(0).Text = Left$(sTemp, Len(sTemp) - 1) & ".cwp"
        Else
            txtProject(0).Text = sTemp & ".cwp"
        End If
        optProject(2).Value = True
        iStep = 1
        iLoaded = 2
    End If
    txtProject(1).Text = GetSetting("LastProject")
End Sub
Private Sub Form_Resize()
Dim i As Integer
    If Not Me.WindowState = 1 Then
        If iStep = 4 Then
            With Me
                cmdHelp.Top = .Height - 1040
                cmdCancel.Top = .Height - 1040
                cmdCancel.Left = .Width - 1380
                cmdNext.Top = .Height - 1040
                cmdNext.Left = .Width - 2700
                cmdBack.Top = .Height - 1040
                cmdBack.Left = .Width - 3900
                picTop.Width = .Width
                lneTop.X2 = .Width
                imgTopIcon.Left = .Width - 1140
                picBottom.Top = .Height - 1400
                picBottom.Width = .Width
                lneBottom(0).X2 = .Width - 320
                lneBottom(1).X2 = .Width - 320
                picWizard(4).Width = .Width - 150
                picWizard(4).Height = .Height - 1200
                lstMenu.Height = .Height - 2790
                For i = 0 To dhePage.Count - 1
                    dhePage(i).Width = .Width - 2310
                    dhePage(i).Height = .Height - 2820
                Next i
                shpEditBorder.Width = .Width - 2280
                shpEditBorder.Height = .Height - 2790
            End With
        End If
    End If
End Sub
Private Sub cmdNext_Click()
Dim oFSO  As Object
Dim sPath As String
    Select Case iStep
    Case 0
        cmdBack.Visible = True
        picTop.Visible = True
        picBottom.Visible = True
        lblStep.Caption = cTrans.GetString(1)
        lblStepInfo.Caption = cTrans.GetString(2)
        picWizard(1).ZOrder
        picBottom.ZOrder
        picTop.ZOrder
        iStep = iStep + 1
    Case 1
        If iLoaded = 0 Or iLoaded = 1 Then
            If optProject(1).Value Then
                If LenB(txtProject(0).Text) = 0 Or CBool(PathIsRelative(txtProject(0).Text)) Or Right$(txtProject(0).Text, 1) = "\" Then
                    MessageBox cTrans.GetString(502), cTrans.GetString(501)
                    txtProject(0).SetFocus
                    Exit Sub
                End If
            End If
            If optProject(2).Value = True Then
                If Not FileExists(txtProject(1).Text) Or LenB(txtProject(1).Text) = 0 Or Right$(txtProject(1).Text, 1) = "\" Then
                    MessageBox cTrans.GetString(503), cTrans.GetString(501)
                    txtProject(1).SetFocus
                    Exit Sub
                End If
                If iLoaded = 1 Then UnloadProject
                LoadProject txtProject(1).Text
                SaveSetting "LastProject", txtProject(1).Text
                iLoaded = 1
            End If
        End If
        cmdBack.Visible = True
        picTop.Visible = True
        picBottom.Visible = True
        lblStep.Caption = cTrans.GetString(3)
        lblStepInfo.Caption = cTrans.GetString(4)
        picWizard(2).ZOrder
        picBottom.ZOrder
        picTop.ZOrder
        iStep = iStep + 1
    Case 2
        lblStep.Caption = cTrans.GetString(5)
        lblStepInfo.Caption = cTrans.GetString(6)
        picWizard(3).ZOrder
        picBottom.ZOrder
        picTop.ZOrder
        txtTitle.SetFocus
        iStep = iStep + 1
    Case 3
        If LenB(txtTitle.Text) = 0 Or LenB(txtAuthor.Text) = 0 Then
            MessageBox cTrans.GetString(505), cTrans.GetString(504)
            txtTitle.SetFocus
        Else
            EnableMaxButton
            lblStep.Caption = cTrans.GetString(7)
            lblStepInfo.Caption = cTrans.GetString(8)
            picWizard(4).ZOrder
            picBottom.ZOrder
            picTop.ZOrder
            lstMenu.ListIndex = 0
            iStep = iStep + 1
        End If
    Case 4
        DisableMaxButton
        lblStep.Caption = cTrans.GetString(9)
        lblStepInfo.Caption = cTrans.GetString(10)
        cmdNext.Caption = cTrans.GetString(103)
        cmdNext.ToolTipText = cTrans.GetString(104)
        picWizard(5).ZOrder
        picBottom.ZOrder
        picTop.ZOrder
        txtDestination.SetFocus
        iStep = iStep + 1
    Case 5
        If LenB(txtDestination.Text) = 0 Then
            MessageBox cTrans.GetString(508), cTrans.GetString(507)
            txtDestination.SetFocus
        ElseIf CBool(PathIsRelative(txtDestination.Text)) Or Left$(txtDestination.Text, 1) = "\" Or Len(txtDestination.Text) = 2 Then
            MessageBox cTrans.GetString(510), cTrans.GetString(509)
            txtDestination.SetFocus
        Else
            'The following code is done by Xpert
            '--------------------------------------------
            sPath = txtDestination.Text & "\files"
            If Dir(txtDestination.Text, vbDirectory) = "" Or Len(txtDestination.Text) = 3 Then
                CreateFolder sPath
            Else
                If MessageBox(txtDestination.Text & cTrans.GetString(512) & vbNewLine & _
                 cTrans.GetString(513), cTrans.GetString(511), True) = vbOK Then
                    Set oFSO = CreateObject("Scripting.FileSystemObject")
                    If Right$(txtDestination.Text, 1) = "\" Then
                        oFSO.DeleteFolder Left$(txtDestination.Text, Len(txtDestination.Text) - 1)
                    Else
                        oFSO.DeleteFolder txtDestination.Text
                    End If
                    CreateFolder sPath
                Else
                    txtDestination.SetFocus
                    Exit Sub
                End If
            End If
            '---------------------------------------------
            Build
            If optProject(1).Value = True Then
                If FileExists(txtProject(0).Text) = False Then
                    SaveProject txtProject(0).Text
                Else
                    Kill txtProject(0).Text
                    SaveProject txtProject(0).Text
                End If
            End If
            If optProject(2).Value = True Then
                If FileExists(txtProject(1).Text) = True Then
                    Kill txtProject(1).Text
                    SaveProject txtProject(1).Text
                End If
            End If
            picFinish.Picture = picStart.Picture
            cmdBack.Enabled = False
            cmdNext.Caption = cTrans.GetString(105)
            cmdNext.ToolTipText = cTrans.GetString(106)
            picTop.Visible = False
            picBottom.Visible = False
            picWizard(6).ZOrder
            iStep = iStep + 1
        End If
    Case 6
        If chkView.Value = 1 Then
            ShellExecute 0, vbNullString, txtDestination.Text & "\index.htm", vbNullString, "", 10
        End If
        Unload Me
    End Select
End Sub
Private Sub cmdBack_Click()
    If iLoaded = 2 Then iLoaded = 1
    Select Case iStep
    Case 1
        cmdBack.Visible = False
        picTop.Visible = False
        picBottom.Visible = False
        picWizard(0).ZOrder
        iStep = iStep - 1
    Case 2
        lblStep.Caption = cTrans.GetString(1)
        lblStepInfo.Caption = cTrans.GetString(2)
        picWizard(1).ZOrder
        picBottom.ZOrder
        picTop.ZOrder
        iStep = iStep - 1
    Case 3
        lblStep.Caption = cTrans.GetString(3)
        lblStepInfo.Caption = cTrans.GetString(4)
        picWizard(2).ZOrder
        picBottom.ZOrder
        picTop.ZOrder
        iStep = iStep - 1
    Case 4
        DisableMaxButton
        lblStep.Caption = cTrans.GetString(5)
        lblStepInfo.Caption = cTrans.GetString(6)
        picWizard(3).ZOrder
        picBottom.ZOrder
        picTop.ZOrder
        txtTitle.SetFocus
        txtTitle.SelLength = 0
        txtTitle.SelStart = Len(txtTitle.Text)
        iStep = iStep - 1
    Case 5
        EnableMaxButton
        lblStep.Caption = cTrans.GetString(7)
        lblStepInfo.Caption = cTrans.GetString(8)
        cmdNext.Caption = cTrans.GetString(101)
        cmdNext.ToolTipText = cTrans.GetString(102)
        picWizard(4).ZOrder
        picBottom.ZOrder
        picTop.ZOrder
        lstMenu.ListIndex = 0
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
Private Sub mnuUndo_Click()
    dhePage(lstMenu.ListIndex).ExecCommand DECMD_UNDO
End Sub
Private Sub mnuRedo_Click()
    dhePage(lstMenu.ListIndex).ExecCommand DECMD_REDO
End Sub
Private Sub mnuCut_Click()
    dhePage(lstMenu.ListIndex).ExecCommand DECMD_CUT
End Sub
Private Sub mnuCopy_Click()
    dhePage(lstMenu.ListIndex).ExecCommand DECMD_COPY
End Sub
Private Sub mnuPaste_Click()
    dhePage(lstMenu.ListIndex).ExecCommand DECMD_PASTE
End Sub
Private Sub mnuDelete_Click()
    dhePage(lstMenu.ListIndex).ExecCommand DECMD_DELETE
End Sub
Private Sub mnuSelectAll_Click()
    dhePage(lstMenu.ListIndex).ExecCommand DECMD_SELECTALL
End Sub
Private Sub mnuAddPage_Click()
    frmAdd.Show vbModal, Me
End Sub
Private Sub mnuDeletePage_Click()
    DeleteItem
End Sub
Private Sub mnuRenamePage_Click()
Dim sNewName As String
    sNewName = FieldBox(cTrans.GetString(528), cTrans.GetString(527), lstMenu.List(lstMenu.ListIndex))
    If LenB(sNewName) Then
        lstMenu.List(lstMenu.ListIndex) = sNewName
    End If
End Sub
Private Sub cmdBrowseDes_Click()
Dim sDestPath As String
    sDestPath = ShowBFF(Me.hwnd, cTrans.GetString(107))
    If LenB(sDestPath) Then
        txtDestination.Text = sDestPath
    End If
    txtDestination.SetFocus
End Sub
Private Sub cmdBrowsePro_Click(Index As Integer)
    On Error GoTo ErrHandler
    cmDlg.Filter = "CoolWeb Project (*.cwp)|*.cwp"
    If optProject(2).Value Then
        cmDlg.ShowOpen
    Else
        cmDlg.ShowSave
    End If
    txtProject(Index).Text = cmDlg.FileName
    txtProject(Index).SetFocus
ErrHandler:
    If Err = 32755 Then
        txtProject(Index).SetFocus
    End If
End Sub
Private Sub dhePage_DisplayChanged(Index As Integer)
    With tbrEdit
        CheckButton .Buttons(1), DECMD_BOLD, Index
        CheckButton .Buttons(2), DECMD_ITALIC, Index
        CheckButton .Buttons(3), DECMD_UNDERLINE, Index
        CheckButton .Buttons(5), DECMD_JUSTIFYLEFT, Index
        CheckButton .Buttons(6), DECMD_JUSTIFYCENTER, Index
        CheckButton .Buttons(7), DECMD_JUSTIFYRIGHT, Index
        CheckButton .Buttons(12), DECMD_ORDERLIST, Index
        CheckButton .Buttons(13), DECMD_UNORDERLIST, Index
    End With
End Sub
Private Sub dhePage_ShowContextMenu(Index As Integer, _
                                    ByVal xPos As Long, _
                                    ByVal yPos As Long)
    CheckMenuItem mnuUndo, DECMD_UNDO, Index
    CheckMenuItem mnuRedo, DECMD_REDO, Index
    CheckMenuItem mnuCut, DECMD_CUT, Index
    CheckMenuItem mnuCopy, DECMD_COPY, Index
    CheckMenuItem mnuPaste, DECMD_PASTE, Index
    CheckMenuItem mnuDelete, DECMD_DELETE, Index
    CheckMenuItem mnuSelectAll, DECMD_SELECTALL, Index
    PopupMenu mnuPage, vbPopupMenuRightButton
End Sub
Private Sub lblWebsite_Click()
    ShellExecute 0, "open", lblWebsite.Caption, "", "", 10
End Sub
Public Sub lstMenu_Click()
    tbrMenu.Buttons(2).Enabled = (lstMenu.ListIndex > 0)
    dhePage(lstMenu.ListIndex).ZOrder
End Sub
Private Sub lstMenu_MouseUp(Button As Integer, _
                            Shift As Integer, _
                            x As Single, _
                            y As Single)
Dim lRet As Long
    If Button = 2 Then
        lRet = SendMessage(lstMenu.hwnd, LB_ITEMFROMPOINT, 0, ByVal ((CLng(y / Screen.TwipsPerPixelY) * 65536) + CLng(x / Screen.TwipsPerPixelX)))
        If lRet < lstMenu.ListCount Then
            lstMenu.ListIndex = lRet
            If lRet = 0 Then
                mnuDeletePage.Enabled = False
                mnuRenamePage.Enabled = False
            Else
                mnuDeletePage.Enabled = True
                mnuRenamePage.Enabled = True
            End If
            PopupMenu mnuItem
        Else
            mnuDeletePage.Enabled = False
            mnuRenamePage.Enabled = False
            PopupMenu mnuItem
        End If
    End If
End Sub
Private Sub lstTemplates_ItemClick(ByVal Item As ComctlLib.ListItem)
    If FileExists(Item.Key & "\screen_full.jpg") Then
        picSelTemp.Picture = LoadPicture(Item.Key & "\screen_full.jpg")
    Else
        picSelTemp.Picture = imgNoPreview(0)
    End If
    lblAuthor.Caption = ReadINI("Data", "Author", Item.Key & "\data.ini")
    lblWebsite.Caption = ReadINI("Data", "URL", Item.Key & "\data.ini")
    lblDescription.Caption = ReadINI("Data", "Description", Item.Key & "\data.ini")
End Sub
Private Sub optProject_Click(Index As Integer)
    Select Case Index
    Case 0
        cmdBrowsePro(0).Enabled = False
        cmdBrowsePro(1).Enabled = False
        txtProject(0).Enabled = False
        txtProject(1).Enabled = False
    Case 1
        cmdBrowsePro(0).Enabled = True
        cmdBrowsePro(1).Enabled = False
        txtProject(0).Enabled = True
        txtProject(1).Enabled = False
    Case 2
        cmdBrowsePro(0).Enabled = False
        cmdBrowsePro(1).Enabled = True
        txtProject(0).Enabled = False
        txtProject(1).Enabled = True
    End Select
End Sub
Private Sub tbrEdit_ButtonClick(ByVal Button As ComctlLib.Button)
Dim sSysDir As String
    Select Case Button.Key
    Case "Bold"
        dhePage(lstMenu.ListIndex).ExecCommand DECMD_BOLD
    Case "Italic"
        dhePage(lstMenu.ListIndex).ExecCommand DECMD_ITALIC
    Case "Underline"
        dhePage(lstMenu.ListIndex).ExecCommand DECMD_UNDERLINE
    Case "Left"
        dhePage(lstMenu.ListIndex).ExecCommand DECMD_JUSTIFYLEFT
    Case "Center"
        dhePage(lstMenu.ListIndex).ExecCommand DECMD_JUSTIFYCENTER
    Case "Right"
        dhePage(lstMenu.ListIndex).ExecCommand DECMD_JUSTIFYRIGHT
    Case "Indent"
        dhePage(lstMenu.ListIndex).ExecCommand DECMD_INDENT
    Case "Outdent"
        dhePage(lstMenu.ListIndex).ExecCommand DECMD_OUTDENT
    Case "Ordered"
        dhePage(lstMenu.ListIndex).ExecCommand DECMD_ORDERLIST
    Case "Unordered"
        dhePage(lstMenu.ListIndex).ExecCommand DECMD_UNORDERLIST
    Case "Hyperlink"
        dhePage(lstMenu.ListIndex).ExecCommand DECMD_HYPERLINK
    Case "Image"
        dhePage(lstMenu.ListIndex).ExecCommand DECMD_IMAGE
    Case "Font"
        dhePage(lstMenu.ListIndex).ExecCommand DECMD_FONT
    Case "DelFormat"
        dhePage(lstMenu.ListIndex).ExecCommand DECMD_REMOVEFORMAT
    Case "CharMap"
        sSysDir = Space$(255)
        sSysDir = Left$(sSysDir, GetSystemDirectory(sSysDir, 255))
        If FileExists(sSysDir & "\charmap.exe") Then
            ShellExecute Me.hwnd, "open", sSysDir & "\charmap.exe", vbNullString, "", 10
        Else
            MessageBox cTrans.GetString(515), cTrans.GetString(514)
        End If
    End Select
End Sub
Private Sub tbrMenu_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Key
    Case "Add"
        frmAdd.Show vbModal, Me
    Case "Delete"
        DeleteItem
    End Select
End Sub
Private Sub txtAuthor_GotFocus()
    txtAuthor.SelStart = 0
    txtAuthor.SelLength = Len(txtAuthor.Text)
End Sub
Private Sub txtDestination_GotFocus()
    txtDestination.SelStart = 0
    txtDestination.SelLength = Len(txtDestination.Text)
End Sub
Private Sub txtProject_GotFocus(Index As Integer)
    txtProject(Index).SelStart = 0
    txtProject(Index).SelLength = Len(txtProject(Index).Text)
End Sub
Private Sub txtTitle_GotFocus()
    txtTitle.SelStart = 0
    txtTitle.SelLength = Len(txtTitle.Text)
End Sub
Private Sub EnableMaxButton()
    With Me
        SetWindowLong .hwnd, GWL_STYLE, GetWindowLong(.hwnd, GWL_STYLE) Or WS_MAXIMIZEBOX
        SetWindowPos .hwnd, hWnd_NOTOPMOST, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_FRAMECHANGED
    End With
End Sub
Private Sub DisableMaxButton()
    With Me
        SetWindowLong .hwnd, GWL_STYLE, GetWindowLong(.hwnd, GWL_STYLE) And Not WS_MAXIMIZEBOX
        SetWindowPos .hwnd, hWnd_NOTOPMOST, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_FRAMECHANGED
        .WindowState = 0
    End With
End Sub
Private Function ReadINI(Section As String, _
                         KeyName As String, _
                         FileName As String, _
                         Optional Length As Long = 255) As String
Dim sRet As String
    sRet = String$(Length, vbNullChar)
    ReadINI = Left$(sRet, GetPrivateProfileString(Section, ByVal KeyName, vbNullString, sRet, Len(sRet), FileName))
End Function
Private Sub CheckButton(ByVal Button As ComctlLib.Button, _
                        TheFunction As DHTMLEDITCMDID, _
                        ByVal PageIndex As Integer)
    If dhePage(PageIndex).QueryStatus(TheFunction) = DECMDF_LATCHED Then
        Button.Value = tbrPressed
    Else
        Button.Value = tbrUnpressed
    End If
End Sub
Private Sub CheckMenuItem(MenuItem As Menu, _
                          TheFunction As DHTMLEDITCMDID, _
                          ByVal PageIndex As Integer)
    MenuItem.Enabled = dhePage(PageIndex).QueryStatus(TheFunction) >= DECMDF_ENABLED
End Sub
Private Sub DeleteItem()
Dim sTemp As String
Dim j     As Integer
    If lstMenu.ListIndex > -1 Then
        If MessageBox(cTrans.GetString(517) & lstMenu.Text & cTrans.GetString(518), cTrans.GetString(516), True) = vbOK Then
            If lstMenu.ListIndex = lstMenu.ListCount - 1 Then
                Unload dhePage(lstMenu.ListCount - 1)
                lstMenu.RemoveItem lstMenu.ListIndex
                lstMenu.ListIndex = 0
            Else
                For j = lstMenu.ListIndex + 1 To lstMenu.ListCount - 1
                    sTemp = dhePage(j).DocumentHTML
                    Unload dhePage(j - 1)
                    Load dhePage(j - 1)
                    dhePage(j - 1).Visible = True
                    dhePage(j - 1).DocumentHTML = sTemp
                    dhePage(lstMenu.ListIndex).ZOrder
                Next j
                Unload dhePage(dhePage.Count - 1)
                lstMenu.RemoveItem lstMenu.ListIndex
                lstMenu.ListIndex = 0
            End If
        End If
    End If
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
'The following function was provided by Pietro Cecchi
Private Function TakeCareOfUnicode(ByVal ss As String) As String
Dim ss1 As String
Dim a   As Long
Dim B   As String
Dim bb  As Long
    ss1 = vbNullString
    For a = 1 To Len(ss)
        B = Mid$(ss, a, 1)
        bb = AscW(B)
        If bb <= 255 Then
            ss1 = ss1 & B
        Else
            ss1 = ss1 & "&#" & bb & ";"
        End If
    Next a
    TakeCareOfUnicode = ss1
End Function
Private Sub LoadTemplates()
Dim i As Integer
    dirTemplates.Path = App.Path & "\Templates"
    For i = 0 To dirTemplates.ListCount
        If FileExists(dirTemplates.List(i) & "\data.ini") Then
            If FileExists(dirTemplates.List(i) & "\screen_thumb.jpg") Then
                imlTemplates.ListImages.Add , , LoadPicture(dirTemplates.List(i) & "\screen_thumb.jpg")
            Else
                imlTemplates.ListImages.Add , , imgNoPreview(1).Picture
            End If
            lstTemplates.ListItems.Add , dirTemplates.List(i), ReadINI("Data", "Name", dirTemplates.List(i) & "\data.ini"), imlTemplates.ListImages.Count
        End If
    Next i
    lstTemplates_ItemClick lstTemplates.ListItems(1)
    lstTemplates.SelectedItem = lstTemplates.ListItems.Item(1)
End Sub
Private Sub SaveProject(ByVal Project As String)
Dim oDom       As New MSXML.DOMDocument
Dim oParent    As MSXML.IXMLDOMNode
Dim oChild     As MSXML.IXMLDOMNode
Dim oAttribute As MSXML.IXMLDOMElement
Dim i          As Integer
    With oDom
        Set oParent = .appendChild(.createElement("Project"))
        Set oChild = oParent.appendChild(.createElement("Template"))
        oChild.Text = lstTemplates.SelectedItem.Text
        Set oChild = oParent.appendChild(.createElement("Title"))
        oChild.Text = txtTitle.Text
        Set oChild = oParent.appendChild(.createElement("Author"))
        oChild.Text = txtAuthor.Text
        Set oChild = oParent.appendChild(.createElement("Destination"))
        oChild.Text = txtDestination.Text
        Set oChild = oParent.appendChild(.createElement("Unicode"))
        oChild.Text = chkUnicode.Value
        Set oChild = oParent.appendChild(.createElement("Pages"))
        oChild.Text = lstMenu.ListCount - 1
        If chkUnicode.Value = 0 Then
            Set oChild = oParent.appendChild(.createElement("Home"))
            oChild.Text = dhePage(0).DocumentHTML
            For i = 1 To lstMenu.ListCount - 1
                Set oChild = oParent.appendChild(.createElement("Page" & i))
                Set oAttribute = oDom.documentElement.selectSingleNode("Page" & i)
                oAttribute.setAttribute "Name", lstMenu.List(i)
                oChild.Text = dhePage(i).DocumentHTML
            Next i
        Else
            Set oChild = oParent.appendChild(.createElement("Home"))
            oChild.Text = TakeCareOfUnicode(dhePage(0).DocumentHTML)
            For i = 1 To lstMenu.ListCount - 1
                Set oChild = oParent.appendChild(.createElement("Page" & i))
                Set oAttribute = oDom.documentElement.selectSingleNode("Page" & i)
                oAttribute.setAttribute "Name", lstMenu.List(i)
                oChild.Text = TakeCareOfUnicode(dhePage(i).DocumentHTML)
            Next i
        End If
        .save Project
    End With
    Set oDom = Nothing
    Set oParent = Nothing
    Set oChild = Nothing
    Set oAttribute = Nothing
End Sub
Private Sub LoadProject(ByVal Project As String)
Dim oDom       As New MSXML.DOMDocument
Dim oNode      As MSXML.IXMLDOMNode
Dim oAttribute As MSXML.IXMLDOMElement
Dim i          As Integer
    oDom.Load Project
    With oDom.documentElement
        Set oNode = .selectSingleNode("Template")
        For i = 1 To lstTemplates.ListItems.Count
            If lstTemplates.ListItems(i).Text = oNode.Text Then
                lstTemplates.ListItems(i).Selected = True
                lstTemplates_ItemClick lstTemplates.ListItems(i)
            End If
        Next i
        Set oNode = .selectSingleNode("Title")
        txtTitle.Text = oNode.Text
        Set oNode = .selectSingleNode("Author")
        txtAuthor.Text = oNode.Text
        Set oNode = .selectSingleNode("Destination")
        txtDestination.Text = oNode.Text
        Set oNode = .selectSingleNode("Unicode")
        chkUnicode.Value = oNode.Text
        Set oNode = .selectSingleNode("Home")
        dhePage(0).DocumentHTML = oNode.Text
        Set oNode = .selectSingleNode("Pages")
        If oNode.Text <> 0 Then
            For i = 1 To oNode.Text
                Set oNode = .selectSingleNode("Page" & i)
                Set oAttribute = .selectSingleNode("Page" & i)
                Load dhePage(dhePage.Count)
                dhePage(dhePage.Count - 1).Visible = True
                dhePage(dhePage.Count - 1).DocumentHTML = oNode.Text
                lstMenu.AddItem oAttribute.getAttribute("Name"), lstMenu.ListCount
                lstMenu.ListIndex = lstMenu.ListCount - 1
            Next i
        End If
    End With
    Set oDom = Nothing
    Set oNode = Nothing
    Set oAttribute = Nothing
End Sub
Private Sub UnloadProject()
Dim i As Integer
    lstTemplates.SelectedItem = lstTemplates.ListItems.Item(1)
    lstTemplates_ItemClick lstTemplates.ListItems(1)
    For i = 1 To lstMenu.ListCount - 1
        Unload dhePage(i)
        lstMenu.RemoveItem 1
    Next i
    dhePage(0).DocumentHTML = ""
End Sub
Private Sub Build()
Dim SHFO  As SHFILEOPSTRUCT
Dim sTemp As String
Dim sMain As String
Dim sMenu As String
Dim i     As Integer
    Open lstTemplates.SelectedItem.Key & "\menuitem.txt" For Input As #1
    sTemp = Input$(LOF(1), 1)
    Close #1
    For i = 1 To lstMenu.ListCount
        sMenu = sMenu & sTemp
        sMenu = Replace$(sMenu, "<!--ItemName-->", lstMenu.List(i - 1))
        If lstMenu.List(i - 1) = "Home" Then
            sMenu = Replace$(sMenu, "<!--ItemPath-->", "index.htm")
        Else
            sMenu = Replace$(sMenu, "<!--ItemPath-->", LCase$(lstMenu.List(i - 1)) & ".htm")
        End If
    Next i
    For i = 1 To lstMenu.ListCount
        Open lstTemplates.SelectedItem.Key & "\main.txt" For Input As #1
        sMain = Input$(LOF(1), 1)
        Close #1
        sMain = Replace$(sMain, "<!--Title-->", txtTitle.Text)
        sMain = Replace$(sMain, "<!--Author-->", txtAuthor.Text)
        sMain = Replace$(sMain, "<!--MenuItems-->", sMenu)
        sMain = Replace$(sMain, "<!--PageName-->", lstMenu.List(i - 1))
        If chkUnicode.Value = 0 Then
            sMain = Replace$(sMain, "<!--PageContent-->", dhePage(i - 1).DOM.body.innerHTML)
        Else
            sMain = Replace$(sMain, "<!--PageContent-->", TakeCareOfUnicode(dhePage(i - 1).DOM.body.innerHTML))
        End If
        sMain = Replace$(sMain, "<!--Year-->", Year(Date))
        sMain = Replace$(sMain, "<!--Date-->", Date)
        If lstMenu.List(i - 1) = "Home" Then
            sTemp = txtDestination.Text & "\index.htm"
        Else
            sTemp = txtDestination.Text & "\" & LCase$(lstMenu.List(i - 1) & ".htm")
        End If
        Open sTemp For Output As #1
        Print #1, sMain
        Close #1
    Next i
    With SHFO
        .iFlags = FOF_NOCONFIRMATION
        .lFunc = FO_COPY
        .sFrom = lstTemplates.SelectedItem.Key & "\files\*.*"
        .sTo = txtDestination.Text & "\files\"
        If LenB(.sTo) Then
            SHFileOperation SHFO
        End If
    End With
End Sub
