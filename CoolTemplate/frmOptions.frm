VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmOptions"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5940
   Icon            =   "frmOptions.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   320
      Left            =   5280
      TabIndex        =   9
      Top             =   2280
      Width           =   405
   End
   Begin VB.TextBox txtWebShot 
      Height          =   285
      Left            =   600
      TabIndex        =   8
      Top             =   2280
      Width           =   4575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "cmdOK"
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
      Left            =   3480
      TabIndex        =   6
      Top             =   3120
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
      Left            =   4680
      TabIndex        =   5
      Top             =   3120
      Width           =   1125
   End
   Begin VB.ComboBox cmbTrans 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1250
      Width           =   1695
   End
   Begin VB.PictureBox picTop 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   -120
      ScaleHeight     =   885
      ScaleWidth      =   8295
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caption"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   435
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Width           =   1530
      End
      Begin VB.Line lneTop 
         BorderColor     =   &H0099A8AC&
         X1              =   0
         X2              =   6600
         Y1              =   855
         Y2              =   855
      End
      Begin VB.Image imgTopIcon 
         Height          =   840
         Left            =   240
         Picture         =   "frmOptions.frx":000C
         Top             =   10
         Width           =   840
      End
      Begin VB.Label lblShadow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caption"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   1230
         TabIndex        =   2
         Top             =   270
         Width           =   1515
      End
   End
   Begin VB.Label lblWebShot 
      Caption         =   "lblWebShot"
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label lblLanguage 
      Caption         =   "lblLanguage"
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
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000011&
      X1              =   120
      X2              =   5760
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   120
      X2              =   5760
      Y1              =   3015
      Y2              =   3000
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
Dim i As Integer
    cTrans.SetTranslation Me
    lblCaption.Caption = Me.Caption
    lblShadow.Caption = Me.Caption
    With frmMain.filTemp
        .Pattern = "*.lng"
        .Path = App.Path & "\Translations"
        For i = 0 To .ListCount - 1
            cmbTrans.AddItem GetFileName(.List(i))
            If LCase$(cmbTrans.List(i)) = LCase$(GetSetting("CTLanguage")) Then
                cmbTrans.ListIndex = i
            End If
        Next i
    End With
    txtWebShot.Text = GetSetting("WebShot")
End Sub
Private Sub cmdOK_Click()
    SaveSetting "WebShot", txtWebShot.Text
    SaveSetting "CTLanguage", cmbTrans.List(cmbTrans.ListIndex)
    If MessageBox(cTrans.GetString(512), cTrans.GetString(511), True) = vbOK Then
        cTrans.Translation = App.Path & "\Translations\" & cmbTrans.List(cmbTrans.ListIndex) & ".lng"
        Unload frmMain
        frmMain.Show
        cTrans.LoadStrings
        cTrans.TranslateAll
    End If
    Unload Me
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdBrowse_Click()
    On Error GoTo ErrHandler
    With frmMain.cmDlg
        .Filter = "webshotcmd.exe|webshotcmd.exe"
        .ShowOpen
        txtWebShot.Text = .FileName
        cmdOK.SetFocus
    End With
ErrHandler:
    If Err = 32755 Then
        txtWebShot.Text = "Not Installed"
        cmdOK.SetFocus
    End If
End Sub
Private Function GetFileName(File As String) As String
Dim i As Integer
    i = InStrRev(File, "\") + 1
    GetFileName = Mid$(File, i, InStrRev(File, ".") - i)
End Function
