VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmOptions"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6360
   Icon            =   "frmOptions.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkAssociate 
      Caption         =   "chkAssociate"
      Height          =   195
      Left            =   360
      TabIndex        =   13
      Top             =   1920
      Width           =   5895
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
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
      TabIndex        =   12
      Top             =   3465
      Width           =   405
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
      Left            =   3840
      TabIndex        =   11
      Top             =   4080
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
      Left            =   5040
      TabIndex        =   10
      Top             =   4080
      Width           =   1125
   End
   Begin VB.TextBox txtCustDes 
      Height          =   285
      Left            =   960
      TabIndex        =   9
      Top             =   3480
      Width           =   4575
   End
   Begin VB.OptionButton optDestination 
      Caption         =   "optDestination(2)"
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
      Left            =   720
      TabIndex        =   7
      Top             =   3240
      Width           =   3615
   End
   Begin VB.OptionButton optDestination 
      Caption         =   "optDestination(1)"
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
      Index           =   1
      Left            =   720
      TabIndex        =   6
      Top             =   2940
      Width           =   3495
   End
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
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   2280
      Width           =   5895
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
   Begin VB.Label lblDefaultDes 
      Caption         =   "lblDefaultDes"
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
      TabIndex        =   8
      Top             =   2640
      Width           =   3000
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
      X2              =   6240
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   120
      X2              =   6240
      Y1              =   3975
      Y2              =   3975
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
        .Path = App.Path & "\Translations"
        .Pattern = "*.lng"
        For i = 0 To .ListCount - 1
            cmbTrans.AddItem GetFileName(.List(i))
            If LCase$(cmbTrans.List(i)) = LCase$(GetSetting("CWLanguage")) Then
                cmbTrans.ListIndex = i
            End If
        Next i
    End With
    chkUnicode.Value = GetSetting("Unicode")
    optDestination(GetSetting("DestDir")).Value = True
    txtCustDes.Text = GetSetting("CustDes")
    If GetSetting("Associate") = "" Then
        SaveSetting "Associate", 0
    Else
        chkAssociate.Value = GetSetting("Associate")
    End If

End Sub

Private Sub cmdOK_Click()

    SaveSetting "CWLanguage", cmbTrans.List(cmbTrans.ListIndex)
    SaveSetting "Associate", chkAssociate.Value
    SaveSetting "Unicode", chkUnicode.Value
    SaveSetting "CustDes", txtCustDes.Text
    If optDestination(1).Value Then
        SaveSetting "DestDir", 1
    Else
        SaveSetting "DestDir", 2
    End If
    If chkAssociate.Value = 1 Then
        Associate
    Else
        Dissociate
    End If
    If MessageBox(cTrans.GetString(530), cTrans.GetString(529), True) = vbOK Then
        cTrans.Translation = App.Path & "\Translations\" & cmbTrans.List(cmbTrans.ListIndex) & ".lng"
        Unload frmMain
        Load frmMain
        cTrans.LoadStrings
        cTrans.TranslateAll
        frmMain.Show
    End If
    Unload Me

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub optDestination_Click(Index As Integer)

    txtCustDes.Enabled = Not optDestination(1).Value
    cmdBrowse.Enabled = Not optDestination(1).Value

End Sub

Private Sub cmdBrowse_Click()

Dim sCustDes As String

    sCustDes = ShowBFF(Me.hwnd, cTrans.GetString(112))
    If LenB(sCustDes) Then
        txtCustDes.Text = sCustDes
    End If
    txtCustDes.SetFocus

End Sub

Private Sub txtCustDes_GotFocus()

    txtCustDes.SelStart = 0
    txtCustDes.SelLength = Len(txtCustDes.Text)

End Sub

