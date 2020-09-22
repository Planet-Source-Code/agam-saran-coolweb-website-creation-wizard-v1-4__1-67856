VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Begin VB.Form frmAdd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmMain"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5640
   Icon            =   "frmAdd.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar tbrTabs 
      Height          =   390
      Left            =   60
      TabIndex        =   3
      Top             =   30
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   688
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      ImageList       =   "imlToolbar"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   3
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Blank"
            Object.Tag             =   ""
            ImageIndex      =   1
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Open"
            Object.Tag             =   ""
            ImageIndex      =   2
            Style           =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Template"
            Object.Tag             =   ""
            ImageIndex      =   3
            Style           =   1
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1710
      Left            =   4440
      Picture         =   "frmAdd.frx":000C
      ScaleHeight     =   1710
      ScaleWidth      =   975
      TabIndex        =   6
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
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
      Left            =   4320
      TabIndex        =   2
      Top             =   2880
      Width           =   1125
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "cmdAdd"
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
      Left            =   3120
      TabIndex        =   1
      Top             =   2880
      Width           =   1125
   End
   Begin ComctlLib.ListView lstPages 
      Height          =   2055
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   3625
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDragMode     =   1
      _Version        =   327682
      SmallIcons      =   "imlToolbar"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      OLEDragMode     =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "lstPages"
         Object.Width           =   6703
      EndProperty
   End
   Begin VB.CommandButton cmdBrowse 
      Cancel          =   -1  'True
      Caption         =   "cmdBrowse"
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
      Left            =   3200
      TabIndex        =   9
      Top             =   1320
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   1320
      Width           =   3495
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   840
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label lblFile 
      AutoSize        =   -1  'True
      Caption         =   "lblFile"
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
      Left            =   390
      TabIndex        =   7
      Top             =   1005
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "lblName"
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
      TabIndex        =   5
      Top             =   1350
      Width           =   555
   End
   Begin VB.Label lblTab 
      Caption         =   "lblTab"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1320
      TabIndex        =   4
      Top             =   120
      Width           =   4170
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0099A8AC&
      X1              =   0
      X2              =   6120
      Y1              =   420
      Y2              =   420
   End
   Begin ComctlLib.ImageList imlToolbar 
      Left            =   960
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmAdd.frx":0B76
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmAdd.frx":0EC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmAdd.frx":121A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmAdd.frx":156C
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                            ByVal Msg As Long, _
                                                                            ByVal wParam As Long, _
                                                                            ByVal lParam As Long) As Long
                                                                            
Private Const LVM_FIRST As Long = &H1000
Private Const LVM_SETEXTENDEDLISTVIEWSTYLE As Double = LVM_FIRST + 54
Private Const LVM_GETEXTENDEDLISTVIEWSTYLE As Double = LVM_FIRST + 55
Private Const LVS_EX_FULLROWSELECT  As Long = &H20

Private iMode As Integer
Private Sub Form_Load()
Dim i As Integer
    cTrans.SetTranslation Me
    txtName.Text = cTrans.GetString(108) & iNewPage
    lblTab.Caption = cTrans.GetString(109)
    iMode = 1
    SendMessageLong lstPages.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, SendMessageLong(lstPages.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&) Or LVS_EX_FULLROWSELECT
    With frmMain.filTemp
        .Path = App.Path & "\Pages"
        .Pattern = "*.htm;*.html"
        For i = 0 To .ListCount - 1
            lstPages.ListItems.Add , App.Path & "\Pages\" & .List(i), GetFileName(.List(i)), , 4
        Next i
    End With
End Sub
Private Sub cmdAdd_Click()
    Select Case iMode
    Case 1
        If LenB(txtName.Text) = 0 Then
            MessageBox cTrans.GetString(520), cTrans.GetString(519)
            txtName.SetFocus
        Else
            If Left$(txtName.Text, Len(cTrans.GetString(108))) = cTrans.GetString(108) Then
                iNewPage = iNewPage + 1
            End If
            AddPage txtName.Text
        End If
    Case 2
        If LenB(txtFile.Text) = 0 Or LenB(txtName.Text) = 0 Then
            MessageBox cTrans.GetString(522), cTrans.GetString(521)
            txtFile.SetFocus
        ElseIf Not FileExists(txtFile.Text) Then
            MessageBox cTrans.GetString(524), cTrans.GetString(523)
            txtFile.SetFocus
        Else
            If Left$(txtName.Text, Len(cTrans.GetString(108))) = cTrans.GetString(108) Then
                iNewPage = iNewPage + 1
            End If
            AddPage txtName.Text, txtFile.Text
        End If
    Case 3
        AddPage GetFileName(lstPages.SelectedItem.Key), lstPages.SelectedItem.Key
    End Select
End Sub
Private Sub tbrTabs_ButtonClick(ByVal Button As ComctlLib.Button)
    With tbrTabs
        .Buttons(1).Value = tbrUnpressed
        .Buttons(2).Value = tbrUnpressed
        .Buttons(3).Value = tbrUnpressed
    End With
    Button.Value = tbrPressed
    Select Case Button.Key
    Case "Blank"
        lblTab.Caption = cTrans.GetString(109)
        lblFile.Visible = False
        txtFile.Visible = False
        cmdBrowse.Visible = False
        lblName.Visible = True
        txtName.Visible = True
        lstPages.Visible = False
        txtName.Top = 1320
        lblName.Top = 1350
        txtName.SetFocus
        iMode = 1
    Case "Open"
        lblTab.Caption = cTrans.GetString(110)
        lblFile.Visible = True
        txtFile.Visible = True
        cmdBrowse.Visible = True
        lblName.Visible = True
        txtName.Visible = True
        lstPages.Visible = False
        txtName.Top = 1920
        lblName.Top = 1950
        txtFile.SetFocus
        iMode = 2
    Case "Template"
        lblTab.Caption = cTrans.GetString(111)
        lstPages.Visible = True
        lstPages.ListItems(1).Selected = True
        iMode = 3
    End Select
End Sub
Private Sub cmdBrowse_Click()
    On Error GoTo ErrHandler
    With frmMain.cmDlg
        .Filter = "HTML Files(*.htm,*.html)|*.htm;*.html"
        .ShowOpen
        txtFile.Text = .FileName
        txtName.Text = GetFileName(.FileName)
        txtName.SetFocus
    End With
ErrHandler:
    If Err = 32755 Then
        txtFile.SetFocus
    End If
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub txtFile_GotFocus()
    txtFile.SelStart = 0
    txtFile.SelLength = Len(txtFile.Text)
End Sub
Private Sub txtName_GotFocus()
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName.Text)
End Sub
Private Sub AddPage(ByVal PageName As String, _
                    Optional ByVal LoadFile As String = vbNullString)
Dim bExists As Boolean
Dim i       As Integer
    With frmMain
        For i = 0 To .lstMenu.ListCount
            If LCase$(.lstMenu.List(i)) = LCase$(PageName) Then
                bExists = True
            End If
        Next i
        If bExists Then
            MessageBox cTrans.GetString(526), cTrans.GetString(525)
            txtName.SetFocus
        Else
            Load .dhePage(.dhePage.Count)
            .dhePage(.dhePage.Count - 1).Visible = True
            .dhePage(.dhePage.Count - 1).ZOrder
            .lstMenu.AddItem PageName, .lstMenu.ListCount
            .lstMenu.ListIndex = .lstMenu.ListCount - 1
            If LenB(LoadFile) Then
                .dhePage(.dhePage.Count - 1).LoadDocument LoadFile
            End If
            Unload Me
        End If
    End With
End Sub
