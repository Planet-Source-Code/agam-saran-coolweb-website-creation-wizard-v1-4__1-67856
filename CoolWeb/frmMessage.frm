VERSION 5.00
Begin VB.Form frmMessage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmMessage"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5085
   Icon            =   "frmMessage.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtInput 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      ToolTipText     =   "Input Value"
      Top             =   720
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
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
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label lblMessage 
      Caption         =   "lblMessage"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   1560
      TabIndex        =   1
      Top             =   480
      Width           =   3375
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000011&
      X1              =   120
      X2              =   4920
      Y1              =   1545
      Y2              =   1545
   End
   Begin VB.Image imgExclame 
      Height          =   945
      Left            =   240
      Picture         =   "frmMessage.frx":000C
      Top             =   240
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   120
      X2              =   4920
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Image imgInput 
      Height          =   1485
      Left            =   240
      Picture         =   "frmMessage.frx":07B6
      Top             =   0
      Visible         =   0   'False
      Width           =   1140
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    If txtInput.Visible Then
        txtInput.SetFocus
    End If
End Sub
Private Sub cmdOK_Click()
    MsgResult = vbOK
    sInputText = txtInput.Text
    Unload Me
End Sub
Private Sub cmdCancel_Click()
    MsgResult = vbCancel
    sInputText = vbNullString
    Unload Me
End Sub
Private Sub txtInput_GotFocus()
    txtInput.SelStart = 0
    txtInput.SelLength = Len(txtInput.Text)
End Sub
