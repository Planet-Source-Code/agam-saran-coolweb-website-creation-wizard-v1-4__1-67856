VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About CoolWeb"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5625
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   345
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picShow 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3420
      Left            =   1170
      ScaleHeight     =   228
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   197
      TabIndex        =   0
      Top             =   1080
      Width           =   2955
   End
   Begin VB.PictureBox picLeft 
      BorderStyle     =   0  'None
      Height          =   1785
      Left            =   0
      Picture         =   "frmAbout.frx":000C
      ScaleHeight     =   1785
      ScaleWidth      =   3585
      TabIndex        =   3
      Top             =   0
      Width           =   3585
      Begin VB.Label lblVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "Version: 1.4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   270
         Left            =   1040
         TabIndex        =   4
         Top             =   720
         Width           =   1365
      End
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   15
      Left            =   240
      Top             =   3360
   End
   Begin VB.PictureBox picBottom 
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   0
      ScaleHeight     =   630
      ScaleWidth      =   6000
      TabIndex        =   1
      Top             =   4560
      Width           =   6000
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
         Height          =   345
         Left            =   2160
         TabIndex        =   2
         Top             =   120
         Width           =   1245
      End
   End
   Begin VB.PictureBox picRight 
      BorderStyle     =   0  'None
      Height          =   1785
      Left            =   3570
      Picture         =   "frmAbout.frx":1B2E
      ScaleHeight     =   1785
      ScaleWidth      =   2055
      TabIndex        =   5
      Top             =   0
      Width           =   2055
   End
   Begin VB.Line lneBottom 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   376
      Y1              =   302
      Y2              =   302
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, _
                                                                  ByVal lpStr As String, _
                                                                  ByVal nCount As Long, _
                                                                  lpRect As RECT, _
                                                                  ByVal wFormat As Long) As Long
                                                                  
Private Type RECT
    lLeft As Long
    lTop As Long
    lRight As Long
    lBottom As Long
End Type

Private iTop         As Integer
Private sCredits()   As String
Private bBold(0 To 19)   As Boolean
Private Sub Form_Load()
Dim i     As Integer
    iTop = picShow.Height
    sCredits = Split("Created & Programmed By:" _
    & vbNewLine & "Agam Saran" _
    & vbNewLine & "www.agamsaran.co.nr" _
    & vbNewLine _
    & vbNewLine & "Templates Designed By:" _
    & vbNewLine & "Respective Users of OSWD" _
    & vbNewLine & "www.oswd.org" _
    & vbNewLine _
    & vbNewLine & "Toolbar-Icons Designed By:" _
    & vbNewLine & "Mark James" _
    & vbNewLine & "famfamfam.com/lab/icons/silk" _
    & vbNewLine _
    & vbNewLine & "Special Thanks To:" _
    & vbNewLine & "Xpert" _
    & vbNewLine & "Pietro Cecchi" _
    & vbNewLine _
    & vbNewLine & "Special, Special Thanks To:" _
    & vbNewLine & "YOU!" _
    & vbNewLine & "For Using CoolWeb", vbNewLine)
    For i = 0 To 16 Step 4
        bBold(i) = True
    Next i
End Sub
Private Sub tmrUpdate_Timer()
Dim Rectangle As RECT
Dim iTextTop  As Integer
Dim iLength   As Integer
Dim lDrawCol  As Long
Dim i         As Integer
    picShow.Cls
    iTextTop = iTop
    For i = 0 To UBound(sCredits)
        If iTextTop > -50 Then
            If iTextTop < picShow.Height Then
                iLength = picShow.Height * (1 / 6)
                If iTextTop <= iLength And iTextTop >= -50 Then
                    lDrawCol = GetShade(vbBlack, &HE0E0E0, (iLength - iTextTop) / (iLength + 20))
                ElseIf iTextTop <= picShow.Height And iTextTop >= picShow.Height * (1 - (1 / 6)) Then
                    lDrawCol = GetShade(&HE0E0E0, vbBlack, (picShow.Height - iTextTop) / iLength)
                Else
                    lDrawCol = vbBlack
                End If
                With Rectangle
                    .lTop = iTextTop
                    .lRight = picShow.Width
                    .lBottom = picShow.Height
                End With
                With picShow
                    .FontBold = bBold(i)
                    .ForeColor = lDrawCol
                    DrawText .hdc, sCredits(i), -1, Rectangle, &H800 Or &H1
                End With
            End If
        End If
        iTextTop = iTextTop + 16
    Next i
    If iTop + 20 < -16 * UBound(sCredits) Then
        iTop = picShow.Height
    End If
    iTop = iTop - 1
End Sub
Private Sub cmdOK_Click()
    Unload Me
End Sub
Private Function GetShade(ByVal StartCol As Long, _
                          ByVal EndCol As Long, _
                          ByVal ColDepth As Double) As Long
Dim dRate As Double
Dim cBlue As Long, cGreen As Long, cRed As Long
Dim sBlue As Long, sGreen As Long, sRed As Long
    On Error Resume Next
    dRate = ColDepth
    GetRGB EndCol, sRed, sGreen, sBlue
    GetRGB StartCol, cRed, cGreen, cBlue
    cRed = cRed + (sRed - cRed) * dRate
    cGreen = cGreen + (sGreen - cGreen) * dRate
    cBlue = cBlue + (sBlue - cBlue) * dRate
    If cRed < 0 Then cRed = -cRed
    If cGreen < 0 Then cGreen = -cGreen
    If cBlue < 0 Then cBlue = -cBlue
    GetShade = RGB(cRed, cGreen, cBlue)
    On Error GoTo 0
End Function
Private Sub GetRGB(ByVal LngCol As Long, _
                   R As Long, _
                   G As Long, _
                   B As Long)
    R = LngCol Mod 256
    G = (LngCol And vbGreen) / 256
    B = (LngCol And vbBlue) / 65536
End Sub
