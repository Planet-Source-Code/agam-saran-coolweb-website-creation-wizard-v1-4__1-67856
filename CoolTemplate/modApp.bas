Attribute VB_Name = "modApp"
Option Explicit
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                                                                      ByVal lpValueName As String, _
                                                                                      ByVal lpReserved As Long, _
                                                                                      lpType As Long, _
                                                                                      lpData As Any, _
                                                                                      lpcbData As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, _
                                                                            ByVal lpSubKey As String, _
                                                                            phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, _
                                                                                  ByVal lpValueName As String, _
                                                                                  ByVal Reserved As Long, _
                                                                                  ByVal dwType As Long, _
                                                                                  lpData As Any, _
                                                                                  ByVal cbData As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, _
                                                                                ByVal lpSubKey As String, _
                                                                                phkResult As Long) As Long
Public Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hWnd As Long, _
                                                                           ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, _
                                                                           ByVal nIndex As Long, _
                                                                           ByVal dwNewLong As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, _
                                                                              ByVal lpOperation As String, _
                                                                              ByVal lpFile As String, _
                                                                              ByVal lpParameters As String, _
                                                                              ByVal lpDirectory As String, _
                                                                              ByVal nShowCmd As Long) As Long
                                                                              
Private Type tagInitCommonControlsEx
    lSize As Long
    lICC As Long
End Type

Private Const HKEY_LOCAL_MACHINE  As Long = &H80000002
Private Const REG_SZ  As Long = 1
Private Const ERROR_SUCCESS  As Long = 0&
Private Const ICC_USEREX_CLASSES  As Long = &H200

Public cTrans As New clsTranslator
Public MsgResult As VbMsgBoxResult
Private Sub Main()
    InitCommonControlsVB
    If LenB(GetSetting("CTLanguage")) = 0 Then
        cTrans.Translation = App.Path & "\Translations\English.lng"
        SaveSetting "CTLanguage", "English"
    Else
        cTrans.Translation = App.Path & "\Translations\" & GetSetting("CTLanguage") & ".lng"
    End If
    cTrans.LoadStrings
    If LenB(GetSetting("")) = 0 Then
        MessageBox cTrans.GetString(510), cTrans.GetString(509)
        End
    End If
    If LenB(GetSetting("WebShot")) = 0 Then
        SaveSetting "WebShot", "Not Installed"
    End If
    frmMain.cmDlg.Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist
    frmMain.Show
End Sub
Private Function InitCommonControlsVB() As Boolean
Dim iccex As tagInitCommonControlsEx
    On Error Resume Next
    With iccex
        .lSize = LenB(iccex)
        .lICC = ICC_USEREX_CLASSES
    End With
    InitCommonControlsEx iccex
    InitCommonControlsVB = (Err.Number = 0)
    On Error GoTo 0
End Function
Public Function FileExists(File As String) As Boolean
    FileExists = LenB(Dir(File))
End Function
Public Function GetSetting(ByVal Value As String) As String
Dim lKeyHand     As Long
Dim lResult      As Long
Dim sBuffer      As String
Dim lDataBufSize As Long
Dim iZeroPos     As Integer
Dim lValueType   As Long
    RegOpenKey HKEY_LOCAL_MACHINE, "Software\CoolWeb", lKeyHand
    lResult = RegQueryValueEx(lKeyHand, Value, 0&, lValueType, ByVal 0&, lDataBufSize)
    If lValueType = REG_SZ Then
        sBuffer = String$(lDataBufSize, " ")
        lResult = RegQueryValueEx(lKeyHand, Value, 0&, 0&, ByVal sBuffer, lDataBufSize)
        If lResult = ERROR_SUCCESS Then
            iZeroPos = InStr(sBuffer, vbNullChar)
            If iZeroPos > 0 Then
                GetSetting = Left$(sBuffer, iZeroPos - 1)
            Else
                GetSetting = sBuffer
            End If
        End If
    End If
End Function
Public Sub SaveSetting(Value As String, _
                       ByVal Data As String)
Dim lKeyHand As Long
    RegCreateKey HKEY_LOCAL_MACHINE, "Software\CoolWeb", lKeyHand
    RegSetValueEx lKeyHand, Value, 0, REG_SZ, ByVal Data, Len(Data)
    RegCloseKey lKeyHand
End Sub
Public Function MessageBox(ByVal Message As String, _
                           Optional ByVal Title As String = vbNullString, _
                           Optional UseCancel As Boolean = False) As VbMsgBoxResult
    With frmMessage
        .lblMessage.Caption = Message
        If LenB(Title) = 0 Then
            .Caption = "CoolTemplate"
        Else
            .Caption = "CoolTemplate - " & Title
        End If
        .cmdCancel.Enabled = UseCancel
        .Show vbModal
        MessageBox = MsgResult
    End With
End Function
