Attribute VB_Name = "modApp"
Option Explicit
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, _
                                                                            ByVal lpSubKey As String, _
                                                                            phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                                                                      ByVal lpValueName As String, _
                                                                                      ByVal lpReserved As Long, _
                                                                                      lpType As Long, _
                                                                                      lpData As Any, _
                                                                                      lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, _
                                                                              ByVal lpSubKey As String, _
                                                                              ByVal dwType As Long, _
                                                                              ByVal lpData As String, _
                                                                              ByVal cbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, _
                                                                                  ByVal lpValueName As String, _
                                                                                  ByVal Reserved As Long, _
                                                                                  ByVal dwType As Long, _
                                                                                  lpData As Any, _
                                                                                  ByVal cbData As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, _
                                                                                ByVal lpSubKey As String, _
                                                                                phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, _
                                                                                ByVal lpSubKey As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, _
                                                                  ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpBI As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, _
                                                            ByVal lpBuffer As String) As Long
Private Declare Sub SHChangeNotify Lib "shell32.dll" (ByVal wEventId As Long, _
                                                      ByVal uFlags As Long, _
                                                      dwItem1 As Any, _
                                                      dwItem2 As Any)
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
                                                                              ByVal lpOperation As String, _
                                                                              ByVal lpFile As String, _
                                                                              ByVal lpParameters As String, _
                                                                              ByVal lpDirectory As String, _
                                                                              ByVal nShowCmd As Long) As Long
                                                                              
Private Const REG_SZ As Long = 1
Private Const ERROR_SUCCESS As Long = 0&
Private Const ICC_USEREX_CLASSES As Long = &H200
Private Const BIF_RETURNONLYFSDIRS As Long = 1
Private Const HKEY_CLASSES_ROOT As Long = &H80000000
Private Const HKEY_LOCAL_MACHINE As Long = &H80000002
Private Const SHCNE_ASSOCCHANGED As Long = &H8000000
Private Const SHCNF_IDLIST As Long = &H0&

Private Type tagInitCommonControlsEx
    lSize As Long
    lICC As Long
End Type
Private Type BrowseInfo
    lHwnd As Long
    lIDLRoot As Long
    lDisplayName As Long
    lTitle As Long
    lFlags As Long
    lCallback As Long
    lParam As Long
    lImage As Long
End Type

Public cTrans As New clsTranslator
Public MsgResult As VbMsgBoxResult
Public sInputText As String
Public iNewPage As Integer
Private Sub Main()
    InitCommonControlsVB
    If LenB(GetSetting("CWLanguage")) = 0 Then
        cTrans.Translation = App.Path & "\Translations\English.lng"
        SaveSetting "CWLanguage", "English"
    Else
        cTrans.Translation = App.Path & "\Translations\" & GetSetting("CWLanguage") & ".lng"
    End If
    cTrans.LoadStrings
    SaveSetting "", App.Path
    If LenB(GetSetting("Unicode")) = 0 Then
        SaveSetting "Unicode", 0
    End If
    If LenB(GetSetting("DestDir")) = 0 Then
        SaveSetting "DestDir", 1
    End If
    iNewPage = 1
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
Public Function ShowBFF(FrmHwnd As Long, _
                        Title As String) As String
Dim iNull    As Integer
Dim lpIDList As Long
Dim sPath    As String
Dim udtBI    As BrowseInfo
    With udtBI
        .lHwnd = FrmHwnd
        .lTitle = lstrcat((Title), "")
        .lFlags = BIF_RETURNONLYFSDIRS
    End With
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(260, 0)
        SHGetPathFromIDList lpIDList, sPath
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If
    ShowBFF = sPath
End Function
Public Function GetString(RegKey As Long, _
                          Path As String, _
                          ByVal Value As String) As String
Dim lKeyHand     As Long
Dim lResult      As Long
Dim sBuffer      As String
Dim lDataBufSize As Long
Dim iZeroPos     As Integer
Dim lValueType   As Long
    RegOpenKey RegKey, Path, lKeyHand
    lResult = RegQueryValueEx(lKeyHand, Value, 0&, lValueType, ByVal 0&, lDataBufSize)
    If lValueType = REG_SZ Then
        sBuffer = String$(lDataBufSize, " ")
        lResult = RegQueryValueEx(lKeyHand, Value, 0&, 0&, ByVal sBuffer, lDataBufSize)
        If lResult = ERROR_SUCCESS Then
            iZeroPos = InStr(sBuffer, vbNullChar)
            If iZeroPos > 0 Then
                GetString = Left$(sBuffer, iZeroPos - 1)
            Else
                GetString = sBuffer
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
Public Function GetSetting(Value As String) As String
    GetSetting = GetString(HKEY_LOCAL_MACHINE, "Software\CoolWeb", Value)
End Function
Public Function FileExists(File As String) As Boolean
    FileExists = LenB(Dir(File))
End Function
Public Function GetFileName(File As String) As String
Dim i As Integer
    i = InStrRev(File, "\") + 1
    GetFileName = Mid$(File, i, InStrRev(File, ".") - i)
End Function
Public Function MessageBox(ByVal Message As String, _
                           Optional ByVal Title As String = vbNullString, _
                           Optional ByVal UseCancel As Boolean = False) As VbMsgBoxResult
    With frmMessage
        .lblMessage.Caption = Message
        If LenB(Title) = 0 Then
            .Caption = "CoolWeb"
        Else
            .Caption = "CoolWeb - " & Title
        End If
        .imgExclame.Visible = True
        .imgInput.Visible = False
        .txtInput.Visible = False
        .cmdCancel.Enabled = UseCancel
        .Show vbModal
        MessageBox = MsgResult
    End With
End Function
Public Function FieldBox(ByVal Message As String, _
                         Optional ByVal Title As String = vbNullString, _
                         Optional ByVal DefaultValue As String = vbNullString) As String
    With frmMessage
        .lblMessage.Caption = Message
        If LenB(Title) = 0 Then
            .Caption = "CoolWeb"
        Else
            .Caption = "CoolWeb - " & Title
        End If
        .imgExclame.Visible = False
        .imgInput.Visible = True
        .cmdCancel.Enabled = True
        .txtInput.Visible = True
        .txtInput.Text = DefaultValue
        .Show vbModal
        FieldBox = sInputText
    End With
End Function
Public Sub Associate()
Dim lKeyHand As Long
    RegCreateKey HKEY_CLASSES_ROOT, "CoolWeb", lKeyHand
    RegSetValue lKeyHand, "", REG_SZ, "CoolWeb Project", 0
    RegCreateKey HKEY_CLASSES_ROOT, ".cwp", lKeyHand
    RegSetValue lKeyHand, "", REG_SZ, "CoolWeb", 0
    RegCreateKey HKEY_CLASSES_ROOT, "CoolWeb", lKeyHand
    RegSetValue lKeyHand, "shell\open\command", REG_SZ, App.Path & "\" & App.EXEName & ".exe %1", 260
    RegCreateKey HKEY_CLASSES_ROOT, "CoolWeb", lKeyHand
    RegSetValue lKeyHand, "DefaultIcon", REG_SZ, App.Path & "\" & App.EXEName & ".exe,1", 260
    SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0
End Sub
Public Sub Dissociate()
    RegDeleteKey HKEY_CLASSES_ROOT, ".cwp"
    RegDeleteKey HKEY_CLASSES_ROOT, "CoolWeb\DefaultIcon"
    RegDeleteKey HKEY_CLASSES_ROOT, "CoolWeb\Shell\Open\Command"
    RegDeleteKey HKEY_CLASSES_ROOT, "CoolWeb\Shell\Open"
    RegDeleteKey HKEY_CLASSES_ROOT, "CoolWeb\Shell"
    RegDeleteKey HKEY_CLASSES_ROOT, "CoolWeb"
    SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0
End Sub
