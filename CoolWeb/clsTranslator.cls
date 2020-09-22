VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTranslator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'written by: Agam Saran
'You can also use this class in your own apps without any hesitation

Option Explicit

'An API used by "ItemIndex" function
Private Declare Sub PokeLong Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, Optional ByVal Length As Long = 4)

'Collection in which all the Strings from the translation are loaded
Private cStrings     As New Collection
'The path of translation file is stored in this variable
Private sTransFile   As String

Public Property Let Translation(ByVal TransFile As String)
'Set the value of "sTransFile"
    sTransFile = TransFile
End Property

Public Property Get Translation() As String
'Return the value given by "sTransFile"
    Translation = sTransFile
End Property

Public Property Get StringCount() As Integer
'Return the number of items in "cStrings" collection
    StringCount = cStrings.Count
End Property

Public Sub TranslateAll()
Dim i As Integer

    'Loop through all the loaded forms
    For i = 0 To Forms.Count - 1
        'and set the translation of each form
        SetTranslation Forms(i)
    Next i
End Sub

Public Function GetString(ID As Integer) As String
    On Error Resume Next
    'Get the index of the String ID and return
    'its value from "cStrings" collection
    GetString = cStrings(ItemIndex(ID, cStrings))
    On Error GoTo 0
End Function

'This function was taken from a submission by Ulli
'It returns the index of the collection item whose key is given
Private Function ItemIndex(ByVal Key As String, _
                           Coll As Collection, _
                           Optional ByVal Compare As VbCompareMethod = vbTextCompare) As Long
Dim Ptr  As Long
Dim sKey As String
Dim aKey As Long

    If Coll Is Nothing Then                             'oops!
        Err.Raise 91                                    'No object
    Else
        If Coll.Count Then
            aKey = StrPtr(sKey)                         'save StrPtr
            PokeLong Ptr, ByVal ObjPtr(Coll) + 24       'first Ptr
            ItemIndex = 1                               'walk items upwards from first
            Do
                PokeLong ByVal VarPtr(sKey), ByVal Ptr + 16
                If StrComp(Key, sKey, Compare) = 0 Then 'equal
                    Exit Do                             'found
                End If
                ItemIndex = ItemIndex + 1               'next Index
                PokeLong Ptr, ByVal Ptr + 24            'next Ptr
            Loop Until Ptr = 0                          'end of chain
            PokeLong ByVal VarPtr(sKey), aKey           'restore original StrPtr
        End If
        If Ptr = 0 Then
            ItemIndex = -1                              'key not found
        End If
    End If

End Function

Public Sub LoadStrings()
Dim sCurrent As String
Dim iPos     As Integer

    'Be sure the translation file is specified and it exists
    If LenB(sTransFile) = 0 Or LenB(Dir(sTransFile)) = 0 Then Exit Sub

    With cStrings
        'If Strings are already loaded in "cStrings"
        If .Count <> 0 Then
            'Loop through each string
            For iPos = 1 To .Count
                'and remove it
                .Remove 1
            Next iPos
        End If
    End With

    'Open the translation file for reading
    Open sTransFile For Input As #1
    'Loop through each line
    Do Until EOF(1)
        'and place its content in a variable
        Line Input #1, sCurrent
        'Check whether current line is a String and if it is
        If Left$(sCurrent, 6) = "String" Then
            'Get the position of "=" symbol
            iPos = InStr(1, sCurrent, "=")
            'Get the value of string and add it as an item
            'of "cStrings" collection, also retrieve
            'the String ID and save it as the key of item
            cStrings.Add Mid$(sCurrent, iPos + 1, Len(sCurrent) - iPos), Mid$(sCurrent, 7, iPos - 7)
        End If
    Loop
    Close #1

End Sub

Public Sub SetTranslation(Frm As Form)
Dim sCurrent  As String, sValue     As String
Dim sSection  As String
Dim iPos      As Integer
Dim ctl       As Control, sCtlType  As String
Dim iCtlIndex As Integer, iObjIndex As Integer
Dim sCaption  As String, sToolTip   As String

On Error Resume Next

    'If translation is file in not specified yet or
    'if it does not exist then exit the function
    If LenB(sTransFile) = 0 Or LenB(Dir(sTransFile)) = 0 Then Exit Sub
    
    'Open the translation file for reading
    Open sTransFile For Input As #1
    'Loop through each line
    Do Until EOF(1)
        'and place it in "sCurrent" variable
        Line Input #1, sCurrent
        'Make sure we are not working on an empty or a comment line
        If sCurrent <> "" And Left$(sCurrent, 1) <> ";" Then
            'Also make sure it is not a String
            If Left$(sCurrent, 6) <> "String" Then
                'Check whether it is a Form section
                If Left$(sCurrent, 1) = "[" And Right$(sCurrent, 1) = "]" Then
                    'If it is, save its name
                    sSection = Mid$(sCurrent, 2, Len(sCurrent) - 2)
                    'If it's not, make sure we are working on the form specified
                ElseIf LCase$(sSection) = LCase$(Frm.Name) Then
                    'Get the position of "="
                    iPos = InStr(1, sCurrent, "=")
                    'Save the right-hand part in "sValue"
                    sValue = Mid$(sCurrent, iPos + 1, Len(sCurrent) - iPos)
                    'Save the left-hand part in "sCurrent"
                    sCurrent = Left$(sCurrent, iPos - 1)
                    'If the object in current line is the Form itself
                    If LCase$(sCurrent) = LCase$(Frm.Name) Then
                        'Then set the caption of the form
                        Frm.Caption = sValue
                    Else    'Otherwise, its a control on the form
                        'If the index is specified and there is a control array
                        If Right$(sCurrent, 1) = ")" Then
                            iPos = InStr(1, sCurrent, "(")
                            'Get the index and save it in "iCtlIndex"
                            iCtlIndex = Mid$(sCurrent, iPos + 1, InStr(iPos + 1, sCurrent, ")") - iPos - 1)
                            'and separate the name of control from the index
                            sCurrent = Left$(sCurrent, iPos - 1)
                        Else
                            'Otherwise, set a flag of "-1" which shows that
                            'there is a single control not a control array
                            iCtlIndex = -1
                        End If

                        'See whether the control has object in it
                        'and the object index is specified
                        If Left$(sCurrent, 1) = "{" Then
                            iPos = InStr(1, sCurrent, "}")
                            'Get the object index and place it in "iObjIndex"
                            iObjIndex = Mid$(sCurrent, 2, iPos - 2)
                            'Separate the control name from the object index
                            sCurrent = Mid$(sCurrent, iPos + 1, Len(sCurrent) - iPos)
                        Else
                            'Othewise, set a flag of "-1" which shows that
                            'the control does not contain objects
                            iObjIndex = -1
                        End If

                        'Go through each control in the Form specified
                        For Each ctl In Frm.Controls
                            'Look if control name in the current line
                            'matches control name in the loop
                            If LCase$(ctl.Name) = LCase$(sCurrent) Then
                                'Get the position of " | " in the value
                                iPos = InStr(1, sValue, " | ")
                                'If there is no " | "
                                If iPos = 0 Then
                                    'All the content of the value is the Caption
                                    sCaption = sValue
                                    'and there is not ToolTip
                                    sToolTip = vbNullString
                                Else    'Otherwise
                                    'Set the left-hand part as the Caption
                                    sCaption = Left$(sValue, iPos - 1)
                                    'Set the right-hand part as the ToolTip
                                    sToolTip = Mid$(sValue, iPos + 3, Len(sValue) - iPos)
                                End If

                                'Get the type of control
                                sCtlType = TypeName(ctl)

                                'If there is no control array
                                If iCtlIndex = -1 Then
                                    Select Case sCtlType
                                        'Set Caption and ToolTip of the following controls
                                    Case "Label", "CommandButton", "OptionButton", "CheckBox", "Frame"
                                        If LenB(sCaption) Then
                                            ctl.Caption = sCaption
                                        End If
                                        If LenB(sToolTip) Then
                                            ctl.ToolTipText = sToolTip
                                        End If
                                        'Set Text and ToolTip of the following controls
                                    Case "TextBox", "ComboBox"
                                        If LenB(sCaption) Then
                                            ctl.Text = sCaption
                                        End If
                                        If LenB(sToolTip) Then
                                            ctl.ToolTipText = sToolTip
                                        End If
                                        'Set just the ToolTip of the following controls
                                    Case "PictureBox", "ListBox", "Slider"
                                        ctl.ToolTipText = sToolTip
                                        'Set just the Caption of the following controls
                                    Case "Menu"
                                        ctl.Caption = sCaption
                                        'If the control is a Toolbar
                                    Case "Toolbar"
                                        'If the Object(here, Button) index is not specified
                                        If iObjIndex = -1 Then
                                            'Set the ToolTip of the Toolbar
                                            ctl.ToolTipText = sToolTip
                                        Else
                                            'Otherwise when the Button index is specified
                                            'Set the Caption and ToolTip of the Button
                                            ctl.Buttons(iObjIndex).Caption = sCaption
                                            ctl.Buttons(iObjIndex).ToolTipText = sToolTip
                                        End If
                                        'If the control is a ListView
                                    Case "ListView"
                                        'If the Object(here, Column-Header) index is not specified
                                        If iObjIndex = -1 Then
                                            'Set the ToolTip of the ListView
                                            ctl.ToolTipText = sToolTip
                                        Else
                                            'Otherwise when the Column-Header index is specified
                                            'Set the Text of the Column-Header
                                            ctl.ColumnHeaders(iObjIndex).Text = sCaption
                                        End If
                                        'If the control is a TabStrip
                                    Case "TabStrip"
                                        'If the Object(here, Tab) index is not specified
                                        If iObjIndex = -1 Then
                                            'Set the ToolTip of the TabStrip
                                            ctl.ToolTipText = sToolTip
                                        Else    'Otherwise when the Tab index is specified
                                            'Set the Caption and ToolTip of the Tab
                                            ctl.Tabs(iObjIndex).Caption = sCaption
                                            ctl.Tabs(iObjIndex).ToolTipText = sToolTip
                                        End If
                                        'If the control is a StatusBar
                                    Case "StatusBar"
                                        'If the Object(here, Panel) index is not specified
                                        If iObjIndex = -1 Then
                                            'Set the SimpleText and ToolTip of the StatusBar
                                            ctl.SimpleText = sCaption
                                            ctl.ToolTipText = sToolTip
                                        Else
                                            'Otherwise when the Panel index is specified
                                            'Set the Text and ToolTip of the Panel
                                            ctl.Panels(iObjIndex).Text = sCaption
                                            ctl.Panels(iObjIndex).ToolTipText = sToolTip
                                        End If
                                    End Select
                                Else    'If there is a control array
                                    'Make sure we are working on the right control
                                    If ctl.Index = iCtlIndex Then

                                        'Same as above

                                        Select Case sCtlType
                                        Case "Label", "CommandButton", "OptionButton", "CheckBox"
                                            If LenB(sCaption) Then
                                                ctl.Caption = sCaption
                                            End If
                                            If LenB(sToolTip) Then
                                                ctl.ToolTipText = sToolTip
                                            End If
                                        Case "TextBox", "ComboBox"
                                            If LenB(sCaption) Then
                                                ctl.Text = sCaption
                                            End If
                                            If LenB(sToolTip) Then
                                                ctl.ToolTipText = sToolTip
                                            End If
                                        Case "PictureBox", "ListBox", "Slider"
                                            ctl.ToolTipText = sToolTip
                                        Case "Menu"
                                            ctl.Caption = sCaption
                                        Case "Toolbar"
                                            If iObjIndex = -1 Then
                                                ctl.ToolTipText = sToolTip
                                            Else
                                                ctl.Buttons(iObjIndex).Caption = sCaption
                                                ctl.Buttons(iObjIndex).ToolTipText = sToolTip
                                            End If
                                        Case "ListView"
                                            If iObjIndex = -1 Then
                                                ctl.ToolTipText = sToolTip
                                            Else
                                                ctl.ColumnHeaders(iObjIndex).Text = sCaption
                                            End If
                                        Case "TabStrip"
                                            If iObjIndex = -1 Then
                                                ctl.ToolTipText = sToolTip
                                            Else
                                                ctl.Tabs(iObjIndex).Caption = sCaption
                                                ctl.Tabs(iObjIndex).ToolTipText = sToolTip
                                            End If
                                        Case "StatusBar"
                                            If iObjIndex = -1 Then
                                                ctl.SimpleText = sCaption
                                                ctl.ToolTipText = sToolTip
                                            Else
                                                ctl.Panels(iObjIndex).Text = sCaption
                                                ctl.Panels(iObjIndex).ToolTipText = sToolTip
                                            End If
                                        End Select
                                    End If
                                End If
                            End If
                        Next ctl
                    End If
                End If
            End If
        End If
    Loop
    Close #1
On Error GoTo 0
End Sub

