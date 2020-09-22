VERSION 5.00
Begin VB.Form frmWait 
   BorderStyle     =   0  'None
   Caption         =   "ShapedForm"
   ClientHeight    =   3750
   ClientLeft      =   4875
   ClientTop       =   2625
   ClientWidth     =   2250
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   Picture         =   "frmWait.frx":0000
   ScaleHeight     =   3750
   ScaleWidth      =   2250
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, _
                                                       ByVal nCount As Long, _
                                                       ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, _
                                                    ByVal y1 As Long, _
                                                    ByVal X2 As Long, _
                                                    ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, _
                                                 ByVal hSrcRgn1 As Long, _
                                                 ByVal hSrcRgn2 As Long, _
                                                 ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "User32" (ByVal hWnd As Long, _
                                                    ByVal hRgn As Long, _
                                                    ByVal bRedraw As Boolean) As Long
                                                    
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private ResultRegion As Long
Private Function CreateFormRegion(ScaleX As Single, _
                                  ScaleY As Single, _
                                  OffsetX As Integer, _
                                  OffsetY As Integer) As Long
Dim ObjectRegion As Long, nRet As Long, i As Integer
Dim PolyPoints() As POINTAPI
    ResultRegion = CreateRectRgn(0, 0, 0, 0)
    ReDim PolyPoints(0 To 151)
    For i = 0 To 151
        PolyPoints(i).X = GP0X(i) * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX
        PolyPoints(i).Y = GP0Y(i) * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY
    Next i
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 152, 1)
    nRet = CombineRgn(ResultRegion, ObjectRegion, ObjectRegion, 5)
    DeleteObject ObjectRegion
    CreateFormRegion = ResultRegion
End Function
Private Function GP0X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP0X = 0
    Case 1
        GP0X = 149
    Case 2
        GP0X = 150
    Case 3
        GP0X = 150
    Case 4
        GP0X = 149
    Case 5
        GP0X = 149
    Case 6
        GP0X = 148
    Case 7
        GP0X = 148
    Case 8
        GP0X = 147
    Case 9
        GP0X = 146
    Case 10
        GP0X = 144
    Case 11
        GP0X = 143
    Case 12
        GP0X = 141
    Case 13
        GP0X = 140
    Case 14
        GP0X = 138
    Case 15
        GP0X = 138
    Case 16
        GP0X = 135
    Case 17
        GP0X = 135
    Case 18
        GP0X = 121
    Case 19
        GP0X = 120
    Case 20
        GP0X = 116
    Case 21
        GP0X = 115
    Case 22
        GP0X = 112
    Case 23
        GP0X = 111
    Case 24
        GP0X = 108
    Case 25
        GP0X = 107
    Case 26
        GP0X = 104
    Case 27
        GP0X = 103
    Case 28
        GP0X = 100
    Case 29
        GP0X = 99
    Case 30
        GP0X = 93
    Case 31
        GP0X = 92
    Case 32
        GP0X = 91
    Case 33
        GP0X = 91
    Case 34
        GP0X = 87
    Case 35
        GP0X = 87
    Case 36
        GP0X = 85
    Case 37
        GP0X = 84
    Case 38
        GP0X = 83
    Case 39
        GP0X = 82
    Case 40
        GP0X = 82
    Case 41
        GP0X = 83
    Case 42
        GP0X = 83
    Case 43
        GP0X = 84
    Case 44
        GP0X = 85
    Case 45
        GP0X = 93
    Case 46
        GP0X = 94
    Case 47
        GP0X = 96
    Case 48
        GP0X = 99
    Case 49
        GP0X = 101
    Case 50
        GP0X = 104
    Case 51
        GP0X = 106
    Case 52
        GP0X = 109
    Case 53
        GP0X = 111
    Case 54
        GP0X = 114
    Case 55
        GP0X = 116
    Case 56
        GP0X = 117
    Case 57
        GP0X = 119
    Case 58
        GP0X = 120
    Case 59
        GP0X = 122
    Case 60
        GP0X = 123
    Case 61
        GP0X = 135
    Case 62
        GP0X = 135
    Case 63
        GP0X = 138
    Case 64
        GP0X = 138
    Case 65
        GP0X = 140
    Case 66
        GP0X = 145
    Case 67
        GP0X = 146
    Case 68
        GP0X = 147
    Case 69
        GP0X = 148
    Case 70
        GP0X = 149
    Case 71
        GP0X = 150
    Case 72
        GP0X = 150
    Case 73
        GP0X = 149
    Case 74
        GP0X = 1
    Case 75
        GP0X = 0
    Case 76
        GP0X = 0
    Case 77
        GP0X = 1
    Case 78
        GP0X = 2
    Case 79
        GP0X = 2
    Case 80
        GP0X = 3
    Case 81
        GP0X = 3
    Case 82
        GP0X = 4
    Case 83
        GP0X = 4
    Case 84
        GP0X = 5
    Case 85
        GP0X = 6
    Case 86
        GP0X = 7
    Case 87
        GP0X = 11
    Case 88
        GP0X = 12
    Case 89
        GP0X = 14
    Case 90
        GP0X = 14
    Case 91
        GP0X = 16
    Case 92
        GP0X = 16
    Case 93
        GP0X = 21
    Case 94
        GP0X = 21
    Case 95
        GP0X = 25
    Case 96
        GP0X = 26
    Case 97
        GP0X = 31
    Case 98
        GP0X = 32
    Case 99
        GP0X = 35
    Case 100
        GP0X = 38
    Case 101
        GP0X = 40
    Case 102
        GP0X = 41
    Case 103
        GP0X = 43
    Case 104
        GP0X = 46
    Case 105
        GP0X = 48
    Case 106
        GP0X = 51
    Case 107
        GP0X = 53
    Case 108
        GP0X = 54
    Case 109
        GP0X = 57
    Case 110
        GP0X = 58
    Case 111
        GP0X = 64
    Case 112
        GP0X = 64
    Case 113
        GP0X = 66
    Case 114
        GP0X = 67
    Case 115
        GP0X = 68
    Case 116
        GP0X = 67
    Case 117
        GP0X = 67
    Case 118
        GP0X = 66
    Case 119
        GP0X = 66
    Case 120
        GP0X = 65
    Case 121
        GP0X = 65
    Case 122
        GP0X = 63
    Case 123
        GP0X = 63
    Case 124
        GP0X = 56
    Case 125
        GP0X = 55
    Case 126
        GP0X = 52
    Case 127
        GP0X = 51
    Case 128
        GP0X = 49
    Case 129
        GP0X = 48
    Case 130
        GP0X = 46
    Case 131
        GP0X = 45
    Case 132
        GP0X = 43
    Case 133
        GP0X = 40
    Case 134
        GP0X = 38
    Case 135
        GP0X = 37
    Case 136
        GP0X = 34
    Case 137
        GP0X = 33
    Case 138
        GP0X = 30
    Case 139
        GP0X = 29
    Case 140
        GP0X = 16
    Case 141
        GP0X = 16
    Case 142
        GP0X = 14
    Case 143
        GP0X = 14
    Case 144
        GP0X = 12
    Case 145
        GP0X = 6
    Case 146
        GP0X = 5
    Case 147
        GP0X = 4
    Case 148
        GP0X = 3
    Case 149
        GP0X = 2
    Case 150
        GP0X = 1
    Case 151
        GP0X = 0
    End Select
End Function
Private Function GP0Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP0Y = 0
    Case 1
        GP0Y = 0
    Case 2
        GP0Y = 1
    Case 3
        GP0Y = 35
    Case 4
        GP0Y = 36
    Case 5
        GP0Y = 40
    Case 6
        GP0Y = 41
    Case 7
        GP0Y = 44
    Case 8
        GP0Y = 45
    Case 9
        GP0Y = 48
    Case 10
        GP0Y = 53
    Case 11
        GP0Y = 56
    Case 12
        GP0Y = 58
    Case 13
        GP0Y = 61
    Case 14
        GP0Y = 63
    Case 15
        GP0Y = 64
    Case 16
        GP0Y = 67
    Case 17
        GP0Y = 68
    Case 18
        GP0Y = 82
    Case 19
        GP0Y = 82
    Case 20
        GP0Y = 86
    Case 21
        GP0Y = 86
    Case 22
        GP0Y = 89
    Case 23
        GP0Y = 89
    Case 24
        GP0Y = 92
    Case 25
        GP0Y = 92
    Case 26
        GP0Y = 95
    Case 27
        GP0Y = 95
    Case 28
        GP0Y = 98
    Case 29
        GP0Y = 98
    Case 30
        GP0Y = 104
    Case 31
        GP0Y = 104
    Case 32
        GP0Y = 105
    Case 33
        GP0Y = 106
    Case 34
        GP0Y = 110
    Case 35
        GP0Y = 111
    Case 36
        GP0Y = 113
    Case 37
        GP0Y = 116
    Case 38
        GP0Y = 119
    Case 39
        GP0Y = 124
    Case 40
        GP0Y = 128
    Case 41
        GP0Y = 129
    Case 42
        GP0Y = 132
    Case 43
        GP0Y = 133
    Case 44
        GP0Y = 136
    Case 45
        GP0Y = 144
    Case 46
        GP0Y = 144
    Case 47
        GP0Y = 146
    Case 48
        GP0Y = 147
    Case 49
        GP0Y = 149
    Case 50
        GP0Y = 150
    Case 51
        GP0Y = 152
    Case 52
        GP0Y = 153
    Case 53
        GP0Y = 155
    Case 54
        GP0Y = 156
    Case 55
        GP0Y = 158
    Case 56
        GP0Y = 158
    Case 57
        GP0Y = 160
    Case 58
        GP0Y = 160
    Case 59
        GP0Y = 162
    Case 60
        GP0Y = 162
    Case 61
        GP0Y = 174
    Case 62
        GP0Y = 175
    Case 63
        GP0Y = 178
    Case 64
        GP0Y = 179
    Case 65
        GP0Y = 181
    Case 66
        GP0Y = 192
    Case 67
        GP0Y = 195
    Case 68
        GP0Y = 198
    Case 69
        GP0Y = 201
    Case 70
        GP0Y = 206
    Case 71
        GP0Y = 211
    Case 72
        GP0Y = 249
    Case 73
        GP0Y = 250
    Case 74
        GP0Y = 250
    Case 75
        GP0Y = 247
    Case 76
        GP0Y = 227
    Case 77
        GP0Y = 227
    Case 78
        GP0Y = 226
    Case 79
        GP0Y = 214
    Case 80
        GP0Y = 213
    Case 81
        GP0Y = 207
    Case 82
        GP0Y = 206
    Case 83
        GP0Y = 202
    Case 84
        GP0Y = 201
    Case 85
        GP0Y = 196
    Case 86
        GP0Y = 193
    Case 87
        GP0Y = 184
    Case 88
        GP0Y = 181
    Case 89
        GP0Y = 179
    Case 90
        GP0Y = 178
    Case 91
        GP0Y = 176
    Case 92
        GP0Y = 175
    Case 93
        GP0Y = 170
    Case 94
        GP0Y = 169
    Case 95
        GP0Y = 165
    Case 96
        GP0Y = 165
    Case 97
        GP0Y = 160
    Case 98
        GP0Y = 160
    Case 99
        GP0Y = 157
    Case 100
        GP0Y = 156
    Case 101
        GP0Y = 154
    Case 102
        GP0Y = 154
    Case 103
        GP0Y = 152
    Case 104
        GP0Y = 151
    Case 105
        GP0Y = 149
    Case 106
        GP0Y = 148
    Case 107
        GP0Y = 146
    Case 108
        GP0Y = 146
    Case 109
        GP0Y = 143
    Case 110
        GP0Y = 143
    Case 111
        GP0Y = 137
    Case 112
        GP0Y = 136
    Case 113
        GP0Y = 134
    Case 114
        GP0Y = 129
    Case 115
        GP0Y = 124
    Case 116
        GP0Y = 123
    Case 117
        GP0Y = 119
    Case 118
        GP0Y = 118
    Case 119
        GP0Y = 116
    Case 120
        GP0Y = 115
    Case 121
        GP0Y = 114
    Case 122
        GP0Y = 112
    Case 123
        GP0Y = 111
    Case 124
        GP0Y = 104
    Case 125
        GP0Y = 104
    Case 126
        GP0Y = 101
    Case 127
        GP0Y = 101
    Case 128
        GP0Y = 99
    Case 129
        GP0Y = 99
    Case 130
        GP0Y = 97
    Case 131
        GP0Y = 97
    Case 132
        GP0Y = 95
    Case 133
        GP0Y = 94
    Case 134
        GP0Y = 92
    Case 135
        GP0Y = 92
    Case 136
        GP0Y = 89
    Case 137
        GP0Y = 89
    Case 138
        GP0Y = 86
    Case 139
        GP0Y = 86
    Case 140
        GP0Y = 73
    Case 141
        GP0Y = 72
    Case 142
        GP0Y = 70
    Case 143
        GP0Y = 69
    Case 144
        GP0Y = 67
    Case 145
        GP0Y = 54
    Case 146
        GP0Y = 51
    Case 147
        GP0Y = 48
    Case 148
        GP0Y = 43
    Case 149
        GP0Y = 38
    Case 150
        GP0Y = 29
    Case 151
        GP0Y = 18
    End Select
End Function
Private Sub Form_Load()
Dim nRet As Long
    nRet = SetWindowRgn(Me.hWnd, CreateFormRegion(1, 1, 0, 0), True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    DeleteObject ResultRegion
End Sub
