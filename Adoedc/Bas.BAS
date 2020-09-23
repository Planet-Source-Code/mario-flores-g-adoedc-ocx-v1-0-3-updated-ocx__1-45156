Attribute VB_Name = "Bas"
Option Explicit

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" ( _
                                             ByVal dwExStyle As Long, _
                                             ByVal lpClassName As String, _
                                             ByVal lpWindowName As String, _
                                             ByVal dwStyle As Long, _
                                             ByVal x As Long, _
                                             ByVal y As Long, _
                                             ByVal nWidth As Long, _
                                             ByVal nHeight As Long, _
                                             ByVal hWndParent As Long, _
                                             ByVal hMenu As Long, _
                                             ByVal hInstance As Long, _
                                             lpParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
                                             ByVal hwnd As Long, _
                                             ByVal wMsg As Long, _
                                             ByVal wParam As Long, _
                                             lParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
                                             ByVal hWndInsertAfter As Long, _
                                             ByVal x As Long, ByVal y As Long, _
                                             ByVal cx As Long, _
                                             ByVal cy As Long, _
                                             ByVal wFlags As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
                                             ByVal x As Long, _
                                             ByVal y As Long, _
                                             ByVal nWidth As Long, _
                                             ByVal nHeight As Long, _
                                             ByVal hSrcDC As Long, _
                                             ByVal xSrc As Long, _
                                             ByVal ySrc As Long, _
                                             ByVal dwRop As Long) As Long


Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type


Private Type TOOLINFO
    lSize As Long
    lFlags As Long
    lHwnd As Long
    lId As Long
    lpRect As RECT
    hInstance As Long
    lpStr As String
    lParam As Long
End Type

Public Indice As Integer

Public Enum ttIconType
    TTNoIcon = 0
    TTIconInfo = 1
    TTIconWarning = 2
    TTIconError = 3
End Enum

Public Enum ttStyleEnum
    TTStandard
    TTBalloon
End Enum

Private lHwnd As Long
Private Tooltip As TOOLINFO

Public Sub Create_Tooltip(Style As ttStyleEnum, Destiny As Object, text As String, Optional Icono As ttIconType, Optional Title As String, Optional Background = Null, Optional Foreground = Null)
    'Crea un tooltip:
    
    Dim lWinStyle As Long
    Dim Region As RECT
    
    lWinStyle = &H1 Or &H2
    
    If Style = TTBalloon Then lWinStyle = lWinStyle + &H40
    
    DestroyWindow lHwnd
    
    
    lHwnd = CreateWindowEx(0&, "tooltips_class32", vbNullString, lWinStyle, 0&, 0&, 0&, 0&, Destiny.hwnd, 0&, App.hInstance, 0&)
    SetWindowPos lHwnd, -1, 0&, 0&, 0&, 0&, &H10 Or &H1 Or &H2
    GetClientRect Destiny.hwnd, Region
    
    With Tooltip
        .lFlags = &H10
        .lHwnd = Destiny.hwnd
        .lId = 0
        .hInstance = App.hInstance
        .lpRect = Region
        .lpStr = text
    End With
    
     SendMessage lHwnd, 1028, 0&, Tooltip
     If Icono <> TTNoIcon Or Title <> vbNullString Then SendMessage lHwnd, 1056, CLng(Icono), ByVal Title
     If Not IsNull(Foreground) Then SendMessage lHwnd, 1044, vbBlack, 0&
     If Not IsNull(Background) Then SendMessage lHwnd, 1043, vbWhite, 0&
    
End Sub

Public Sub Destroy_Tooltip()
        DestroyWindow lHwnd
End Sub

Public Function TextFormatPath(Path As String, Val As Integer) As String
Dim Temp, Temp2, i

For i = Len(Path) To 0 Step -1
    Temp = VBA.Right(VBA.Left(Path, i), 1): If Temp = "\" Then Exit For
    Temp2 = Temp & Temp2
Next

If Val = 1 Then TextFormatPath = Temp2                              '1 Actual Name of The File
If Val = 2 Then TextFormatPath = VBA.Left(Path, Len(Path) - Len(Temp2)) '2 Direction or Folder where File is Located
End Function
