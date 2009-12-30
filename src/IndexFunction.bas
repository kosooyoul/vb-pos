Attribute VB_Name = "IndexFunction"
Option Explicit

Public Declare Function ReleaseCapture Lib "user32" () As Long          '폼이동에 대한 API
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
                                                                        '그림창에 대한 API
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long

Public Const HTCAPTION = 2                                              '폼 이동에 대한 상수
Public Const WM_NCLBUTTONDOWN = &HA1

Public Const RGN_OR = 2
Public Const HWND_TOPMOST = -1                                          '폼 위치 대한 상수
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2

Sub Main()                                                              '프로그램의 시작
    Load Index                                                          '인덱스 로드
    
    SetWindowRgn Index.hwnd, lGetRegion(Index.Pic_Main, RGB(255, 0, 255)), True
    DeleteObject lGetRegion(Index.Pic_Main, RGB(255, 0, 255))
    
    Index.Show                                                          '인덱스 폼 표시
    SetFormPosition Index.hwnd, True                                    '인덱스 폼을 항상 위에 표시
End Sub

Public Sub SetFormPosition(hwnd As Long, TopPosition As Boolean)        '폼 위치 대한 메소드
    If TopPosition Then
         SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
     Else
         SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
     End If
End Sub
                                                                        '그림으로 폼을 표시 : 인덱스 폼
Public Function lGetRegion(pic As PictureBox, lBackColor As Long) As Long
    Dim lRgn As Long
    Dim lSkinRgn As Long
    Dim lStart As Long
    Dim lX As Long, lY As Long
    Dim lHeight As Long, lWidth As Long
    lSkinRgn = CreateRectRgn(0, 0, 0, 0)
    With pic
        lHeight = .Height / Screen.TwipsPerPixelY
        lWidth = .Width / Screen.TwipsPerPixelX
        For lX = 0 To lHeight - 1
            lY = 0
            Do While lY < lWidth
                Do While lY < lWidth And GetPixel(.hdc, lY, lX) = lBackColor
                    lY = lY + 1
                Loop
                If lY < lWidth Then
                    lStart = lY
                    Do While lY < lWidth And GetPixel(.hdc, lY, lX) <> lBackColor
                        lY = lY + 1
                    Loop
                    If lY > lWidth Then lY = lWidth
                    lRgn = CreateRectRgn(lStart, lX, lY, lX + 1)
                    CombineRgn lSkinRgn, lSkinRgn, lRgn, RGN_OR
                    DeleteObject lRgn
                End If
            Loop
        Next
    End With
    lGetRegion = lSkinRgn
End Function
