Attribute VB_Name = "VariableOfPublic"
Option Explicit

Public FunctionMode As Integer                                          '판매:0,손실:1,입고:2
Public TotalCost As Long                                                '판매목록의 총판매액
Public LogDate As String                                                '로그파일날짜
    
Public Const DataBasePassWord = "Library"                               '데이터베이스 암호
Public LossRecord As ADODB.Recordset                                    '손실내역
Public LossConnection As ADODB.Connection
Public ItemRecord As ADODB.Recordset                                    '물품목록
Public ItemConnection As ADODB.Connection
Public SelectRecord As ADODB.Recordset                                  '선택목록
Public SelectConnection As ADODB.Connection
Public SellRecord As ADODB.Recordset                                    '판매내역
Public SellConnection As ADODB.Connection
Public StorageRecord As ADODB.Recordset                                 '입고내역
Public StorageConnection As ADODB.Connection
                                                                        '반투명창에 대한 API
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
                                                                        '웹사이트 이동에 대한 API
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Const GWL_EXSTYLE = (-20)                                        '반투명창에 대한 상수
Public Const LWA_ALPHA = &H2
Public Const WS_EX_LAYERED = &H80000

Public Sub MakeTransparentForm(TheForm As Form, Degree As Long)         '반투명 창
    Dim WStyle As Long
    WStyle = GetWindowLong(TheForm.hwnd, GWL_EXSTYLE)
    WStyle = WStyle Or WS_EX_LAYERED
    SetWindowLong TheForm.hwnd, GWL_EXSTYLE, WStyle
    SetLayeredWindowAttributes TheForm.hwnd, 0, Degree, LWA_ALPHA
End Sub

Public Sub AddLog(Str As String)                                        '로그 저장에 대한 메소드
    Dim LogFileNum As Integer                                           '파일번호
    Dim LogFileName As String                                           '파일이름
    Dim i As Integer

    If LogDate < Date Then                                              '로그날짜 변경확인
         MainForm.LogList.Clear
         LogDate = Date
         MainForm.LogList.AddItem " " & Now & Chr(9) & "새 로그파일 생성>>"
    End If

    MainForm.LogList.AddItem " " & Now & Chr(9) & Str                   '로그 추가
    MainForm.LogList.ListIndex = Abs(MainForm.LogList.ListCount - 1)

    LogFileName = App.Path & "\Log\Log_" & LogDate & ".log"             '로그 파일 이름
    LogFileNum = FreeFile

    Open LogFileName For Output As LogFileNum                           '로그 파일 저장
        For i = 0 To MainForm.LogList.ListCount
            Print #LogFileNum, MainForm.LogList.List(i)
        Next i
    Close LogFileNum
End Sub

Public Sub ViewTip(Str As String)                                       '스태터스바에 팁보이기
    MainForm.Status.Panels(2) = Str
End Sub

Public Sub NoTip()                                                      '스태터스바에 팁지우기
    MainForm.Status.Panels(2) = ""
End Sub

Public Sub SaveConfig()                                                 '설정파일 저장
    Dim ConfigFileNum As Integer
    Dim ConfigFileName As String
    Dim i As Integer

    ConfigFileName = App.Path & "\market.cfg"
    ConfigFileNum = FreeFile                                               '사용가능한 파일번호

    Open ConfigFileName For Output As ConfigFileNum                           '저장 모드로 파일을 읽음
        Print #ConfigFileNum, "[w]" & MainForm.Width
        Print #ConfigFileNum, "[h]" & MainForm.Height
        Print #ConfigFileNum, "[n]" & MainForm.Status.Panels(1).Text
        Print #ConfigFileNum, "[1]" & CBool(MainForm.S_ShowLog.Checked)
        Print #ConfigFileNum, "[2]" & CBool(MainForm.S_ShowStatusBar.Checked)
    Close ConfigFileNum
End Sub
