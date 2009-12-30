VERSION 5.00
Begin VB.Form Index 
   BorderStyle     =   0  '없음
   Caption         =   "편의점 물품 관리 - 시작"
   ClientHeight    =   2160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3345
   Icon            =   "Index.frx":0000
   LinkTopic       =   "Index"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   3345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.Timer Timer2 
      Interval        =   10000
      Left            =   120
      Top             =   120
   End
   Begin VB.PictureBox Pic_Main 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  '없음
      Height          =   2160
      Left            =   0
      Picture         =   "Index.frx":08CA
      ScaleHeight     =   2160
      ScaleWidth      =   3345
      TabIndex        =   0
      Top             =   0
      Width           =   3345
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   120
         Top             =   120
      End
      Begin VB.Label B_ENTER 
         BackStyle       =   0  '투명
         Height          =   420
         Left            =   1440
         MouseIcon       =   "Index.frx":1830C
         MousePointer    =   99  '사용자 정의
         TabIndex        =   2
         Top             =   795
         Width           =   770
      End
      Begin VB.Label URL_Ahyane 
         BackStyle       =   0  '투명
         Height          =   255
         Left            =   0
         MouseIcon       =   "Index.frx":18BD6
         MousePointer    =   99  '사용자 정의
         TabIndex        =   1
         Top             =   1680
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Index"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'사용된 코드 : 최상위폼, 반투명화폼, 그림표시폼, 텍스트입출력, 데이터베이스관리
'주석의 위치 : 각 행의 73열(Tab * 18)부터 표시 ("'로그" 제외)
Option Explicit
Dim Alpha As Long                                                       '불투명도

Private Sub Form_Load()
    Me.Pic_Main.Move 0, 0
    Alpha = 200
    MakeTransparentForm Me, Alpha
    
    TotalCost = 0                                                       '판매가의 초기화
End Sub

Private Sub B_ENTER_Click()
    Dim LogFileNum As Integer                                           '파일번호
    Dim Logs As String                                                  '파일내용
    Dim LogFileName As String                                           '로그파일주소
    Dim i As Integer
    Dim Str As String

    B_ENTER.Enabled = False
    Timer1.Enabled = True
    Timer2.Enabled = False
    
    MainForm.Show                                                       '메인 화면 표시

On Error GoTo LogFileNotFound                                           '로그파일이 없는 경우 LogFileNotFound로 이동
    LogDate = Date                                                      '로그파일 날짜
    LogFileName = App.Path & "\Log\Log_" & LogDate & ".log"             '로그파일 경로
    LogFileNum = FreeFile

    Open LogFileName For Input As LogFileNum                            '파일을 읽기 전용으로 열기
        Do Until EOF(LogFileNum)
            Line Input #LogFileNum, Logs                                '로그를 한줄씩 읽음
            If Logs <> "" Then
                MainForm.LogList.AddItem (Logs)
            End If
        Loop
    Close LogFileNum
                                                                        '정상종료 했는지 체크
    If Right(MainForm.LogList.List(MainForm.LogList.ListCount - 1), 3) = "End" Then
        AddLog ("프로그램 시작 >> Program Start")                       '프로그램 시작 로그
    Else
        Set SelectConnection = New ADODB.Connection                     '선택목록초기화를 위해 데이터베이스 로드
        SelectConnection.CursorLocation = adUseClient
        SelectConnection.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Database\Library.mdb;Jet OLEDB:Database Password=" & DataBasePassWord & ";"
        Str = "DELETE FROM 선택목록"                                    'SQL문 : 선택목록 초기화
        SelectConnection.Execute Str
        SelectRecord.Close                                              '선택목록 데이터베이스 닫기
        SelectConnection.Close
                
        AddLog ("프로그램 비정상적인 종료 >>")                          '비정상종료시 로그
        AddLog ("프로그램 재시작 >> Program Start")
    End If
    
LogFileNotFound:
    If Err.Number = 53 Then                                             '로그파일이 없는 경우 새로 생성
        AddLog ("새 로그파일 생성>>")
        AddLog ("프로그램 시작 >> Program Start")
    End If
End Sub

Private Sub Pic_Main_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
        Case vbLeftButton                                               '인덱스 폼 이동
            Call ReleaseCapture
            Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End Select
End Sub

Private Sub Timer1_Timer()
    Alpha = Alpha - 3                                                   '불투명도 점점 감소
    MakeTransparentForm Me, Alpha
    If Alpha <= 10 Then Unload Me
End Sub

Private Sub Timer2_Timer()
    Call B_ENTER_Click
End Sub

Private Sub URL_Ahyane_Click()
    ShellExecute 0, "open", "http://www.ahyane.net", "", "", 10
End Sub
