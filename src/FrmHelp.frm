VERSION 5.00
Begin VB.Form FrmHelp 
   BorderStyle     =   1  '단일 고정
   Caption         =   "도움말"
   ClientHeight    =   5895
   ClientLeft      =   1365
   ClientTop       =   660
   ClientWidth     =   6375
   Icon            =   "FrmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   6375
   StartUpPosition =   1  '소유자 가운데
   Begin VB.ListBox KeywordList 
      Height          =   3840
      Left            =   45
      TabIndex        =   0
      Top             =   1215
      Width           =   1935
   End
   Begin VB.TextBox HelpText 
      Height          =   3840
      Left            =   2040
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   1
      Top             =   1215
      Width           =   4260
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '없음
      Height          =   735
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8415
      Begin VB.Image Image1 
         Height          =   495
         Left            =   120
         Top             =   120
         Width           =   615
      End
      Begin VB.Label MsgLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '투명
         Caption         =   "편의점 물품 관리 프로그램 도움말"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   840
         TabIndex        =   4
         Top             =   120
         Width           =   2970
      End
      Begin VB.Line Line4 
         BorderColor     =   &H0099A8AC&
         X1              =   -240
         X2              =   8400
         Y1              =   705
         Y2              =   705
      End
      Begin VB.Image BGIMG 
         Height          =   705
         Index           =   1
         Left            =   1560
         Top             =   0
         Width           =   4815
      End
   End
   Begin 편의점_물품관리.isButton B_Close 
      Height          =   360
      Left            =   4680
      TabIndex        =   2
      Top             =   5280
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   635
      Caption         =   "닫기"
      IconAlign       =   1
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   9720
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "목차 :"
      Height          =   180
      Left            =   120
      TabIndex        =   5
      Top             =   945
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   0
      X2              =   9720
      Y1              =   5160
      Y2              =   5160
   End
End
Attribute VB_Name = "FrmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()                                                 '도움말 폼 시작
    Dim LogFileNum As Integer
    Dim LogFileName As String
    Dim Logs As String
    Dim strText As String                                               '도움말 내용을 읽기 위한 버퍼
    Dim i As Integer
    Image1.Picture = MainForm.ImageList1.ListImages(6).Picture
    BGIMG(1).Picture = LoadPicture(App.Path & "\Images\BGIMG_1.bmp")    '상단 배경이미지 설정
    
On Error GoTo NotFound                                                  '도움말 파일 찾지 못한 경우 NotFound로 이동

    LogFileName = App.Path & "\MarketMng.hlp"                           '도움말 파일 경로
    LogFileNum = FreeFile

    Open LogFileName For Input As LogFileNum                            '도움말 파일 읽기
        Do Until EOF(LogFileNum)
            Line Input #LogFileNum, Logs
                                                                            
            If Right(Logs, 5) = "[key]" Then                            '도움말의 키워드 부분
                Logs = Left(Logs, Len(Logs) - 5)
                KeywordList.AddItem Logs
            End If
            strText = strText + Logs + vbCrLf
            
        Loop
        HelpText.Text = strText
        HelpText.Locked = True
    Close LogFileNum
    
NotFound:                                                               '도움말 파일을 찾지 못한 경우
    If Err.Number = 53 Then
        MsgBox ("MarketMng.hlp 파일을 찾을 수 없습니다.")
    End If
End Sub

Private Sub B_Close_Click()                                             '도움말 폼 닫기
    Unload Me
End Sub

Private Sub KeywordList_Click()                                         '선택한 키워드로 도움말 내용 검색
    Dim i As Integer
    Dim KeyLen As Integer                                               '선택/입력한 검색어 길이
    Dim TextTemp As String
    Dim Keyword As String
    
    Keyword = KeywordList.List(KeywordList.ListIndex)                   '검색어
    TextTemp = HelpText.Text                                            '빠른 도움말 검색을 위한 임시공간
    KeyLen = Len(Keyword)
    
    For i = 1 To Len(HelpText.Text)                                     '도움말 내용에서 검색
        If Mid(HelpText, i, KeyLen) = Keyword Then                      '검색 성공시 서브루틴 종료
            HelpText.SelStart = Len(HelpText.Text)
            HelpText.SelStart = i - 1
            HelpText.SelLength = KeyLen
            HelpText.SetFocus
            Exit Sub
        End If
    Next i
    MsgBox """" & Keyword & """을(를) 찾을 수 없습니다."           '검색 실패시
    HelpText.SelStart = 0
End Sub
