VERSION 5.00
Begin VB.Form FrmExactCalculation 
   BorderStyle     =   1  '단일 고정
   Caption         =   "정산"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   ControlBox      =   0   'False
   Icon            =   "FrmExactCalculation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   4815
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '없음
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.Line Line4 
         BorderColor     =   &H0099A8AC&
         X1              =   -240
         X2              =   7200
         Y1              =   705
         Y2              =   705
      End
      Begin VB.Label MsgLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '투명
         Caption         =   "정산 처리"
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
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   840
      End
      Begin VB.Image BGIMG 
         Height          =   705
         Index           =   1
         Left            =   0
         Top             =   0
         Width           =   4815
      End
   End
   Begin 편의점_물품관리.isButton B_Close 
      Height          =   360
      Left            =   3360
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
      _ExtentX        =   2143
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
      X2              =   6690
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label ExactDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "2000년 1월 1일 정산"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   480
      TabIndex        =   5
      Top             =   960
      Width           =   1620
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "- 총 판매 금액 :"
      Height          =   180
      Left            =   360
      TabIndex        =   4
      Top             =   1440
      Width           =   1290
   End
   Begin VB.Label ResultProfit 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "0 \"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   1800
      TabIndex        =   3
      Top             =   1440
      Width           =   315
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   0
      X2              =   6720
      Y1              =   1920
      Y2              =   1920
   End
End
Attribute VB_Name = "FrmExactCalculation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()                                                 '정산 폼 시작
    Dim Str As String
    Dim i As Integer
    Dim TotalProfit As Long                                             '총수익
    
    BGIMG(1).Picture = LoadPicture(App.Path & "\Images\BGIMG_1.bmp")    '상단 배경이미지 설정
    
    Set SellConnection = New ADODB.Connection                           '판매내역 데이터베이스 로드
    SellConnection.CursorLocation = adUseClient
    SellConnection.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Database\Library.mdb;Jet OLEDB:Database Password=" & DataBasePassWord & ";"
    Str = "select 판매코드,판매일시,판매물품,판매수량,판매금액 from 판매내역 Order by 판매코드"
    Set SellRecord = New ADODB.Recordset
    SellRecord.Open Str, SellConnection, adOpenStatic, adLockOptimistic
    
    TotalProfit = 0
    For i = 0 To SellRecord.RecordCount - 1                             '총수익 계산
        TotalProfit = TotalProfit + SellRecord.Fields(4)
        SellRecord.MoveNext
    Next i
              
    Str = "DELETE FROM 판매내역"                                        'SQL문 : 판매목록 초기화
    SellConnection.Execute Str
                                                                        '정산날짜 출력
    ExactDate.Caption = Year(Now) & "년 " & Month(Now) & "월 " & Day(Now) & "일 정산처리 결과"
    ResultProfit.Caption = Format(TotalProfit, "###,##0") & " \"        '총 수익 출력
                                                                        '정산 로그
    AddLog ("정산 >>     수익(" & TotalProfit & " \), 정산시각(" & Now & ")")
End Sub

Private Sub B_Close_Click()                                             '정산 폼 닫기
    SellRecord.Close                                                    '판매내역 데이터베이스 닫기
    SellConnection.Close
    Unload Me
End Sub

