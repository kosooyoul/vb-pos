VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmSellList 
   BorderStyle     =   1  '단일 고정
   Caption         =   "판매 내역"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   ControlBox      =   0   'False
   Icon            =   "FrmSellList.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   6735
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '없음
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10095
      Begin VB.Line Line7 
         BorderColor     =   &H0099A8AC&
         X1              =   0
         X2              =   10300
         Y1              =   705
         Y2              =   705
      End
      Begin VB.Label lblPage 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '투명
         Caption         =   "판매 내역"
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
         TabIndex        =   3
         Top             =   120
         Width           =   840
      End
      Begin VB.Label ItemN 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '투명
         Caption         =   "어떤 물품을 언제 얼마나 판매하였는지 확인합니다."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   4065
      End
      Begin VB.Image BGIMG 
         Height          =   705
         Index           =   1
         Left            =   1920
         Top             =   0
         Width           =   4815
      End
   End
   Begin MSDataGridLib.DataGrid SellList 
      Height          =   2310
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   4075
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BorderStyle     =   0
      ColumnHeaders   =   -1  'True
      HeadLines       =   1.5
      RowHeight       =   14
      TabAction       =   1
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "판매코드"
         Caption         =   "No."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "판매일시"
         Caption         =   "판매일시"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "yy-mm-dd   AM/PM hh:mm"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "판매물품"
         Caption         =   "판매물품"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#0 EA "
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "판매수량"
         Caption         =   "판매수량"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#0 EA "
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "판매금액"
         Caption         =   "판매금액"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,### ""\ """
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   555.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1920.189
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   854.929
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1110.047
         EndProperty
      EndProperty
   End
   Begin 편의점_물품관리.isButton B_Close 
      Height          =   360
      Left            =   5160
      TabIndex        =   0
      Top             =   3840
      Width           =   1335
      _extentx        =   2355
      _extenty        =   635
      caption         =   "닫기"
      iconalign       =   1
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   0
      font            =   "FrmSellList.frx":058A
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "- 총 판매 금액 :"
      Height          =   180
      Left            =   240
      TabIndex        =   6
      Top             =   3360
      Width           =   1290
   End
   Begin VB.Label TotalSell 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      Caption         =   "0 \"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   1680
      TabIndex        =   5
      Top             =   3360
      Width           =   315
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   7000
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   0
      X2              =   7000
      Y1              =   3720
      Y2              =   3720
   End
End
Attribute VB_Name = "FrmSellList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()                                                 '판매내역 폼 시작
    Dim slct As String
    Dim TotalCost As Single
    Dim i As Integer
    
    BGIMG(1).Picture = LoadPicture(App.Path & "\Images\BGIMG_1.bmp")    '상단 배경이미지 설정
    
    Set SellConnection = New ADODB.Connection                           '판매내역 데이터베이스 로드
    SellConnection.CursorLocation = adUseClient
    SellConnection.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Database\Library.mdb;Jet OLEDB:Database Password=" & DataBasePassWord & ";"
    slct = "select 판매코드,판매일시,판매물품,판매수량,판매금액 from 판매내역 Order by 판매코드"
    Set SellRecord = New ADODB.Recordset
    SellRecord.Open slct, SellConnection, adOpenStatic, adLockOptimistic
    
    Set SellList.DataSource = SellRecord                                '표에 판매내역 데이터베이스 연결
    SellList.ReBind
    
    TotalCost = 0
    If SellRecord.RecordCount > 0 Then                                  '총판매금액 계산
        For i = 0 To SellRecord.RecordCount - 1
            TotalCost = TotalCost + SellRecord.Fields(4)
            SellRecord.MoveNext
        Next i
        SellRecord.MoveLast
    End If
    
    TotalSell.Caption = Format(TotalCost, "###,##0") & " \"             '총판매금액 출력
End Sub

Private Sub B_Close_Click()                                             '판매내역 폼 닫기
    Unload Me
    SellRecord.Close                                                    '판매내역 데이터베이스 닫기
    SellConnection.Close
End Sub

Private Sub B_Close_MouseEnter()                                        '판매내역 폼 닫기 팁
    ViewTip ("판매한 물품목록을 닫습니다.")
End Sub

Private Sub B_Close_MouseLeave()                                        '팁 지우기
    NoTip
End Sub
