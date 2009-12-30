VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmLoss 
   BorderStyle     =   1  '단일 고정
   Caption         =   "손실 등록"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   ControlBox      =   0   'False
   Icon            =   "FrmLoss.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   6735
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '없음
      Height          =   735
      Left            =   0
      TabIndex        =   4
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
         Caption         =   "손실된 물품을 선택하여 손실 등록"
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
         TabIndex        =   6
         Top             =   120
         Width           =   2970
      End
      Begin VB.Label ItemN 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '투명
         Caption         =   "파손되거나 분실된 물품을 선택하여 손실등록 해주세요."
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
         TabIndex        =   5
         Top             =   360
         Width           =   4425
      End
      Begin VB.Image BGIMG 
         Height          =   705
         Index           =   1
         Left            =   1920
         Top             =   0
         Width           =   4815
      End
   End
   Begin 편의점_물품관리.isButton B_Del 
      Height          =   360
      Left            =   1680
      TabIndex        =   1
      Top             =   3720
      Width           =   1335
      _extentx        =   2355
      _extenty        =   635
      caption         =   "물품 취소"
      iconalign       =   1
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   0
      font            =   "FrmLoss.frx":058A
   End
   Begin 편의점_물품관리.isButton B_Add 
      Height          =   360
      Left            =   240
      TabIndex        =   0
      Top             =   3720
      Width           =   1335
      _extentx        =   2355
      _extenty        =   635
      caption         =   "물품 선택"
      iconalign       =   1
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   0
      font            =   "FrmLoss.frx":05B2
   End
   Begin 편의점_물품관리.isButton B_Loss 
      Height          =   360
      Left            =   3120
      TabIndex        =   2
      Top             =   3720
      Width           =   1335
      _extentx        =   2355
      _extenty        =   635
      caption         =   "손실등록"
      iconalign       =   1
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   0
      font            =   "FrmLoss.frx":05DA
   End
   Begin 편의점_물품관리.isButton B_Close 
      Height          =   360
      Left            =   5160
      TabIndex        =   3
      Top             =   3720
      Width           =   1335
      _extentx        =   2355
      _extenty        =   635
      caption         =   "닫기"
      iconalign       =   1
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   0
      font            =   "FrmLoss.frx":0602
   End
   Begin MSDataGridLib.DataGrid Sell 
      Height          =   2310
      Left            =   120
      TabIndex        =   9
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "선택물품"
         Caption         =   "손실물품"
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
      BeginProperty Column01 
         DataField       =   "선택수량"
         Caption         =   "손실수량"
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
      BeginProperty Column02 
         DataField       =   "선택금액"
         Caption         =   "손실금액"
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
            ColumnWidth     =   3105.071
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   1440
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1395.213
         EndProperty
      EndProperty
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   9720
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "- 총 손실 금액 :"
      Height          =   180
      Left            =   240
      TabIndex        =   8
      Top             =   3240
      Width           =   1290
   End
   Begin VB.Label TotalSellCost 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      Caption         =   "0 \"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   1680
      TabIndex        =   7
      Top             =   3240
      Width           =   315
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   0
      X2              =   9720
      Y1              =   3600
      Y2              =   3600
   End
End
Attribute VB_Name = "FrmLoss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()                                                 '손실등록 폼 시작
    Dim slct As String

    BGIMG(1).Picture = LoadPicture(App.Path & "\Images\BGIMG_1.bmp")    '상단 배경이미지 설정
    
    Set SelectConnection = New ADODB.Connection                         '선택목록 데이터베이스 로드
        SelectConnection.CursorLocation = adUseClient
        SelectConnection.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Database\Library.mdb;Jet OLEDB:Database Password=" & DataBasePassWord & ";"
        slct = "select 선택코드,선택물품,선택수량,선택금액 from 선택목록 Order by 선택코드"
    Set SelectRecord = New ADODB.Recordset
        SelectRecord.Open slct, SelectConnection, adOpenStatic, adLockOptimistic
    
    Set Sell.DataSource = SelectRecord                                  '표에 선택목록 데이터베이스 연결
    Sell.ReBind

    TotalSellCost.Caption = Format(TotalCost, "###,##0") & " \"         '총손실액 출력
End Sub

Private Sub B_Loss_Click()                                              '손실등록시
    Dim slct As String
    Dim Str As String
    Dim i As Integer, j As Integer
    
    If SelectRecord.RecordCount = 0 Then                                '손실 등록할 물품이 없는 경우
        MsgBox "손실 등록할 물품이 없습니다. 등록할 물품을 추가하십시요.", vbOKOnly, "손실 등록할 물품 없음"
    Else                                                                '손실 등록할 물품이 있는 경우
        Set LossConnection = New ADODB.Connection                       '손실내역 데이터베이스 로드
        LossConnection.CursorLocation = adUseClient
        LossConnection.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Database\Library.mdb;Jet OLEDB:Database Password=" & DataBasePassWord & ";"
        slct = "select 손실코드,손실일시,손실물품,손실수량,손실금액 from 손실내역 Order by 손실코드"
        Set LossRecord = New ADODB.Recordset
        
        SelectRecord.MoveFirst
        For i = 0 To SelectRecord.RecordCount - 1                       '선택물품 수 만큼 반복
            Str = "INSERT INTO 손실내역"                                'SQL문 : 손실내역에 레코드 추가
            Str = Str & "(손실물품,손실수량,손실금액) "
            Str = Str & "VALUES('" & SelectRecord.Fields(1) & "', "         '손실물품
            Str = Str & "'" & Val(SelectRecord.Fields(2)) & "', "           '손실수량
            Str = Str & "'" & Val(SelectRecord.Fields(3)) & "')"            '손실금액
            LossConnection.Execute (Str)
                                                                        '손실등록 로그
            AddLog ("손실등록 >>     손실물품(" & SelectRecord.Fields(1) & "), 손실수량(" & SelectRecord.Fields(2) & " EA), 손실금액(" & SelectRecord.Fields(3) & " \)")
                    
            ItemRecord.MoveFirst                                        '물품 수량 감소
            For j = 0 To ItemRecord.RecordCount - 1
                If SelectRecord.Fields(1) = ItemRecord.Fields(1) Then
                    Str = "UPDATE 물품목록 SET "                        'SQL문 : 물품목록 레코드 수량 수정
                    Str = Str & "물품수량=" & Val(ItemRecord.Fields(3) - SelectRecord.Fields(2))
                    Str = Str & " WHERE 물품이름='" & ItemRecord.Fields(1) & "'"
                    ItemConnection.Execute Str
                End If
                ItemRecord.MoveNext
            Next j
            ItemRecord.MoveFirst
        
            SelectRecord.Delete adAffectCurrent                         '손실 등록된 물품은 선택목록에서 삭제
            SelectRecord.MoveFirst
        Next i

        TotalCost = 0                                                   '선택한 물품들의 총손실금액 초기화
        TotalSellCost.Caption = "0 \"
        
        MsgBox "손실등록이 완료되었습니다.", vbOKOnly, "손실 등록 완료"
        FrmLossList.Show vbModal                                        '손실내역 폼 표시
    End If
End Sub

Private Sub B_Add_Click()                                               '손실등록할 물품 추가 폼 표시
    FrmSelectSell.Left = Me.Left                                            '물품선택 폼 표시
    FrmSelectSell.Top = Me.Top
    FrmSelectSell.Show
    Unload Me
End Sub

Private Sub B_Close_Click()                                             '손실등록 폼 닫기
    Dim i As Integer
    TotalCost = 0
    For i = 0 To SelectRecord.RecordCount - 1                           '선택목록 초기화
        SelectRecord.Delete adAffectCurrent
        SelectRecord.MoveNext
    Next i
   
    SelectRecord.Close                                                  '선택목록 데이터베이스 닫기
    SelectConnection.Close
    MainForm.Enabled = True
    Unload Me
End Sub

Private Sub B_Del_Click()                                               '선택목록에서 항목 삭제
On Error GoTo BrankTable
    TotalCost = TotalCost - SelectRecord.Fields(3)                      '총손실금액 수정
    TotalSellCost.Caption = Format(TotalCost, "###,##0") & " \"
    SelectRecord.Delete adAffectCurrent
    Sell.ReBind
BrankTable:
    If Err.Number = 3021 Then                                           '삭제할 품목이 없는 경우
        MsgBox "취소할 물품이 없습니다.", vbOKOnly, "취소할 물품 없음"
    End If
End Sub

Private Sub B_Add_MouseEnter()                                          '물품추가 팁
    ViewTip ("손실 등록할 물품을 추가합니다.")
End Sub

Private Sub B_Add_MouseLeave()                                          '팁 지우기
    NoTip
End Sub
    
Private Sub B_Close_MouseEnter()                                        '손실등록 폼 닫기
    ViewTip ("손실 등록 창을 닫습니다.")
End Sub

Private Sub B_Close_MouseLeave()                                        '팁 지우기
    NoTip
End Sub

Private Sub B_Del_MouseEnter()                                          '선택한 물품 표에서 삭제
    ViewTip ("선택한 물품을 취소합니다.")
End Sub

Private Sub B_Del_MouseLeave()                                          '팁 지우기
    NoTip
End Sub

Private Sub B_Loss_MouseEnter()                                         '손실등록 팁
    ViewTip ("선택한 물품 목록의 물품들을 모두 손실등록합니다.")
End Sub

Private Sub B_Loss_MouseLeave()                                         '팁 지우기
    NoTip
End Sub
