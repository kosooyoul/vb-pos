VERSION 5.00
Begin VB.Form FrmSelectSell 
   BorderStyle     =   1  '단일 고정
   Caption         =   "물품 선택"
   ClientHeight    =   4695
   ClientLeft      =   4965
   ClientTop       =   2955
   ClientWidth     =   6255
   ControlBox      =   0   'False
   Icon            =   "FrmSelectSell.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   6255
   Begin VB.ListBox ItemList 
      Height          =   2580
      ItemData        =   "FrmSelectSell.frx":014A
      Left            =   240
      List            =   "FrmSelectSell.frx":014C
      TabIndex        =   0
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '없음
      Height          =   735
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   10095
      Begin VB.Label ItemN 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '투명
         Caption         =   "왼쪽 물품목록에서 물품을 선택하고 수량을 입력하십시요."
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
         TabIndex        =   13
         Top             =   360
         Width           =   4605
      End
      Begin VB.Label lblPage 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '투명
         Caption         =   "물품 선택"
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
         TabIndex        =   12
         Top             =   120
         Width           =   840
      End
      Begin VB.Line Line7 
         BorderColor     =   &H0099A8AC&
         X1              =   0
         X2              =   10300
         Y1              =   705
         Y2              =   705
      End
      Begin VB.Image BGIMG 
         Height          =   705
         Index           =   1
         Left            =   1440
         Top             =   0
         Width           =   4815
      End
   End
   Begin VB.TextBox SellQuantity 
      Height          =   270
      Left            =   4680
      MaxLength       =   5
      TabIndex        =   1
      Text            =   "0"
      Top             =   2955
      Width           =   1335
   End
   Begin 편의점_물품관리.isButton B_Submit 
      Height          =   360
      Left            =   2760
      TabIndex        =   2
      Top             =   4080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   635
      Caption         =   "물품 추가"
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
   Begin 편의점_물품관리.isButton B_Cancel 
      Height          =   360
      Left            =   4440
      TabIndex        =   3
      Top             =   4080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   635
      Caption         =   "취소"
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
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "물품 목록 :"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   240
      TabIndex        =   10
      Top             =   960
      Width           =   900
   End
   Begin VB.Label SellCost 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      Caption         =   "0 \"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   4680
      TabIndex        =   9
      Top             =   3360
      Width           =   315
   End
   Begin VB.Label Cost 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      Caption         =   "0 \"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   4680
      TabIndex        =   8
      Top             =   2400
      Width           =   315
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   $"FrmSelectSell.frx":014E
      Height          =   540
      Left            =   3360
      TabIndex        =   7
      Top             =   3000
      Width           =   1230
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   $"FrmSelectSell.frx":0174
      Height          =   900
      Left            =   3360
      TabIndex        =   6
      Top             =   1680
      Width           =   1230
   End
   Begin VB.Label Quantity 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      Caption         =   "0 EA"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   4680
      TabIndex        =   5
      Top             =   2040
      Width           =   390
   End
   Begin VB.Label ItemName 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      Caption         =   "None"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   4680
      TabIndex        =   4
      Top             =   1680
      Width           =   450
   End
   Begin VB.Line Line13 
      BorderColor     =   &H80000014&
      X1              =   3120
      X2              =   6120
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   -120
      X2              =   9600
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   -120
      X2              =   9600
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line12 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   3120
      X2              =   6120
      Y1              =   2760
      Y2              =   2760
   End
End
Attribute VB_Name = "FrmSelectSell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()                                                 '물품선택 폼 시작
    Dim slct As String

    BGIMG(1).Picture = LoadPicture(App.Path & "\Images\BGIMG_1.bmp")    '상단 배경이미지 설정
    
    Set ItemConnection = New ADODB.Connection                           '물품목록 데이터베이스 읽음
    ItemConnection.CursorLocation = adUseClient
    ItemConnection.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Database\Library.mdb;Jet OLEDB:Database Password=" & DataBasePassWord & ";"
    slct = "select 물품코드,물품이름,물품단가,물품수량 from 물품목록 Order by 물품코드"
    Set ItemRecord = New ADODB.Recordset
    ItemRecord.Open slct, ItemConnection, adOpenStatic, adLockOptimistic
    
    Set ItemList.DataSource = ItemRecord                                '리스트박스에 물품목록 데이터베이스 연결
    Call SetItemList
End Sub

Private Sub B_Submit_Click()                                            '물품추가 선택
    Dim Str As String
    Dim MsgResult As VbMsgBoxResult
    Dim i As Integer
    Dim ResultCost As Single
    
    If ItemName.Caption = "None" Then                                   '물품선택 안한 경우
        MsgBox "물품을 선택하지 않았습니다.", vbOKOnly, "물품 선택"
    ElseIf Val(SellQuantity.Text) = 0 Then                              '수량을 0으로한 경우
        MsgBox "0 개의 물품을 선택할 수 없습니다.", vbOKOnly, "수량 입력"
    Else                                                                '물품추가부분
        On Error GoTo OverLap
        ResultCost = Val(SellQuantity.Text) * Val(Cost.Caption)
        Str = "INSERT INTO 선택목록"                                    'SQL문 : 선택목록에 레코드 추가
        Str = Str & "(선택물품,선택수량,선택금액) "
        Str = Str & "VALUES('" & ItemName.Caption & "', "                   '선택물품
        Str = Str & "'" & Val(SellQuantity.Text) & "', "                    '선택수량
        Str = Str & "'" & ResultCost & "')"                                 '선택금액
        SelectConnection.Execute (Str)
        TotalCost = TotalCost + ResultCost                              '총금액 증가
        
OverLap:                                                                '선택한 물품이 이미 있는 경우
        If Err.Number = -2147467259 Then
            MsgResult = MsgBox("현재 선택한 물품은 이미 목록에 등록되어있습니다." & Chr(13) & "물품의 수량을 변경하시겠습니까?", vbYesNo, "선택 중복")
            If MsgResult = vbNo Then                                    '수량 수정안할시
                Exit Sub                                                    '서브루틴종료
            Else                                                        '수량 수정
                SelectRecord.MoveFirst
                For i = 0 To SelectRecord.RecordCount - 1
                    If Trim(SelectRecord.Fields(1)) = Trim(ItemName.Caption) Then
                        TotalCost = TotalCost - SelectRecord.Fields(3)  '총금액 계산
                        Exit For
                    End If
                    SelectRecord.MoveNext
                Next i
                
                Str = "UPDATE 선택목록 SET "                            'SQL문 : 선택목록의 물품을 수정
                Str = Str & "선택수량='" & SellQuantity.Text & "',"         '물품수량
                Str = Str & "선택금액=" & ResultCost                        '물품금액
                Str = Str & " WHERE 선택물품='" & ItemName.Caption & "'"    '선택물품으로 식별
                SelectConnection.Execute (Str)
                
                TotalCost = TotalCost + ResultCost                      '총 금액변경
            End If
        End If
        
        Call ReturnForm                                                 '이전 폼으로 복귀
    End If
End Sub

Private Sub B_Cancel_Click()                                            '취소
    Call ReturnForm                                                     '이전 폼으로 복귀
End Sub

Private Sub ItemList_Click()                                            '목록에서 물품선택
    Dim i As Integer
    ItemRecord.MoveFirst
    For i = 0 To ItemRecord.RecordCount - 1                             '선택물품을 검색하여 정보를 출력
        If Trim(ItemList.List(ItemList.ListIndex)) = ItemRecord.Fields(1) Then
            ItemName.Caption = Trim(ItemList.List(ItemList.ListIndex))
            Quantity.Caption = ItemRecord.Fields(3) & " EA"                 '물품수량출력
            Cost.Caption = ItemRecord.Fields(2) & " \"                      '물품단가출력
            If Val(SellQuantity.Text) > Val(Quantity.Caption) Then
                SellQuantity.Text = Quantity.Caption
            End If
                                                                            '선택물품의 금액 계산(물품단가*입력수량)
            SellCost.Caption = Format(Val(SellQuantity.Text) * Val(Cost.Caption), "###,##0") & " \"
            Exit For
        End If
    ItemRecord.MoveNext
    Next i
End Sub

Private Sub SellQuantity_Change()                                       '수량입력시
    SellQuantity.Text = Abs(Val(SellQuantity.Text))                         '자연수로 변환
    If FunctionMode = 2 Then
        If 30000 - Val(Quantity.Caption) < Val(SellQuantity.Text) Then      '선택수량제한 : 최대 30000 - 남은 수량
            SellQuantity.Text = 30000 - Val(Quantity.Caption)
        End If

    Else
        If Val(Quantity.Caption) < Val(SellQuantity.Text) Then              '수량제한 - 판매/손실
            SellQuantity.Text = Val(Quantity.Caption)
        End If
    End If                                                                  '선택물품의 금액 계산(물품단가*입력수량)
    SellCost.Caption = Format(Val(SellQuantity.Text) * Val(Cost.Caption), "###,##0") & " \"
    
    If SellQuantity.Text = 0 Then                                       '입력수량이 0이면 텍스트블럭
        Call SellQuantity_GotFocus
    End If
End Sub

Private Sub SellQuantity_GotFocus()
    SellQuantity.SelStart = 0                                           '편의상 텍스트블럭
    SellQuantity.SelLength = Len(SellQuantity.Text)
End Sub

Private Sub B_Cancel_MouseEnter()                                       '취소 팁
    ViewTip ("물품을 목록에 추가하지않고 돌아갑니다.")
End Sub

Private Sub B_Cancel_MouseLeave()                                       '팁 지우기
    NoTip
End Sub

Private Sub B_Submit_MouseEnter()                                       '물품추가 팁
    ViewTip ("선택한 물품을 선택목록에 추가합니다.")
End Sub

Private Sub B_Submit_MouseLeave()                                       '팁 지우기
    NoTip
End Sub

Function ReturnForm()
    Select Case FunctionMode                                            '복귀 폼 선택
        Case 0
            FrmSell.Left = Me.Left                                          '판매로 돌아감
            FrmSell.Top = Me.Top
            FrmSell.Show
        Case 1
            FrmLoss.Left = Me.Left                                          '손실등록으로 돌아감
            FrmLoss.Top = Me.Top
            FrmLoss.Show
        Case 2
            FrmStorage.Left = Me.Left                                       '입고로 돌아감
            FrmStorage.Top = Me.Top
            FrmStorage.Show
    End Select
    Unload Me
End Function

Private Function SetItemList()                                          '리스트박스에 물품목록 로드
    Dim i As Integer
    ItemList.Clear
    If ItemRecord.RecordCount > 0 Then
        ItemRecord.MoveFirst
        For i = 0 To ItemRecord.RecordCount - 1
            ItemList.AddItem ItemRecord.Fields(1)
            ItemRecord.MoveNext
        Next i
        ItemRecord.MoveFirst
    End If
End Function

