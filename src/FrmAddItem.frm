VERSION 5.00
Begin VB.Form FrmAddItem 
   BorderStyle     =   1  '단일 고정
   Caption         =   "물품 추가"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3735
   ControlBox      =   0   'False
   Icon            =   "FrmAddItem.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   3735
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '없음
      Height          =   975
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6015
      Begin VB.Label ItemN 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '투명
         Caption         =   $"FrmAddItem.frx":014A
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Width           =   2595
      End
      Begin VB.Label lblPage 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '투명
         Caption         =   "물품 목록에 새 물품 추가"
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
         Width           =   2190
      End
      Begin VB.Line Line4 
         BorderColor     =   &H0099A8AC&
         X1              =   -120
         X2              =   7320
         Y1              =   945
         Y2              =   945
      End
      Begin VB.Image BGIMG 
         Height          =   945
         Index           =   2
         Left            =   -960
         Top             =   0
         Width           =   4695
      End
   End
   Begin VB.TextBox Cost 
      Height          =   300
      Left            =   960
      MaxLength       =   10
      TabIndex        =   2
      Text            =   "0"
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox Code 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "00000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1042
         SubFormatType   =   0
      EndProperty
      Height          =   300
      Left            =   960
      MaxLength       =   5
      TabIndex        =   0
      Text            =   "00000"
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox ItemName 
      Height          =   300
      Left            =   960
      MaxLength       =   20
      TabIndex        =   1
      Text            =   "품명"
      Top             =   1440
      Width           =   2055
   End
   Begin 편의점_물품관리.isButton B_Submit 
      Height          =   360
      Left            =   720
      TabIndex        =   3
      Top             =   2400
      Width           =   1335
      _extentx        =   3625
      _extenty        =   635
      caption         =   "물품 추가"
      iconalign       =   1
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   0
      font            =   "FrmAddItem.frx":0181
   End
   Begin 편의점_물품관리.isButton B_Cancel 
      Height          =   360
      Left            =   2160
      TabIndex        =   4
      Top             =   2400
      Width           =   1335
      _extentx        =   2143
      _extenty        =   635
      caption         =   "취소"
      iconalign       =   1
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   0
      font            =   "FrmAddItem.frx":01A9
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   $"FrmAddItem.frx":01D1
      ForeColor       =   &H00000000&
      Height          =   900
      Left            =   345
      TabIndex        =   8
      Top             =   1110
      Width           =   480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   6690
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   0
      X2              =   6720
      Y1              =   2280
      Y2              =   2280
   End
End
Attribute VB_Name = "FrmAddItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()                                                 '물품 추가 폼의 시작
    BGIMG(2).Picture = LoadPicture(App.Path & "\Images\BGIMG_2.bmp")    '상단 배경이미지 설정
End Sub

Private Sub B_Submit_Click()                                            '물품 추가
    Dim Str As String
    
On Error GoTo OverLap                                                   '물품명,코드 중복 에러발생시 OverLap으로 이동

    Str = "INSERT INTO 물품목록"                                        'SQL문 : 레코드 추가
    Str = Str & "(물품코드,물품이름,물품단가) "
    Str = Str & "VALUES('" & Code.Text & "', "                              '물품코드
    Str = Str & "'" & Trim(ItemName.Text) & "', "                           '물품이름
    Str = Str & "'" & Val(Cost.Text) & "')"                                 '물품단가
    ItemConnection.Execute (Str)
                                                                        '물품 추가 로그
    AddLog ("물품목록 > 품목추가 >>     코드(" & Code.Text & "), 품명(" & ItemName.Text & "), 단가(" & Cost.Text & " \)")
    
    FrmItemList.Left = Me.Left                                          '물품명 관리 폼 표시
    FrmItemList.Top = Me.Top
    FrmItemList.Show
    Unload Me

OverLap:                                                                '코드나 물품명이 중복되는 경우
    If Err.Number = -2147467259 Then
        MsgBox "같은 상품이 존재합니다.", vbOKOnly, "중복에 대한 경고"
    End If
End Sub

Private Sub B_Cancel_Click()                                            '물품 추가 취소
    FrmItemList.Left = Me.Left
    FrmItemList.Top = Me.Top
    FrmItemList.Show
    Unload Me
End Sub

Private Sub B_Submit_MouseEnter()                                       '물품 추가 팁
    ViewTip ("물품목록에 입력한 물품을 추가합니다.")
End Sub

Private Sub B_Submit_MouseLeave()                                       '팁 지우기
    NoTip
End Sub

Private Sub B_Cancel_MouseEnter()                                       '취소 팁
    ViewTip ("물품을 추가하지 않고 물품목록으로 돌아갑니다.")
End Sub

Private Sub B_Cancel_MouseLeave()                                       '팁 지우기
    NoTip
End Sub

Private Sub Code_GotFocus()                                             '코드 입력시 텍스트 블럭지정
    Code.SelStart = 0
    Code.SelLength = Len(Code.Text)
End Sub

Private Sub Code_LostFocus()                                            '코드를 숫자 5자리로 변환
    Code.Text = Format(Val(Code.Text), "00000")
End Sub

Private Sub Cost_GotFocus()                                             '단가 입력시 텍스트 블럭지정
    Cost.SelStart = 0
    Cost.SelLength = Len(Cost.Text)
End Sub

Private Sub Cost_LostFocus()                                            '단가를 자연수로 변환
    Cost.Text = Abs(Val(Cost.Text))
End Sub

Private Sub ItemName_GotFocus()                                         '물품명 입력시 텍스트 블럭지정
    ItemName.SelStart = 0
    ItemName.SelLength = Len(ItemName.Text)
End Sub
    
