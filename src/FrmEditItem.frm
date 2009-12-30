VERSION 5.00
Begin VB.Form FrmEditItem 
   BorderStyle     =   1  '단일 고정
   Caption         =   "물품 수정"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3735
   ControlBox      =   0   'False
   Icon            =   "FrmEditItem.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   3735
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '없음
      Height          =   735
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   3855
      Begin VB.Label ItemN 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '투명
         Caption         =   "물품의 이름과, 단가를 수정합니다."
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
         TabIndex        =   7
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label lblPage 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '투명
         Caption         =   "선택한 물품에 대한 정보 수정"
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
         Width           =   2580
      End
      Begin VB.Line Line4 
         BorderColor     =   &H0099A8AC&
         X1              =   -120
         X2              =   7320
         Y1              =   705
         Y2              =   705
      End
      Begin VB.Image BGIMG 
         Height          =   705
         Index           =   1
         Left            =   -1080
         Top             =   0
         Width           =   4815
      End
   End
   Begin VB.TextBox Cost 
      Height          =   300
      Left            =   960
      MaxLength       =   10
      TabIndex        =   2
      Text            =   "0"
      Top             =   1680
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
      Enabled         =   0   'False
      Height          =   300
      Left            =   960
      MaxLength       =   5
      ScrollBars      =   1  '수평
      TabIndex        =   0
      Text            =   "00000"
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox ItemName 
      Height          =   300
      Left            =   960
      MaxLength       =   20
      TabIndex        =   1
      Text            =   "품명"
      Top             =   1320
      Width           =   2055
   End
   Begin 편의점_물품관리.isButton B_Submit 
      Height          =   360
      Left            =   720
      TabIndex        =   3
      Top             =   2280
      Width           =   1335
      _extentx        =   3625
      _extenty        =   635
      caption         =   "물품 수정"
      iconalign       =   1
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   0
      font            =   "FrmEditItem.frx":058A
   End
   Begin 편의점_물품관리.isButton B_Cancel 
      Height          =   360
      Left            =   2160
      TabIndex        =   4
      Top             =   2280
      Width           =   1335
      _extentx        =   3625
      _extenty        =   635
      caption         =   "취소"
      iconalign       =   1
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   0
      font            =   "FrmEditItem.frx":05B2
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   $"FrmEditItem.frx":05DA
      ForeColor       =   &H00000000&
      Height          =   900
      Left            =   345
      TabIndex        =   8
      Top             =   990
      Width           =   480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   6690
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   0
      X2              =   6720
      Y1              =   2160
      Y2              =   2160
   End
End
Attribute VB_Name = "FrmEditItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()                                                 '물품 수정 폼의 시작
    BGIMG(1).Picture = LoadPicture(App.Path & "\Images\BGIMG_1.bmp")    '상단 배경이미지 설정
End Sub

Private Sub B_Submit_Click()                                            '물품 수정
    Dim Str As String
    
On Error GoTo OverLap                                                   '물품명 중복 에러발생시 OverLap으로 이동

    Str = "UPDATE 물품목록 SET "                                        'SQL문 : 레코드 수정
    Str = Str & "물품이름='" & Trim(ItemName.Text) & "',"                   '물품이름
    Str = Str & "물품단가=" & Val(Cost.Text) & ","                          '물품단가
    Str = Str & "물품코드='" & Code.Text & "'"                              '물품코드
    Str = Str & " WHERE 물품코드='" & Code.Text & "'"                       '물품코드로 식별
    ItemConnection.Execute (Str)
                                                                        '물품 삭제 로그
    AddLog ("물품목록 > 품목수정 >>     코드(" & Code.Text & "), 품명(" & ItemName.Text & "), 단가(" & Cost.Text & " \)")
        
    FrmItemList.Left = Me.Left                                          '물품명 관리 폼 표시
    FrmItemList.Top = Me.Top
    FrmItemList.Show
    Unload Me

OverLap:                                                                '물품명 중복시
    If Err.Number = -2147467259 Then
        MsgBox "같은 상품이 존재합니다.", vbOKOnly, "중복에 대한 경고"
    End If
End Sub

Private Sub B_Cancel_Click()                                            '물품 수정 취소
    FrmItemList.Left = Me.Left
    FrmItemList.Top = Me.Top
    FrmItemList.Show
    Unload Me
End Sub

Private Sub B_Cancel_MouseEnter()                                       '수정 취소 팁
    ViewTip ("물품을 수정하지 않고 물품목록으로 돌아갑니다.")
End Sub

Private Sub B_Cancel_MouseLeave()                                       '팁 지우기
    NoTip
End Sub

Private Sub B_Submit_MouseEnter()                                       '물품 수정 팁
    ViewTip ("물품의 품명, 단가를 수정합니다.")
End Sub

Private Sub B_Submit_MouseLeave()                                       '팁 지우기
    NoTip
End Sub

Private Sub Cost_GotFocus()                                             '단가 입력시 텍스트 블럭
    Cost.SelStart = 0
    Cost.SelLength = Len(Cost.Text)
End Sub

Private Sub Cost_LostFocus()                                            '단가 입력후 자연수로 변환
    Cost.Text = Abs(Val(Cost.Text))
End Sub

Private Sub ItemName_GotFocus()                                         '물품명 입력시 텍스트 블럭
    ItemName.SelStart = 0
    ItemName.SelLength = Len(ItemName.Text)
End Sub
