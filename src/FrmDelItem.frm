VERSION 5.00
Begin VB.Form FrmDelItem 
   BorderStyle     =   1  '단일 고정
   Caption         =   "물품 삭제"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   ControlBox      =   0   'False
   Icon            =   "FrmDelItem.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   5415
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '없음
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6135
      Begin VB.Line Line4 
         BorderColor     =   &H0099A8AC&
         X1              =   0
         X2              =   7440
         Y1              =   705
         Y2              =   705
      End
      Begin VB.Label lblPage 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '투명
         Caption         =   "물품 목록에서 선택한 물품을 제거"
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
         TabIndex        =   4
         Top             =   120
         Width           =   2970
      End
      Begin VB.Label ItemN 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '투명
         Caption         =   "다음 물품을 삭제하시겠습니까?"
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
         TabIndex        =   3
         Top             =   360
         Width           =   2505
      End
      Begin VB.Image BGIMG 
         Height          =   945
         Index           =   2
         Left            =   720
         Top             =   -240
         Width           =   4695
      End
   End
   Begin 편의점_물품관리.isButton B_Submit 
      Height          =   360
      Left            =   2400
      TabIndex        =   0
      Top             =   1680
      Width           =   1335
      _extentx        =   3625
      _extenty        =   635
      caption         =   "물품 삭제"
      iconalign       =   1
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   0
      font            =   "FrmDelItem.frx":058A
   End
   Begin 편의점_물품관리.isButton B_Cancel 
      Height          =   360
      Left            =   3840
      TabIndex        =   1
      Top             =   1680
      Width           =   1335
      _extentx        =   3625
      _extenty        =   635
      caption         =   "취소"
      iconalign       =   1
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   0
      font            =   "FrmDelItem.frx":05B2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "물품단가 :"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   3120
      TabIndex        =   9
      Top             =   1230
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   $"FrmDelItem.frx":05DA
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   345
      TabIndex        =   8
      Top             =   870
      Width           =   840
   End
   Begin VB.Label Cost 
      AutoSize        =   -1  'True
      Caption         =   "0 \"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   4080
      TabIndex        =   7
      Top             =   1230
      Width           =   315
   End
   Begin VB.Label ItemName 
      AutoSize        =   -1  'True
      Caption         =   "품명"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   1320
      TabIndex        =   6
      Top             =   1230
      Width           =   360
   End
   Begin VB.Label Code 
      AutoSize        =   -1  'True
      Caption         =   "000000"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   1320
      TabIndex        =   5
      Top             =   870
      Width           =   540
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   6690
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   0
      X2              =   6720
      Y1              =   1560
      Y2              =   1560
   End
End
Attribute VB_Name = "FrmDelItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()                                                 '물품 삭제 폼의 시작
    BGIMG(2).Picture = LoadPicture(App.Path & "\Images\BGIMG_2.bmp")    '상단 배경이미지 설정
End Sub

Private Sub B_Cancel_Click()                                            '삭제 취소
    FrmItemList.Left = Me.Left
    FrmItemList.Top = Me.Top
    FrmItemList.Show
    Unload Me
End Sub

Private Sub B_Submit_Click()                                            '물품 삭제
    ItemRecord.Delete adAffectCurrent
                                                                        '물품 삭제 로그
    AddLog ("물품목록 > 품목삭제 >>     코드(" & Code.Caption & "), 품명(" & ItemName.Caption & "), 단가(" & Cost.Caption & ")")
     
    FrmItemList.Left = Me.Left                                          '물품명 관리 폼 표시
    FrmItemList.Top = Me.Top
    FrmItemList.Show
    Unload Me
End Sub

Private Sub B_Cancel_MouseEnter()                                       '삭제 취소 팁
    ViewTip ("물품을 삭제하지 않고 물품목록으로 돌아갑니다.")
End Sub

Private Sub B_Cancel_MouseLeave()                                       '팁 지우기
    NoTip
End Sub

Private Sub B_Submit_MouseEnter()                                       '물품 삭제 팁
    ViewTip ("물품 [" & ItemName.Caption & "]을 삭제합니다.")
End Sub

Private Sub B_Submit_MouseLeave()                                       '팁 지우기
    NoTip
End Sub
