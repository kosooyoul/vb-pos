VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "편의점 물품 관리 정보"
   ClientHeight    =   3015
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5415
   ClipControls    =   0   'False
   Icon            =   "FrmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2081.007
   ScaleMode       =   0  '사용자
   ScaleWidth      =   5084.965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '없음
      Height          =   2295
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6015
      Begin VB.PictureBox picIcon 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  '없음
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   180
         Picture         =   "FrmAbout.frx":058A
         ScaleHeight     =   337.12
         ScaleMode       =   0  '사용자
         ScaleWidth      =   337.12
         TabIndex        =   2
         Top             =   120
         Width           =   480
      End
      Begin VB.Label URL_SMUniv 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "http://smics.semyung.ac.kr"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   1080
         MouseIcon       =   "FrmAbout.frx":0E54
         MousePointer    =   99  '사용자 정의
         TabIndex        =   5
         Top             =   960
         Width           =   2400
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   $"FrmAbout.frx":171E
         ForeColor       =   &H80000008&
         Height          =   1440
         Left            =   960
         TabIndex        =   4
         Top             =   600
         Width           =   4080
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "+ 편의점 물품 관리 1.0"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   240
         Width           =   3735
      End
      Begin VB.Line Line4 
         BorderColor     =   &H0099A8AC&
         X1              =   -120
         X2              =   7320
         Y1              =   2265
         Y2              =   2265
      End
   End
   Begin 편의점_물품관리.isButton B_Close 
      Height          =   360
      Left            =   3840
      TabIndex        =   0
      Top             =   2400
      Width           =   1335
      _ExtentX        =   2143
      _ExtentY        =   635
      Caption         =   "확인"
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
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub B_Close_Click()                                             '프로그램정보 폼 닫기
    Unload Me
End Sub

Private Sub URL_SMUniv_Click()                                          '세명대학교 정통 홈페이지로 이동
    ShellExecute 0, "open", "http://smics.semyung.ac.kr", "", "", 10
End Sub
