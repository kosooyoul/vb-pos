VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmItemList 
   BorderStyle     =   1  '단일 고정
   Caption         =   "물품명 관리"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   ControlBox      =   0   'False
   Icon            =   "FrmItemList.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   6015
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '없음
      Height          =   735
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6255
      Begin VB.Label ItemN 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '투명
         Caption         =   "물품 목록에 새 물품에 대한 정보를 추가하거나 수정, 삭제합니다."
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
         TabIndex        =   6
         Top             =   360
         Width           =   5160
      End
      Begin VB.Label lblPage 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '투명
         Caption         =   "물품명 관리"
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
         TabIndex        =   5
         Top             =   120
         Width           =   1035
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
         Left            =   1200
         Top             =   0
         Width           =   4815
      End
   End
   Begin MSDataGridLib.DataGrid ItemList 
      Height          =   3215
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   5662
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BorderStyle     =   0
      ColumnHeaders   =   -1  'True
      HeadLines       =   1.5
      RowHeight       =   14
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "물품코드"
         Caption         =   "물품코드"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "물품이름"
         Caption         =   "물품이름"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "물품단가"
         Caption         =   "물품단가"
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
      BeginProperty Column03 
         DataField       =   "물품수량"
         Caption         =   "남은수량"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0 EA "
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
            ColumnWidth     =   854.929
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2099.906
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1049.953
         EndProperty
      EndProperty
   End
   Begin 편의점_물품관리.isButton B_Add 
      Height          =   360
      Left            =   240
      TabIndex        =   0
      Top             =   4320
      Width           =   1215
      _extentx        =   2143
      _extenty        =   635
      caption         =   "물품 추가"
      iconalign       =   1
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   0
      font            =   "FrmItemList.frx":058A
   End
   Begin 편의점_물품관리.isButton Close 
      Height          =   360
      Left            =   4560
      TabIndex        =   3
      Top             =   4320
      Width           =   1215
      _extentx        =   2143
      _extenty        =   635
      caption         =   "닫기"
      iconalign       =   1
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   0
      font            =   "FrmItemList.frx":05B2
   End
   Begin 편의점_물품관리.isButton B_Edit 
      Height          =   360
      Left            =   1560
      TabIndex        =   1
      Top             =   4320
      Width           =   1215
      _extentx        =   2143
      _extenty        =   635
      caption         =   "물품 수정"
      iconalign       =   1
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   0
      font            =   "FrmItemList.frx":05DA
   End
   Begin 편의점_물품관리.isButton B_Del 
      Height          =   360
      Left            =   2880
      TabIndex        =   2
      Top             =   4320
      Width           =   1215
      _extentx        =   2143
      _extenty        =   635
      caption         =   "물품 삭제"
      iconalign       =   1
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   0
      font            =   "FrmItemList.frx":0602
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   6690
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   0
      X2              =   6720
      Y1              =   4200
      Y2              =   4200
   End
End
Attribute VB_Name = "FrmItemList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strCode As String                                                   '물품코드
Dim strItemName As String                                               '물품이름
Dim longCost As Long                                                    '물품단가
Dim intQuantity As Integer                                              '물품남은수량

Private Sub Form_Load()                                                 '물품명 관리 폼 시작
    Dim slct As String
    
    BGIMG(1).Picture = LoadPicture(App.Path & "\Images\BGIMG_1.bmp")    '상단 배경이미지 설정
    
    Set ItemConnection = New ADODB.Connection                           '물품목록 데이터베이스 로드
    ItemConnection.CursorLocation = adUseClient
    ItemConnection.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Database\Library.mdb;Jet OLEDB:Database Password=" & DataBasePassWord & ";"
    slct = "select 물품코드,물품이름,물품단가,물품수량 from 물품목록 Order by 물품코드"
    Set ItemRecord = New ADODB.Recordset
    ItemRecord.Open slct, ItemConnection, adOpenStatic, adLockOptimistic
    
    Set ItemList.DataSource = ItemRecord                                '표에 데이터베이스 연결
    ItemList.ReBind
    ItemRecord.MoveFirst
End Sub

Private Sub B_Add_Click()                                               '물품 추가
    FrmAddItem.Left = Me.Left
    FrmAddItem.Top = Me.Top
    FrmAddItem.Show
    Unload Me
End Sub

Private Sub B_Edit_Click()                                              '물품 수정
    FrmEditItem.Left = Me.Left
    FrmEditItem.Top = Me.Top
    FrmEditItem.Show
    FrmEditItem.Code.Text = strCode
    FrmEditItem.ItemName.Text = strItemName
    FrmEditItem.Cost.Text = longCost
    Unload Me
End Sub

Private Sub B_Del_Click()                                               '물품 삭제
    If intQuantity = 0 Then                                             '물품 수량이 0인 경우 삭제 가능
        FrmDelItem.Left = Me.Left
        FrmDelItem.Top = Me.Top
        FrmDelItem.Show
        FrmDelItem.Code.Caption = strCode
        FrmDelItem.ItemName.Caption = strItemName
        FrmDelItem.Cost.Caption = longCost & " \"
        Unload Me
    Else                                                                '물품 수량이 있는 경우
        MsgBox "선택한 물품의 수량이 한 개 이상 남아있으면 삭제할 수 없습니다.", vbOKOnly, "품목 삭제 불가능"
    End If
End Sub

Private Sub Close_Click()                                               '물품명 관리 폼 닫기
    ItemRecord.Close                                                    '물품목록 데이터베이스 닫기
    ItemConnection.Close
    MainForm.Enabled = True
    Unload Me
End Sub

Private Sub B_Add_MouseEnter()                                          '물품 추가 팁
    ViewTip ("새 물품을 추가합니다.")
End Sub

Private Sub B_Add_MouseLeave()                                          '팁 지우기
    NoTip
End Sub

Private Sub B_Del_MouseEnter()                                          '물품 삭제 팁
    ViewTip ("선택한 물품을 삭제합니다.")
End Sub

Private Sub B_Del_MouseLeave()                                          '팁 지우기
    NoTip
End Sub
Private Sub B_Edit_MouseEnter()                                         '물품 수정 팁
    ViewTip ("선택한 물품을 수정합니다.")
End Sub

Private Sub B_Edit_MouseLeave()                                         '팁 지우기
    NoTip
End Sub

Private Sub Close_MouseEnter()                                          '물품명 관리 폼 닫기 팁
    ViewTip ("물품목록을 닫습니다.")
End Sub

Private Sub Close_MouseLeave()                                          '팁 지우기
    NoTip
End Sub

Private Sub ItemList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    strCode = ItemRecord.Fields(0)                                      '표에서 선택한 물품
    strItemName = ItemRecord.Fields(1)
    longCost = ItemRecord.Fields(2)
    intQuantity = ItemRecord.Fields(3)
End Sub

Private Sub ItemList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    strCode = ItemRecord.Fields(0)                                      '표에서 선택한 물품
    strItemName = ItemRecord.Fields(1)
    longCost = ItemRecord.Fields(2)
    intQuantity = ItemRecord.Fields(3)
End Sub
