VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.MDIForm MainForm 
   BackColor       =   &H00000000&
   Caption         =   "편의점 물품 관리"
   ClientHeight    =   6105
   ClientLeft      =   225
   ClientTop       =   540
   ClientWidth     =   9300
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MainForm"
   StartUpPosition =   1  '소유자 가운데
   Begin VB.PictureBox BottomFrame 
      Align           =   2  '아래 맞춤
      BorderStyle     =   0  '없음
      Height          =   2290
      Left            =   0
      ScaleHeight     =   2295
      ScaleWidth      =   9300
      TabIndex        =   1
      Top             =   3555
      Width           =   9300
      Begin VB.PictureBox VScrL 
         BackColor       =   &H00000000&
         BorderStyle     =   0  '없음
         Height          =   60
         Left            =   0
         ScaleHeight     =   60
         ScaleWidth      =   9735
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   9735
      End
      Begin VB.PictureBox VScroll 
         BorderStyle     =   0  '없음
         Height          =   60
         Left            =   0
         MousePointer    =   7  'N S크기 조정
         ScaleHeight     =   60
         ScaleWidth      =   10335
         TabIndex        =   3
         Top             =   0
         Width           =   10335
         Begin VB.Line LineC 
            BorderColor     =   &H8000000E&
            Index           =   0
            X1              =   0
            X2              =   9420
            Y1              =   30
            Y2              =   30
         End
         Begin VB.Line LineC 
            BorderColor     =   &H80000010&
            Index           =   1
            X1              =   0
            X2              =   9420
            Y1              =   20
            Y2              =   20
         End
      End
      Begin VB.ListBox LogList 
         Height          =   2220
         ItemData        =   "MDIForm1.frx":058A
         Left            =   0
         List            =   "MDIForm1.frx":058C
         TabIndex        =   2
         Top             =   60
         Width           =   9735
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3240
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":058E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0E68
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1742
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":201C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":28F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":31D0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  '아래 맞춤
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   5850
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
            Text            =   "편의점 이름"
            TextSave        =   "편의점 이름"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8017
            Text            =   "프로그램이 시작하였습니다."
            TextSave        =   "프로그램이 시작하였습니다."
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Text            =   "현재 시간"
            TextSave        =   "오전 12:00"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  '위 맞춤
      Height          =   825
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   1455
      BandCount       =   1
      _CBWidth        =   9300
      _CBHeight       =   825
      _Version        =   "6.0.8169"
      Child1          =   "MenuTool"
      MinWidth1       =   8100
      MinHeight1      =   765
      Width1          =   9240
      NewRow1         =   0   'False
      BandStyle1      =   1
      Begin MSComctlLib.Toolbar MenuTool 
         Height          =   765
         Left            =   30
         TabIndex        =   6
         Top             =   30
         Width           =   9240
         _ExtentX        =   16298
         _ExtentY        =   1349
         ButtonWidth     =   1773
         ButtonHeight    =   1349
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "정산"
               Key             =   "정산"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "물품판매"
               Key             =   "물품판매"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "물품명관리"
               Key             =   "물품명관리"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "물품입고"
               Key             =   "물품입고"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "손실등록"
               Key             =   "손실등록"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "도움말"
               Key             =   "도움말"
               ImageIndex      =   6
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu F_Main 
      Caption         =   "메인(&Main)"
      Begin VB.Menu S_ExactCalculation 
         Caption         =   "정산(&Exact Calculation)..."
         Shortcut        =   ^E
      End
      Begin VB.Menu S_Bar0 
         Caption         =   "-"
      End
      Begin VB.Menu S_Exit 
         Caption         =   "프로그램 종료(E&xit)"
         Shortcut        =   ^{F4}
      End
   End
   Begin VB.Menu F_Depot 
      Caption         =   "재고(&Depot)"
      Begin VB.Menu S_ItemList 
         Caption         =   "물품명 관리(&Item Management)..."
         Shortcut        =   ^I
      End
      Begin VB.Menu S_Bar1 
         Caption         =   "-"
      End
      Begin VB.Menu S_Storage 
         Caption         =   "물품 입고(&Storage)..."
         Shortcut        =   ^T
      End
      Begin VB.Menu S_Loss 
         Caption         =   "손실 등록(&Loss)..."
         Shortcut        =   ^L
      End
      Begin VB.Menu S_Bar2 
         Caption         =   "-"
      End
      Begin VB.Menu S_StorageList 
         Caption         =   "물품 입고 내역(S&torage List)..."
         Shortcut        =   ^R
      End
      Begin VB.Menu S_LossList 
         Caption         =   "손실 등록 내역(L&oss List)..."
         Shortcut        =   ^K
      End
   End
   Begin VB.Menu F_Selling 
      Caption         =   "판매(&Sell)"
      Begin VB.Menu S_Selling 
         Caption         =   "물품 판매(&Sell)..."
         Shortcut        =   ^S
      End
      Begin VB.Menu S_Bar3 
         Caption         =   "-"
      End
      Begin VB.Menu S_SellList 
         Caption         =   "물품 판매 내역(&Sell List)..."
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu F_Option 
      Caption         =   "설정(&Option)"
      Begin VB.Menu S_ShowLog 
         Caption         =   "로그 보기(Show &Log)"
         Checked         =   -1  'True
      End
      Begin VB.Menu S_ShowStatusBar 
         Caption         =   "상태표시줄 보기(Show &Status Bar)"
         Checked         =   -1  'True
      End
      Begin VB.Menu S_Bar4 
         Caption         =   "-"
      End
      Begin VB.Menu S_SetMarketName 
         Caption         =   "편의점 이름 설정(Set &Market Name)"
      End
   End
   Begin VB.Menu F_Help 
      Caption         =   "도움말(&Help)"
      Begin VB.Menu S_Help 
         Caption         =   "도움말 항목(&Help)..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu S_Bar5 
         Caption         =   "-"
      End
      Begin VB.Menu S_Information 
         Caption         =   "프로그램 정보(&Information)..."
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Y2 As Single                                                        '로그부분 사이즈 조절
Dim Moving As Boolean                                                   '로그 사이즈 조정중인지

Private Sub MDIForm_Load()                                              '폼 로딩시
    Dim LogFileNum As Integer
    Dim LogFileName As String                                           '로그 파일 이름
    Dim Logs As String
    
On Error GoTo ConfigNotFound                                            '설정파일이 없는 경우 ConfigNotFound로 이동

    LogFileName = App.Path & "\market.cfg"                              '설정파일 이름
    LogFileNum = FreeFile

    Open LogFileName For Input As LogFileNum                            '설정 파일 로드
    Do Until EOF(LogFileNum)
        Line Input #LogFileNum, Logs                                    '설정을 읽고 적용
        Select Case Left(Logs, 3)
            Case "[w]"                                                  '폼 너비 설정
                Me.Width = Abs(Mid(Logs, 4))
            Case "[h]"                                                  '폼 높이 설정
                Me.Height = Abs(Mid(Logs, 4))
            Case "[n]"                                                  '편의점 이름 설정
                Status.Panels(1).Text = Mid(Logs, 4)
            Case "[1]"                                                  '로그 표시 여부
                S_ShowLog.Checked = CBool(Mid(Logs, 4))
                BottomFrame.Visible = S_ShowLog.Checked
            Case "[2]"                                                  '상태바 표시 여부
                S_ShowStatusBar.Checked = CBool(Mid(Logs, 4))
                Status.Visible = S_ShowStatusBar.Checked
        End Select
    Loop
    Close LogFileNum
    
ConfigNotFound:                                                         '설정파일을 찾을 수 없을 경우
    If Err.Number = 53 Then
        MsgBox ("market.cfg 파일을 찾을 수 없습니다.")
    End If
End Sub

Private Sub MDIForm_Resize()                                            '폼 사이즈 조절하는 경우
    LogList.Width = Abs(BottomFrame.Width - 60)                         '로그목록 사이즈 조절
    VScrL.Width = BottomFrame.Width
    VScroll.Width = BottomFrame.Width
    LineC(0).X2 = Me.Width
    LineC(1).X2 = LineC(0).X2
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)                           '메인 폼 종료시
    Dim MsgResult As VbMsgBoxResult
    MsgResult = MsgBox("프로그램을 종료하시겠습니까?", vbYesNo, "프로그램 종료 선택")
    If MsgResult = vbNo Then
        Cancel = 1
    Else
        AddLog ("프로그램 종료 >> Program End")
        Cancel = 1
        Call SaveConfig
        End
    End If
End Sub

Private Sub MenuTool_ButtonClick(ByVal Button As MSComctlLib.Button)    '툴바버튼을 클릭시
    Select Case Button.Key
        Case "정산"
            Call S_ExactCalculation_Click
        Case "물품명관리"
            Call S_ItemList_Click
        Case "물품입고"
            Call S_Storage_Click
        Case "손실등록"
            Call S_Loss_Click
        Case "물품판매"
            Call S_Selling_Click
        Case "도움말"
            Call S_Help_Click
    End Select
End Sub

Private Sub S_ExactCalculation_Click()                                  '정산처리
    Dim MsgResult As VbMsgBoxResult
    MsgResult = MsgBox("정산처리를 하겠습니까?", vbYesNo, "정산처리 선택")
    If MsgResult = vbYes Then FrmExactCalculation.Show vbModal
End Sub

Private Sub S_Exit_Click()                                              '프로그램 종료 시도
    Unload Me
End Sub

Private Sub S_Help_Click()                                              '도움말 폼 표시
    FrmHelp.Show
End Sub

Private Sub S_Information_Click()                                       '프로그램 정보 폼 표시
    frmAbout.Show vbModal
End Sub

Private Sub S_ItemList_Click()                                          '물품명 관리 폼 표시
    FrmItemList.Show
    MainForm.Enabled = False
End Sub

Private Sub S_Loss_Click()                                              '손실등록 폼 표시
    FunctionMode = 1                                                    '1 : 손실(판매:0,입고:2)
    FrmLoss.Show
    MainForm.Enabled = False
End Sub

Private Sub S_LossList_Click()                                          '손실내역 폼 표시
    FrmLossList.Show vbModal
End Sub

Private Sub S_Selling_Click()                                           '물품판매 폼 표시
    FunctionMode = 0                                                    '0 : 판매(손실:1,입고:2)
    FrmSell.Show
    MainForm.Enabled = False
End Sub

Private Sub S_SellList_Click()                                          '판매내역 폼 표시
    FrmSellList.Show vbModal
End Sub

Private Sub S_SetMarketName_Click()                                     '편의점 이름 설정 : 상태표시줄.첫번째칸
    Dim Str As String
    Str = InputBox("편의점 이름을 입력하세요", "편의점 이름 설정", Status.Panels(1).Text)
    If Str <> "" Then Status.Panels(1).Text = Str
End Sub

Private Sub S_ShowLog_Click()                                           '로그 표시 On/Off
    S_ShowLog.Checked = Not (S_ShowLog.Checked)
    BottomFrame.Visible = S_ShowLog.Checked
End Sub

Private Sub S_ShowStatusBar_Click()                                     '상태표시줄 표시 On/Off
    S_ShowStatusBar.Checked = Not (S_ShowStatusBar.Checked)
    Status.Visible = S_ShowStatusBar.Checked
    Status.Top = Me.Height + 1000
End Sub

Private Sub S_Storage_Click()                                           '물품입고 폼 표시
    FunctionMode = 2                                                    '2 : 입고(판매:0,손실:1)
    FrmStorage.Show
    MainForm.Enabled = False
End Sub

Private Sub S_StorageList_Click()                                       '입고내역 폼 표시
    FrmStorageList.Show vbModal
End Sub

Private Sub VScroll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)    '로그 프레임 크기 조절 시작
    Y2 = Y
    VScrL.Visible = True
    VScrL.Top = VScroll.Top
    Moving = True
End Sub

Private Sub VScroll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)    '로그 프레임 크기 조절 중
    If Moving Then
        VScrL.Top = VScrL.Top + (Y - Y2)
        Y2 = Y
    End If
End Sub

Private Sub VScroll_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)      '로그 프레임 크기 조절 끝
        If Moving Then
        If Abs(BottomFrame.Height - Y) > 3000 Then                    '로그 프레임 최대 크기 제한
            BottomFrame.Height = 2900
        Else
            BottomFrame.Height = Abs(BottomFrame.Height - Y)
        End If
        LogList.Height = BottomFrame.Height
        BottomFrame.Height = LogList.Height + 80
        Status.Top = BottomFrame.Top + BottomFrame.Height
        VScroll.Top = 0
        VScrL.Visible = False
        Moving = False
    End If
End Sub
