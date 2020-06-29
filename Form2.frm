VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  '단일 고정
   Caption         =   "출석체크 현황"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9525
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   9525
   StartUpPosition =   2  '화면 가운데
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   2760
      Top             =   1800
   End
   Begin VB.CommandButton Command4 
      Caption         =   "설정"
      Height          =   255
      Left            =   8640
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "조회"
      Height          =   255
      Left            =   7680
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Height          =   1335
      Left            =   4560
      TabIndex        =   2
      Top             =   480
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4335
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   120
      Style           =   2  '드롭다운 목록
      TabIndex        =   0
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StartTime As String
Dim EndTime As String
Private Sub RefreshTime(ByVal Type_ As String)
    If Combo1.ListIndex > 0 Then
        If Type_ <> "" Then setBeacon Combo1.List(Combo1.ListIndex), Type_

        spt = Split(getbeaconTime(Combo1.List(Combo1.ListIndex)), "/")
        
        StartTime = spt(0)
        EndTime = spt(1)
        
        Command1.Caption = "입실" & vbCrLf & vbCrLf & StartTime
        Command2.Caption = "퇴실" & vbCrLf & vbCrLf & EndTime
    End If
End Sub

Private Sub Command1_Click()
    If MsgBox("주의!!" & vbCrLf & vbCrLf & "입실 처리 하겠습니까?", vbInformation + vbYesNo, "질문") = vbYes Then
        '//attendScd ( I : 입실, O : 퇴실, R : 외출, K : 복귀
     
        If Len(StartTime) > 1 And MsgBox("주의!!" & vbCrLf & vbCrLf & "입실 처리가 이미 적용되어 있는 상태 입니다." & vbCrLf & vbCrLf & "다시 처리하겠습니까?", vbExclamation + vbYesNo, "질문") = vbNo Then Exit Sub
        
        RefreshTime "I"
        
        MsgBox StartTime & " 에 입실 처리가 완료 되었습니다. 이제부터 해당 윈도우 창은 퇴실 시간 전까지 숨겨집니다.", vbInformation, "알림"
        Me.Visible = False
    End If
End Sub

Private Sub Command2_Click()
    If MsgBox("주의!!" & vbCrLf & vbCrLf & "퇴실 처리 하겠습니까?", vbInformation + vbYesNo, "질문") = vbYes Then
        '//attendScd ( I : 입실, O : 퇴실, R : 외출, K : 복귀
        
        If Len(EndTime) > 1 And MsgBox("주의!!" & vbCrLf & vbCrLf & "퇴실 처리가 이미 적용되어 있는 상태 입니다." & vbCrLf & vbCrLf & "다시 처리하겠습니까?", vbExclamation + vbYesNo, "질문") = vbNo Then Exit Sub
        RefreshTime "O"
    End If
End Sub

Private Sub Command3_Click()
    Dim spt() As String
    RefreshTime vbNullString
End Sub

Private Sub Command4_Click()
    Form3.Show vbModal, Me
End Sub

Private Sub Form_Activate()

    Dim tm As Date
    tm = Now
    
    BringWindowToTop Me.hwnd
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE
    
    If Len(StartTime) > 1 Then
        Me.Visible = False
    End If
    
    If Len(EndTime) > 1 Or Hour(tm) > 17 Or (Hour(tm) >= 17 And Minute(tm) >= 35) Then
        Me.Visible = True
    End If

End Sub

Private Sub Form_Load()
    Combo1.AddItem "선택하세요"
    
    For i = 0 To 999
        Combo1.AddItem Format(DateAdd("d", -i, Now), "yyyy-mm-dd")
    Next
    
    Combo1.ListIndex = 1
    Call Command3_Click
    
        
    smDx = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not smDx Then End
End Sub

Private Sub Timer1_Timer()
    Dim tm As Date
    tm = Now

    If Len(EndTime) > 1 Or Hour(tm) > 17 Or (Hour(tm) >= 17 And Minute(tm) >= 35) Then
        Me.Visible = True
    End If
End Sub
