VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  '���� ����
   Caption         =   "�⼮üũ ��Ȳ"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9525
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   9525
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   2760
      Top             =   1800
   End
   Begin VB.CommandButton Command4 
      Caption         =   "����"
      Height          =   255
      Left            =   8640
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��ȸ"
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
      Style           =   2  '��Ӵٿ� ���
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
        
        Command1.Caption = "�Խ�" & vbCrLf & vbCrLf & StartTime
        Command2.Caption = "���" & vbCrLf & vbCrLf & EndTime
    End If
End Sub

Private Sub Command1_Click()
    If MsgBox("����!!" & vbCrLf & vbCrLf & "�Խ� ó�� �ϰڽ��ϱ�?", vbInformation + vbYesNo, "����") = vbYes Then
        '//attendScd ( I : �Խ�, O : ���, R : ����, K : ����
     
        If Len(StartTime) > 1 And MsgBox("����!!" & vbCrLf & vbCrLf & "�Խ� ó���� �̹� ����Ǿ� �ִ� ���� �Դϴ�." & vbCrLf & vbCrLf & "�ٽ� ó���ϰڽ��ϱ�?", vbExclamation + vbYesNo, "����") = vbNo Then Exit Sub
        
        RefreshTime "I"
        
        MsgBox StartTime & " �� �Խ� ó���� �Ϸ� �Ǿ����ϴ�. �������� �ش� ������ â�� ��� �ð� ������ �������ϴ�.", vbInformation, "�˸�"
        Me.Visible = False
    End If
End Sub

Private Sub Command2_Click()
    If MsgBox("����!!" & vbCrLf & vbCrLf & "��� ó�� �ϰڽ��ϱ�?", vbInformation + vbYesNo, "����") = vbYes Then
        '//attendScd ( I : �Խ�, O : ���, R : ����, K : ����
        
        If Len(EndTime) > 1 And MsgBox("����!!" & vbCrLf & vbCrLf & "��� ó���� �̹� ����Ǿ� �ִ� ���� �Դϴ�." & vbCrLf & vbCrLf & "�ٽ� ó���ϰڽ��ϱ�?", vbExclamation + vbYesNo, "����") = vbNo Then Exit Sub
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
    Combo1.AddItem "�����ϼ���"
    
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
