VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  '단일 고정
   Caption         =   "월드잡 출석체크"
   ClientHeight    =   1230
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   4230
   StartUpPosition =   2  '화면 가운데
   Begin VB.CheckBox Check2 
      Caption         =   "쿠키저장"
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   840
      Value           =   1  '확인
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "자동 로그인"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "로그인"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   270
      IMEMode         =   3  '사용 못함
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "비밀번호 : "
      Height          =   180
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "아이디 : "
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Winhttp As New Winhttp.WinHttpRequest

Private Sub checkSession(typerx As Long)
On Error Resume Next

    Dim bc As Long
    Dim quot As Long
    
    Dim src As String

    src = firstProc
    bc = InStr(1, src, "/indvdl/epmtSprtMng/beaconAtendenceAuto.do")
    
    If bc < 1 Then
        If typerx = 1 Then MsgBox "로그인 실패", vbExclamation, "알림"
    Else
        quot = InStr(bc, src, """")
        If quot > bc Then
            beaconAddr = Mid(src, bc, quot - bc)
            
            If Check2.value = 1 Then
                SaveSetting "JKJ", "SSE", "SSS", getCookieAll
            Else
                DeleteSetting "JKJ"
            End If
            
            Form2.Show
            Unload Me
        Else
            If typerx = 1 Then MsgBox "Error", vbCritical, "ERROR"
            End
        End If
    End If
End Sub

Private Sub Command1_Click()
    setCookie Login(Text1.Text, Text2.Text)
    Call checkSession(1)
End Sub

Private Sub Form_Activate()
    Dim areg As String
    
    areg = GetSetting("JKJ", "SSE", "SSS")
    
    If areg <> "" Then
        setCookie "Cookie:" & areg & vbCrLf
        Text3 = "Cookie:" & areg & vbCrLf
    End If
    
    Call checkSession(0)
    Command1.Enabled = True
End Sub
