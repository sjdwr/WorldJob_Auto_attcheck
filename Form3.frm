VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  '���� ����
   Caption         =   "����"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5445
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   5445
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.CheckBox Check3 
      Caption         =   "��� �� �ý��� ���� �մϴ�."
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   3135
   End
   Begin VB.CheckBox Check1 
      Caption         =   "�ڵ� ���"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�ڵ� �α����� �����մϴ�."
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    DeleteSetting "JKJ"
    beaconAddr = vbNullString
    UnloadCookie
    Unload Me
    Form1.Show
    smDx = True
    Unload Form2
    MsgBox "���� �Ϸ�. �ٽ� �����ϼ���.", vbInformation, ""
    End
End Sub

Private Sub Form_Activate()
    BringWindowToTop Me.hwnd
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE
End Sub
