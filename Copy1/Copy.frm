VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "�������U��"
   ClientHeight    =   6375
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8415
   LinkTopic       =   "�������U��"
   ScaleHeight     =   6375
   ScaleWidth      =   8415
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command4 
      Caption         =   "���U��"
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   1800
      TabIndex        =   6
      Top             =   600
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   1800
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2640
      Width           =   4935
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   1800
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1680
      Width           =   4935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ѡ�񱣴��ļ��ĵ�ַ��"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   5880
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ʾʱ��"
      Height          =   495
      Index           =   1
      Left            =   5160
      TabIndex        =   1
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      Height          =   495
      Index           =   0
      Left            =   1800
      TabIndex        =   0
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���ƽ����"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   2640
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click(Index As Integer)
Timer1.Enabled = True
Timer1.Interval = 2000
Rem Text1.Text = Text1.Text & vbCrLf & Now
End Sub

Private Sub Command2_Click(Index As Integer)
Dim str As String
str = Text2.Text
If (str = "") Then
MsgBox "��ѡ��ַΪ��,��ѡ��洢��ַ��"
Else
Set fso = CreateObject("scripting.filesystemobject")
Set ws = CreateObject("wscript.shell")
On Error Resume Next
'����U�̵�����
Dim U_name$
U_name = Text3.Text
If fso.DriveExists(U_name) Then
fso.copyfile U_name & "*", str
Rem fso.copyfile "G:\CopyForTest01\*", str
fso.copyfolder U_name & "*", str
Rem fso.copyfolder "G:\CopyForTest01\*", str
MsgBox "copy success!"
'�ڽ��������ʾ
Text1.Text = Text1.Text & Now & "       �ɹ�" & vbCrLf
Else
MsgBox "file is not exist!"
End If
End If
End Sub

Private Sub Command3_Click()
'shell ����
Dim objDol
'�ļ�ѡ��Ի������
Dim objF
Dim DstPath
Dim I
Set objDlg = CreateObject("shell.Application")
Set objF = objDlg.BrowseForFolder(&H0, "ѡ����λ�ã�", &H1)
If InStr(1, TypeName(objF), "Folder", vbTextCompare) > 0 Then
    DstPath = objF.self.Path
    MsgBox "·��Ϊ��" & vbCrLf & DstPath
    Text2.Text = DstPath
Else
MsgBox "Ŀ¼��Ч"
End If
End Sub

'���U���Ƿ����
Private Sub Command4_Click()
Set fso = CreateObject("scripting.filesystemobject")
Set ws = CreateObject("wscript.shell")
On Error Resume Next
Dim u_file(7) As String
'����E�̿�ʼ��K�̣��ж��Ƿ����U��
u_file(0) = "E:\"
u_file(1) = "F:\"
u_file(2) = "G:\"
u_file(3) = "H:\"
u_file(4) = "I:\"
u_file(5) = "J:\"
u_file(6) = "K:\"

End Sub

Private Sub Form_Load()
Text1.Text = Text1.Text & "     ʱ��                 " & "״̬" & vbCrLf
End Sub


Private Sub Timer1_Timer()
Text1.Text = Text1.Text & vbCrLf & Now
End Sub
