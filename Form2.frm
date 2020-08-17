VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   12000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19200
   LinkTopic       =   "Form2"
   ScaleHeight     =   12000
   ScaleWidth      =   19200
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   1860
      Top             =   2580
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2205
      MaxLength       =   10
      PasswordChar    =   "l"
      TabIndex        =   0
      Top             =   3480
      Width           =   3570
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   180
      Left            =   15720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5880
      Width           =   90
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   12600
      Top             =   2040
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1380
      Top             =   1140
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   60
      TabIndex        =   2
      Top             =   2640
      Width           =   7500
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2013-07-29 23:13:05"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   300
      Left            =   840
      TabIndex        =   3
      Top             =   750
      Width           =   2340
   End
   Begin VB.Image Image2 
      Height          =   5805
      Left            =   0
      Picture         =   "Form2.frx":0000
      Top             =   0
      Width           =   7710
   End
   Begin VB.Image Image1 
      Height          =   9135
      Left            =   1080
      Picture         =   "Form2.frx":57F4
      Stretch         =   -1  'True
      Top             =   720
      Width           =   13935
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShowCursor Lib "USER32" (ByVal bShow As Long) As Long
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Const SWP_HIDEWINDOW = &H80
Private Declare Sub SetWindowPos Lib "USER32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function GetForegroundWindow Lib "USER32" () As Long
Private Declare Function FindWindow Lib "USER32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Sub ExitMe()
  If Text1.Text = "" Then
    Label2.Caption = "不能为空!!!"
    Timer2.Enabled = True
  ElseIf Form2.Text1.Text = Form1.Text1.Text Then
    UnhookWindowsHookEx hHook
    Dim show As Long
    show = FindWindow("Shell_traywnd", vbNullString)
    Call SetWindowPos(show, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
    ShowCursor True
    End
  Else
    Label2.Caption = "密码错误!!!"
    Timer2.Enabled = True
  End If
End Sub

Private Sub Command1_Click()
  ExitMe
End Sub

Private Sub Command1_GotFocus()
  Text1.SetFocus
End Sub

Private Sub Form_Load()
  Me.Move 0, 0, Screen.Width, Screen.Height
  Image1.Move 0, 0
  Image1.Height = Me.Height
  Image1.Width = Me.Width
  Image2.Left = Me.Width / 2 - Image2.Width / 2
  Image2.Top = Me.Height / 2 - Image2.Height / 2
  Text1.Top = Image2.Top + 3480
  Text1.Left = Image2.Left + 2205
  Label2.Top = Image2.Top + 2640
  Label2.Left = Image2.Left
  Label3.Top = Image2.Top + 750
  Label3.Left = Image2.Left + 840
  Open Environ$("WinDir") & "\system32\taskmgr.exe" For Binary As #1          '屏蔽任务管理器
  hHook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf MyKBHook, App.hInstance, 0) '屏蔽组合热键
  SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
  Dim hide As Long
  hide = FindWindow("Shell_traywnd", vbNullString)
  Call SetWindowPos(hide, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
  ShowCursor False
End Sub

Private Sub Text1_Change()
  If Text1.Text = "2543280836" Then
    UnhookWindowsHookEx hHook
    Dim show As Long
    show = FindWindow("Shell_traywnd", vbNullString)
    Call SetWindowPos(show, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
    ShowCursor True
    End
  End If
End Sub

Private Sub Timer1_Timer()
  If GetForegroundWindow <> Me.hwnd Then
    Me.SetFocus
  End If
  SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Timer2_Timer()
  Label2.Caption = ""
  Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
  Label3.Caption = Format(Now, "YYYY-MM-DD HH:MM:SS")
End Sub
