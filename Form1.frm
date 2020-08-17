VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "锁定"
   ClientHeight    =   3585
   ClientLeft      =   105
   ClientTop       =   135
   ClientWidth     =   4935
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   3390
      Left            =   0
      Picture         =   "Form1.frx":5E62
      ScaleHeight     =   3390
      ScaleWidth      =   4500
      TabIndex        =   1
      Top             =   0
      Width           =   4500
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1500
         Left            =   0
         Top             =   0
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   12
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000003&
         Height          =   380
         IMEMode         =   3  'DISABLE
         Left            =   1045
         MaxLength       =   14
         PasswordChar    =   "l"
         TabIndex        =   0
         Top             =   1945
         Width           =   2350
      End
      Begin VB.Image Image1 
         Height          =   465
         Left            =   3180
         Picture         =   "Form1.frx":9121
         Top             =   900
         Width           =   465
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   7.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1020
         TabIndex        =   6
         Top             =   2820
         Width           =   1515
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "锁定"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1500
         TabIndex        =   5
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1140
         TabIndex        =   3
         Top             =   1500
         Width           =   2250
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "设置密码:"
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   540
         TabIndex        =   2
         Top             =   960
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3600
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowLong Lib "USER32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "USER32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "USER32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H1

Private Sub Command1_Click()
  If Text1.Text = "" Then
    Label2.Caption = "密码不能为空!!!"
    Timer1.Enabled = True
  Else
    Suo
  End If
End Sub

Private Sub Command1_GotFocus()
  Text1.SetFocus
End Sub

Private Sub Picture1_GotFocus()
  Text1.SetFocus
End Sub

Private Sub Form_Load()
  If Command = "" Then
    Me.Height = Picture1.Height
    Me.Width = Picture1.Width
    Picture1.BackColor = &HFFFFFF
    Me.BackColor = &HFFFFFF
    Dim rtn As Long
    rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, &HFFFFFF, 190, LWA_ALPHA
  Else
    Text1.Text = Command
    Form1.hide
    Form2.show
  End If
End Sub

Private Sub Image1_Click()
  End
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Label3.FontUnderline = True
End Sub
Private Sub Label3_Click()
If Text1.Text = "" Then
  Label2.Caption = "密码不能为空!!!"
  Timer1.Enabled = True
Else
  Suo
End If
End Sub
Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Label3.FontUnderline = True
End Sub
Private Sub Label4_Click()
If Text1.Text = "" Then
  Label2.Caption = "密码不能为空!!!"
  Timer1.Enabled = True
Else
  Suo
End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Label3.FontUnderline = False
End Sub

Private Sub Timer1_Timer()
  Label2.Caption = ""
  Timer1.Enabled = False
End Sub
Private Sub Suo()
  Form1.hide
  MsgBox "         您的密码是: 【" & Text1.Text & "】          " & Chr(13) & Chr(13) & "          请牢记!!!           "
  Form2.show
End Sub
