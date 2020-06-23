VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "小游戏"
   ClientHeight    =   3915
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5955
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   5955
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox Text5 
      Height          =   1575
      Left            =   1800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   2040
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   9
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1800
      TabIndex        =   7
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1800
      TabIndex        =   6
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton Command6 
      Caption         =   "清空提示框"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "清除数字"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "显示答案"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "退出"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "确定"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始游戏/重置"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4560
      TabIndex        =   11
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Dim h As Integer, j As Integer
h = 0
j = 0
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Then
Form2.Show
Else
Open "D:\smallgame\游戏.txt" For Input As #1
d = Input(1, #1)
e = Input(1, #1)
f = Input(1, #1)
g = Input(1, #1)
If d <> Text1.Text Then
h = h
Else: h = h + 1
End If
If e <> Text2.Text Then
h = h
Else: h = h + 1
End If
If f <> Text3.Text Then
h = h
Else: h = h + 1
End If
If g <> Text4.Text Then
h = h
Else: h = h + 1
End If
If d = Text1.Text Or d = Text2.Text Or d = Text3.Text Or d = Text4.Text Then
j = j + 1
End If
If e = Text1.Text Or e = Text2.Text Or e = Text3.Text Or e = Text4.Text Then
j = j + 1
End If
If f = Text1.Text Or f = Text2.Text Or f = Text3.Text Or f = Text4.Text Then
j = j + 1
End If
If g = Text1.Text Or g = Text2.Text Or g = Text3.Text Or g = Text4.Text Then
j = j + 1
End If
Text5.Text = Text5.Text & "有" & h & "个位置对了，" & "包含" & j & "个相同数字" & vbCrLf
End If
Close
End Sub
Private Sub Command3_Click()
End
End Sub
Private Sub Command1_Click()
Dim w As Integer
Dim x As Integer
Dim y As Integer
Dim z As Integer
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Label1 = ""
Text5.Text = ""
Dim b As String
a = a + 1
Randomize
If a = 0 Then
Command2.Enabled = False
Else
Command2.Enabled = True
End If
w = Int(Rnd() * 9 + 1)
x = Int(Rnd() * 9 + 1)
y = Int(Rnd() * 9 + 1)
z = Int(Rnd() * 9 + 1)
Do While w = x Or w = y Or w = z Or x = y Or x = z Or y = z
w = Int(Rnd() * 9 + 1)
x = Int(Rnd() * 9 + 1)
y = Int(Rnd() * 9 + 1)
z = Int(Rnd() * 9 + 1)
Loop
b = w * 1000 + x * 100 + y * 10 + z
Open "D:\smallgame\game.txt" For Output As #1
Print #1, b
Close
End Sub
Private Sub Command4_Click()
Dim i As Single
Open "D:\smallgame\game.txt" For Input As #1
i = Input(4, #1)
Label1 = i
Close
End Sub
Private Sub Command5_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
End Sub
Private Sub Command6_Click()
Label1 = ""
Text5.Text = ""
End Sub
Private Sub Form_Load()
Dim a As Integer
If a = 0 Then
Command2.Enabled = False
Else
Command2.Enabled = True
End If
End Sub
