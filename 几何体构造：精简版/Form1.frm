VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7005
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   12720
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   12720
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "fill"
      Height          =   495
      Left            =   2160
      TabIndex        =   35
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "structure_void"
      Height          =   375
      Index           =   7
      Left            =   8040
      TabIndex        =   34
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "glass"
      Height          =   375
      Index           =   3
      Left            =   6960
      TabIndex        =   33
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "iron_block"
      Height          =   375
      Index           =   2
      Left            =   6960
      TabIndex        =   32
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "netherite_block"
      Height          =   375
      Index           =   1
      Left            =   5880
      TabIndex        =   31
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "air"
      Height          =   375
      Index           =   0
      Left            =   5880
      TabIndex        =   30
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   29
      Text            =   "Form1.frx":0000
      Top             =   4320
      Width           =   7935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "obsidian"
      Height          =   375
      Index           =   6
      Left            =   6960
      TabIndex        =   28
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "glass"
      Height          =   375
      Index           =   5
      Left            =   6960
      TabIndex        =   27
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "sea_lantern"
      Height          =   375
      Index           =   4
      Left            =   6960
      TabIndex        =   26
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "light"
      Height          =   375
      Index           =   3
      Left            =   5880
      TabIndex        =   25
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "netherite_block"
      Height          =   375
      Index           =   2
      Left            =   5880
      TabIndex        =   24
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "air"
      Height          =   375
      Index           =   1
      Left            =   8040
      TabIndex        =   23
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "填充物名称"
      Height          =   495
      Left            =   2400
      TabIndex        =   22
      Top             =   6000
      Width           =   1335
   End
   Begin VB.TextBox Text9 
      Height          =   270
      Left            =   3360
      TabIndex        =   21
      Text            =   "D:\MC\.minecraft\saves\虚空\datapacks\魔法构造\data\create\functions"
      Top             =   120
      Width           =   10455
   End
   Begin VB.CommandButton Command9 
      Caption         =   "替换模式"
      Height          =   615
      Left            =   3720
      TabIndex        =   19
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox Text8 
      Height          =   495
      Left            =   600
      TabIndex        =   18
      Text            =   "create"
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   375
      Left            =   2160
      TabIndex        =   17
      Top             =   240
      Width           =   495
   End
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   600
      TabIndex        =   16
      Text            =   "20"
      Top             =   5400
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "上下限"
      Height          =   495
      Left            =   1440
      TabIndex        =   15
      Top             =   5400
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "输出方式"
      Height          =   615
      Left            =   840
      TabIndex        =   14
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "iron_block"
      Height          =   375
      Index           =   0
      Left            =   5880
      TabIndex        =   13
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      Height          =   6615
      Left            =   10440
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   960
      Width           =   4575
   End
   Begin VB.TextBox Text5 
      Height          =   615
      Left            =   4440
      TabIndex        =   9
      Text            =   "air"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   4440
      TabIndex        =   7
      Text            =   "light"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "构造"
      Height          =   495
      Left            =   840
      TabIndex        =   6
      Top             =   3600
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   5
      Left            =   840
      TabIndex        =   5
      Text            =   "23"
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   4
      Left            =   840
      TabIndex        =   4
      Text            =   "3"
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   3
      Left            =   840
      TabIndex        =   3
      Text            =   "23"
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   840
      TabIndex        =   2
      Text            =   "-23"
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   840
      TabIndex        =   1
      Text            =   "-6"
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Text            =   "-23"
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "路径"
      Height          =   255
      Left            =   2880
      TabIndex        =   20
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "扫描终值点"
      Height          =   1095
      Left            =   360
      TabIndex        =   11
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "扫描起始点"
      Height          =   1095
      Left            =   360
      TabIndex        =   10
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "填充物"
      Height          =   615
      Left            =   3720
      TabIndex        =   8
      Top             =   1080
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xia(1 To 166410), shang(1 To 166410) As Long
Dim x0(0 To 5), x, y, z, m, k, e(0 To 17) As Long
Dim a(0 To 23)  As Double
Dim hh, ff, outtxt, huimie As Boolean
Dim t5, gg As String



Private Sub Command1_Click()

t5 = Text5.Text

If t5 <> "" And huimie = False Then t5 = " replace " + t5
If huimie = True Then t5 = " destroy"

For i = 0 To 5
x0(i) = Val(Text1(i).Text)
If i < 3 And x0(i) < -128 Then x0(i) = -128
If i > 2 And x0(i) > 128 Then x0(i) = 128
Next i

txt = ""

For i = 1 To 3
If x0(i - 1) > x0(i + 2) Then
gg = x0(i - 1)
x0(i - 1) = x0(i + 2)
x0(i + 2) = gg
End If
Next i

Text2.Text = "执行计算中"

k = 1
m = x0(3) - x0(0) + 1
ff = True
step1 = 3
For i = 1 To m * (x0(4) - x0(1) + 1) Step step1 '大头戏登场
x = x0(0) + (i - 1) Mod m
y = x0(1) + Int((i - 1) / m)
ff = True
'mo = 10
xia(k) = x0(2)
  For z = x0(2) To x0(5) Step step1
  hh = ff
  'ff为是否满足条件；是则有方块，无则没方块
    'ff = False
    'L = Int(Sqr(x ^ 2 + z ^ 2))
    'Mx = x Mod mo
    'My = y Mod mo
    'Mz = z Mod mo
    'D = Int(Sqr(x ^ 2 + z ^ 2 + (y + 2) ^ 2))
    'S = (Abs(x + z) + Abs(x - z) - Abs(Abs(x + z) - Abs(x - z))) / 2
    'G = (Abs(x) + Abs(z) - Abs(Abs(x) - Abs(z))) / 2
    'D2 = 0
    'If Mx = 0 And Mz = 0 Then ff = True
    'If L <= 23 Then ff = True '输入xyz对应条件
    If Command3.Caption = "setblock" And ff = True Then txt = txt + "setblock " + CStr(gg) + CStr(x) + " " + CStr(gg) + CStr(y) + " " + CStr(gg) + CStr(z) + " " + Text4.Text + t5 + vbCrLf
  If hh = False And ff = True Then xia(k) = z
  If z = x0(5) Then ff = False
  If (hh = True And ff = False) Then
  shang(k) = z - 1
  k = k + 1
  If Command3.Caption = "fill" Then txt = txt + "fill " + CStr(gg) + CStr(x) + " " + CStr(gg) + CStr(y) + " " + CStr(gg) + CStr(xia(k - 1)) + " " + CStr(gg) + CStr(x) + " " + CStr(gg) + CStr(y) + " " + CStr(gg) + CStr(shang(k - 1)) + " " + Text4.Text + t5 + vbCrLf
  End If
Next z
Next i

 If outtxt = False Then
  Text6.Text = txt
 Else
  sss = "\" + Text8.Text + ".mcfunction"
  Open Text9.Text & sss For Output As #1
  Print #1, txt 'ssssss
  Close #1
 End If

Text2.Text = "共输出" + CStr(k - 1) + "条命令"

If outtxt = False Then
  Text2.Text = Text2.Text + vbCrLf + "本次输出至文本框中展示，不予直接生成mcfunction"
Else
  Text2.Text = Text2.Text + vbCrLf + "本次输出生成mcfunction," + "地址为" + Text9.Text & sss
End If
 


End Sub




Private Sub Command2_Click()

Text8.Text = Text4.Text
End Sub

Private Sub Command3_Click()
If Command3.Caption = "fill" Then
Command3.Caption = "setblock"
Else
Command3.Caption = "fill"
End If

End Sub

Private Sub Command4_Click(Index As Integer)
Text5.Text = Command4(Index).Caption
End Sub



Private Sub Command5_Click(Index As Integer)
Text4.Text = Command5(Index).Caption
End Sub

Private Sub Command6_Click()
outtxt = Not outtxt
If outtxt = False Then Command6.Caption = "输出到文本框"
If outtxt = True Then Command6.Caption = "只输出mcfunction"
End Sub

Private Sub Command7_Click()
For i = 0 To 5
If i < 3 Then Text1(i).Text = "-" + CStr(Abs(Val(Text7.Text)))
If i > 2 Then Text1(i).Text = CStr(Abs(Val(Text7.Text)))
Next i
Text7.Text = Val(Text7.Text) Mod 30 + 5
End Sub

Private Sub Command8_Click()
If Command8.Caption = "~" Then Command8.Caption = "^" Else: Command8.Caption = "~"
gg = Command8.Caption

End Sub

Private Sub Command9_Click()
huimie = Not huimie
If huimie = False Then Command9.Caption = "替换模式" Else Command9.Caption = "毁灭模式"
End Sub

Private Sub Form_Load()

Call Command2_Click


Command8.Caption = "~"
gg = Command8.Caption

For i = 0 To 2
    Text1(i).Text = "-100"
Next i

For i = 3 To 5
    Text1(i).Text = "100"
Next i

Text1(1).Text = "45"
Text1(4).Text = "45"

Text9.Text = "D:\MC\1.20\.minecraft\saves\§dOlostland 1.0\datapacks\魔法构造\data\create\functions"

outtxt = True
Command6.Caption = "输出mcfunction"

End Sub

