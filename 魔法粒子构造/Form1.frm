VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8655
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   16380
   LinkTopic       =   "Form1"
   ScaleHeight     =   8655
   ScaleWidth      =   16380
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "cherry_leaves"
      Height          =   255
      Index           =   38
      Left            =   3120
      TabIndex        =   85
      Top             =   5160
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   3
      Left            =   9480
      TabIndex        =   83
      Text            =   "1"
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   7560
      TabIndex        =   81
      Text            =   "-5"
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   7560
      TabIndex        =   80
      Text            =   "0"
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   7560
      TabIndex        =   79
      Text            =   "0"
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "dust 0 0 0 0.3"
      Height          =   255
      Index           =   37
      Left            =   8640
      TabIndex        =   78
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "squid_ink"
      Height          =   255
      Index           =   36
      Left            =   4680
      TabIndex        =   77
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton Command16 
      Caption         =   "右太极"
      Height          =   495
      Left            =   1560
      TabIndex        =   76
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Command15 
      Caption         =   "左太极"
      Height          =   495
      Left            =   1080
      TabIndex        =   75
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "wax_on"
      Height          =   255
      Index           =   30
      Left            =   5640
      TabIndex        =   74
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "dust_color_transition 0.9 0.9 0.9 1 1 1 1"
      Height          =   255
      Index           =   35
      Left            =   7080
      TabIndex        =   73
      Top             =   5280
      Width           =   4095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "snowflake"
      Height          =   255
      Index           =   34
      Left            =   4680
      TabIndex        =   72
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "scrape"
      Height          =   255
      Index           =   33
      Left            =   5640
      TabIndex        =   71
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "glow"
      Height          =   255
      Index           =   32
      Left            =   5640
      TabIndex        =   70
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "wax_off"
      Height          =   255
      Index           =   31
      Left            =   4680
      TabIndex        =   69
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton Command14 
      Caption         =   "录入长方体侧面"
      Height          =   495
      Left            =   1080
      TabIndex        =   68
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox Text16 
      Height          =   495
      Left            =   1920
      TabIndex        =   67
      Text            =   "0 1.5 15"
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox Text15 
      Height          =   495
      Left            =   1080
      TabIndex        =   66
      Text            =   "0 1.5 3"
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton Command13 
      Caption         =   "录入线段"
      Height          =   495
      Left            =   240
      TabIndex        =   65
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton Command12 
      Caption         =   "录入矩形点阵"
      Height          =   495
      Left            =   1080
      TabIndex        =   62
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox Text12 
      Height          =   495
      Left            =   2520
      TabIndex        =   61
      Text            =   "20"
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox Text11 
      Height          =   495
      Left            =   2040
      TabIndex        =   60
      Text            =   "8"
      Top             =   2880
      Width           =   495
   End
   Begin VB.CommandButton Command11 
      Caption         =   "录入圆柱侧面"
      Height          =   495
      Left            =   1080
      TabIndex        =   59
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Command9 
      Caption         =   "录入经纬球"
      Height          =   495
      Left            =   240
      TabIndex        =   58
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "录入笼球"
      Height          =   495
      Left            =   240
      TabIndex        =   57
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "dust 1 1 0.5 1"
      Height          =   255
      Index           =   29
      Left            =   8640
      TabIndex        =   56
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "dust 1 1 0.8 1"
      Height          =   255
      Index           =   28
      Left            =   8640
      TabIndex        =   55
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "dust 1 0 0 1"
      Height          =   255
      Index           =   27
      Left            =   7080
      TabIndex        =   54
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "dust 1 0.9 0.9 1"
      Height          =   255
      Index           =   26
      Left            =   7080
      TabIndex        =   53
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "dripping_lava"
      Height          =   255
      Index           =   24
      Left            =   3120
      TabIndex        =   52
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "dripping_water"
      Height          =   255
      Index           =   25
      Left            =   3120
      TabIndex        =   51
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "flame"
      Height          =   255
      Index           =   23
      Left            =   5640
      TabIndex        =   50
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "block minecraft:cobweb"
      Height          =   255
      Index           =   15
      Left            =   3120
      TabIndex        =   49
      Top             =   6240
      Width           =   3735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "item minecraft:diamond_sword"
      Height          =   255
      Index           =   22
      Left            =   7200
      TabIndex        =   48
      Top             =   6720
      Width           =   3135
   End
   Begin VB.CommandButton Command4 
      Caption         =   "block minecraft:magma_block"
      Height          =   255
      Index           =   21
      Left            =   3120
      TabIndex        =   47
      Top             =   6960
      Width           =   3735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "item minecraft:emerald"
      Height          =   255
      Index           =   20
      Left            =   7200
      TabIndex        =   46
      Top             =   6960
      Width           =   3135
   End
   Begin VB.CommandButton Command4 
      Caption         =   "block minecraft:glass"
      Height          =   255
      Index           =   19
      Left            =   3120
      TabIndex        =   45
      Top             =   6720
      Width           =   3735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "item minecraft:diamond"
      Height          =   255
      Index           =   18
      Left            =   7200
      TabIndex        =   44
      Top             =   6480
      Width           =   3135
   End
   Begin VB.CommandButton Command4 
      Caption         =   "block minecraft:structure_void"
      Height          =   255
      Index           =   17
      Left            =   3120
      TabIndex        =   43
      Top             =   6480
      Width           =   3735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "item minecraft:experience_bottle"
      Height          =   255
      Index           =   16
      Left            =   7200
      TabIndex        =   42
      Top             =   6240
      Width           =   3135
   End
   Begin VB.CommandButton Command4 
      Caption         =   "flash"
      Height          =   255
      Index           =   12
      Left            =   4680
      TabIndex        =   41
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "dust 0.9 1 1 1"
      Height          =   255
      Index           =   14
      Left            =   7080
      TabIndex        =   40
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "dust 0.5 0 0 1"
      Height          =   255
      Index           =   13
      Left            =   7080
      TabIndex        =   39
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "dust 0 0 0 1"
      Height          =   255
      Index           =   11
      Left            =   8640
      TabIndex        =   38
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "dust 1 0.8 0.8 1"
      Height          =   255
      Index           =   10
      Left            =   7080
      TabIndex        =   37
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "dust 1 1 1 1"
      Height          =   255
      Index           =   9
      Left            =   8640
      TabIndex        =   36
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "dust 1 0.8 1 1"
      Height          =   255
      Index           =   8
      Left            =   8640
      TabIndex        =   35
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "录入圆"
      Height          =   495
      Left            =   240
      TabIndex        =   34
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   3600
      TabIndex        =   32
      Text            =   "0.01 0.01 0.01"
      Top             =   840
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "soul"
      Height          =   255
      Index           =   7
      Left            =   4680
      TabIndex        =   31
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "electric_spark"
      Height          =   255
      Index           =   6
      Left            =   3120
      TabIndex        =   30
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "nautilus"
      Height          =   255
      Index           =   5
      Left            =   5640
      TabIndex        =   29
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "lava"
      Height          =   255
      Index           =   4
      Left            =   4680
      TabIndex        =   28
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "instant_effect"
      Height          =   255
      Index           =   3
      Left            =   3120
      TabIndex        =   27
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "end_rod"
      Height          =   255
      Index           =   2
      Left            =   5640
      TabIndex        =   26
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "enchant"
      Height          =   255
      Index           =   1
      Left            =   4680
      TabIndex        =   25
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "soul_fire_flame"
      Height          =   255
      Index           =   0
      Left            =   3120
      TabIndex        =   24
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   5280
      TabIndex        =   22
      Text            =   "0"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   5280
      TabIndex        =   20
      Text            =   "0.6"
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   3600
      TabIndex        =   18
      Text            =   "20"
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "垂直z轴"
      Height          =   495
      Left            =   7080
      TabIndex        =   17
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   3600
      TabIndex        =   16
      Text            =   "end_rod"
      Top             =   3240
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3600
      TabIndex        =   13
      Text            =   "0.3"
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "清空"
      Height          =   735
      Left            =   9600
      TabIndex        =   12
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      Caption         =   "录入八星"
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   8
      Left            =   1200
      TabIndex        =   9
      Text            =   "0"
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   7
      Left            =   1200
      TabIndex        =   8
      Text            =   "2"
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   6
      Left            =   1200
      TabIndex        =   7
      Text            =   "0"
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "随机名称"
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      Top             =   6960
      Width           =   615
   End
   Begin VB.TextBox Text9 
      Height          =   270
      Left            =   3360
      TabIndex        =   5
      Text            =   "D:\MC\1.20\.minecraft\saves\§dOlostland 1.0\datapacks\魔法构造\data\miracle\functions"
      Top             =   120
      Width           =   10455
   End
   Begin VB.TextBox Text8 
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Text            =   "miracle"
      Top             =   6960
      Width           =   1695
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   240
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "输出mcfunction"
      Height          =   615
      Left            =   840
      TabIndex        =   1
      Top             =   5760
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      Height          =   6615
      Left            =   11160
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   4575
   End
   Begin VB.Label Label11 
      Caption         =   "浓度0：点度变为速度矢量"
      Height          =   1095
      Left            =   6120
      TabIndex        =   86
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "放大率"
      Height          =   1095
      Index           =   2
      Left            =   9000
      TabIndex        =   84
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "目标矢量"
      Height          =   1095
      Index           =   1
      Left            =   7080
      TabIndex        =   82
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label10 
      Caption         =   "精度"
      Height          =   375
      Left            =   2520
      TabIndex        =   64
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "长度"
      Height          =   375
      Left            =   2040
      TabIndex        =   63
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label9 
      Caption         =   "点度"
      Height          =   375
      Left            =   3120
      TabIndex        =   33
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "浓度"
      Height          =   375
      Left            =   4800
      TabIndex        =   23
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "速度"
      Height          =   375
      Left            =   4800
      TabIndex        =   21
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "精度"
      Height          =   375
      Left            =   3120
      TabIndex        =   19
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "粒子"
      Height          =   375
      Left            =   3120
      TabIndex        =   15
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "半径"
      Height          =   375
      Left            =   3120
      TabIndex        =   14
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "中心"
      Height          =   1095
      Index           =   0
      Left            =   720
      TabIndex        =   10
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "路径"
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim moshi As Integer
Dim gg, txt As String
Const yan = 5





Private Sub Command1_Click()
moshi = (moshi + 1) Mod 3
If moshi = 0 Then Command1.Caption = "垂直z轴"
If moshi = 1 Then Command1.Caption = "垂直y轴"
If moshi = 2 Then Command1.Caption = "垂直x轴"
End Sub

Private Sub Command10_Click()
r = Val(Text2.Text)
x = Val(Text1(8).Text)
y = Val(Text1(7).Text)
z = Val(Text1(6).Text)
n = Val(Text4.Text)

For i = 1 To n
If moshi = 0 Then
X1 = x + r * (-1 + 2 * i / n)
Y1 = y + r * (1 - Abs(-1 + 2 * i / n))
z1 = z
End If
If moshi = 1 Then
X1 = x + r * (-1 + 2 * i / n)
z1 = z + r * (1 - Abs(-1 + 2 * i / n))
Y1 = y
End If
If moshi = 2 Then
z1 = z + r * (-1 + 2 * i / n)
Y1 = y + r * (1 - Abs(-1 + 2 * i / n))
X1 = x
End If
X1 = Int(X1 * 100000) / 100000
Y1 = Int(Y1 * 100000) / 100000
z1 = Int(z1 * 100000) / 100000
txt = txt + "particle minecraft:" + Text3.Text + " " + CStr(gg) + CStr(X1) + " " + CStr(gg) + CStr(Y1) + " " + CStr(gg) + CStr(z1) + " " + Text5.Text + " " + Text7.Text + " " + Text10.Text + vbCrLf
Next i

For i = 1 To n
If moshi = 0 Then
X1 = x + r * (-1 + 2 * i / n)
Y1 = y - r * (1 - Abs(-1 + 2 * i / n))
z1 = z
End If
If moshi = 1 Then
X1 = x + r * (-1 + 2 * i / n)
z1 = z - r * (1 - Abs(-1 + 2 * i / n))
Y1 = y
End If
If moshi = 2 Then
z1 = z + r * (-1 + 2 * i / n)
Y1 = y - r * (1 - Abs(-1 + 2 * i / n))
X1 = x
End If
X1 = Int(X1 * 100000) / 100000
Y1 = Int(Y1 * 100000) / 100000
z1 = Int(z1 * 100000) / 100000
txt = txt + "particle minecraft:" + Text3.Text + " " + CStr(gg) + CStr(X1) + " " + CStr(gg) + CStr(Y1) + " " + CStr(gg) + CStr(z1) + " " + Text5.Text + " " + Text7.Text + " " + Text10.Text + vbCrLf
Next i

n = n / 2

For i = 1 To n
If moshi = 0 Then
X1 = x + r * 0.707106
Y1 = y + r * (-0.707106 + 1.414213 * i / n)
z1 = z
End If
If moshi = 1 Then
X1 = x + r * 0.707106
z1 = z + r * (-0.707106 + 1.414213 * i / n)
Y1 = y
End If
If moshi = 2 Then
z1 = z + r * 0.707106
Y1 = y + r * (-0.707106 + 1.414213 * i / n)
X1 = x
End If
X1 = Int(X1 * 100000) / 100000
Y1 = Int(Y1 * 100000) / 100000
z1 = Int(z1 * 100000) / 100000
txt = txt + "particle minecraft:" + Text3.Text + " " + CStr(gg) + CStr(X1) + " " + CStr(gg) + CStr(Y1) + " " + CStr(gg) + CStr(z1) + " " + Text5.Text + " " + Text7.Text + " " + Text10.Text + vbCrLf
Next i

For i = 1 To n
If moshi = 0 Then
X1 = x - r * 0.707106
Y1 = y + r * (-0.707106 + 1.414213 * i / n)
z1 = z
End If
If moshi = 1 Then
X1 = x - r * 0.707106
z1 = z + r * (-0.707106 + 1.414213 * i / n)
Y1 = y
End If
If moshi = 2 Then
z1 = z - r * 0.707106
Y1 = y + r * (-0.707106 + 1.414213 * i / n)
X1 = x
End If
X1 = Int(X1 * 100000) / 100000
Y1 = Int(Y1 * 100000) / 100000
z1 = Int(z1 * 100000) / 100000
txt = txt + "particle minecraft:" + Text3.Text + " " + CStr(gg) + CStr(X1) + " " + CStr(gg) + CStr(Y1) + " " + CStr(gg) + CStr(z1) + " " + Text5.Text + " " + Text7.Text + " " + Text10.Text + vbCrLf
Next i

For i = 1 To n
If moshi = 0 Then
X1 = x + r * (-0.707106 + 1.414213 * i / n)
Y1 = y + r * 0.707106
z1 = z
End If
If moshi = 1 Then
X1 = x + r * (-0.707106 + 1.414213 * i / n)
z1 = z + r * 0.707106
Y1 = y
End If
If moshi = 2 Then
z1 = z + r * (-0.707106 + 1.414213 * i / n)
Y1 = y + r * 0.707106
X1 = x
End If
X1 = Int(X1 * 100000) / 100000
Y1 = Int(Y1 * 100000) / 100000
z1 = Int(z1 * 100000) / 100000
txt = txt + "particle minecraft:" + Text3.Text + " " + CStr(gg) + CStr(X1) + " " + CStr(gg) + CStr(Y1) + " " + CStr(gg) + CStr(z1) + " " + Text5.Text + " " + Text7.Text + " " + Text10.Text + vbCrLf
Next i

For i = 1 To n
If moshi = 0 Then
X1 = x + r * (-0.707106 + 1.414213 * i / n)
Y1 = y - r * 0.707106
z1 = z
End If
If moshi = 1 Then
X1 = x + r * (-0.707106 + 1.414213 * i / n)
z1 = z - r * 0.707106
Y1 = y
End If
If moshi = 2 Then
z1 = z + r * (-0.707106 + 1.414213 * i / n)
Y1 = y - r * 0.707106
X1 = x
End If
X1 = Int(X1 * 100000) / 100000
Y1 = Int(Y1 * 100000) / 100000
z1 = Int(z1 * 100000) / 100000
txt = txt + "particle minecraft:" + Text3.Text + " " + CStr(gg) + CStr(X1) + " " + CStr(gg) + CStr(Y1) + " " + CStr(gg) + CStr(z1) + " " + Text5.Text + " " + Text7.Text + " " + Text10.Text + vbCrLf
Next i


Text6.Text = txt
End Sub
Private Sub Command11_Click()
r = Val(Text2.Text)
x = Val(Text1(8).Text)
y = Val(Text1(7).Text)
z = Val(Text1(6).Text)
n = Val(Text4.Text)
h = Val(Text11.Text)
nm = Val(Text12.Text)

For j = 0 To nm
l = j / nm * h
For i = 1 To n
phi = i / n * 6.283185307
If moshi = 0 Then
X1 = x + r * Sin(phi)
Y1 = y + r * Cos(phi)
z1 = z - h / 2 + l
End If
If moshi = 1 Then
X1 = x + r * Sin(phi) * Sin(theta)
z1 = z + r * Cos(phi) * Sin(theta)
Y1 = y - h / 2 + l
End If
If moshi = 2 Then
z1 = z + r * Sin(phi)
Y1 = y + r * Cos(phi)
X1 = x - h / 2 + l
End If
X1 = Int(X1 * 100000) / 100000
Y1 = Int(Y1 * 100000) / 100000
z1 = Int(z1 * 100000) / 100000
txt = txt + "particle minecraft:" + Text3.Text + " " + CStr(gg) + CStr(X1) + " " + CStr(gg) + CStr(Y1) + " " + CStr(gg) + CStr(z1) + " " + Text5.Text + " " + Text7.Text + " " + Text10.Text + vbCrLf
Next i
Next j
Text6.Text = txt

End Sub

Private Sub Command12_Click()
r = Val(Text2.Text)
x = Val(Text1(8).Text)
y = Val(Text1(7).Text)
z = Val(Text1(6).Text)
n = Val(Text4.Text)
h = Val(Text11.Text)
nm = Val(Text12.Text)

For j = 1 To nm
l = j / nm * h
For i = 1 To n
ll = i / n * 2 * r
If moshi = 0 Then
X1 = x - r + ll
Y1 = y - h / 2 + l
z1 = z
End If
If moshi = 1 Then
X1 = x - r + ll
z1 = z - h / 2 + l
Y1 = y
End If
If moshi = 2 Then
z1 = z - r + ll
Y1 = y - h / 2 + l
X1 = x
End If
X1 = Int(X1 * 100000) / 100000
Y1 = Int(Y1 * 100000) / 100000
z1 = Int(z1 * 100000) / 100000
txt = txt + "particle minecraft:" + Text3.Text + " " + CStr(gg) + CStr(X1) + " " + CStr(gg) + CStr(Y1) + " " + CStr(gg) + CStr(z1) + " " + Text5.Text + " " + Text7.Text + " " + Text10.Text + vbCrLf
Next i
Next j
Text6.Text = txt

End Sub

Private Sub Command13_Click()
Dim a(0 To 2) As Integer

mark = 0
For i = 1 To Len(Text15.Text)
If Mid(Text15.Text, i, 1) = " " Then mark = mark + 1
a(mark) = i
Next i
x = Val(Mid(Text15.Text, 1, a(0)))
y = Val(Mid(Text15.Text, a(0) + 2, a(1) - a(0) - 1))
z = Val(Mid(Text15.Text, a(1) + 2, a(2) - a(1) - 1))

mark = 0
For i = 1 To Len(Text16.Text)
If Mid(Text16.Text, i, 1) = " " Then mark = mark + 1
a(mark) = i
Next i
delx = Val(Mid(Text16.Text, 1, a(0))) - x
dely = Val(Mid(Text16.Text, a(0) + 2, a(1) - a(0) - 1)) - y
delz = Val(Mid(Text16.Text, a(1) + 2, a(2) - a(1) - 1)) - z

n = Val(Text4.Text)

For i = 1 To n
X1 = x + i / n * delx
Y1 = y + i / n * dely
z1 = z + i / n * delz
X1 = Int(X1 * 100000) / 100000
Y1 = Int(Y1 * 100000) / 100000
z1 = Int(z1 * 100000) / 100000
txt = txt + "particle minecraft:" + Text3.Text + " " + CStr(gg) + CStr(X1) + " " + CStr(gg) + CStr(Y1) + " " + CStr(gg) + CStr(z1) + " " + Text5.Text + " " + Text7.Text + " " + Text10.Text + vbCrLf
Next i
Text6.Text = txt
End Sub

Private Sub Command14_Click()
r = Val(Text2.Text)
x = Val(Text1(8).Text)
y = Val(Text1(7).Text)
z = Val(Text1(6).Text)
n = Val(Text4.Text)
h = Val(Text11.Text)
nm = Val(Text12.Text)

For j = 0 To nm
l = j / nm * h
For i = 1 To n
ll = i / n * 2 * r
If moshi = 0 Then
X1 = x - r + ll
Y1 = y - r
z1 = z - h / 2 + l
End If
If moshi = 1 Then
X1 = x - r + ll
z1 = z - r
Y1 = y - h / 2 + l
End If
If moshi = 2 Then
z1 = z - r + ll
Y1 = y - r
X1 = x - h / 2 + l
End If
X1 = Int(X1 * 100000) / 100000
Y1 = Int(Y1 * 100000) / 100000
z1 = Int(z1 * 100000) / 100000
txt = txt + "particle minecraft:" + Text3.Text + " " + CStr(gg) + CStr(X1) + " " + CStr(gg) + CStr(Y1) + " " + CStr(gg) + CStr(z1) + " " + Text5.Text + " " + Text7.Text + " " + Text10.Text + vbCrLf
Next i
Next j

For j = 0 To nm
l = j / nm * h
For i = 1 To n
ll = i / n * 2 * r
If moshi = 0 Then
X1 = x - r + ll
Y1 = y + r
z1 = z - h / 2 + l
End If
If moshi = 1 Then
X1 = x - r + ll
z1 = z + r
Y1 = y - h / 2 + l
End If
If moshi = 2 Then
z1 = z - r + ll
Y1 = y + r
X1 = x - h / 2 + l
End If
X1 = Int(X1 * 100000) / 100000
Y1 = Int(Y1 * 100000) / 100000
z1 = Int(z1 * 100000) / 100000
txt = txt + "particle minecraft:" + Text3.Text + " " + CStr(gg) + CStr(X1) + " " + CStr(gg) + CStr(Y1) + " " + CStr(gg) + CStr(z1) + " " + Text5.Text + " " + Text7.Text + " " + Text10.Text + vbCrLf
Next i
Next j

For j = 0 To nm
l = j / nm * h
For i = 1 To n
ll = i / n * 2 * r
If moshi = 0 Then
X1 = x - r
Y1 = y - r + ll
z1 = z - h / 2 + l
End If
If moshi = 1 Then
X1 = x - r
z1 = z - r + ll
Y1 = y - h / 2 + l
End If
If moshi = 2 Then
z1 = z - r
Y1 = y - r + ll
X1 = x - h / 2 + l
End If
X1 = Int(X1 * 100000) / 100000
Y1 = Int(Y1 * 100000) / 100000
z1 = Int(z1 * 100000) / 100000
txt = txt + "particle minecraft:" + Text3.Text + " " + CStr(gg) + CStr(X1) + " " + CStr(gg) + CStr(Y1) + " " + CStr(gg) + CStr(z1) + " " + Text5.Text + " " + Text7.Text + " " + Text10.Text + vbCrLf
Next i
Next j

For j = 0 To nm
l = j / nm * h
For i = 1 To n
ll = i / n * 2 * r
If moshi = 0 Then
X1 = x + r
Y1 = y - r + ll
z1 = z - h / 2 + l
End If
If moshi = 1 Then
X1 = x + r
z1 = z - r + ll
Y1 = y - h / 2 + l
End If
If moshi = 2 Then
z1 = z + r
Y1 = y - r + ll
X1 = x - h / 2 + l
End If
X1 = Int(X1 * 100000) / 100000
Y1 = Int(Y1 * 100000) / 100000
z1 = Int(z1 * 100000) / 100000
txt = txt + "particle minecraft:" + Text3.Text + " " + CStr(gg) + CStr(X1) + " " + CStr(gg) + CStr(Y1) + " " + CStr(gg) + CStr(z1) + " " + Text5.Text + " " + Text7.Text + " " + Text10.Text + vbCrLf
Next i
Next j

Text6.Text = txt

End Sub

Private Sub Command15_Click()

r = Val(Text2.Text)
x = Val(Text1(8).Text)
y = Val(Text1(7).Text)
z = Val(Text1(6).Text)
n = Val(Text4.Text)


For j = 1 To n
l = j / n * 2 * r
For i = 1 To n
ll = i / n * 2 * r
xx = l - r
yy = ll - r
taiji = False
If (xx <= 0 And yy >= 0 And xx ^ 2 + (yy - r / 2) ^ 2 >= (r / yan) ^ 2 And xx ^ 2 + yy ^ 2 <= r ^ 2) Then taiji = True
If (xx <= 0 And yy <= 0 And xx ^ 2 + (yy + r / 2) ^ 2 >= (r / 2) ^ 2 And xx ^ 2 + yy ^ 2 <= r ^ 2) Then taiji = True
If (yy <= 0 And xx ^ 2 + (yy + r / 2) ^ 2 <= (r / yan) ^ 2) Then taiji = True
If (xx >= 0 And yy >= 0 And xx ^ 2 + (yy - r / 2) ^ 2 <= (r / 2) ^ 2) And xx ^ 2 + (yy - r / 2) ^ 2 >= (r / yan) ^ 2 Then taiji = True

If taiji Then
If moshi = 0 Then
X1 = x - r + ll
Y1 = y - r + l
z1 = z
End If
If moshi = 1 Then
X1 = x - r + ll
z1 = z - r + l
Y1 = y
End If
If moshi = 2 Then
z1 = z - r + ll
Y1 = y - r + l
X1 = x
End If
X1 = Int(X1 * 100000) / 100000
Y1 = Int(Y1 * 100000) / 100000
z1 = Int(z1 * 100000) / 100000
txt = txt + "particle minecraft:" + Text3.Text + " " + CStr(gg) + CStr(X1) + " " + CStr(gg) + CStr(Y1) + " " + CStr(gg) + CStr(z1) + " " + Text5.Text + " " + Text7.Text + " " + Text10.Text + vbCrLf
End If

Next i
Next j
Text6.Text = txt

End Sub

Private Sub Command16_Click()

r = Val(Text2.Text)
x = Val(Text1(8).Text)
y = Val(Text1(7).Text)
z = Val(Text1(6).Text)
n = Val(Text4.Text)


For j = 1 To n
l = j / n * 2 * r
For i = 1 To n
ll = i / n * 2 * r
xx = r - l
yy = r - ll
taiji = False
If (xx <= 0 And yy >= 0 And xx ^ 2 + (yy - r / 2) ^ 2 >= (r / yan) ^ 2 And xx ^ 2 + yy ^ 2 <= r ^ 2) Then taiji = True
If (xx <= 0 And yy <= 0 And xx ^ 2 + (yy + r / 2) ^ 2 >= (r / 2) ^ 2 And xx ^ 2 + yy ^ 2 <= r ^ 2) Then taiji = True
If (yy <= 0 And xx ^ 2 + (yy + r / 2) ^ 2 <= (r / yan) ^ 2) Then taiji = True
If (xx >= 0 And yy >= 0 And xx ^ 2 + (yy - r / 2) ^ 2 <= (r / 2) ^ 2) And xx ^ 2 + (yy - r / 2) ^ 2 >= (r / yan) ^ 2 Then taiji = True

If taiji Then
If moshi = 0 Then
X1 = x - r + ll
Y1 = y - r + l
z1 = z
End If
If moshi = 1 Then
X1 = x - r + ll
z1 = z - r + l
Y1 = y
End If
If moshi = 2 Then
z1 = z - r + ll
Y1 = y - r + l
X1 = x
End If
X1 = Int(X1 * 100000) / 100000
Y1 = Int(Y1 * 100000) / 100000
z1 = Int(z1 * 100000) / 100000
txt = txt + "particle minecraft:" + Text3.Text + " " + CStr(gg) + CStr(X1) + " " + CStr(gg) + CStr(Y1) + " " + CStr(gg) + CStr(z1) + " " + Text5.Text + " " + Text7.Text + " " + Text10.Text + vbCrLf
End If

Next i
Next j
Text6.Text = txt

End Sub

Private Sub Command2_Click()
abc = Text2.Text
For i = 1 To 5
abc = abc + Chr(Asc("a") + Int(26 * Rnd()))
Next i
abc = abc + CStr(Int(64000 * Rnd() - 32000))
Text8.Text = abc
End Sub


Private Sub Command3_Click()
txt = ""
Text6.Text = ""
End Sub






Private Sub Command4_Click(Index As Integer)
Text3.Text = Command4(Index).Caption
End Sub

Private Sub Command5_Click()
r = Val(Text2.Text)
x = Val(Text1(8).Text)
y = Val(Text1(7).Text)
z = Val(Text1(6).Text)
n = Val(Text4.Text)
nm = 4 * n
For i = 1 To nm
phi = i / nm * 6.283185307
If moshi = 0 Then
X1 = x + r * Sin(phi)
Y1 = y + r * Cos(phi)
z1 = z
End If
If moshi = 1 Then
X1 = x + r * Sin(phi)
z1 = z + r * Cos(phi)
Y1 = y
End If
If moshi = 2 Then
z1 = z + r * Sin(phi)
Y1 = y + r * Cos(phi)
X1 = x
End If
X1 = Int(X1 * 100000) / 100000
Y1 = Int(Y1 * 100000) / 100000
z1 = Int(z1 * 100000) / 100000

If Val(Text10.Text) = 0 And moshi = 0 Then
txt = txt + "particle minecraft:" + Text3.Text + " " + CStr(gg) + CStr(x) + " " + CStr(gg) + CStr(y) + " " + CStr(gg) + CStr(z) + " " + CStr(Int(r * Sin(phi) * 100000) / 100000) + " " + CStr(Int(r * Cos(phi) * 100000) / 100000) + " " + "0 " + Text7.Text + " " + Text10.Text + " force" + vbCrLf
ElseIf Val(Text10.Text) = 0 And moshi = 1 Then
txt = txt + "particle minecraft:" + Text3.Text + " " + CStr(gg) + CStr(x) + " " + CStr(gg) + CStr(y) + " " + CStr(gg) + CStr(z) + " " + CStr(Int(r * Sin(phi) * 100000) / 100000) + " " + "0 " + CStr(Int(r * Cos(phi) * 100000) / 100000) + " " + Text7.Text + " " + Text10.Text + " force" + vbCrLf
ElseIf Val(Text10.Text) = 0 And moshi = 2 Then
txt = txt + "particle minecraft:" + Text3.Text + " " + CStr(gg) + CStr(x) + " " + CStr(gg) + CStr(y) + " " + CStr(gg) + CStr(z) + " " + "0 " + CStr(Int(r * Sin(phi) * 100000) / 100000) + " " + CStr(Int(r * Cos(phi) * 100000) / 100000) + " " + Text7.Text + " " + Text10.Text + " force" + vbCrLf
Else
txt = txt + "particle minecraft:" + Text3.Text + " " + CStr(gg) + CStr(X1) + " " + CStr(gg) + CStr(Y1) + " " + CStr(gg) + CStr(z1) + " " + Text5.Text + " " + Text7.Text + " " + Text10.Text + vbCrLf
End If
Next i
Text6.Text = txt
End Sub

Private Sub Command6_Click()

sss = "\" + Text8.Text + ".mcfunction"
  Open Text9.Text & sss For Output As #1
  Print #1, txt 'ssssss
  Close #1

End Sub

Private Sub Command7_Click()
r = Val(Text2.Text)
x = Val(Text1(8).Text)
y = Val(Text1(7).Text)
z = Val(Text1(6).Text)
n = Val(Text4.Text)

If Val(Text1(3).Text) = 0 Then

For j = 0 To n
theta = j / n * 3.14150926
nn = n
If j = 0 Or j = n Then nn = 2
If j = 1 Or j = n - 1 Then nn = Int(Sqr(n)) * 2
If j = 2 Or j = n - 2 Then nn = Int(Sqr(n)) * 3
For i = 1 To 2 * nn
phi = i / (2 * nn) * 6.283185307
If moshi = 0 Then
X1 = x + r * Sin(phi) * Sin(theta)
Y1 = y + r * Cos(phi) * Sin(theta)
z1 = z + r * Cos(theta)
End If
If moshi = 1 Then
X1 = x + r * Sin(phi) * Sin(theta)
z1 = z + r * Cos(phi) * Sin(theta)
Y1 = y + r * Cos(theta)
End If
If moshi = 2 Then
z1 = z + r * Sin(phi) * Sin(theta)
Y1 = y + r * Cos(phi) * Sin(theta)
X1 = x + r * Cos(theta)
End If
X1 = Int(X1 * 100000) / 100000
Y1 = Int(Y1 * 100000) / 100000
z1 = Int(z1 * 100000) / 100000

If Val(Text10.Text) = 0 And moshi = 0 Then
txt = txt + "particle minecraft:" + Text3.Text + " " + CStr(gg) + CStr(x) + " " + CStr(gg) + CStr(y) + " " + CStr(gg) + CStr(z) + " " + CStr(Int(r * Sin(phi) * Sin(theta) * 100000) / 100000) + " " + CStr(Int(r * Cos(phi) * Sin(theta) * 100000) / 100000) + " " + CStr(Int(r * Cos(theta) * 100000) / 100000) + " " + Text7.Text + " " + Text10.Text + " force" + vbCrLf
ElseIf Val(Text10.Text) = 0 And moshi = 1 Then
txt = txt + "particle minecraft:" + Text3.Text + " " + CStr(gg) + CStr(x) + " " + CStr(gg) + CStr(y) + " " + CStr(gg) + CStr(z) + " " + CStr(Int(r * Sin(phi) * Sin(theta) * 100000) / 100000) + " " + CStr(Int(r * Cos(theta) * 100000) / 100000) + " " + CStr(Int(r * Cos(phi) * Sin(theta) * 100000) / 100000) + " " + Text7.Text + " " + Text10.Text + " force" + vbCrLf
ElseIf Val(Text10.Text) = 0 And moshi = 2 Then
txt = txt + "particle minecraft:" + Text3.Text + " " + CStr(gg) + CStr(x) + " " + CStr(gg) + CStr(y) + " " + CStr(gg) + CStr(z) + " " + CStr(Int(r * Cos(theta) * Sin(theta) * 100000) / 100000) + " " + CStr(Int(r * Sin(phi) * Sin(theta) * 100000) / 100000) + " " + CStr(Int(r * Cos(phi) * Sin(theta) * 100000) / 100000) + " " + Text7.Text + " " + Text10.Text + " force" + vbCrLf
Else
txt = txt + "particle minecraft:" + Text3.Text + " " + CStr(gg) + CStr(X1) + " " + CStr(gg) + CStr(Y1) + " " + CStr(gg) + CStr(z1) + " " + Text5.Text + " " + Text7.Text + " " + Text10.Text + vbCrLf
End If
Next i
Next j

Else


X2 = Val(Text1(0).Text)
Y2 = Val(Text1(1).Text)
z2 = Val(Text1(2).Text)
mag = Val(Text1(3).Text)

For j = 0 To n
theta = j / n * 3.14150926
nn = n
If j = 0 Or j = n Then nn = 2
If j = 1 Or j = n - 1 Then nn = Int(Sqr(n)) * 2
If j = 2 Or j = n - 2 Then nn = Int(Sqr(n)) * 3
For i = 1 To 2 * nn
phi = i / (2 * nn) * 6.283185307
X1 = x + r * Sin(phi) * Sin(theta)
Y1 = y + r * Cos(phi) * Sin(theta)
z1 = z + r * Cos(theta)

X1 = Int(X1 * 100000) / 100000
Y1 = Int(Y1 * 100000) / 100000
z1 = Int(z1 * 100000) / 100000

txt = txt + "particle minecraft:" + Text3.Text + " " + CStr(gg) + CStr(X1) + " " + CStr(gg) + CStr(Y1) + " " + CStr(gg) + CStr(z1) + " " + CStr((Int(r * Sin(phi) * Sin(theta) * 100000) / 100000) * (mag - 1) + X2) + " " + CStr((Int(r * Cos(phi) * Sin(theta) * 100000) / 100000) * (mag - 1) + Y2) + " " + CStr((Int(r * Cos(theta) * 100000) / 100000) * (mag - 1) + z2) + " " + Text7.Text + " " + "0" + " force" + vbCrLf



Next i
Next j

End If


Text6.Text = txt
End Sub


Private Sub Command8_Click()
If Command8.Caption = "~" Then Command8.Caption = "^" Else: Command8.Caption = "~"
gg = Command8.Caption

End Sub



Private Sub Command9_Click()
r = Val(Text2.Text)
x = Val(Text1(8).Text)
y = Val(Text1(7).Text)
z = Val(Text1(6).Text)
n = Val(Text4.Text)
n = n / 3
For j = 0 To 3 * n
theta = j / (3 * n) * 3.14150926
nn = n
If j = 0 Or j = n Then nn = 2
If j = 1 Or j = n - 1 Then nn = Int(Sqr(n)) * 2
If j = 2 Or j = n - 2 Then nn = Int(Sqr(n)) * 3
If j Mod 3 = 0 Then nn = 3 * nn
For i = 1 To 2 * nn
phi = i / (2 * nn) * 6.283185307
If moshi = 0 Then
X1 = x + r * Sin(phi) * Sin(theta)
Y1 = y + r * Cos(phi) * Sin(theta)
z1 = z + r * Cos(theta)
End If
If moshi = 1 Then
X1 = x + r * Sin(phi) * Sin(theta)
z1 = z + r * Cos(phi) * Sin(theta)
Y1 = y + r * Cos(theta)
End If
If moshi = 2 Then
z1 = z + r * Sin(phi) * Sin(theta)
Y1 = y + r * Cos(phi) * Sin(theta)
X1 = x + r * Cos(theta)
End If
X1 = Int(X1 * 100000) / 100000
Y1 = Int(Y1 * 100000) / 100000
z1 = Int(z1 * 100000) / 100000
txt = txt + "particle minecraft:" + Text3.Text + " " + CStr(gg) + CStr(X1) + " " + CStr(gg) + CStr(Y1) + " " + CStr(gg) + CStr(z1) + " " + Text5.Text + " " + Text7.Text + " " + Text10.Text + vbCrLf
Next i
Next j
Text6.Text = txt
End Sub

Private Sub Form_Load()


Command8.Caption = "^"
gg = Command8.Caption


End Sub

Private Sub Timer1_Timer()
Text6.Text = txt
End Sub

