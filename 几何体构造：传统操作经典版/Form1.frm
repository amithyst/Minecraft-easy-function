VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7005
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   13950
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   13950
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text9 
      Height          =   270
      Left            =   3360
      TabIndex        =   87
      Text            =   "D:\MC\.minecraft\saves\虚空\datapacks\魔法构造\data\create\functions"
      Top             =   120
      Width           =   10455
   End
   Begin VB.CommandButton Command12 
      Caption         =   "空气"
      Height          =   375
      Left            =   9720
      TabIndex        =   85
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton Command11 
      Caption         =   "标准构造：空心抛物面"
      Height          =   495
      Left            =   7920
      TabIndex        =   84
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton Command10 
      Caption         =   "标准构造：抛物面"
      Height          =   495
      Left            =   7920
      TabIndex        =   83
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton Command9 
      Caption         =   "替换模式"
      Height          =   615
      Left            =   7680
      TabIndex        =   82
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox Text8 
      Height          =   495
      Left            =   600
      TabIndex        =   81
      Text            =   "create"
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   375
      Left            =   2160
      TabIndex        =   80
      Top             =   240
      Width           =   495
   End
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   600
      TabIndex        =   79
      Text            =   "20"
      Top             =   5400
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "上下限"
      Height          =   495
      Left            =   1440
      TabIndex        =   78
      Top             =   5400
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   615
      Left            =   840
      TabIndex        =   77
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "铁块"
      Height          =   495
      Left            =   9720
      TabIndex        =   76
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "空气"
      Height          =   495
      Left            =   9720
      TabIndex        =   75
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "标准构造：空心球"
      Height          =   495
      Left            =   7920
      TabIndex        =   74
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "标准构造：球"
      Height          =   495
      Left            =   7920
      TabIndex        =   73
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      Height          =   6615
      Left            =   10440
      MultiLine       =   -1  'True
      TabIndex        =   72
      Top             =   960
      Width           =   4575
   End
   Begin VB.TextBox Text5 
      Height          =   615
      Left            =   8400
      TabIndex        =   69
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   8400
      TabIndex        =   67
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Index           =   17
      Left            =   5880
      TabIndex        =   66
      Text            =   "Text3"
      Top             =   4800
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Index           =   16
      Left            =   4680
      TabIndex        =   65
      Text            =   "Text3"
      Top             =   4800
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Index           =   15
      Left            =   3480
      TabIndex        =   64
      Text            =   "Text3"
      Top             =   4800
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   23
      Left            =   6480
      TabIndex        =   63
      Text            =   "Text2"
      Top             =   5040
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   22
      Left            =   5040
      TabIndex        =   61
      Text            =   "Text2"
      Top             =   5040
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   21
      Left            =   3840
      TabIndex        =   59
      Text            =   "Text2"
      Top             =   5040
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   20
      Left            =   2640
      TabIndex        =   57
      Text            =   "Text2"
      Top             =   5040
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Index           =   14
      Left            =   5880
      TabIndex        =   56
      Text            =   "Text3"
      Top             =   3960
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Index           =   13
      Left            =   4680
      TabIndex        =   55
      Text            =   "Text3"
      Top             =   3960
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Index           =   12
      Left            =   3480
      TabIndex        =   54
      Text            =   "Text3"
      Top             =   3960
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   19
      Left            =   6480
      TabIndex        =   53
      Text            =   "Text2"
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   18
      Left            =   5040
      TabIndex        =   51
      Text            =   "Text2"
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   17
      Left            =   3840
      TabIndex        =   49
      Text            =   "Text2"
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   16
      Left            =   2640
      TabIndex        =   47
      Text            =   "Text2"
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Index           =   11
      Left            =   5880
      TabIndex        =   46
      Text            =   "Text3"
      Top             =   3120
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Index           =   10
      Left            =   4680
      TabIndex        =   45
      Text            =   "Text3"
      Top             =   3120
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Index           =   9
      Left            =   3480
      TabIndex        =   44
      Text            =   "Text3"
      Top             =   3120
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   15
      Left            =   6480
      TabIndex        =   43
      Text            =   "Text2"
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   14
      Left            =   5040
      TabIndex        =   41
      Text            =   "Text2"
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   13
      Left            =   3840
      TabIndex        =   39
      Text            =   "Text2"
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   12
      Left            =   2640
      TabIndex        =   37
      Text            =   "Text2"
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Index           =   8
      Left            =   5880
      TabIndex        =   36
      Text            =   "Text3"
      Top             =   2280
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Index           =   7
      Left            =   4680
      TabIndex        =   35
      Text            =   "Text3"
      Top             =   2280
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Index           =   6
      Left            =   3480
      TabIndex        =   34
      Text            =   "Text3"
      Top             =   2280
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   11
      Left            =   6480
      TabIndex        =   33
      Text            =   "Text2"
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   10
      Left            =   5040
      TabIndex        =   31
      Text            =   "Text2"
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   9
      Left            =   3840
      TabIndex        =   29
      Text            =   "Text2"
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   8
      Left            =   2640
      TabIndex        =   27
      Text            =   "Text2"
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Index           =   5
      Left            =   5880
      TabIndex        =   26
      Text            =   "Text3"
      Top             =   1440
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Index           =   4
      Left            =   4680
      TabIndex        =   25
      Text            =   "Text3"
      Top             =   1440
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Index           =   3
      Left            =   3480
      TabIndex        =   24
      Text            =   "Text3"
      Top             =   1440
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   7
      Left            =   6480
      TabIndex        =   23
      Text            =   "Text2"
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   6
      Left            =   5040
      TabIndex        =   21
      Text            =   "Text2"
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   5
      Left            =   3840
      TabIndex        =   19
      Text            =   "Text2"
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   4
      Left            =   2640
      TabIndex        =   17
      Text            =   "Text2"
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Index           =   2
      Left            =   5880
      TabIndex        =   16
      Text            =   "Text3"
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Index           =   1
      Left            =   4680
      TabIndex        =   15
      Text            =   "Text3"
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Index           =   0
      Left            =   3480
      TabIndex        =   14
      Text            =   "Text3"
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   3
      Left            =   6480
      TabIndex        =   13
      Text            =   "Text2"
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   2
      Left            =   5040
      TabIndex        =   11
      Text            =   "Text2"
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   1
      Left            =   3840
      TabIndex        =   9
      Text            =   "Text2"
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   0
      Left            =   2640
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   840
      Width           =   735
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
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   4
      Left            =   840
      TabIndex        =   4
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   3
      Left            =   840
      TabIndex        =   3
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   840
      TabIndex        =   2
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   840
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "路径"
      Height          =   255
      Left            =   2880
      TabIndex        =   86
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "扫描终值点"
      Height          =   1095
      Left            =   360
      TabIndex        =   71
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "扫描起始点"
      Height          =   1095
      Left            =   360
      TabIndex        =   70
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "填充物"
      Height          =   615
      Left            =   7680
      TabIndex        =   68
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "z >="
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
      Index           =   17
      Left            =   5760
      TabIndex        =   62
      Top             =   5040
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "y +"
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
      Index           =   16
      Left            =   4560
      TabIndex        =   60
      Top             =   5040
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "x +"
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
      Index           =   15
      Left            =   3360
      TabIndex        =   58
      Top             =   5040
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "z >="
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
      Index           =   14
      Left            =   5760
      TabIndex        =   52
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "y +"
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
      Index           =   13
      Left            =   4560
      TabIndex        =   50
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "x +"
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
      Index           =   12
      Left            =   3360
      TabIndex        =   48
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "z >="
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
      Index           =   11
      Left            =   5760
      TabIndex        =   42
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "y +"
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
      Index           =   10
      Left            =   4560
      TabIndex        =   40
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "x +"
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
      Index           =   9
      Left            =   3360
      TabIndex        =   38
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "z >="
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
      Index           =   8
      Left            =   5760
      TabIndex        =   32
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "y +"
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
      Index           =   7
      Left            =   4560
      TabIndex        =   30
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "x +"
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
      Index           =   6
      Left            =   3360
      TabIndex        =   28
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "z >="
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
      Index           =   5
      Left            =   5760
      TabIndex        =   22
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "y +"
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
      Index           =   4
      Left            =   4560
      TabIndex        =   20
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "x +"
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
      Index           =   3
      Left            =   3360
      TabIndex        =   18
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "z >="
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
      Index           =   2
      Left            =   5760
      TabIndex        =   12
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "y +"
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
      Index           =   1
      Left            =   4560
      TabIndex        =   10
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "x +"
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
      Index           =   0
      Left            =   3360
      TabIndex        =   8
      Top             =   840
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
If x0(i - 1) >= x0(i + 2) Then
gg = x0(i - 1)
x0(i - 1) = x0(i + 2)
x0(i + 2) = gg
End If
Next i

For i = 1 To 24
a(i - 1) = Val(Text2(i - 1).Text)
If i < 19 Then e(i - 1) = Val(Text3(i - 1).Text)
Next i

k = 1
m = x0(3) - x0(0) + 1

For i = 1 To m * (x0(4) - x0(1) + 1) '大头戏登场
x = x0(0) + (i - 1) Mod m
y = x0(1) + Int((i - 1) / m)
ff = False
xia(k) = x0(2)
  For z = x0(2) To x0(5)
  hh = ff
  ff = True
    For j = 1 To 6
    If a(4 * j - 4) * x ^ e(3 * j - 3) + a(4 * j - 3) * y ^ e(3 * j - 2) + a(4 * j - 2) * z ^ e(3 * j - 1) < a(4 * j - 1) Then ff = False
    Next j
  If hh = False And ff = True Then xia(k) = z
  If z = x0(5) Then ff = False
  If (hh = True And ff = False) Then
  shang(k) = z - 1
  k = k + 1
txt = txt + "fill " + CStr(gg) + CStr(x) + " " + CStr(gg) + CStr(y) + " " + CStr(gg) + CStr(xia(k - 1)) + " " + CStr(gg) + CStr(x) + " " + CStr(gg) + CStr(y) + " " + CStr(gg) + CStr(shang(k - 1)) + " " + Text4.Text + t5 + vbCrLf
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





End Sub

Private Sub Command10_Click()

m = 3
For i = 1 To 2
Text2(i - 5 + 4 * m).Text = "-0.4"
Text3(i - 4 + 3 * m).Text = "2"
Next i
Text2(i - 5 + 4 * m).Text = "-1"
Text3(i - 4 + 3 * m).Text = "1"
i = 4
Text2(i - 5 + 4 * m).Text = "-10"


End Sub

Private Sub Command11_Click()

m = 4
For i = 1 To 2
Text2(i - 5 + 4 * m).Text = "0.15"
Text3(i - 4 + 3 * m).Text = "2"
Next i
Text2(i - 5 + 4 * m).Text = "1"
Text3(i - 4 + 3 * m).Text = "1"
i = 4
Text2(i - 5 + 4 * m).Text = "2"

End Sub

Private Sub Command12_Click()
Text4.Text = "air"
End Sub

Private Sub Command2_Click()

m = 1
For i = 1 To 3
Text2(i - 5 + 4 * m).Text = "-1"
Text3(i - 4 + 3 * m).Text = "2"
Next i
Text2(i - 5 + 4 * m).Text = "-90"

End Sub

Private Sub Command3_Click()

m = 2
For i = 1 To 3
Text2(i - 5 + 4 * m).Text = "1"
Text3(i - 4 + 3 * m).Text = "2"
Next i
Text2(i - 5 + 4 * m).Text = "70"

End Sub

Private Sub Command4_Click()
Text5.Text = "air"
End Sub

Private Sub Command5_Click()
Text4.Text = "iron_block"
End Sub

Private Sub Command6_Click()
outtxt = Not outtxt
If outtxt = False Then Command6.Caption = "输出文本框"
If outtxt = True Then Command6.Caption = "输出mcfunction"
End Sub

Private Sub Command7_Click()
For i = 0 To 5
If i < 3 Then Text1(i).Text = "-" + CStr(Abs(Val(Text7.Text)))
If i > 2 Then Text1(i).Text = CStr(Abs(Val(Text7.Text)))
Next i
Text7.Text = Val(Text7.Text) Mod 40 + 10
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

Command8.Caption = "~"
gg = Command8.Caption

For i = 1 To 24
Text2(i - 1).Text = "0"
If i < 19 Then Text3(i - 1).Text = "1"
Next i

For i = 0 To 5
If i < 3 Then Text1(i).Text = "-10"
If i > 2 Then Text1(i).Text = "10"
Next i

outtxt = False
Command6.Caption = "输出文本框"

End Sub

