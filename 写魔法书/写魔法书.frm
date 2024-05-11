VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8310
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15615
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   15615
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   11640
      Top             =   6720
   End
   Begin VB.CommandButton Command27 
      Caption         =   "一键更新"
      Height          =   855
      Left            =   11640
      TabIndex        =   115
      Top             =   5400
      Width           =   495
   End
   Begin VB.CommandButton Command26 
      Caption         =   "删除本书"
      Height          =   855
      Left            =   11520
      TabIndex        =   112
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton Command25 
      Caption         =   "导入本页"
      Height          =   255
      Left            =   13800
      TabIndex        =   111
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command24 
      Caption         =   "转移"
      Height          =   855
      Left            =   10800
      TabIndex        =   110
      Top             =   6120
      Width           =   495
   End
   Begin VB.CommandButton Command23 
      Caption         =   "插入空页"
      Height          =   615
      Left            =   8880
      TabIndex        =   109
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton Command22 
      Caption         =   "导入下页"
      Height          =   255
      Left            =   13800
      TabIndex        =   108
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton Command21 
      Caption         =   "导出"
      Height          =   255
      Left            =   8880
      TabIndex        =   107
      Top             =   3240
      Width           =   495
   End
   Begin VB.TextBox Text14 
      Height          =   270
      Left            =   9360
      TabIndex        =   106
      Text            =   "D:\VB\MC—程序\写魔法书\书本目录\godbook1.txt"
      Top             =   3240
      Width           =   4455
   End
   Begin VB.TextBox Text13 
      Height          =   375
      Left            =   6240
      TabIndex        =   104
      Top             =   3960
      Width           =   495
   End
   Begin VB.CommandButton Command20 
      Caption         =   "翻页"
      Height          =   495
      Left            =   4440
      TabIndex        =   103
      Top             =   4440
      Width           =   375
   End
   Begin VB.CommandButton Command9 
      Caption         =   "下一页"
      Height          =   495
      Left            =   3960
      TabIndex        =   102
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton Command19 
      Caption         =   "写入下面并换行"
      Height          =   855
      Left            =   8400
      TabIndex        =   101
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton Command18 
      Caption         =   "覆盖这并换行"
      Height          =   855
      Left            =   7920
      TabIndex        =   100
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton Command17 
      Caption         =   "/trigger trigger set 1 "
      Height          =   255
      Index           =   2
      Left            =   4920
      TabIndex        =   99
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton Command17 
      Caption         =   "/function give:reload"
      Height          =   255
      Index           =   1
      Left            =   4920
      TabIndex        =   98
      Top             =   4320
      Width           =   2055
   End
   Begin VB.CommandButton Command17 
      Height          =   255
      Index           =   0
      Left            =   4920
      TabIndex        =   97
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton Command16 
      Caption         =   "换行"
      Height          =   495
      Left            =   7440
      TabIndex        =   96
      Top             =   4440
      Width           =   495
   End
   Begin VB.ListBox List2 
      Height          =   2760
      Left            =   12600
      TabIndex        =   95
      Top             =   4200
      Width           =   2295
   End
   Begin VB.CommandButton Command15 
      Caption         =   "下一元素"
      Height          =   495
      Left            =   3480
      TabIndex        =   94
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton Command14 
      Caption         =   "上一元素"
      Height          =   495
      Left            =   3000
      TabIndex        =   93
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton Command13 
      Caption         =   "上一页"
      Height          =   495
      Left            =   2520
      TabIndex        =   92
      Top             =   4440
      Width           =   495
   End
   Begin VB.TextBox Text12 
      Height          =   375
      Left            =   840
      TabIndex        =   91
      Text            =   "@p"
      Top             =   3360
      Width           =   2535
   End
   Begin VB.CommandButton Command12 
      Caption         =   "entity"
      Height          =   375
      Left            =   120
      TabIndex        =   90
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Index           =   5
      Left            =   4320
      TabIndex        =   88
      Text            =   "key.inventory"
      Top             =   3600
      Width           =   4575
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Index           =   4
      Left            =   4320
      TabIndex        =   86
      Text            =   "block.minecraft.netherite_block"
      Top             =   3240
      Width           =   4575
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Index           =   3
      Left            =   4320
      TabIndex        =   84
      Text            =   "@e[type=zombie]"
      Top             =   2880
      Width           =   4575
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Index           =   2
      Left            =   3960
      TabIndex        =   82
      Text            =   "{\""name\"":\""*\"",\""objective\"":\""armor\""}"
      Top             =   2520
      Width           =   4935
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Index           =   1
      Left            =   3960
      TabIndex        =   80
      Text            =   "Attributes"
      Top             =   2160
      Width           =   4935
   End
   Begin VB.TextBox Text11 
      Height          =   735
      Left            =   6600
      TabIndex        =   78
      Top             =   120
      Width           =   1215
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   255
      Left            =   6360
      Max             =   255
      TabIndex        =   74
      Top             =   1440
      Value           =   255
      Width           =   2415
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   6360
      Max             =   255
      TabIndex        =   73
      Top             =   1200
      Value           =   255
      Width           =   2415
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   6360
      Max             =   255
      TabIndex        =   72
      Top             =   960
      Value           =   255
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "white"
      Height          =   375
      Index           =   15
      Left            =   1320
      TabIndex        =   71
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "black"
      Height          =   375
      Index           =   14
      Left            =   840
      TabIndex        =   70
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "light_purple"
      Height          =   375
      Index           =   13
      Left            =   1800
      TabIndex        =   69
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "blue"
      Height          =   375
      Index           =   12
      Left            =   4200
      TabIndex        =   68
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "green"
      Height          =   375
      Index           =   11
      Left            =   4800
      TabIndex        =   67
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "gray"
      Height          =   375
      Index           =   10
      Left            =   5880
      TabIndex        =   66
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "aqua"
      Height          =   375
      Index           =   9
      Left            =   3000
      TabIndex        =   65
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "yellow"
      Height          =   375
      Index           =   8
      Left            =   5400
      TabIndex        =   64
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "gold"
      Height          =   375
      Index           =   7
      Left            =   5400
      TabIndex        =   63
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "dark_blue"
      Height          =   375
      Index           =   6
      Left            =   4200
      TabIndex        =   62
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "red"
      Height          =   375
      Index           =   5
      Left            =   3600
      TabIndex        =   61
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "dark_aqua"
      Height          =   375
      Index           =   4
      Left            =   3000
      TabIndex        =   60
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "dark_purple"
      Height          =   375
      Index           =   3
      Left            =   1800
      TabIndex        =   59
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "dark_green"
      Height          =   375
      Index           =   2
      Left            =   4800
      TabIndex        =   58
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "dark_gray"
      Height          =   375
      Index           =   1
      Left            =   5880
      TabIndex        =   57
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Command11 
      Caption         =   "添加文本"
      Height          =   735
      Left            =   8040
      TabIndex        =   56
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command10 
      Caption         =   "覆盖这一元素"
      Height          =   855
      Left            =   8880
      TabIndex        =   55
      Top             =   4080
      Width           =   495
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Index           =   1
      Left            =   4080
      TabIndex        =   53
      Text            =   "远古神明"
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Index           =   0
      Left            =   4080
      TabIndex        =   51
      Text            =   "旧日秘经"
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      Caption         =   "写入下一元素"
      Height          =   855
      Left            =   9360
      TabIndex        =   49
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "删除本页"
      Height          =   855
      Left            =   10320
      TabIndex        =   48
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "删除本元素"
      Height          =   855
      Left            =   9840
      TabIndex        =   47
      Top             =   4080
      Width           =   495
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   45
      Text            =   "1"
      Top             =   4440
      Width           =   615
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   43
      Text            =   "1"
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "挤入本元素后"
      Height          =   855
      Left            =   10800
      TabIndex        =   42
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Height          =   375
      Index           =   14
      Left            =   3000
      TabIndex        =   41
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "false"
      Height          =   375
      Index           =   13
      Left            =   2400
      TabIndex        =   40
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "true"
      Height          =   375
      Index           =   12
      Left            =   1800
      TabIndex        =   39
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Height          =   375
      Index           =   11
      Left            =   3000
      TabIndex        =   38
      Top             =   2400
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "false"
      Height          =   375
      Index           =   10
      Left            =   2400
      TabIndex        =   37
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "true"
      Height          =   375
      Index           =   9
      Left            =   1800
      TabIndex        =   36
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Height          =   375
      Index           =   8
      Left            =   3000
      TabIndex        =   35
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "false"
      Height          =   375
      Index           =   7
      Left            =   2400
      TabIndex        =   34
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "true"
      Height          =   375
      Index           =   6
      Left            =   1800
      TabIndex        =   33
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Height          =   375
      Index           =   5
      Left            =   3000
      TabIndex        =   32
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "false"
      Height          =   375
      Index           =   4
      Left            =   2400
      TabIndex        =   31
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "true"
      Height          =   375
      Index           =   3
      Left            =   1800
      TabIndex        =   30
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Height          =   375
      Index           =   2
      Left            =   3000
      TabIndex        =   29
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "false"
      Height          =   375
      Index           =   1
      Left            =   2400
      TabIndex        =   28
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "true"
      Height          =   375
      Index           =   0
      Left            =   1800
      TabIndex        =   27
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   600
      TabIndex        =   25
      Top             =   3960
      Width           =   4215
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   840
      TabIndex        =   24
      Text            =   "D:\VB\MC—程序\实体构造\物品目录\trade\trade.txt"
      Top             =   6120
      Width           =   4815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "导入标签路径"
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   6120
      Width           =   735
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   6120
      TabIndex        =   22
      Top             =   6120
      Width           =   4095
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Index           =   0
      Left            =   3960
      TabIndex        =   20
      Text            =   "附近的实体有"
      Top             =   1800
      Width           =   4935
   End
   Begin VB.ListBox List1 
      Height          =   2760
      Left            =   8880
      TabIndex        =   19
      Top             =   360
      Width           =   5895
   End
   Begin VB.TextBox Text4 
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   18
      Top             =   5040
      Width           =   11175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   5
      Left            =   1200
      TabIndex        =   17
      Top             =   2880
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   4
      Left            =   1320
      TabIndex        =   15
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   3
      Left            =   1080
      TabIndex        =   13
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   720
      TabIndex        =   11
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   720
      TabIndex        =   9
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "dark_red"
      Height          =   375
      Index           =   0
      Left            =   3600
      TabIndex        =   7
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   6
      Text            =   "真随机呀"
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   8640
      TabIndex        =   3
      Text            =   "book"
      Top             =   6600
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   6600
      Width           =   6735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "构造"
      Height          =   855
      Left            =   10320
      TabIndex        =   0
      Top             =   6120
      Width           =   495
   End
   Begin VB.Label Label13 
      Caption         =   "书本显示形态"
      Height          =   255
      Left            =   13080
      TabIndex        =   114
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label12 
      Caption         =   "书本源码形态"
      Height          =   255
      Left            =   10920
      TabIndex        =   113
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label11 
      Caption         =   "页码跳转"
      Height          =   375
      Left            =   5760
      TabIndex        =   105
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label10 
      Caption         =   "keybind"
      Height          =   255
      Index           =   5
      Left            =   3480
      TabIndex        =   89
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label10 
      Caption         =   "translate"
      Height          =   255
      Index           =   4
      Left            =   3480
      TabIndex        =   87
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "selector"
      Height          =   255
      Index           =   3
      Left            =   3480
      TabIndex        =   85
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "分数"
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   83
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label10 
      Caption         =   "nbt"
      Height          =   255
      Index           =   1
      Left            =   3480
      TabIndex        =   81
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label10 
      Caption         =   "文本"
      Height          =   255
      Index           =   0
      Left            =   3480
      TabIndex        =   79
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "蓝：255"
      Height          =   255
      Index           =   2
      Left            =   5520
      TabIndex        =   77
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "绿：255"
      Height          =   255
      Index           =   1
      Left            =   5520
      TabIndex        =   76
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "红：255"
      Height          =   255
      Index           =   0
      Left            =   5520
      TabIndex        =   75
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "光标"
      Height          =   375
      Left            =   120
      TabIndex        =   54
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label Label8 
      Caption         =   "作者"
      Height          =   375
      Index           =   1
      Left            =   3480
      TabIndex        =   52
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "标题"
      Height          =   375
      Index           =   0
      Left            =   3480
      TabIndex        =   50
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "元素"
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   46
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "页"
      Height          =   255
      Index           =   0
      Left            =   1320
      TabIndex        =   44
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label Label6 
      Caption         =   "添加指令"
      Height          =   375
      Left            =   120
      TabIndex        =   26
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "其他属性"
      Height          =   375
      Left            =   5760
      TabIndex        =   21
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "obfuscated"
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   16
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "strikethrough"
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "underlined"
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "italic"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "bold"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "color"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "名称"
      Height          =   375
      Left            =   7920
      TabIndex        =   4
      Top             =   6600
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "路径"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   6600
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(1 To 100000), b(1 To 100000), atxt(1 To 100000), txt, page(1 To 1000), s, sss, xxxl, s1 As String
Dim wr(1 To 100000) As Boolean '一千页，一百行
Dim yixie, guangbiao, ksksk As Long
Dim out, red, green, blue, nbts, x1 As Integer
Dim time0 As Integer



Private Sub Command1_Click()

ppag = ""

For i = yixie To 1 Step -1
If wr(i) = True Then Exit For
Next
yixie = i

For p = 1 To p0(yixie)
  page(p) = ""
  For j = 100 To 1 Step -1
    i = l(p, j)
    If wr(i) = True Then Exit For
  Next
  For i = l(p, 1) To l(p, j)
    b(i) = a(i)
    If b(i) = "" Then b(i) = "{\""text\"":\""\""}"
    page(p) = page(p) + "," + b(i)
  Next
  page(p) = """[" + Mid(page(p), 2) + "]"""
  ppag = ppag + "," + page(p)
Next


ppag = Mid(ppag, 2)
nbt = "{""title"":""" + Text10(0).Text + """,""author"":""" + Text10(1).Text + """,pages:[" + ppag + "]"
If Text6.Text <> "" Then nbt = nbt + "," + Text6.Text
nbt = nbt + "}"

sss = "give @p written_book" + nbt
s = Text2.Text & "\" + Text3.Text + ".mcfunction"


's1 = "D:\VB\temp\" + Text3.Text + ".mcfunction"
s1 = "D:\VB\temp\temp.mcfunction"


Open s1 For Output As #1
Print #1, sss 'ssssss
Close #1

Text4.Text = sss

'/c /k /q

Shell "cmd /c D:\VB\temp\change.py"


End Sub

Private Sub Command10_Click()
Call sentence
End Sub

Private Sub Command11_Click()
out = (out + 1) Mod 6
If out = 0 Then Command11.Caption = "添加文本"
If out = 1 Then Command11.Caption = "添加路径标签"
If out = 2 Then Command11.Caption = "添加分数"
If out = 3 Then Command11.Caption = "添加selector"
If out = 4 Then Command11.Caption = "添加translate"
If out = 5 Then Command11.Caption = "添加keybind"
End Sub

Private Sub Command12_Click()
nbts = (nbts + 1) Mod 3
If nbts = 0 Then
Command12.Caption = "entity"
Text12.Text = "@e[limit=1,distance=1..10]"
Text5(1).Text = "Health"
ElseIf nbts = 1 Then
Command12.Caption = "block"
Text12.Text = "~ ~-1 ~"
Text5(1).Text = "Item"
ElseIf nbts = 2 Then
Command12.Caption = "storage"
Text12.Text = "text:a"
Text5(1).Text = "a"
End If
End Sub

Private Sub Command13_Click()
guangbiao = dada(guangbiao - 100, 1)
kjkjk = guangbiao
Text9(0).Text = CStr(p0(kjkjk))
Text9(1).Text = CStr(l0(kjkjk))
End Sub

Private Sub Command14_Click()
guangbiao = dada(guangbiao - 1, 1)
kjkjk = guangbiao
Text9(0).Text = CStr(p0(kjkjk))
Text9(1).Text = CStr(l0(kjkjk))
End Sub

Private Sub Command15_Click()
guangbiao = guangbiao + 1
kjkjk = guangbiao
Text9(0).Text = CStr(p0(kjkjk))
Text9(1).Text = CStr(l0(kjkjk))
End Sub

Private Sub Command16_Click() '------------------------------------------------------------换行----------------------------
Call enter
End Sub

Public Sub enter()
Call jiru
wr(guangbiao) = True
a(guangbiao) = "{\""text\"":\""\\n\""}"
atxt(guangbiao) = "\\n"
Call LiST
Text9(0).Text = CStr(p0(guangbiao))
Text9(1).Text = CStr(l0(guangbiao))
End Sub


Private Sub Command17_Click(Index As Integer)
Text8.Text = Command17(Index).Caption
End Sub







Private Sub Command18_Click()
Call sentence
Call enter
End Sub

Private Sub Command19_Click()
guangbiao = guangbiao + 1
Call sentence
Call enter
End Sub

Private Sub Command2_Click(Index As Integer)
Text1(0).Text = Command2(Index).Caption
End Sub

Private Sub Command20_Click()
guangbiao = l(p0(guangbiao) + 1, 1)
Text9(0).Text = CStr(p0(guangbiao))
Text9(1).Text = CStr(1)
End Sub

Private Sub Command21_Click()
Open Text14.Text For Output As #1
For i = 1 To yixie
If wr(i) = True Then
Print #1, CStr(i)
Print #1, atxt(i)
Print #1, a(i)
End If
Next
Close #1
End Sub

Private Sub Command22_Click() '----------------------------------------------------------导入-----------------------------
Open Text14.Text For Input As #1
chushi = 0
wenben = ""
xulie = 1
Do While Not EOF(1)
Line Input #1, wenben
If chushi Mod 3 = 0 Then
xulie = Val(wenben)
wr(p0(guangbiao) * 100 + xulie) = True
End If
If chushi Mod 3 = 1 Then atxt(p0(guangbiao) * 100 + xulie) = wenben
If chushi Mod 3 = 2 Then a(p0(guangbiao) * 100 + xulie) = wenben
chushi = chushi + 1
'Text4.Text = Text4.Text + wenben + ">>>>" + Str(xulie) + a(101) + b(101) + "<<<<"
wenben = ""
Loop
Close #1
yixie = yixie + xulie + 100
'Text4.Text = CStr(yixie)
Call LiST
End Sub

Private Sub Command23_Click()
ksksk = guangbiao

For i = yixie + 1 To guangbiao Step -1
If wr(i) = True Then Exit For
Next

k = l(p0(i) + 1, 100)
'Text4.Text = Text4.Text + Str(k)--------------------------------------------------调试栏----------------------------
'Text4.Text = Text4.Text + Str(yixie)

For i = k To l(p0(guangbiao) + 1, 1) Step -1
a(i + 100) = a(i)
atxt(i + 100) = atxt(i)
wr(i + 100) = wr(i)
Next
If yixie < k + 100 Then yixie = k + 100

For jkjk = l(p0(guangbiao) + 1, 1) To l(p0(guangbiao) + 1, 100)
wr(jkjk) = False
a(jkjk) = ""
atxt(jkjk) = ""
Next

'Text4.Text = Text4.Text + Str(yixie)
guangbiao = ksksk
Call LiST
End Sub

Private Sub Command24_Click()

s = Text2.Text & "\" + Text3.Text + ".mcfunction"


's1 = "D:\VB\temp\" + Text3.Text + ".mcfunction"
s1 = "D:\VB\temp\temp.mcfunction"
Shell "cmd /c D:\VB\temp\change.py"

FileCopy s1, s
End Sub

Private Sub Command25_Click()
Open Text14.Text For Input As #1
chushi = 0
wenben = ""
xulie = 1
Do While Not EOF(1)
Line Input #1, wenben
If chushi Mod 3 = 0 Then
xulie = Val(wenben)
wr(p0(guangbiao) * 100 - 100 + xulie) = True
End If
If chushi Mod 3 = 1 Then atxt(p0(guangbiao) * 100 - 100 + xulie) = wenben
If chushi Mod 3 = 2 Then a(p0(guangbiao) * 100 - 100 + xulie) = wenben
chushi = chushi + 1
'Text4.Text = Text4.Text + wenben + ">>>>" + Str(xulie) + a(101) + b(101) + "<<<<"
wenben = ""
Loop
Close #1
yixie = yixie + xulie
'Text4.Text = CStr(yixie)
Call LiST
End Sub

Private Sub Command26_Click()
For i = yixie To 1 Step -1
wr(i) = False
a(i) = ""
atxt(i) = ""
Next

guangbiao = 1
guangbiao = dada(guangbiao - 1, 1)
kjkjk = guangbiao

Text9(0).Text = CStr(p0(kjkjk))
Text9(1).Text = CStr(l0(kjkjk))

Call LiST

End Sub

Private Sub Command27_Click()
Timer1.Enabled = True
time0 = 0
End Sub

Private Sub Command3_Click()
Open Text7.Text For Input As #1
Do While Not EOF(1)
Line Input #1, jy
Loop
Close #1

For i = 1 To 60
If Mid(jy, i, 6) = ",tag:{" Then
jy = Mid(jy, i + 6)
Exit For
End If
Next
jy = Mid(jy, 1, Len(jy) - 2)
Text6.Text = jy
End Sub

Private Sub Command4_Click(Index As Integer)
m = Index \ 3 + 1
Text1(m).Text = Command4(Index).Caption
End Sub

Public Sub sentence() '_________________________________________________------------------------------写入一个元素
If yixie <= guangbiao Then yixie = guangbiao
wr(guangbiao) = True
Call fresh
a(guangbiao) = txt
atxt(guangbiao) = xxxl
Call LiST
Text9(0).Text = CStr(p0(guangbiao))
Text9(1).Text = CStr(l0(guangbiao))
End Sub

Public Sub delsentence() '_________________________________________________------------------------------删除一个元素------------
wr(guangbiao) = False
a(guangbiao) = ""
atxt(guangbiao) = ""
Call LiST
Text9(0).Text = CStr(p0(guangbiao))
Text9(1).Text = CStr(l0(guangbiao))
End Sub

Public Function p0(x) As Integer '------------------------------------------------------计算光标对应的页码-----------------
p0 = (x - 1) \ 100 + 1
End Function
Public Function l0(x) As Integer '------------------------------------------------------计算光标对应的单页上的元素序数-----------------
l0 = (x - 1) Mod 100 + 1
End Function
Public Function l(ye, hang) As Long '------------------------------------------------------计算页码对应的光标-----------------
l = (ye - 1) * 100 + hang
End Function
Public Function tianjia(x) As String
If x = 0 Then tianjia = "text"
If x = 1 Then tianjia = "nbt"
If x = 2 Then tianjia = "score"
If x = 3 Then tianjia = "selector"
If x = 4 Then tianjia = "translate"
If x = 5 Then tianjia = "keybind"
End Function
Public Function sy(x As String) As String '_________________________________________________打双引号-----------------------------
If out = 2 Then
sy = x
Else
sy = "\""" + x + "\"""
End If
End Function
Public Function sljz(x) As String '---------------------------------------转成十六进制--------------------------------------
Dim sl(0 To 1) As Integer
Dim sll(0 To 1) As String
sl(0) = x Mod 16
sl(1) = x \ 16
For i = 0 To 1
If sl(i) >= 0 And sl(i) <= 9 Then sll(i) = CStr(sl(i))
If sl(i) >= 10 Then sll(i) = Chr(sl(i) + 55)
Next
sljz = sll(1) + sll(0)

End Function
Public Function dada(x, y) As Long '---------------------------------------取较大数--------------------------------------
dada = ((x + y) + Abs(x - y)) / 2
End Function

Public Sub fresh() '_______________________________________刷新刷新刷新刷新刷新刷新刷新刷新_______________txt句子元素刷新-------------
Dim xxx As String
xxx = Text5(out).Text
If out = 0 And Text5(0).Text = "真随机呀" Then
Randomize
X2 = Int(Rnd() * 3)
If x1 = 0 And X2 = 0 Then xxx = "hi，我是你的智慧助手\\n"
If x1 = 0 And X2 = 1 Then xxx = "尊贵的客户您好，我是您的智慧手机助手\\n"
If x1 = 0 And X2 = 2 Then xxx = "Hello!I'm your guider! \\n"
If x1 = 1 And X2 = 0 Then xxx = "欢迎使用新的iphone系列\\n"
If x1 = 1 And X2 = 1 Then xxx = "欢迎购买全新版本的iphone手机\\n"
If x1 = 1 And X2 = 2 Then xxx = "欢迎您使用这款新的iphone系列\\n"
If x1 = 2 And X2 = 0 Then xxx = "本系列手机采用魔能驱动\\n"
If x1 = 2 And X2 = 1 Then xxx = "本系列产品是魔法造物\\n"
If x1 = 2 And X2 = 2 Then xxx = "本系列是炼金物品\\n"
If x1 = 3 Then xxx = "您的操作系统已是当前的最新版本\\n"
If x1 = 4 Then xxx = "当前版本号为iPh10.7.9\\n"
If x1 = 5 And X2 = 0 Then xxx = "在此由衷感谢您对我们的支持！\\n为此，我们给您提供了免费会员机会！\\n"
If x1 = 5 And X2 = 1 Then xxx = "于此向您对我们的支持表示感谢！\\n为此，您将有一次免费会员抽奖机会！\\n"
If x1 = 5 And X2 = 2 Then xxx = "感谢您的购买与信任！\\n为此，您将有机会免费成为会员！\\n"
If x1 = 6 And X2 = 0 Then xxx = "恭喜您成为普通注册会员\\n"
If x1 = 6 And X2 = 1 Then xxx = "恭喜您抽中白金会员\\n"
If x1 = 6 And X2 = 2 Then xxx = "恭喜您抽中超级会员\\n"
If x1 = 7 Then
zcm = ""
For i = 1 To 16
zcm = zcm + Chr(50 + Int(70 * Rnd()))
Next
xx = "您的注册码为" + zcm + "\\n"
End If
If x1 = 8 Then xxx = "本产品最后更新时间为" + Str(Now()) + "\\n"
If x1 = 9 Then xxx = "最后，欢迎来到“乐园”！\\n享受吧！\\n"
x1 = x1 + 1
End If

col = Text1(0).Text
If out = 0 And Text1(0).Text = "真随机呀" Then
Randomize
col = "#" + sljz(Int(256 * Rnd())) + sljz(Int(256 * Rnd())) + sljz(Int(256 * Rnd()))
End If

txt = "{\""" + tianjia(out) + "\"":" + sy(xxx)
If Text1(0).Text <> "" Then txt = txt + ",\""color\"":\""" + col + "\"""
For i = 1 To 5
If Text1(i).Text <> "" Then txt = txt + ",\""" + Label1(i).Caption + "\"":\""" + Text1(i).Text + "\"""
Next
If out = 1 Then txt = txt + ",\""" + Command12.Caption + "\"":\""" + Text12.Text + "\"""
If Text8.Text <> "" Then txt = txt + ",\""clickEvent\"":{\""action\"":\""run_command\"",\""value\"":\""" + Text8.Text + "\""}"
If Text13.Text <> "" Then txt = txt + ",\""clickEvent\"":{\""action\"":\""change_page\"",\""value\"":" + Text13.Text + "}"
txt = txt + "}"

xxxl = xxx
End Sub

Public Sub LiST() '----------------------------------------------------------------------列表刷新与预览--------------------------
List1.Clear
List2.Clear
aaatxt = ""

For i = 1 To yixie
If wr(i) = True Then

If p0(i) > kkk Then
List1.AddItem ""
List1.AddItem "第" + CStr(p0(i)) + "页内容如下"
List2.AddItem aaatxt
aaatxt = ""
List2.AddItem ""
List2.AddItem "第" + CStr(p0(i)) + "页"
End If
kkk = p0(i)
List1.AddItem "第" + CStr(l0(i)) + "元素：" + a(i)
End If

For j = 1 To Len(atxt(i))
If Mid(atxt(i), j, 3) = "\\n" Then
List2.AddItem aaatxt
aaatxt = ""
ElseIf Mid(atxt(i), dada(j - 1, 1), 3) <> "\\n" And Mid(atxt(i), dada(j - 2, 1), 3) <> "\\n" Then
aaatxt = aaatxt + Mid(atxt(i), j, 1)
End If
Next

Next
List2.AddItem aaatxt
End Sub

Private Sub Command5_Click()
Call jiru
Call sentence
End Sub

Public Sub jiru() '-----------------------------------------------------------------挤进去-------------------------------
guangbiao = guangbiao + 1
For i = guangbiao To yixie + 1
If wr(i) = False Then Exit For
Next
k = i
wr(k) = True
For i = k To guangbiao + 1 Step -1
a(i) = a(i - 1)
atxt(i) = atxt(i - 1)
Next
If yixie < k Then yixie = k
Call delsentence
End Sub



Private Sub Command6_Click() '--------------------------------------------------删除一个元素--------------------------------=
Call delsentence
End Sub

Private Sub Command7_Click() '----------------------------------------------删除一页--------------------------------------------------
For guangbiao = l(p0(guangbiao), 100) To l(p0(guangbiao), 1) Step -1
wr(guangbiao) = False
a(guangbiao) = ""
atxt(guangbiao) = ""
Next
guangbiao = guangbiao + 1
End Sub

Private Sub Command8_Click()
guangbiao = guangbiao + 1
Call sentence
End Sub

Private Sub Command9_Click()
guangbiao = guangbiao + 100
kjkjk = guangbiao
Text9(0).Text = CStr(p0(kjkjk))
Text9(1).Text = CStr(l0(kjkjk))
End Sub

Private Sub Form_Load()
red = 255
green = 255
blue = 255
guangbiao = 1
yixie = 3

Text2.Text = "D:\MC\1.20\.minecraft\saves\§dOlostland 1.0\datapacks\魔法构造\data\give\functions"

Text6.Text = "display:{Name:'{""text"":""旧日秘经"",""color"":""yellow"",""bold"":true}',Lore:['{""text"":""ID:O_000"",""color"":""gold"",""bold"":false}','{""text"":""无上 超神器/先天灵宝/道器"",""color"":""aqua"",""bold"":true}','{""text"":""古老旧日的造物权柄，蕴含着无法理解的力量，被魔法女神锻造成了一本书"",""color"":""yellow"",""bold"":true}','{""text"":""「造物回响」："",""color"":""yellow"",""bold"":true}','{""text"":""祂昔在，今在，祂亦将永在！"",""color"":""dark_red"",""bold"":true}','{""text"":""祂是唯一，祂亦是万有！"",""color"":""dark_red"",""bold"":true}']}"
End Sub

Private Sub HScroll1_Change()
red = HScroll1.Value
Label4(0).Caption = "红：" + CStr(red)
Call colour
End Sub

Private Sub HScroll2_Change()
green = HScroll2.Value
Label4(1).Caption = "绿：" + CStr(green)
Call colour
End Sub

Private Sub HScroll3_Change()
blue = HScroll3.Value
Label4(2).Caption = "蓝：" + CStr(blue)
Call colour
End Sub
Public Sub colour() '------------------------------------------------------可爱的改变颜色环节---------------------------
Text11.BackColor = RGB(red, green, blue)
Text1(0).Text = "#" + sljz(red) + sljz(green) + sljz(blue)
End Sub



Private Sub Text9_Change(Index As Integer)
guangbiao = l(Val(Text9(0).Text), Val(Text9(1).Text))
End Sub

Private Sub Timer1_Timer()
time0 = time0 + 1
If time0 = 1 Then Call Command26_Click
If time0 = 2 Then Call Command25_Click
If time0 = 3 Then Call Command1_Click
If time0 = 5 Then Call Command24_Click
If time0 = 7 Then Timer1.Enabled = False

End Sub
