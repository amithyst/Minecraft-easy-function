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
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   11640
      Top             =   6720
   End
   Begin VB.CommandButton Command27 
      Caption         =   "һ������"
      Height          =   855
      Left            =   11640
      TabIndex        =   115
      Top             =   5400
      Width           =   495
   End
   Begin VB.CommandButton Command26 
      Caption         =   "ɾ������"
      Height          =   855
      Left            =   11520
      TabIndex        =   112
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton Command25 
      Caption         =   "���뱾ҳ"
      Height          =   255
      Left            =   13800
      TabIndex        =   111
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command24 
      Caption         =   "ת��"
      Height          =   855
      Left            =   10800
      TabIndex        =   110
      Top             =   6120
      Width           =   495
   End
   Begin VB.CommandButton Command23 
      Caption         =   "�����ҳ"
      Height          =   615
      Left            =   8880
      TabIndex        =   109
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton Command22 
      Caption         =   "������ҳ"
      Height          =   255
      Left            =   13800
      TabIndex        =   108
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton Command21 
      Caption         =   "����"
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
      Text            =   "D:\VB\MC������\дħ����\�鱾Ŀ¼\godbook1.txt"
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
      Caption         =   "��ҳ"
      Height          =   495
      Left            =   4440
      TabIndex        =   103
      Top             =   4440
      Width           =   375
   End
   Begin VB.CommandButton Command9 
      Caption         =   "��һҳ"
      Height          =   495
      Left            =   3960
      TabIndex        =   102
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton Command19 
      Caption         =   "д�����沢����"
      Height          =   855
      Left            =   8400
      TabIndex        =   101
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton Command18 
      Caption         =   "�����Ⲣ����"
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
      Caption         =   "����"
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
      Caption         =   "��һԪ��"
      Height          =   495
      Left            =   3480
      TabIndex        =   94
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton Command14 
      Caption         =   "��һԪ��"
      Height          =   495
      Left            =   3000
      TabIndex        =   93
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton Command13 
      Caption         =   "��һҳ"
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
      Caption         =   "����ı�"
      Height          =   735
      Left            =   8040
      TabIndex        =   56
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command10 
      Caption         =   "������һԪ��"
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
      Text            =   "Զ������"
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Index           =   0
      Left            =   4080
      TabIndex        =   51
      Text            =   "�����ؾ�"
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      Caption         =   "д����һԪ��"
      Height          =   855
      Left            =   9360
      TabIndex        =   49
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "ɾ����ҳ"
      Height          =   855
      Left            =   10320
      TabIndex        =   48
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "ɾ����Ԫ��"
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
      Caption         =   "���뱾Ԫ�غ�"
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
      Text            =   "D:\VB\MC������\ʵ�幹��\��ƷĿ¼\trade\trade.txt"
      Top             =   6120
      Width           =   4815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "�����ǩ·��"
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
      Text            =   "������ʵ����"
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
      Text            =   "�����ѽ"
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
      Caption         =   "����"
      Height          =   855
      Left            =   10320
      TabIndex        =   0
      Top             =   6120
      Width           =   495
   End
   Begin VB.Label Label13 
      Caption         =   "�鱾��ʾ��̬"
      Height          =   255
      Left            =   13080
      TabIndex        =   114
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label12 
      Caption         =   "�鱾Դ����̬"
      Height          =   255
      Left            =   10920
      TabIndex        =   113
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label11 
      Caption         =   "ҳ����ת"
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
      Caption         =   "����"
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
      Caption         =   "�ı�"
      Height          =   255
      Index           =   0
      Left            =   3480
      TabIndex        =   79
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "����255"
      Height          =   255
      Index           =   2
      Left            =   5520
      TabIndex        =   77
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "�̣�255"
      Height          =   255
      Index           =   1
      Left            =   5520
      TabIndex        =   76
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "�죺255"
      Height          =   255
      Index           =   0
      Left            =   5520
      TabIndex        =   75
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "���"
      Height          =   375
      Left            =   120
      TabIndex        =   54
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label Label8 
      Caption         =   "����"
      Height          =   375
      Index           =   1
      Left            =   3480
      TabIndex        =   52
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "����"
      Height          =   375
      Index           =   0
      Left            =   3480
      TabIndex        =   50
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "Ԫ��"
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   46
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "ҳ"
      Height          =   255
      Index           =   0
      Left            =   1320
      TabIndex        =   44
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label Label6 
      Caption         =   "���ָ��"
      Height          =   375
      Left            =   120
      TabIndex        =   26
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "��������"
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
      Caption         =   "����"
      Height          =   375
      Left            =   7920
      TabIndex        =   4
      Top             =   6600
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "·��"
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
Dim wr(1 To 100000) As Boolean 'һǧҳ��һ����
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
If out = 0 Then Command11.Caption = "����ı�"
If out = 1 Then Command11.Caption = "���·����ǩ"
If out = 2 Then Command11.Caption = "��ӷ���"
If out = 3 Then Command11.Caption = "���selector"
If out = 4 Then Command11.Caption = "���translate"
If out = 5 Then Command11.Caption = "���keybind"
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

Private Sub Command16_Click() '------------------------------------------------------------����----------------------------
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

Private Sub Command22_Click() '----------------------------------------------------------����-----------------------------
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
'Text4.Text = Text4.Text + Str(k)--------------------------------------------------������----------------------------
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

Public Sub sentence() '_________________________________________________------------------------------д��һ��Ԫ��
If yixie <= guangbiao Then yixie = guangbiao
wr(guangbiao) = True
Call fresh
a(guangbiao) = txt
atxt(guangbiao) = xxxl
Call LiST
Text9(0).Text = CStr(p0(guangbiao))
Text9(1).Text = CStr(l0(guangbiao))
End Sub

Public Sub delsentence() '_________________________________________________------------------------------ɾ��һ��Ԫ��------------
wr(guangbiao) = False
a(guangbiao) = ""
atxt(guangbiao) = ""
Call LiST
Text9(0).Text = CStr(p0(guangbiao))
Text9(1).Text = CStr(l0(guangbiao))
End Sub

Public Function p0(x) As Integer '------------------------------------------------------�������Ӧ��ҳ��-----------------
p0 = (x - 1) \ 100 + 1
End Function
Public Function l0(x) As Integer '------------------------------------------------------�������Ӧ�ĵ�ҳ�ϵ�Ԫ������-----------------
l0 = (x - 1) Mod 100 + 1
End Function
Public Function l(ye, hang) As Long '------------------------------------------------------����ҳ���Ӧ�Ĺ��-----------------
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
Public Function sy(x As String) As String '_________________________________________________��˫����-----------------------------
If out = 2 Then
sy = x
Else
sy = "\""" + x + "\"""
End If
End Function
Public Function sljz(x) As String '---------------------------------------ת��ʮ������--------------------------------------
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
Public Function dada(x, y) As Long '---------------------------------------ȡ�ϴ���--------------------------------------
dada = ((x + y) + Abs(x - y)) / 2
End Function

Public Sub fresh() '_______________________________________ˢ��ˢ��ˢ��ˢ��ˢ��ˢ��ˢ��ˢ��_______________txt����Ԫ��ˢ��-------------
Dim xxx As String
xxx = Text5(out).Text
If out = 0 And Text5(0).Text = "�����ѽ" Then
Randomize
X2 = Int(Rnd() * 3)
If x1 = 0 And X2 = 0 Then xxx = "hi����������ǻ�����\\n"
If x1 = 0 And X2 = 1 Then xxx = "���Ŀͻ����ã����������ǻ��ֻ�����\\n"
If x1 = 0 And X2 = 2 Then xxx = "Hello!I'm your guider! \\n"
If x1 = 1 And X2 = 0 Then xxx = "��ӭʹ���µ�iphoneϵ��\\n"
If x1 = 1 And X2 = 1 Then xxx = "��ӭ����ȫ�°汾��iphone�ֻ�\\n"
If x1 = 1 And X2 = 2 Then xxx = "��ӭ��ʹ������µ�iphoneϵ��\\n"
If x1 = 2 And X2 = 0 Then xxx = "��ϵ���ֻ�����ħ������\\n"
If x1 = 2 And X2 = 1 Then xxx = "��ϵ�в�Ʒ��ħ������\\n"
If x1 = 2 And X2 = 2 Then xxx = "��ϵ����������Ʒ\\n"
If x1 = 3 Then xxx = "���Ĳ���ϵͳ���ǵ�ǰ�����°汾\\n"
If x1 = 4 Then xxx = "��ǰ�汾��ΪiPh10.7.9\\n"
If x1 = 5 And X2 = 0 Then xxx = "�ڴ����Ը�л�������ǵ�֧�֣�\\nΪ�ˣ����Ǹ����ṩ����ѻ�Ա���ᣡ\\n"
If x1 = 5 And X2 = 1 Then xxx = "�ڴ����������ǵ�֧�ֱ�ʾ��л��\\nΪ�ˣ�������һ����ѻ�Ա�齱���ᣡ\\n"
If x1 = 5 And X2 = 2 Then xxx = "��л���Ĺ��������Σ�\\nΪ�ˣ������л�����ѳ�Ϊ��Ա��\\n"
If x1 = 6 And X2 = 0 Then xxx = "��ϲ����Ϊ��ͨע���Ա\\n"
If x1 = 6 And X2 = 1 Then xxx = "��ϲ�����а׽��Ա\\n"
If x1 = 6 And X2 = 2 Then xxx = "��ϲ�����г�����Ա\\n"
If x1 = 7 Then
zcm = ""
For i = 1 To 16
zcm = zcm + Chr(50 + Int(70 * Rnd()))
Next
xx = "����ע����Ϊ" + zcm + "\\n"
End If
If x1 = 8 Then xxx = "����Ʒ������ʱ��Ϊ" + Str(Now()) + "\\n"
If x1 = 9 Then xxx = "��󣬻�ӭ��������԰����\\n���ܰɣ�\\n"
x1 = x1 + 1
End If

col = Text1(0).Text
If out = 0 And Text1(0).Text = "�����ѽ" Then
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

Public Sub LiST() '----------------------------------------------------------------------�б�ˢ����Ԥ��--------------------------
List1.Clear
List2.Clear
aaatxt = ""

For i = 1 To yixie
If wr(i) = True Then

If p0(i) > kkk Then
List1.AddItem ""
List1.AddItem "��" + CStr(p0(i)) + "ҳ��������"
List2.AddItem aaatxt
aaatxt = ""
List2.AddItem ""
List2.AddItem "��" + CStr(p0(i)) + "ҳ"
End If
kkk = p0(i)
List1.AddItem "��" + CStr(l0(i)) + "Ԫ�أ�" + a(i)
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

Public Sub jiru() '-----------------------------------------------------------------����ȥ-------------------------------
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



Private Sub Command6_Click() '--------------------------------------------------ɾ��һ��Ԫ��--------------------------------=
Call delsentence
End Sub

Private Sub Command7_Click() '----------------------------------------------ɾ��һҳ--------------------------------------------------
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

Text2.Text = "D:\MC\1.20\.minecraft\saves\��dOlostland 1.0\datapacks\ħ������\data\give\functions"

Text6.Text = "display:{Name:'{""text"":""�����ؾ�"",""color"":""yellow"",""bold"":true}',Lore:['{""text"":""ID:O_000"",""color"":""gold"",""bold"":false}','{""text"":""���� ������/�����鱦/����"",""color"":""aqua"",""bold"":true}','{""text"":""���Ͼ��յ�����Ȩ�����̺����޷�������������ħ��Ů��������һ����"",""color"":""yellow"",""bold"":true}','{""text"":""��������졹��"",""color"":""yellow"",""bold"":true}','{""text"":""�k���ڣ����ڣ��k�ཫ���ڣ�"",""color"":""dark_red"",""bold"":true}','{""text"":""�k��Ψһ���k�������У�"",""color"":""dark_red"",""bold"":true}']}"
End Sub

Private Sub HScroll1_Change()
red = HScroll1.Value
Label4(0).Caption = "�죺" + CStr(red)
Call colour
End Sub

Private Sub HScroll2_Change()
green = HScroll2.Value
Label4(1).Caption = "�̣�" + CStr(green)
Call colour
End Sub

Private Sub HScroll3_Change()
blue = HScroll3.Value
Label4(2).Caption = "����" + CStr(blue)
Call colour
End Sub
Public Sub colour() '------------------------------------------------------�ɰ��ĸı���ɫ����---------------------------
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
