VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   11220
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   21255
   LinkTopic       =   "Form1"
   ScaleHeight     =   11220
   ScaleWidth      =   21255
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command32 
      Height          =   495
      Index           =   8
      Left            =   12360
      TabIndex        =   332
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Command35 
      Height          =   735
      Index           =   19
      Left            =   9840
      TabIndex        =   331
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Command35 
      Height          =   735
      Index           =   18
      Left            =   9600
      TabIndex        =   330
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Command35 
      Height          =   735
      Index           =   17
      Left            =   9360
      TabIndex        =   329
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Command35 
      Height          =   735
      Index           =   16
      Left            =   9120
      TabIndex        =   328
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Command35 
      Height          =   735
      Index           =   15
      Left            =   8880
      TabIndex        =   327
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Command35 
      Height          =   735
      Index           =   14
      Left            =   8640
      TabIndex        =   326
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Command35 
      Height          =   735
      Index           =   13
      Left            =   8400
      TabIndex        =   325
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Command35 
      Height          =   735
      Index           =   12
      Left            =   8160
      TabIndex        =   324
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Command35 
      Height          =   735
      Index           =   11
      Left            =   7920
      TabIndex        =   323
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Command35 
      Height          =   735
      Index           =   10
      Left            =   7680
      TabIndex        =   322
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Command35 
      Height          =   735
      Index           =   9
      Left            =   7440
      TabIndex        =   321
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Command35 
      Height          =   735
      Index           =   8
      Left            =   7200
      TabIndex        =   320
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Command35 
      Height          =   735
      Index           =   7
      Left            =   6960
      TabIndex        =   319
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Command35 
      Height          =   735
      Index           =   6
      Left            =   6720
      TabIndex        =   318
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Command35 
      Height          =   735
      Index           =   5
      Left            =   6480
      TabIndex        =   317
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Command35 
      Height          =   735
      Index           =   4
      Left            =   6240
      TabIndex        =   316
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Command35 
      Height          =   735
      Index           =   3
      Left            =   6000
      TabIndex        =   315
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Command35 
      Height          =   735
      Index           =   2
      Left            =   5760
      TabIndex        =   314
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Command35 
      Height          =   735
      Index           =   1
      Left            =   5520
      TabIndex        =   313
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Command35 
      Height          =   735
      Index           =   0
      Left            =   5280
      TabIndex        =   312
      Top             =   2640
      Width           =   255
   End
   Begin VB.TextBox Text29 
      Height          =   270
      Left            =   12120
      TabIndex        =   311
      Text            =   "ϡ��"
      Top             =   2160
      Width           =   495
   End
   Begin VB.HScrollBar HScroll4 
      Height          =   255
      Left            =   11280
      Max             =   7
      TabIndex        =   309
      Top             =   9720
      Width           =   1455
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   255
      Left            =   11160
      Max             =   3000
      Min             =   500
      SmallChange     =   100
      TabIndex        =   308
      Top             =   10680
      Value           =   500
      Width           =   1575
   End
   Begin VB.CommandButton Command8 
      Caption         =   "�̱�ʯ"
      Height          =   255
      Index           =   23
      Left            =   9840
      TabIndex        =   307
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "��ʯ"
      Height          =   255
      Index           =   22
      Left            =   10080
      TabIndex        =   306
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton Command8 
      Caption         =   "��"
      Height          =   255
      Index           =   21
      Left            =   9840
      TabIndex        =   305
      Top             =   360
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   10680
      Top             =   10560
   End
   Begin VB.CommandButton Command43 
      Caption         =   "����"
      Height          =   375
      Left            =   12240
      TabIndex        =   304
      Top             =   10320
      Width           =   495
   End
   Begin VB.TextBox Text38 
      Height          =   375
      Left            =   12240
      TabIndex        =   303
      Text            =   "Text38"
      Top             =   9960
      Width           =   495
   End
   Begin VB.CommandButton Command42 
      Caption         =   "����"
      Height          =   375
      Left            =   12240
      TabIndex        =   302
      Top             =   8520
      Width           =   495
   End
   Begin VB.CommandButton Command41 
      Caption         =   "���"
      Height          =   495
      Left            =   12240
      TabIndex        =   301
      Top             =   8040
      Width           =   495
   End
   Begin VB.ListBox List4 
      Height          =   3120
      Left            =   12840
      TabIndex        =   300
      Top             =   7560
      Width           =   8295
   End
   Begin VB.CommandButton Command40 
      Caption         =   "���"
      Height          =   375
      Left            =   12240
      TabIndex        =   299
      Top             =   7680
      Width           =   495
   End
   Begin VB.TextBox Text35 
      Height          =   270
      Index           =   3
      Left            =   3840
      TabIndex        =   298
      Top             =   1680
      Width           =   6615
   End
   Begin VB.CommandButton Command39 
      Caption         =   "������ԣ�װ��"
      Height          =   495
      Index           =   1
      Left            =   4680
      TabIndex        =   296
      Top             =   8400
      Width           =   855
   End
   Begin VB.CommandButton Command39 
      Caption         =   "������ԣ�����"
      Height          =   495
      Index           =   0
      Left            =   4680
      TabIndex        =   295
      Top             =   7920
      Width           =   855
   End
   Begin VB.CommandButton Command38 
      Caption         =   "\\\"
      Height          =   255
      Index           =   1
      Left            =   11640
      TabIndex        =   294
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton Command38 
      Height          =   255
      Index           =   0
      Left            =   11400
      TabIndex        =   293
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Text37 
      Height          =   270
      Left            =   10560
      TabIndex        =   292
      Text            =   "Text37"
      Top             =   825
      Width           =   855
   End
   Begin VB.CommandButton Command37 
      Caption         =   "green"
      Height          =   375
      Index           =   34
      Left            =   20040
      TabIndex        =   291
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command37 
      Caption         =   "green"
      Height          =   375
      Index           =   33
      Left            =   20040
      TabIndex        =   290
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Command37 
      Caption         =   "green"
      Height          =   375
      Index           =   32
      Left            =   20040
      TabIndex        =   289
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command37 
      Caption         =   "green"
      Height          =   375
      Index           =   31
      Left            =   20040
      TabIndex        =   288
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton Command37 
      Caption         =   "green"
      Height          =   375
      Index           =   30
      Left            =   20040
      TabIndex        =   287
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton Command37 
      Caption         =   "green"
      Height          =   375
      Index           =   29
      Left            =   19560
      TabIndex        =   286
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command37 
      Caption         =   "green"
      Height          =   375
      Index           =   28
      Left            =   19560
      TabIndex        =   285
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Command37 
      Caption         =   "green"
      Height          =   375
      Index           =   27
      Left            =   19560
      TabIndex        =   284
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command37 
      Caption         =   "green"
      Height          =   375
      Index           =   26
      Left            =   19560
      TabIndex        =   283
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton Command37 
      Caption         =   "green"
      Height          =   375
      Index           =   25
      Left            =   19560
      TabIndex        =   282
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton Command37 
      Caption         =   "green"
      Height          =   375
      Index           =   24
      Left            =   19080
      TabIndex        =   281
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command37 
      Caption         =   "green"
      Height          =   375
      Index           =   23
      Left            =   19080
      TabIndex        =   280
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Command37 
      Caption         =   "green"
      Height          =   375
      Index           =   22
      Left            =   19080
      TabIndex        =   279
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command37 
      Caption         =   "green"
      Height          =   375
      Index           =   21
      Left            =   19080
      TabIndex        =   278
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton Command37 
      Caption         =   "green"
      Height          =   375
      Index           =   20
      Left            =   19080
      TabIndex        =   277
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton Command37 
      Caption         =   "green"
      Height          =   375
      Index           =   19
      Left            =   18600
      TabIndex        =   276
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command37 
      Caption         =   "green"
      Height          =   375
      Index           =   18
      Left            =   18600
      TabIndex        =   275
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Command37 
      Caption         =   "green"
      Height          =   375
      Index           =   17
      Left            =   18600
      TabIndex        =   274
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command37 
      Caption         =   "green"
      Height          =   375
      Index           =   16
      Left            =   18600
      TabIndex        =   273
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton Command37 
      Caption         =   "green"
      Height          =   375
      Index           =   15
      Left            =   18600
      TabIndex        =   272
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton Command37 
      Caption         =   "green"
      Height          =   375
      Index           =   9
      Left            =   17640
      TabIndex        =   271
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command37 
      Caption         =   "green"
      Height          =   375
      Index           =   8
      Left            =   17640
      TabIndex        =   270
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Command37 
      Caption         =   "green"
      Height          =   375
      Index           =   7
      Left            =   17640
      TabIndex        =   269
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command37 
      Caption         =   "green"
      Height          =   375
      Index           =   6
      Left            =   17640
      TabIndex        =   268
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton Command37 
      Caption         =   "green"
      Height          =   375
      Index           =   5
      Left            =   17640
      TabIndex        =   267
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton Command37 
      Caption         =   "green"
      Height          =   375
      Index           =   14
      Left            =   18120
      TabIndex        =   266
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command37 
      Caption         =   "green"
      Height          =   375
      Index           =   13
      Left            =   18120
      TabIndex        =   265
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Command37 
      Caption         =   "green"
      Height          =   375
      Index           =   12
      Left            =   18120
      TabIndex        =   264
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command37 
      Caption         =   "green"
      Height          =   375
      Index           =   11
      Left            =   18120
      TabIndex        =   263
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton Command37 
      Caption         =   "green"
      Height          =   375
      Index           =   10
      Left            =   18120
      TabIndex        =   262
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton Command37 
      Caption         =   "green"
      Height          =   375
      Index           =   4
      Left            =   17160
      TabIndex        =   261
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command37 
      Caption         =   "green"
      Height          =   375
      Index           =   3
      Left            =   17160
      TabIndex        =   260
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Command37 
      Caption         =   "green"
      Height          =   375
      Index           =   2
      Left            =   17160
      TabIndex        =   259
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command37 
      Caption         =   "green"
      Height          =   375
      Index           =   1
      Left            =   17160
      TabIndex        =   258
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton Command37 
      Caption         =   "green"
      Height          =   375
      Index           =   0
      Left            =   17160
      TabIndex        =   257
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton Command36 
      Caption         =   "���������Ʒ��ʷ����Ϣ"
      Height          =   855
      Left            =   4800
      TabIndex        =   256
      Top             =   6960
      Width           =   735
   End
   Begin VB.TextBox Text36 
      Height          =   375
      Index           =   4
      Left            =   16440
      TabIndex        =   255
      Text            =   "Text36"
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox Text34 
      Height          =   375
      Index           =   4
      Left            =   12600
      TabIndex        =   254
      Text            =   "Text34"
      Top             =   1920
      Width           =   3855
   End
   Begin VB.TextBox Text36 
      Height          =   375
      Index           =   3
      Left            =   16440
      TabIndex        =   253
      Text            =   "Text36"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text36 
      Height          =   375
      Index           =   2
      Left            =   16440
      TabIndex        =   252
      Text            =   "Text36"
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox Text36 
      Height          =   375
      Index           =   1
      Left            =   16440
      TabIndex        =   251
      Text            =   "Text36"
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox Text36 
      Height          =   375
      Index           =   0
      Left            =   16440
      TabIndex        =   250
      Text            =   "Text36"
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox Text34 
      Height          =   375
      Index           =   3
      Left            =   12600
      TabIndex        =   249
      Text            =   "Text34"
      Top             =   1560
      Width           =   3855
   End
   Begin VB.TextBox Text34 
      Height          =   375
      Index           =   2
      Left            =   12600
      TabIndex        =   248
      Text            =   "Text34"
      Top             =   1200
      Width           =   3855
   End
   Begin VB.TextBox Text34 
      Height          =   375
      Index           =   1
      Left            =   12600
      TabIndex        =   247
      Text            =   "Text34"
      Top             =   840
      Width           =   3855
   End
   Begin VB.TextBox Text35 
      Height          =   270
      Index           =   2
      Left            =   8280
      TabIndex        =   246
      Text            =   "Text35"
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox Text35 
      Height          =   270
      Index           =   1
      Left            =   5640
      TabIndex        =   245
      Text            =   "Text35"
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox Text35 
      Height          =   270
      Index           =   0
      Left            =   5640
      TabIndex        =   244
      Text            =   "��"
      Top             =   2160
      Width           =   4335
   End
   Begin VB.TextBox Text34 
      Height          =   375
      Index           =   0
      Left            =   12600
      TabIndex        =   243
      Text            =   "Text34"
      Top             =   480
      Width           =   3855
   End
   Begin VB.CommandButton Command34 
      Caption         =   "������л꡹"
      Height          =   855
      Index           =   7
      Left            =   11880
      TabIndex        =   242
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Command34 
      Caption         =   "������֮�ۡ�"
      Height          =   855
      Index           =   6
      Left            =   11640
      TabIndex        =   241
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Command34 
      Caption         =   "��Զ�Ż��졹"
      Height          =   855
      Index           =   5
      Left            =   11160
      TabIndex        =   240
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Command34 
      Caption         =   "��ǰ�����Ρ�"
      Height          =   855
      Index           =   4
      Left            =   10440
      TabIndex        =   239
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Command34 
      Caption         =   "���ۻ���"
      Height          =   855
      Index           =   3
      Left            =   11400
      TabIndex        =   238
      Top             =   2640
      Width           =   255
   End
   Begin VB.TextBox Text33 
      Height          =   270
      Left            =   6240
      TabIndex        =   237
      Text            =   "��ħ��ʱ��"
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton Command34 
      Caption         =   "������֮�"
      Height          =   855
      Index           =   2
      Left            =   10920
      TabIndex        =   236
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Command34 
      Caption         =   "������֮�"
      Height          =   855
      Index           =   1
      Left            =   10200
      TabIndex        =   235
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Command34 
      Caption         =   "������߽�"
      Height          =   855
      Index           =   0
      Left            =   10680
      TabIndex        =   234
      Top             =   2640
      Width           =   255
   End
   Begin VB.TextBox Text32 
      Height          =   855
      Left            =   12120
      MultiLine       =   -1  'True
      TabIndex        =   233
      Text            =   "ħ�����߹���.frx":0000
      Top             =   2640
      Width           =   495
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   9960
      Max             =   999
      TabIndex        =   232
      Top             =   1920
      Width           =   2415
   End
   Begin VB.TextBox Text31 
      Height          =   270
      Left            =   9720
      TabIndex        =   231
      Text            =   $"ħ�����߹���.frx":000D
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton Command33 
      Caption         =   "��ʯ"
      Height          =   495
      Index           =   7
      Left            =   11880
      TabIndex        =   230
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Command33 
      Caption         =   "����"
      Height          =   495
      Index           =   6
      Left            =   11640
      TabIndex        =   229
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Command33 
      Caption         =   "����"
      Height          =   495
      Index           =   5
      Left            =   11400
      TabIndex        =   228
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Command33 
      Caption         =   "����"
      Height          =   495
      Index           =   4
      Left            =   11160
      TabIndex        =   227
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Command33 
      Caption         =   "ҩˮ"
      Height          =   495
      Index           =   3
      Left            =   10920
      TabIndex        =   226
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Command33 
      Caption         =   "����"
      Height          =   495
      Index           =   2
      Left            =   10680
      TabIndex        =   225
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Command33 
      Caption         =   "װ��"
      Height          =   495
      Index           =   1
      Left            =   10440
      TabIndex        =   224
      Top             =   2160
      Width           =   255
   End
   Begin VB.TextBox Text30 
      Height          =   270
      Left            =   12120
      TabIndex        =   223
      Text            =   "����"
      Top             =   2385
      Width           =   495
   End
   Begin VB.CommandButton Command33 
      Caption         =   "����"
      Height          =   495
      Index           =   0
      Left            =   10200
      TabIndex        =   222
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Command32 
      Caption         =   " "
      Height          =   495
      Index           =   7
      Left            =   12120
      TabIndex        =   221
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Command32 
      Caption         =   "����"
      Height          =   495
      Index           =   6
      Left            =   11880
      TabIndex        =   220
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Command32 
      BackColor       =   &H8000000E&
      Caption         =   "��"
      Height          =   495
      Index           =   5
      Left            =   11640
      TabIndex        =   219
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Command32 
      Caption         =   "��˵"
      Height          =   495
      Index           =   4
      Left            =   11400
      TabIndex        =   218
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Command32 
      Caption         =   "ʷʫ"
      Height          =   495
      Index           =   3
      Left            =   11160
      TabIndex        =   217
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Command32 
      Caption         =   "ϡ��"
      Height          =   495
      Index           =   2
      Left            =   10920
      TabIndex        =   216
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Command32 
      Caption         =   "����"
      Height          =   495
      Index           =   1
      Left            =   10680
      TabIndex        =   215
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Command32 
      BackColor       =   &H8000000D&
      Caption         =   "��ͨ"
      Height          =   495
      Index           =   0
      Left            =   10440
      TabIndex        =   214
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "swift_sneak"
      Height          =   465
      Index           =   38
      Left            =   5400
      TabIndex        =   213
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "vanishing_curse"
      Height          =   420
      Index           =   37
      Left            =   5280
      TabIndex        =   212
      Top             =   5760
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "binding_curse"
      Height          =   420
      Index           =   36
      Left            =   5280
      TabIndex        =   211
      Top             =   5400
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "��"
      Height          =   495
      Index           =   20
      Left            =   8160
      TabIndex        =   210
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton Command31 
      Caption         =   "ת��"
      Height          =   495
      Left            =   720
      TabIndex        =   209
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text28 
      Height          =   375
      Left            =   1200
      TabIndex        =   208
      Top             =   5040
      Width           =   855
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Index           =   2
      Left            =   16560
      Max             =   255
      TabIndex        =   207
      Top             =   4800
      Width           =   1695
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Index           =   1
      Left            =   14880
      Max             =   255
      TabIndex        =   206
      Top             =   4800
      Value           =   255
      Width           =   1695
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Index           =   0
      Left            =   13200
      Max             =   255
      TabIndex        =   205
      Top             =   4800
      Width           =   1695
   End
   Begin VB.TextBox Text27 
      Height          =   615
      Left            =   18720
      TabIndex        =   204
      Top             =   4560
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "��"
      Height          =   255
      Index           =   19
      Left            =   8400
      TabIndex        =   203
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton Command30 
      Caption         =   "¼�봢��"
      Height          =   495
      Left            =   1800
      TabIndex        =   202
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton Command29 
      Caption         =   "���"
      Height          =   375
      Left            =   9960
      TabIndex        =   201
      Top             =   10200
      Width           =   495
   End
   Begin VB.TextBox Text26 
      Height          =   375
      Left            =   7680
      TabIndex        =   199
      Top             =   10200
      Width           =   1575
   End
   Begin VB.CommandButton Command28 
      Caption         =   "����"
      Height          =   375
      Left            =   9360
      TabIndex        =   198
      Top             =   10200
      Width           =   495
   End
   Begin VB.TextBox Text25 
      Height          =   495
      Left            =   14040
      TabIndex        =   197
      Text            =   "1"
      Top             =   6960
      Width           =   375
   End
   Begin VB.CommandButton Command27 
      Caption         =   "��������������"
      Height          =   495
      Left            =   13080
      TabIndex        =   196
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "������"
      Height          =   255
      Index           =   11
      Left            =   5880
      TabIndex        =   195
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "���ɵ�"
      Height          =   735
      Index           =   18
      Left            =   8400
      TabIndex        =   194
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox Text24 
      Height          =   375
      Left            =   9840
      TabIndex        =   193
      Text            =   "summon"
      Top             =   9720
      Width           =   975
   End
   Begin VB.CommandButton Command26 
      Caption         =   "����"
      Height          =   375
      Left            =   9360
      TabIndex        =   192
      Top             =   9720
      Width           =   495
   End
   Begin VB.TextBox Text23 
      Height          =   375
      Left            =   7680
      TabIndex        =   191
      Top             =   9720
      Width           =   1575
   End
   Begin VB.CommandButton Command25 
      Caption         =   "�½�"
      Height          =   255
      Left            =   7560
      TabIndex        =   189
      Top             =   3600
      Width           =   495
   End
   Begin VB.TextBox Text22 
      Height          =   270
      Left            =   6360
      TabIndex        =   187
      Text            =   "trade"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command24 
      Caption         =   """feet"""
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   186
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command24 
      Caption         =   """legs"""
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   185
      Top             =   9480
      Width           =   855
   End
   Begin VB.CommandButton Command24 
      Caption         =   """chest"""
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   184
      Top             =   9240
      Width           =   855
   End
   Begin VB.CommandButton Command24 
      Caption         =   """head"""
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   183
      Top             =   9000
      Width           =   855
   End
   Begin VB.CommandButton Command24 
      Caption         =   """offhand"""
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   182
      Top             =   8760
      Width           =   975
   End
   Begin VB.CommandButton Command24 
      Caption         =   """mainhand"""
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   181
      Top             =   8520
      Width           =   1095
   End
   Begin VB.TextBox Text21 
      Height          =   270
      Left            =   360
      TabIndex        =   180
      Top             =   8160
      Width           =   855
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Index           =   11
      Left            =   2880
      TabIndex        =   176
      Top             =   10320
      Width           =   735
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Index           =   10
      Left            =   2880
      TabIndex        =   175
      Top             =   9840
      Width           =   735
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Index           =   9
      Left            =   2880
      TabIndex        =   173
      Top             =   9360
      Width           =   735
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Index           =   8
      Left            =   2880
      TabIndex        =   171
      Top             =   8880
      Width           =   735
   End
   Begin VB.CommandButton Command23 
      Caption         =   "ʷʫ����"
      Height          =   495
      Left            =   3600
      TabIndex        =   170
      Top             =   10080
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   495
      Left            =   240
      TabIndex        =   169
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton Command20 
      Caption         =   "���Ե���"
      Height          =   495
      Left            =   3600
      TabIndex        =   168
      Top             =   9000
      Width           =   855
   End
   Begin VB.TextBox Text20 
      Height          =   495
      Left            =   960
      TabIndex        =   166
      Text            =   "0"
      Top             =   4320
      Width           =   735
   End
   Begin VB.TextBox Text12 
      Height          =   270
      Left            =   5640
      TabIndex        =   165
      Text            =   "�̺���ǿ��������"
      Top             =   1920
      Width           =   4335
   End
   Begin VB.CommandButton Command12 
      Caption         =   "�µ����� (1)"
      Height          =   225
      Index           =   1
      Left            =   9600
      TabIndex        =   164
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      Caption         =   "���"
      Height          =   255
      Index           =   0
      Left            =   9120
      TabIndex        =   163
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Index           =   12
      Left            =   17880
      TabIndex        =   160
      Text            =   "healing"
      Top             =   6600
      Width           =   975
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Index           =   11
      Left            =   12000
      TabIndex        =   151
      Text            =   "7"
      Top             =   5400
      Width           =   2535
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Index           =   10
      Left            =   13680
      TabIndex        =   150
      Text            =   "28"
      Top             =   5760
      Width           =   615
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Index           =   9
      Left            =   13560
      TabIndex        =   149
      Text            =   "1000000"
      Top             =   6240
      Width           =   735
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Index           =   8
      Left            =   15120
      TabIndex        =   148
      Text            =   "0"
      Top             =   6600
      Width           =   735
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Index           =   7
      Left            =   16560
      TabIndex        =   147
      Text            =   "1"
      Top             =   6600
      Width           =   735
   End
   Begin VB.ListBox List3 
      Height          =   1140
      Left            =   14520
      TabIndex        =   146
      Top             =   5400
      Width           =   5535
   End
   Begin VB.CommandButton Command22 
      Caption         =   "���"
      Height          =   375
      Left            =   18840
      TabIndex        =   145
      Top             =   6600
      Width           =   735
   End
   Begin VB.CommandButton Command21 
      Caption         =   "���"
      Height          =   375
      Left            =   19560
      TabIndex        =   144
      Top             =   6600
      Width           =   495
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Index           =   6
      Left            =   13560
      TabIndex        =   143
      Text            =   "1"
      Top             =   6600
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "����"
      Height          =   495
      Index           =   10
      Left            =   5280
      TabIndex        =   142
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton Command6 
      Caption         =   "�罦"
      Height          =   495
      Index           =   9
      Left            =   5520
      TabIndex        =   141
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton Command8 
      Caption         =   "ҩˮ"
      Height          =   495
      Index           =   17
      Left            =   8160
      TabIndex        =   140
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Index           =   0
      Left            =   1680
      TabIndex        =   130
      Top             =   7920
      Width           =   855
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   129
      Top             =   8400
      Width           =   855
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Index           =   2
      Left            =   1680
      TabIndex        =   128
      Top             =   8880
      Width           =   855
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Index           =   3
      Left            =   1680
      TabIndex        =   127
      Top             =   9360
      Width           =   855
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Index           =   4
      Left            =   1680
      TabIndex        =   126
      Top             =   9840
      Width           =   855
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Index           =   5
      Left            =   1680
      TabIndex        =   125
      Top             =   10320
      Width           =   855
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Index           =   6
      Left            =   2880
      TabIndex        =   124
      Top             =   7920
      Width           =   735
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Index           =   7
      Left            =   2880
      TabIndex        =   123
      Top             =   8400
      Width           =   735
   End
   Begin VB.CommandButton Command13 
      Caption         =   "��������"
      Height          =   495
      Left            =   3600
      TabIndex        =   122
      Top             =   7920
      Width           =   855
   End
   Begin VB.TextBox Text17 
      Height          =   495
      Left            =   4080
      TabIndex        =   121
      Text            =   "0"
      Top             =   9600
      Width           =   375
   End
   Begin VB.CommandButton Command18 
      Caption         =   "��������"
      Height          =   495
      Left            =   3600
      TabIndex        =   120
      Top             =   8400
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "��"
      Height          =   255
      Index           =   16
      Left            =   7680
      TabIndex        =   119
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox Text19 
      Height          =   270
      Left            =   9120
      TabIndex        =   118
      Text            =   "gray"
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox Text18 
      Height          =   270
      Left            =   8880
      TabIndex        =   117
      Text            =   "gold"
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton Command19 
      Caption         =   "�ռ�"
      Height          =   495
      Left            =   1440
      TabIndex        =   116
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton Command8 
      Caption         =   "����"
      Height          =   495
      Index           =   15
      Left            =   7920
      TabIndex        =   115
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton Command17 
      Caption         =   "�������UUID"
      Height          =   375
      Left            =   1320
      TabIndex        =   114
      Top             =   7560
      Width           =   975
   End
   Begin VB.TextBox Text16 
      Height          =   375
      Left            =   2280
      TabIndex        =   113
      Text            =   "���"
      Top             =   7560
      Width           =   2175
   End
   Begin VB.CommandButton Command16 
      Caption         =   "����"
      Height          =   255
      Left            =   5040
      TabIndex        =   112
      Top             =   3600
      Width           =   615
   End
   Begin VB.CommandButton Command15 
      Caption         =   "trade"
      Height          =   255
      Index           =   2
      Left            =   4440
      TabIndex        =   111
      Top             =   3600
      Width           =   615
   End
   Begin VB.CommandButton Command15 
      Caption         =   "wupin"
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   110
      Top             =   3600
      Width           =   615
   End
   Begin VB.CommandButton Command15 
      Caption         =   "magic"
      Height          =   255
      Index           =   0
      Left            =   3240
      TabIndex        =   109
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Index           =   5
      Left            =   13920
      TabIndex        =   105
      Text            =   "1"
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton Command14 
      Caption         =   "��ɫ����"
      Height          =   615
      Left            =   18240
      TabIndex        =   104
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox Text15 
      Height          =   375
      Index           =   2
      Left            =   16920
      TabIndex        =   102
      Text            =   "0"
      Top             =   4440
      Width           =   615
   End
   Begin VB.TextBox Text15 
      Height          =   375
      Index           =   1
      Left            =   15360
      TabIndex        =   100
      Text            =   "255"
      Top             =   4440
      Width           =   615
   End
   Begin VB.TextBox Text15 
      Height          =   375
      Index           =   0
      Left            =   13800
      TabIndex        =   98
      Text            =   "0"
      Top             =   4440
      Width           =   615
   End
   Begin VB.TextBox Text14 
      Height          =   270
      Left            =   2400
      TabIndex        =   97
      Top             =   3960
      Width           =   6735
   End
   Begin VB.TextBox Text13 
      Height          =   270
      Left            =   2400
      TabIndex        =   94
      Text            =   "magic"
      Top             =   3600
      Width           =   855
   End
   Begin VB.TextBox Text10 
      Height          =   270
      Left            =   5760
      TabIndex        =   93
      Text            =   "��������"
      Top             =   1440
      Width           =   3135
   End
   Begin VB.TextBox Text11 
      Height          =   495
      Left            =   7680
      TabIndex        =   92
      Text            =   "0"
      Top             =   5280
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "�̻����"
      Height          =   495
      Index           =   14
      Left            =   9120
      TabIndex        =   89
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton Command8 
      Caption         =   "��"
      Height          =   255
      Index           =   13
      Left            =   7440
      TabIndex        =   88
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton Command11 
      Caption         =   "���"
      Height          =   615
      Left            =   19440
      TabIndex        =   87
      Top             =   4560
      Width           =   255
   End
   Begin VB.CommandButton Command10 
      Caption         =   "���"
      Height          =   615
      Left            =   19680
      TabIndex        =   86
      Top             =   4560
      Width           =   255
   End
   Begin VB.ListBox List2 
      Height          =   2040
      Left            =   15240
      TabIndex        =   85
      Top             =   2400
      Width           =   4695
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Index           =   4
      Left            =   13920
      TabIndex        =   83
      Text            =   "1"
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Index           =   3
      Left            =   13920
      TabIndex        =   81
      Text            =   "���"
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Index           =   2
      Left            =   13920
      TabIndex        =   79
      Text            =   "1"
      Top             =   3120
      Width           =   495
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Index           =   1
      Left            =   13920
      TabIndex        =   77
      Text            =   "0"
      Top             =   2760
      Width           =   495
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Index           =   0
      Left            =   13920
      TabIndex        =   75
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Command9"
      Height          =   735
      Left            =   360
      TabIndex        =   74
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   3360
      TabIndex        =   73
      Text            =   "@a"
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Height          =   495
      Index           =   8
      Left            =   5040
      TabIndex        =   72
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton Command7 
      Caption         =   "�����ظ�"
      Height          =   255
      Left            =   2880
      TabIndex        =   71
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "���ħ��"
      Height          =   255
      Left            =   1680
      TabIndex        =   70
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "�ʳ�"
      Height          =   495
      Index           =   12
      Left            =   7920
      TabIndex        =   69
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton Command8 
      Caption         =   "�����"
      Height          =   495
      Index           =   11
      Left            =   8640
      TabIndex        =   68
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton Command8 
      Caption         =   "�����"
      Height          =   495
      Index           =   10
      Left            =   8640
      TabIndex        =   67
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "����"
      Height          =   495
      Index           =   35
      Left            =   5760
      TabIndex        =   66
      Top             =   4920
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��֮���"
      Height          =   495
      Index           =   33
      Left            =   5280
      TabIndex        =   65
      Top             =   4920
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "����"
      Height          =   495
      Index           =   34
      Left            =   4920
      TabIndex        =   64
      Top             =   4920
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "�������"
      Height          =   495
      Index           =   32
      Left            =   3360
      TabIndex        =   63
      Top             =   6600
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��͸"
      Height          =   495
      Index           =   31
      Left            =   3840
      TabIndex        =   62
      Top             =   6600
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "����װ��"
      Height          =   495
      Index           =   30
      Left            =   2880
      TabIndex        =   61
      Top             =   6600
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "����"
      Height          =   495
      Index           =   29
      Left            =   4200
      TabIndex        =   60
      Top             =   4920
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "����"
      Height          =   495
      Index           =   28
      Left            =   4440
      TabIndex        =   59
      Top             =   4920
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "�ҳ�"
      Height          =   495
      Index           =   27
      Left            =   4680
      TabIndex        =   58
      Top             =   4920
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "���"
      Height          =   495
      Index           =   26
      Left            =   4200
      TabIndex        =   57
      Top             =   5400
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ˮ�º���"
      Height          =   495
      Index           =   25
      Left            =   4680
      TabIndex        =   56
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��˪����"
      Height          =   495
      Index           =   24
      Left            =   4920
      TabIndex        =   55
      Top             =   4320
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��꼲��"
      Height          =   495
      Index           =   23
      Left            =   4440
      TabIndex        =   54
      Top             =   4320
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��ʸ"
      Height          =   495
      Index           =   22
      Left            =   4920
      TabIndex        =   53
      Top             =   5400
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "����"
      Height          =   495
      Index           =   21
      Left            =   4680
      TabIndex        =   52
      Top             =   5400
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "����"
      Height          =   495
      Index           =   20
      Left            =   4440
      TabIndex        =   51
      Top             =   5400
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "�����ﱣ��"
      Height          =   615
      Index           =   19
      Left            =   2640
      TabIndex        =   50
      Top             =   4920
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "����"
      Height          =   495
      Index           =   18
      Left            =   2880
      TabIndex        =   49
      Top             =   4320
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ˮ���پ�"
      Height          =   495
      Index           =   17
      Left            =   4200
      TabIndex        =   48
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "�̽����"
      Height          =   615
      Index           =   16
      Left            =   3960
      TabIndex        =   47
      Top             =   4320
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "����"
      Height          =   495
      Index           =   15
      Left            =   3120
      TabIndex        =   46
      Top             =   5400
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��ը����"
      Height          =   495
      Index           =   14
      Left            =   3600
      TabIndex        =   45
      Top             =   5400
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ˤ�䱣��"
      Height          =   495
      Index           =   13
      Left            =   3120
      TabIndex        =   44
      Top             =   4920
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "���汣��"
      Height          =   495
      Index           =   12
      Left            =   3600
      TabIndex        =   43
      Top             =   4920
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "����"
      Height          =   495
      Index           =   11
      Left            =   3120
      TabIndex        =   42
      Top             =   4320
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "�����޲�"
      Height          =   495
      Index           =   10
      Left            =   3360
      TabIndex        =   41
      Top             =   4320
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "�;�"
      Height          =   495
      Index           =   9
      Left            =   2640
      TabIndex        =   40
      Top             =   4320
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ч��"
      Height          =   495
      Index           =   8
      Left            =   3600
      TabIndex        =   39
      Top             =   6000
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ʱ��"
      Height          =   495
      Index           =   7
      Left            =   3840
      TabIndex        =   38
      Top             =   6000
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��׼�ɼ�"
      Height          =   495
      Index           =   6
      Left            =   3120
      TabIndex        =   37
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "ʯ"
      Height          =   255
      Index           =   7
      Left            =   4800
      TabIndex        =   36
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton Command6 
      Caption         =   "ľ"
      Height          =   255
      Index           =   6
      Left            =   4560
      TabIndex        =   35
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Ƥ��"
      Height          =   495
      Index           =   5
      Left            =   4560
      TabIndex        =   34
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton Command6 
      Caption         =   "����"
      Height          =   495
      Index           =   4
      Left            =   4800
      TabIndex        =   33
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton Command6 
      Caption         =   "��"
      Height          =   255
      Index           =   3
      Left            =   4800
      TabIndex        =   32
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton Command6 
      Caption         =   "��"
      Height          =   255
      Index           =   2
      Left            =   4560
      TabIndex        =   31
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton Command8 
      Caption         =   "����"
      Height          =   495
      Index           =   9
      Left            =   7440
      TabIndex        =   30
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton Command8 
      Caption         =   "ѥ��"
      Height          =   495
      Index           =   8
      Left            =   7680
      TabIndex        =   29
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton Command8 
      Caption         =   "ͷ��"
      Height          =   495
      Index           =   7
      Left            =   6960
      TabIndex        =   28
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton Command8 
      Caption         =   "�ؼ�"
      Height          =   495
      Index           =   6
      Left            =   7200
      TabIndex        =   27
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton Command8 
      Caption         =   "��"
      Height          =   255
      Index           =   5
      Left            =   7440
      TabIndex        =   26
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton Command8 
      Caption         =   "��"
      Height          =   255
      Index           =   4
      Left            =   6960
      TabIndex        =   25
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton Command8 
      Caption         =   "��"
      Height          =   255
      Index           =   3
      Left            =   7200
      TabIndex        =   24
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton Command8 
      Caption         =   "��"
      Height          =   255
      Index           =   2
      Left            =   7680
      TabIndex        =   23
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "����"
      Height          =   495
      Index           =   5
      Left            =   2160
      TabIndex        =   22
      Top             =   4800
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��ɨ֮��"
      Height          =   495
      Index           =   4
      Left            =   2160
      TabIndex        =   21
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "���渽��"
      Height          =   495
      Index           =   3
      Left            =   2160
      TabIndex        =   20
      Top             =   4320
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��֫ɱ��"
      Height          =   495
      Index           =   2
      Left            =   2160
      TabIndex        =   19
      Top             =   6240
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "����ɱ��"
      Height          =   495
      Index           =   1
      Left            =   2160
      TabIndex        =   18
      Top             =   5760
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "�½�Ͻ�"
      Height          =   495
      Index           =   1
      Left            =   5280
      TabIndex        =   17
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton Command8 
      Caption         =   "��"
      Height          =   255
      Index           =   1
      Left            =   7200
      TabIndex        =   16
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   3360
      TabIndex        =   14
      Text            =   "1"
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "��"
      Height          =   255
      Index           =   0
      Left            =   6960
      TabIndex        =   13
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton Command6 
      Caption         =   "��ʯ"
      Height          =   495
      Index           =   0
      Left            =   5040
      TabIndex        =   12
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "���ɴݻ�"
      Height          =   735
      Left            =   3960
      TabIndex        =   11
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   120
      TabIndex        =   10
      Text            =   "Text5"
      Top             =   0
      Width           =   19815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "����"
      Height          =   495
      Index           =   0
      Left            =   2400
      TabIndex        =   9
      Top             =   4800
      Width           =   255
   End
   Begin VB.ListBox List1 
      Height          =   1140
      Left            =   240
      TabIndex        =   8
      Top             =   2400
      Width           =   4935
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Text            =   "Text3"
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "���벢���"
      Height          =   615
      Left            =   2280
      TabIndex        =   3
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Text            =   "Text4"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label20 
      Caption         =   "Label20"
      Height          =   495
      Left            =   11280
      TabIndex        =   310
      Top             =   10080
      Width           =   975
   End
   Begin VB.Label Label19 
      Caption         =   "Label19"
      Height          =   2175
      Left            =   5520
      TabIndex        =   297
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label Label18 
      Caption         =   "����������"
      Height          =   495
      Index           =   1
      Left            =   7080
      TabIndex        =   200
      Top             =   10200
      Width           =   615
   End
   Begin VB.Label Label18 
      Caption         =   "ʵ������"
      Height          =   495
      Index           =   0
      Left            =   7320
      TabIndex        =   190
      Top             =   9720
      Width           =   375
   End
   Begin VB.Label Label9 
      Caption         =   "�ļ���"
      Height          =   375
      Index           =   2
      Left            =   5760
      TabIndex        =   188
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "������λ"
      Height          =   255
      Left            =   360
      TabIndex        =   179
      Top             =   7800
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "�����Ծ"
      Height          =   375
      Index           =   11
      Left            =   2520
      TabIndex        =   178
      Top             =   9840
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "��ʬ�ٻ�"
      Height          =   375
      Index           =   10
      Left            =   2520
      TabIndex        =   177
      Top             =   10320
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "���ٷ�Χ"
      Height          =   375
      Index           =   9
      Left            =   2520
      TabIndex        =   174
      Top             =   9360
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "��������"
      Height          =   495
      Index           =   8
      Left            =   2520
      TabIndex        =   172
      Top             =   8880
      Width           =   375
   End
   Begin VB.Label Label17 
      Caption         =   "��ٳ̶�"
      Height          =   495
      Left            =   360
      TabIndex        =   167
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label Label16 
      Caption         =   $"ħ�����߹���.frx":001C
      Height          =   2295
      Left            =   360
      TabIndex        =   162
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label Label15 
      Caption         =   "��������"
      Height          =   375
      Left            =   6840
      TabIndex        =   161
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "���"
      Height          =   375
      Index           =   12
      Left            =   17400
      TabIndex        =   159
      Top             =   6600
      Width           =   375
   End
   Begin VB.Label Label14 
      Caption         =   $"ħ�����߹���.frx":00AA
      Height          =   5175
      Left            =   8880
      TabIndex        =   158
      Top             =   4440
      Width           =   2655
   End
   Begin VB.Label Label5 
      Caption         =   "Id"
      Height          =   255
      Index           =   11
      Left            =   11640
      TabIndex        =   157
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "ʱ��"
      Height          =   375
      Index           =   10
      Left            =   13080
      TabIndex        =   156
      Top             =   6240
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "����(127/28����/29߱��)"
      Height          =   735
      Index           =   9
      Left            =   12840
      TabIndex        =   155
      Top             =   5640
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "��ʾ����Ч��"
      Height          =   495
      Index           =   8
      Left            =   14520
      TabIndex        =   154
      Top             =   6600
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Ч��ͼ����ʾ"
      Height          =   495
      Index           =   7
      Left            =   15960
      TabIndex        =   153
      Top             =   6600
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "��ɫ��â"
      Height          =   495
      Index           =   6
      Left            =   13080
      TabIndex        =   152
      Top             =   6600
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "��������"
      Height          =   375
      Index           =   0
      Left            =   1320
      TabIndex        =   139
      Top             =   7920
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "���˿���"
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   138
      Top             =   8400
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "�����˺�"
      Height          =   375
      Index           =   2
      Left            =   1320
      TabIndex        =   137
      Top             =   9360
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "�ƶ��ٶ�"
      Height          =   375
      Index           =   3
      Left            =   1320
      TabIndex        =   136
      Top             =   8880
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "��������"
      Height          =   375
      Index           =   4
      Left            =   1320
      TabIndex        =   135
      Top             =   9840
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "���ױ���"
      Height          =   375
      Index           =   5
      Left            =   1320
      TabIndex        =   134
      Top             =   10320
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "����"
      Height          =   375
      Index           =   6
      Left            =   2520
      TabIndex        =   133
      Top             =   8400
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "�����ٶ�"
      Height          =   375
      Index           =   7
      Left            =   2520
      TabIndex        =   132
      Top             =   7920
      Width           =   495
   End
   Begin VB.Label Label13 
      Caption         =   "������ʽ"
      Height          =   495
      Left            =   3600
      TabIndex        =   131
      Top             =   9600
      Width           =   495
   End
   Begin VB.Label Label12 
      Caption         =   "0=С�� 1=���� 2=���� 3=������ͷ   4=����"
      Height          =   1215
      Left            =   14520
      TabIndex        =   108
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label11 
      Caption         =   $"ħ�����߹���.frx":0351
      Height          =   3975
      Left            =   6720
      TabIndex        =   107
      Top             =   5880
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "��˸"
      Height          =   375
      Index           =   5
      Left            =   13320
      TabIndex        =   106
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label10 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   16560
      TabIndex        =   103
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label10 
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   15000
      TabIndex        =   101
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label10 
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   13440
      TabIndex        =   99
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label9 
      Caption         =   "·����"
      Height          =   375
      Index           =   1
      Left            =   1800
      TabIndex        =   96
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "�ļ���"
      Height          =   375
      Index           =   0
      Left            =   1800
      TabIndex        =   95
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label8 
      Caption         =   "����"
      Height          =   255
      Index           =   1
      Left            =   5160
      TabIndex        =   91
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "����"
      Height          =   375
      Index           =   0
      Left            =   5400
      TabIndex        =   90
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "����"
      Height          =   375
      Index           =   4
      Left            =   13320
      TabIndex        =   84
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "��ɫ"
      Height          =   255
      Index           =   3
      Left            =   13320
      TabIndex        =   82
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "��״"
      Height          =   255
      Index           =   2
      Left            =   13320
      TabIndex        =   80
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "�켣"
      Height          =   255
      Index           =   1
      Left            =   13320
      TabIndex        =   78
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "���"
      Height          =   375
      Index           =   0
      Left            =   13320
      TabIndex        =   76
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "��������"
      Height          =   495
      Left            =   3000
      TabIndex        =   15
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "��ħ�ȼ�"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "��ħ��"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   1800
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(1 To 100), fumomanji As Long '��i����ħ���Ե���ߵȼ�����
Dim b(1 To 100), n As Long '��i����ħ���Ե����յȼ�
Dim t, bukechongfu As Boolean
Dim c(1 To 100), e(1 To 100), x As String 'c�����ַ�����1����2����3��ħ��ǩ
'e�ǵ�i����ħ���Ե�����
Dim k, j, buu As Long
Dim d(1 To 10) As Integer '������������
Dim yanhua, yaoshui As String '��c���
Dim st(0 To 20), st0, sx1(0 To 20), sx2(0 To 20), s1 As String 's1�Ƕ�Ӧ�ļ�·��
Dim shuxingkaiqi, mcfun As Boolean
Dim Lore(0 To 20), LoreC(0 To 20), ID(0 To 5), describe As String
Dim yuansu(1 To 100) As Boolean
Dim baoku As String
Dim mofaquan(0 To 100) As Single
Dim choujiangcishu As Integer
Dim baojushu As Long
Dim pinjiming(0 To 20) As String
Dim pinjicolor(0 To 20) As ColorConstants
Dim laiyuan(0 To 100, 0 To 1), laiy As String
Const kkkk = 11
Const manji = 1024
Const mo = 5






Private Sub Combo1_Change()

End Sub

Private Sub Command1_Click()

k = 1
x = Text1.Text
caizhi = x
If x = "ľ" Then caizhi = "wooden"
If x = "ʯ" Then caizhi = "stone"
If x = "����" Then caizhi = "chainmail"
If x = "��" Then caizhi = "golden"
If x = "��" Then caizhi = "iron"
If x = "��ʯ" Then caizhi = "diamond"
If x = "Ƥ��" Then caizhi = "leather"
If x = "�½�Ͻ�" Then caizhi = "netherite"
If x = "����" Then caizhi = "lingering"
If x = "�罦" Then caizhi = "splash"
If x = "������" Then caizhi = "polar_bear"
c(1) = caizhi '����
If x <> "" Then c(1) = c(1) + "_"
c(2) = wuzhi(Text2.Text) '����



st(0) = "{"

If Mid(c(3), 2) <> "" Then
st(1) = "Enchantments:[" + Mid(c(3), 2) + "],"
Else
st(1) = ""
End If

If Text8(0) <> "" Then
st(2) = "Fireworks:{Flight:" + CStr(Val(Text8(0).Text)) + "b,Explosions:[" + yanhua + "]},"
Else
st(2) = ""
End If

If Text10.Text = "" And Text12.Text = "" Then '������������������������������������ע�ʹ���������������BEGIN������������������������
st(3) = ""
Else
st(3) = "display:{Name:'{""text"":""" + Text10.Text + """,""color"":""" + Text18.Text + """,""italic"":false,""bold"":true}',"
st(3) = st(3) + "Lore:["
'ID,Ʒ�ף�����������֮��
Lore(0) = "'{""text"":""ID:" + Text31.Text + """,""color"":""gold"",""italic"":false,""bold"":false}',"

LoreC(1) = "gray"
If Text29.Text = "��ͨ" Then
LoreC(1) = "white"
ElseIf Text29.Text = "����" Then
LoreC(1) = "green"
ElseIf Text29.Text = "ϡ��" Then
LoreC(1) = "blue"
ElseIf Text29.Text = "ʷʫ" Then
LoreC(1) = "dark_purple"
ElseIf Text29.Text = "��˵" Then
LoreC(1) = "gold"
ElseIf Text29.Text = "��" Then
LoreC(1) = "red"
ElseIf Text29.Text = "����" Then
LoreC(1) = "yellow"
ElseIf Text29.Text = "����" Then
LoreC(1) = "aqua"
End If
Lore(1) = "'{""text"":""" + Text29.Text + " " + Text30.Text + """,""color"":""" + LoreC(1) + """,""italic"":false,""bold"":true}'," 'Ʒ��

If Text35(3).Text <> "" Then
LoreC(2) = Text19.Text '����
Lore(2) = "'{""text"":""" + Text35(3).Text + """,""color"":""" + LoreC(2) + """,""italic"":false,""bold"":true}',"
Else
Lore(2) = ""
End If

LoreC(3) = Text19.Text '����
Lore(3) = "'{""text"":""" + Text12.Text + Text35(0).Text + "," + Text35(1).Text + Text33.Text + Text35(2).Text + """,""color"":""" + LoreC(3) + """,""italic"":false,""bold"":true}',"


LoreC(4) = "dark_gray"
If Text32.Text = "������֮�" Then
LoreC(4) = "yellow"
ElseIf Text32.Text = "��ǰ�����Ρ�" Then
LoreC(4) = "green"
ElseIf Text32.Text = "������߽�" Then
LoreC(4) = "blue"
ElseIf Text32.Text = "������֮�" Then
LoreC(4) = "aqua"
ElseIf Text32.Text = "��Զ�Ż��졹" Then
LoreC(4) = "dark_green"
ElseIf Text32.Text = "���ۻ���" Then
LoreC(4) = "red"
ElseIf Text32.Text = "������֮�ۡ�" Then
LoreC(4) = "dark_red"
ElseIf Text32.Text = "������л꡹" Then
LoreC(4) = "white"
End If
Lore(4) = "'{""text"":""" + Text32.Text + "��"",""color"":""" + LoreC(4) + """,""italic"":false,""bold"":true}',"

For i = 0 To 4

If Text34(i).Text <> "" Then
LoreC(i + 5) = Text36(i).Text
Lore(i + 5) = "'{""text"":""" + Text34(i).Text + """,""color"":""" + LoreC(i + 5) + """,""italic"":false,""bold"":true}',"
Else
Lore(i + 5) = ""
End If

Next i


For i = 0 To 9
st(3) = st(3) + Lore(i)
Next i

st(3) = st(3) + "]},"



End If '������������������������������������ע�ʹ�����END����������������������������������

sxkqlm = False
For i = 0 To kkkk
If Text9(i).Text <> "" Then sxkqlm = True
Next
If shuxingkaiqi And sxkqlm Then
st(4) = "AttributeModifiers:["
tst = ""
If Text21.Text <> "" Then
slot = ",Slot:" + Text21.Text
Else
slot = ""
End If
For i = 0 To kkkk
If Text16.Text <> "���" Then
uuid = Text16.Text
Else
uuid = "[I;"
For ii = 1 To 3
uuid = uuid + CStr(Int(200000000 * Rnd()) - 100000000) + ","
Next
uuid = uuid + CStr(Int(200000000 * Rnd()) - 100000000) + "]"
End If
'If Text9(i).Text <> "" Then tst = tst + ",{Operation:" + Text17.Text + ",UUID:" + uuid + ",Amount:" + Text9(i).Text + ",Name:" + Chr(34) + sx2(i) + Chr(34) + slot + "}"
If Text9(i).Text <> "" Then tst = tst + ",{Operation:" + Text17.Text + ",UUID:" + uuid + ",Amount:" + Text9(i).Text + ",AttributeName:" + Chr(34) + sx1(i) + Chr(34) + ",Name:" + Chr(34) + sx2(i) + Chr(34) + slot + "}"
Next
st(4) = st(4) + Mid(tst, 2) + "],"
Else
st(4) = ""
End If

If Text11.Text = "" Or Val(Text11.Text) = 0 Then
st(5) = ""
Else
st(5) = "HideFlags:" + Text11.Text + ","
End If

If yaoshui <> "" Then 'ҩˮ
st(6) = "CustomPotionEffects:[" + yaoshui + "],CustomPotionColor:[I;" + CStr(65536 * Val(Text15(0).Text) + 256 * Val(Text15(1).Text) + Val(Text15(2).Text)) + "],Color:[I;" + CStr(65536 * Val(Text15(0).Text) + 256 * Val(Text15(1).Text) + Val(Text15(2).Text)) + "],Potion:" + Chr(34) + "minecraft:" + Text8(12).Text + Chr(34) + ","
Else
st(6) = ""
End If

If Text23.Text <> "" Then st(7) = "EntityTag:" + Text23.Text + "," Else st(7) = ""

If Text26.Text <> "" And c(2) = "crossbow" Then st(8) = "ChargedProjectiles:[" + Text26.Text + "],Charged:1b," Else st(8) = "" '������


If Text20.Text = "" Or Val(Text20.Text) = 0 Then
st(9) = ""
Else
st(9) = "Damage:" + Text20.Text + ","
End If

If d(1) = 1 Then
st(20) = "Unbreakable:" + CStr(d(1))
Else
st(20) = ""
End If

For i = 1 To 20
st(0) = st(0) + st(i)
Next i

If Mid(st(0), Len(st(0))) = "," Then st(0) = Mid(st(0), 1, Len(st(0)) - 1)

st(0) = st(0) + "}"

If mcfun = True Then
st0 = "give " + Text7.Text + " " + c(1) + c(2) + st(0) + " " + Text6.Text 'nbt��ǩ��start
Else
st0 = "{id:" + Chr(34) + "minecraft:" + c(1) + c(2) + Chr(34) + ",Count:" + Text6.Text + "b,tag:" + st(0) + "}"
End If

Dim st00 As String
st00 = st0
st0 = ""
For i = 1 To Len(st00)
If Mid(st00, i, 1) = Chr(34) Then
st0 = st0 + Text37.Text + Chr(34)
Else
st0 = st0 + Mid(st00, i, 1)
End If
Next i

Text5.Text = st0
's1 = "D:\VB\temp\" + Text3.Text + ".mcfunction"
s1 = "D:\VB\temp\temp.mcfunction"
If mcfun = True Then
Open s1 For Output As #1
Print #1, st0
Close #1
ElseIf mcfun = False Then
Open Text14.Text & "\" + Text22.Text + "\" + Text13.Text + ".txt" For Output As #1
Print #1, st0
Close #1
End If


Shell "cmd /c D:\VB\temp\change.py"
'/c /k /q


End Sub

Private Sub Command10_Click()
Dim yanse, xingzhuang As String
If Text8(1).Text <> "���" Then xingzhuang = Text8(1).Text
If Text8(3).Text <> "���" Then yanse = Text8(3).Text
For i = 1 To Val(Text8(4).Text)
If Text8(1).Text = "���" Then xingzhuang = CStr(Int(Rnd() * 5))
If Text8(3).Text = "���" Then yanse = CStr(Int(Rnd() * 65536 * 256 - 1))
If yanhua = "" Then
yanhua = "{Type:" + xingzhuang + "b,Trail:" + Text8(2).Text + "b,Flicker:" + Text8(5).Text + "b,Colors:[I;" + yanse + "]}"
Else
yanhua = yanhua + ",{Type:" + xingzhuang + "b,Trail:" + Text8(2).Text + "b,Flicker:" + Text8(5).Text + "b,Colors:[I;" + yanse + "]}"
End If
Next i
List2.AddItem Text8(1).Text + bu(Text8(1).Text) + Text8(2).Text + bu(Text8(2).Text) + Text8(5).Text + bu(Text8(5).Text) + Text8(3).Text + bu(Text8(3).Text) + Text8(4).Text + bu(Text8(4).Text)
End Sub

Private Sub Command11_Click()
yanhua = ""
List2.Clear
List2.AddItem "��״       �켣       ��˸       ��ɫ       ����"
End Sub


Private Sub Command12_Click(Index As Integer)
Text14.Text = "D:\MC\1.20���°汾\.minecraft\saves\" + Command12(Index).Caption + "\datapacks\ħ������\data\give\functions"
End Sub


Private Sub Command13_Click()
shuxingkaiqi = Not shuxingkaiqi
If shuxingkaiqi = True Then Command13.Caption = "��������" Else Command13.Caption = "�ر�����"
End Sub

Private Sub Command14_Click()
Text8(3).Text = CStr(65536 * Val(Text15(0).Text) + 256 * Val(Text15(1).Text) + Val(Text15(2).Text))
Text27.BackColor = RGB((Text15(0).Text), Val(Text15(1).Text), Val(Text15(2).Text))
End Sub

Private Sub Command15_Click(Index As Integer)
Text13.Text = Command15(Index).Caption
Text22.Text = Command15(Index).Caption
If Index = 0 Then Text14.Text = "D:\MC\.minecraft\saves\" + Command12(Index).Caption + "\datapacks\ħ������\data\temp\functions"
If Index >= 1 And Index <= 2 Then Text14.Text = "D:\VB\MC������\ʵ�幹��\��ƷĿ¼"

End Sub

Private Sub Command16_Click()
Text13.Text = Text22.Text
End Sub

Private Sub Command17_Click()
Text16.Text = "[I;"
For i = 1 To 3
Text16.Text = Text16.Text + CStr(Int(200000000 * Rnd()) - 100000000) + ","
Next
Text16.Text = Text16.Text + CStr(Int(200000000 * Rnd()) - 100000000) + "]"
End Sub

Private Sub Command18_Click()
For i = 0 To kkkk
Text9(i).Text = ""
Next i
End Sub

Private Sub Command19_Click()
Text3.Text = CStr(manji)
End Sub

Private Sub Command2_Click()

k = k + 1
a(k) = fumomanji

Label2.Caption = Text4.Text

If Text4.Text = "ˮ���پ�" Then
Label2.Caption = "aqua_affinity"
yuansu(1) = True 'ˮϵ
a(k) = 1
End If
If Text4.Text = "��ը����" Then
Label2.Caption = "blast_protection"
a(k) = 10
End If
If Text4.Text = "�̽����" Then
Label2.Caption = "depth_strider"
yuansu(1) = True 'ˮϵ
a(k) = 3
End If
If Text4.Text = "ˤ�䱣��" Then
Label2.Caption = "feather_falling"
a(k) = 7
yuansu(6) = True '��ϵ
End If
If Text4.Text = "���汣��" Then
Label2.Caption = "fire_protection"
a(k) = 10
yuansu(2) = True '��ϵ
End If
If Text4.Text = "�����ﱣ��" Then
Label2.Caption = "projectile_protection"
yuansu(6) = True '��ϵ
a(k) = 10
End If
If Text4.Text = "��˪����" Then
Label2.Caption = "frost_walker"
a(k) = 14
yuansu(3) = True '��ϵ
End If
If Text4.Text = "�����޲�" Then
Label2.Caption = "mending"
a(k) = 1
End If
If Text4.Text = "����" Then
Label2.Caption = "protection"
a(k) = 20
End If
If Text4.Text = "ˮ�º���" Then
Label2.Caption = "respiration"
yuansu(1) = True 'ˮϵ
End If
If Text4.Text = "��꼲��" Then Label2.Caption = "soul_speed"
If Text4.Text = "����" Then
Label2.Caption = "thorns"
yuansu(7) = True 'ľϵ
End If
If Text4.Text = "�;�" Then Label2.Caption = "unbreaking"


If Text4.Text = "���渽��" Then
Label2.Caption = "fire_aspect"
yuansu(2) = True '��ϵ
End If
If Text4.Text = "��֫ɱ��" Then Label2.Caption = "bane_of_arthropods"
If Text4.Text = "����ɱ��" Then
Label2.Caption = "smite"
yuansu(4) = True '��ϵ
End If
If Text4.Text = "����" Then
Label2.Caption = "knockback"
yuansu(6) = True '��ϵ
End If
If Text4.Text = "����" Then Label2.Caption = "looting"
If Text4.Text = "����" Then Label2.Caption = "sharpness"
If Text4.Text = "��ɨ֮��" Then
Label2.Caption = "sweeping"
yuansu(6) = True '��ϵ
End If

If Text4.Text = "��ʸ" Then
Label2.Caption = "flame"
yuansu(2) = True '��ϵ
a(k) = 1
End If
If Text4.Text = "����" Then
Label2.Caption = "infinity"
a(k) = 1
End If
If Text4.Text = "����" Then Label2.Caption = "power"
If Text4.Text = "���" Then
Label2.Caption = "punch"
yuansu(6) = True '��ϵ
End If

If Text4.Text = "�������" Then
Label2.Caption = "multishot"
a(k) = 1
End If
If Text4.Text = "��͸" Then
Label2.Caption = "piercing"
a(k) = 127
End If
If Text4.Text = "����װ��" Then
Label2.Caption = "quick_charge"
a(k) = 5
End If

If Text4.Text = "����" Then
Label2.Caption = "channeling"
yuansu(5) = True '��ϵ
a(k) = 1
End If
If Text4.Text = "�ҳ�" Then
Label2.Caption = "loyalty"
a(k) = 127
End If
If Text4.Text = "����" Then
Label2.Caption = "riptide"
yuansu(1) = True 'ˮϵ
End If
If Text4.Text = "����" Then Label2.Caption = "impaling"

If Text4.Text = "��׼�ɼ�" Then
Label2.Caption = "silk_touch"
yuansu(8) = True '��ϵ
a(k) = 1
End If
If Text4.Text = "Ч��" Then
Label2.Caption = "efficiency"
yuansu(8) = True '��ϵ
End If
If Text4.Text = "ʱ��" Then
Label2.Caption = "fortune"
yuansu(8) = True '��ϵ
End If


If Text4.Text = "��֮���" Then
Label2.Caption = "luck_of_the_sea"
yuansu(1) = True 'ˮϵ
End If
If Text4.Text = "����" Then
Label2.Caption = "lure"
yuansu(1) = True 'ˮϵ
a(k) = 5
End If
e(k) = Text4.Text
t = False
For i = 1 To k - 1
If e(k) = e(i) Then t = bukechongfu '���غϻ�ʧЧ
Next i
If t = True Then k = k - 1

n = Val(Text3.Text)
If t = False Then List1.AddItem Text4.Text + bu(Text4.Text) + Text3.Text + bu(Text3.Text) + Str(a(k)) + bu(Str(a(k))) + Str((n + a(k) - Abs(n - a(k))) / 2)
c(3) = c(3) + ",{id:" + Label2.Caption + ",lvl:" + CStr((n + a(k) - Abs(n - a(k))) / 2) + "}"
End Sub




Private Sub Command20_Click()
For i = 0 To kkkk
Text9(i).Text = CStr(manji)
Next i
Text9(1).Text = "1"
Text9(2).Text = "0.4" 'ʵ����1024
Text9(3).Text = "2048"
Text9(4).Text = "20"
Text9(5).Text = "30"
Text9(8).Text = ""
Text9(9).Text = ""
Text9(10).Text = ""
Text9(11).Text = ""
End Sub

Private Sub Command21_Click()
yaoshui = ""
List3.Clear
List3.AddItem "ID         ����       ʱ��       ����       ����       ͼ��"
End Sub

Private Sub Command22_Click()

If yaoshui = "" Then
yaoshui = "{Id:" + Text8(11).Text + ",Amplifier:" + Text8(10).Text + ",Duration:" + Text8(9).Text + ",Ambient:" + Text8(6).Text + ",ShowParticles:" + Text8(8).Text + ",ShowIcon:" + Text8(7).Text + "}"
Else
yaoshui = yaoshui + ",{Id:" + Text8(11).Text + ",Amplifier:" + Text8(10).Text + ",Duration:" + Text8(9).Text + ",Ambient:" + Text8(6).Text + ",ShowParticles:" + Text8(8).Text + ",ShowIcon:" + Text8(7).Text + "}"
End If

List3.AddItem Text8(11).Text + bu(Text8(11).Text) + Text8(10).Text + bu(Text8(10).Text) + Text8(9).Text + bu(Text8(9).Text) + Text8(6).Text + bu(Text8(6).Text) + Text8(8).Text + bu(Text8(8).Text) + Text8(7).Text + bu(Text8(7).Text)

End Sub

Private Sub Command23_Click()
For i = 0 To kkkk
Text9(i).Text = "200"
Next i
Text9(1).Text = "1"
Text9(2).Text = "0.4"
Text9(4).Text = "20"
Text9(5).Text = "20"
Text9(8).Text = ""
Text9(9).Text = ""
Text9(10).Text = ""
Text9(11).Text = ""
End Sub

Private Sub Command24_Click(Index As Integer)
Text21.Text = Command24(Index).Caption
End Sub

Private Sub Command25_Click()
If mcfun = False Then MkDir Text14.Text + "\" + Text22.Text + "\"
End Sub

Private Sub Command26_Click()
Open "d:\VB\MC������\ʵ�幹��\ʵ��Ŀ¼\" + Text24.Text + ".txt" For Input As #1
Do While Not EOF(1)
Line Input #1, shiti
Loop
Close #1
Text23.Text = shiti

End Sub

Private Sub Command27_Click()
If yaoshui <> "" Then 'ҩˮ
x = "Effects:[" + yaoshui + "],Potion:" + Chr(34) + "minecraft:" + Text8(12).Text + Chr(34) + ""
Else
x = ""
End If

Text5.Text = x
Open "d:\VB\MC������\ʵ�幹��\��ƷĿ¼\ҩˮ\" + Text25.Text + ".txt" For Output As #1
Print #1, x
Close #1

End Sub

Private Sub Command28_Click()
Text26.Text = Text5.Text
End Sub

Private Sub Command29_Click()
Text26.Text = ""
End Sub

Private Sub Command3_Click(Index As Integer)
Text4.Text = Command3(Index).Caption
End Sub

Private Sub Command30_Click()
If Mid(c(3), 2) <> "" Then
st(10) = "StoredEnchantments:[" + Mid(c(3), 2) + "],"
Else
st(10) = ""
End If
End Sub

Private Sub Command31_Click()


FileCopy s1, Text14.Text & "\" + Text13.Text + ".mcfunction"
End Sub

Private Sub Command32_Click(Index As Integer)
Text29.Text = Command32(Index).Caption
Text29.ForeColor = pinjicolor(Index)

If Command32(Index).Caption = "��ͨ" Then
ID(2) = "N"
ElseIf Command32(Index).Caption = "����" Then
ID(2) = "S"
ElseIf Command32(Index).Caption = "ϡ��" Then
ID(2) = "R"
ElseIf Command32(Index).Caption = "ʷʫ" Then
ID(2) = "E"
ElseIf Command32(Index).Caption = "��˵" Then
ID(2) = "L"
ElseIf Command32(Index).Caption = "��" Then
ID(2) = "M"
ElseIf Command32(Index).Caption = "����" Then
ID(2) = "I"
ElseIf Command32(Index).Caption = "����" Then
ID(2) = "O"
ElseIf Command32(Index).Caption = "����" Then
ID(2) = "G"
End If
Text31.Text = ID(0) + ID(1) + ID(2) + ID(3)

End Sub

Private Sub Command33_Click(Index As Integer)
Text30.Text = Command33(Index).Caption

If Command33(Index).Caption = "����" Then
ID(1) = "W"
ElseIf Command33(Index).Caption = "װ��" Then
ID(1) = "E"
ElseIf Command33(Index).Caption = "����" Then
ID(1) = "F"
ElseIf Command33(Index).Caption = "ҩˮ" Then
ID(1) = "S"
ElseIf Command33(Index).Caption = "����" Then
ID(1) = "M"
ElseIf Command33(Index).Caption = "����" Then
ID(1) = "C"
ElseIf Command33(Index).Caption = "����" Then
ID(1) = "A"
ElseIf Command33(Index).Caption = "��ʯ" Then
ID(1) = "G"
End If

Text31.Text = ID(0) + ID(1) + ID(2) + ID(3)

End Sub

Private Sub Command34_Click(Index As Integer)
Text32.Text = Command34(Index).Caption
End Sub

Private Sub Command35_Click(Index As Integer)
Text33.Text = Command35(Index).Caption
laiy = laiyuan(Index, 1)
End Sub

Private Sub Command36_Click() '������Ʒ��Ϣ

allpoint = 0
For i = 0 To 11
multi = 1
If i = 0 Then multi = 3
If i = 2 Then multi = 0.15
If i = 6 Then multi = 0.8
allpoint = allpoint + Val(Text9(i).Text) / multi
Next i
multi = 1
If Val(Text17.Text) > 0 Then multi = 4
If Val(Text17.Text) > 1 Then Call Command33_Click(6)
allpoint = allpoint * multi + 1
pinji = 0
pinji = Int(Log(allpoint / 5) / Log(2))
If pinji > 8 Then pinji = 8
If pinji < 0 Then pinji = 0

Call Command32_Click(Int(pinji))

HScroll2.Value = Int(888 * Rnd() + 3)

If Text29.Text = "��" Or Text29.Text = "����" Or Text29.Text = "����" Or Text30.Text = "����" Then
d(1) = 1
Else
d(1) = 0
End If

poxian = 0.2
weiyi = 0.4
chaotuo = 0.8


If d(1) = 0 Then Command5.Caption = "���Դݻ�"
If d(1) = 1 Then Command5.Caption = "���ɴݻ�"

Text35(0).Text = Text1.Text + Text2.Text

iname = Text2.Text

aliangci = ""

Key = Int(200 * Rnd()) 'ǿ��ָ��������

If iname = "��" Then
    
    If Key < 60 Then
    iname = "ս��"
    ElseIf Key < 120 Then iname = "��"
    End If
    aliangci = "��"

ElseIf iname = "����" Then
    If Key < 60 Then
    iname = "�޶�"
    ElseIf Key < 110 Then iname = "ս��"
    ElseIf Key < 160 Then iname = "��"
    End If
    aliangci = "��"

ElseIf iname = "�ʳ�" Then
    If Key < 50 Then
    iname = "ս��"
    ElseIf Key < 90 Then iname = "������"
    ElseIf Key < 130 Then iname = "ս��"
    ElseIf Key < 160 Then iname = "���"
    End If
    aliangci = "��"

ElseIf iname = "�����" Then
    If Key < 60 Then
    iname = "���컭�"
    ElseIf Key < 110 Then iname = "ս�"
    ElseIf Key < 160 Then iname = "�"
    End If
    aliangci = "��"

ElseIf iname = "��" Then
    If Key < 40 Then
    iname = "����"
    ElseIf Key < 80 Then iname = "ս��"
    ElseIf Key < 110 Then iname = "����"
    ElseIf Key < 160 Then iname = "��"
    End If
    aliangci = "��"

ElseIf iname = "��" Then
    If Key < 60 Then
    iname = "ħ����"
    ElseIf Key < 110 Then iname = "����"
    ElseIf Key < 160 Then iname = "ħ���鼮"
    End If
    aliangci = "��"

ElseIf iname = "��" Then
    If Key < 30 Then
    iname = "����"
    ElseIf Key < 60 Then iname = "����"
    ElseIf Key < 90 Then iname = "ħ��"
    End If
    aliangci = "��"

ElseIf iname = "�ؼ�" Then
    If Key < 40 Then
    iname = "����"
    ElseIf Key < 80 Then iname = "����"
    ElseIf Key < 120 Then iname = "ս��"
    ElseIf Key < 160 Then iname = "��"
    End If
    
    aliangci = "��"

ElseIf iname = "ͷ��" Then
    If Key < 60 Then
    iname = "��"
    ElseIf Key < 90 Then iname = "ñ��"
    ElseIf Key < 120 Then iname = "��ñ"
    End If
    aliangci = "��"

ElseIf iname = "ѥ��" Then
    If Key < 50 Then
    iname = "սѥ"
    ElseIf Key < 90 Then iname = "ħ��սѥ"
    ElseIf Key < 150 Then iname = "ѥ"
    End If
    aliangci = "˫"

ElseIf iname = "��" Then
    If Key < 60 Then
    iname = "�乭"
    ElseIf Key < 110 Then iname = "��"
    ElseIf Key < 160 Then iname = "����"
    End If
    aliangci = "��"

ElseIf iname = "ҩˮ" Then
    If Key < 60 Then
    iname = "ħҩ"
    ElseIf Key < 110 Then iname = "��ҩ"
    ElseIf Key < 160 Then iname = "ħ��ҩ��"
    End If
    aliangci = "ƿ"

ElseIf iname = "�̻����" Then
    If Key < 60 Then
    iname = "�̻�"
    ElseIf Key < 120 Then iname = "�̻�"
    End If
    aliangci = "��"


ElseIf iname = "�����" Then
    If Key < 60 Then
    iname = "���"
    ElseIf Key < 120 Then iname = "����"
    End If
    aliangci = "��"

End If



If Text2.Text = "��" Or Text2.Text = "��" Or Text2.Text = "��" Or Text2.Text = "��" Or Text2.Text = "��" Or Text2.Text = "��" Or Text2.Text = "��" Or Text2.Text = "�����" Or Text2.Text = "�����" Or Text2.Text = "����" Or Text2.Text = "�̻����" Or Text2.Text = "��" Or Text2.Text = "��" Then
Text30.Text = "����"
If aliangci = "" Then aliangci = "��"
ElseIf Text2.Text = "ͷ��" Or Text2.Text = "�ؼ�" Or Text2.Text = "����" Or Text2.Text = "ѥ��" Or Text2.Text = "�ʳ�" Or Text2.Text = "��" Then
Text30.Text = "װ��"
If aliangci = "" Then aliangci = "��"
End If

If Command33(Index).Caption = "����" Then
ID(1) = "W"
ElseIf Command33(Index).Caption = "װ��" Then ID(1) = "E"
ElseIf Command33(Index).Caption = "����" Then ID(1) = "F"
ElseIf Command33(Index).Caption = "ҩˮ" Then ID(1) = "S"
ElseIf Command33(Index).Caption = "����" Then ID(1) = "M"
ElseIf Command33(Index).Caption = "����" Then ID(1) = "C"
ElseIf Command33(Index).Caption = "����" Then ID(1) = "A"
ElseIf Command33(Index).Caption = "��ʯ" Then ID(1) = "G"
End If

Text31.Text = ID(0) + ID(1) + ID(2) + ID(3)

whetherdes = False
For i = 0 To 7
If Command32(i).Caption = Text29.Text Then whetherdes = True
Next i

If whetherdes = False Then describe = Text29.Text 'Ʒ��

Dim quanjuyuansu(1 To 8) As Boolean
yuansunum = ""
yuansudes = ""
If (Val(Text9(2).Text) > 1.2 And Val(Text17.Text) = 0) Or (Val(Text9(2).Text) > 3 And Val(Text17.Text) > 0) Or (Val(Text9(6).Text) > 10 And Val(Text17.Text) = 0) Or (Val(Text9(6).Text) > 3 And Val(Text17.Text) > 0) Then quanjuyuansu(6) = True
'��Ԫ��
If (Val(Text9(0).Text) > 100 And Val(Text17.Text) = 0) Or (Val(Text9(0).Text) > 5 And Val(Text17.Text) > 0) Then quanjuyuansu(7) = True
'ľԪ��
If (Val(Text9(5).Text) > 20 And Val(Text17.Text) = 0) Or (Val(Text9(5).Text) > 5 And Val(Text17.Text) > 0) Then quanjuyuansu(8) = True

For i = 1 To 8: quanjuyuansu(i) = quanjuyuansu(i) Or yuansu(i): Next i

If quanjuyuansu(1) Then yuansudes = yuansudes + "ˮ"
If quanjuyuansu(2) Then yuansudes = yuansudes + "��"
If quanjuyuansu(3) Then yuansudes = yuansudes + "��"
If quanjuyuansu(4) Then yuansudes = yuansudes + "��"
If quanjuyuansu(5) Then yuansudes = yuansudes + "��"
If quanjuyuansu(6) Then yuansudes = yuansudes + "��"
If quanjuyuansu(7) Then yuansudes = yuansudes + "ľ"
If quanjuyuansu(8) Then yuansudes = yuansudes + "��"
yuansudesc = yuansudes

If Len(yuansudes) = 1 Then yuansunum = ""
If Len(yuansudes) = 2 Then yuansunum = "��"
If Len(yuansudes) = 3 Then yuansunum = "��"
If Len(yuansudes) = 4 Then yuansunum = "��"
If Len(yuansudes) = 5 Then yuansunum = "��"
If Len(yuansudes) = 6 Then yuansunum = "��"
If Len(yuansudes) = 7 Then yuansunum = "��"
If Len(yuansudes) = 8 Then yuansunum = "��"
yuansudes = yuansudes + yuansunum

If Len(yuansudes) > 0 Then yuansudes = yuansudes + "ϵ"




Key = Int(200 * Rnd()) 'ǿ��ָ��������

cname = Text1.Text

If Text1.Text = "ľ" Then
If Key < 50 Then
cname = "��ľ" '_____________--------------------------------����----------------------------
ElseIf Key < 100 Then cname = "��Ȫľ"
ElseIf Key < 120 Then cname = "����"
ElseIf Key < 160 Then cname = "��ľ"
Else: cname = "��̴ľ"
End If

ElseIf Text1.Text = "ʯ" Then
If Key < 40 Then
cname = "��ʯ"
ElseIf Key < 80 Then cname = "������"
ElseIf Key < 110 Then cname = "ʯ��"
ElseIf Key < 140 Then cname = "������"
ElseIf Key < 170 Then cname = "����ʯ"
Else: cname = "����"
End If

ElseIf Text1.Text = "��" Then
If Key < 40 Then
cname = "��"
ElseIf Key < 80 Then cname = "����"
ElseIf Key < 110 Then cname = "��"
ElseIf Key < 140 Then cname = "����"
ElseIf Key < 170 Then cname = "�Ͻ�"
Else: cname = "����"
End If

ElseIf Text1.Text = "��" Then
If Key < 40 Then
cname = "�ƽ�"
ElseIf Key < 80 Then cname = "����"
ElseIf Key < 110 Then cname = "ҫ��"
ElseIf Key < 140 Then cname = "�ǽ�"
ElseIf Key < 170 Then cname = "ħ��"
Else: cname = "��"
End If

ElseIf Text1.Text = "��ʯ" Then
If Key < 40 Then
cname = "��ʯ"
ElseIf Key < 80 Then cname = "����"
ElseIf Key < 110 Then cname = "����"
ElseIf Key < 140 Then cname = "ˮ��"
ElseIf Key < 170 Then cname = "����ʯ"
Else: cname = "��ʯ"
End If

ElseIf Text1.Text = "�½�Ͻ�" Then
If Key < 40 Then
cname = "�ڽ�"
ElseIf Key < 80 Then cname = "�ڸ�"
ElseIf Key < 110 Then cname = "�ھ�"
ElseIf Key < 140 Then cname = "���"
ElseIf Key < 170 Then cname = "�ڰ��Ͻ�"
Else: cname = "��Ԩ�Ͻ�"
End If

ElseIf Text1.Text = "����" Then
If Key < 50 Then
cname = "����"
ElseIf Key < 100 Then cname = "����"
ElseIf Key < 150 Then cname = "����"
Else: cname = "������"
End If

ElseIf Text1.Text = "Ƥ��" Then
If Key < 50 Then
cname = "��Ƥ"
ElseIf Key < 100 Then cname = "Ƥ��"
ElseIf Key < 150 Then cname = "�޲�"
Else: cname = "����"
End If

End If '_____________---------------------------------------����---------------------

Key = Int(200 * Rnd()) '������
Text34(2).Text = ""


If Mid(laiy, 1, 2) = "����" Then

    If yuansunum = "��" Then yuansudes = "ȫԪ������"


    If Text29.Text = "��ͨ" Then '_____________---------Ʒ��---------------------------------------------------
        
        If Key < 50 Then
        describe = "�洦�ɼ���"
        Text34(0).Text = "\\""��" + aliangci + iname + "��" + cname + "���ģ����Ǻܰ��ġ�\\"""
        Text34(1).Text = "���������˹���������"
        Text34(2).Text = ""
        ElseIf Key < 100 Then
        describe = "���Ƶ�"
        Text34(0).Text = "\\""��" + aliangci + iname + "��ĺ�һ�㣡\\"""
        If yuansudes <> "" Then
        Text34(1).Text = "\\""���£����ƺ����̺���" + yuansudes + "��������\\"""
        Else
        Text34(1).Text = "\\""̫�ֲ��ˣ�\\"""
        End If
        Text34(2).Text = "����������ʦ���Ŷ�"
        ElseIf Key < 150 Then
        describe = "���Ƶ�"
        Text34(0).Text = "\\""��" + aliangci + iname + "��ĺ��ã���Ҳ�е��á�\\"""
        Text34(1).Text = "������ð���ߡ���¬˼"
        Text34(2).Text = ""
        Else
        describe = "���޳�����"
        Text34(0).Text = "\\""һ" + aliangci + "�ܳ�����" + iname + "�������������ͨ��\\"""
        Text34(1).Text = "�������ؾ����ˡ�����˹"
        Text34(2).Text = ""
        End If
        Text36(0).Text = "white"
        
        Text18.Text = "#"
        For i = 1 To 3
        Text18.Text = Text18.Text + CStr(86 + Int(Rnd() * 14))
        Next i
        If yuansudes = "" Then
        Text12.Text = "ƽ����"
        Else
        Text12.Text = "�����Ѻ۵�" + yuansudes
        describe = "���Ƶ�"
        End If
    
    ElseIf Text29.Text = "����" Then
        If Key < 50 Then
        describe = "������"
        Text34(0).Text = "\\""��" + aliangci + iname + "�Һ�ϲ����\\"""
        Text34(1).Text = "������սʿ��ϯ��"
        ElseIf Key < 100 Then
        describe = "���Ƶ�"
        Text34(0).Text = "\\""���������˹���ŵ���ʽװ����\\"""
        Text34(1).Text = "���������ų���������˹"
        ElseIf Key < 150 Then
        describe = "��ĥ����"
        Text34(0).Text = "\\""��" + aliangci + iname + "��������\\"""
        Text34(1).Text = "������Ӣ�۶���ʦ����˹��"
        Else
        describe = "�����"
        Text34(0).Text = "\\""�����" + iname + "��\\"""
        Text34(1).Text = "���������ܽ�������������˹"
        End If
        Text36(0).Text = "dark_green"
        Text18.Text = "#"
        Text18.Text = Text18.Text + CStr(56 + Int(Rnd() * 14))
        Text18.Text = Text18.Text + "EE"
        Text18.Text = Text18.Text + CStr(56 + Int(Rnd() * 14))
        If yuansudes = "" Then
        Text12.Text = "�ǳ������"
        Else
        Text12.Text = "��������" + yuansudes
        End If
        
        ElseIf Text29.Text = "ϡ��" Then
        If Key < 50 Then
        describe = "������"
        Text34(0).Text = "\\""��" + aliangci + iname + "����������һ��\\"""
        Text34(1).Text = "��������ɫ�о�������˹"
        Text34(2).Text = ""
        ElseIf Key < 100 Then
        describe = "�Ҵ���"
        Text34(0).Text = "\\""�����Ҽ���������������" + cname + iname + "��\\"""
        If yuansudes <> "" Then
        Text34(1).Text = "\\""���ĸ����㣡�����̺���" + yuansudes + "��������\\"""
        Else
        Text34(1).Text = "\\""�ƺ��Ǵ�ĳ����ս���ϼ�����ģ�\\"""
        End If
        Text34(2).Text = "�������ƽ���ʿ��������˹"
        ElseIf Key < 150 Then
        describe = "���ѵ�"
        Text34(0).Text = "\\""������" + aliangci + iname + "����̫���ˣ�\\"""
        Text34(1).Text = "\\""�뵽��" + cname + "ȥ��������һ������ţ�\\"""
        Text34(2).Text = "������ս�����ܼҡ���˹��"
        Else
        describe = "ǿ����"
        Text34(0).Text = "\\""����������" + iname + "����������\\"""
        Text34(1).Text = "����������ʿ��ķ˹���"
        Text34(2).Text = ""
        End If
        Text36(0).Text = "dark_blue"
        Text18.Text = "#"
        Text18.Text = Text18.Text + CStr(7 + Int(Rnd() * 14))
        Text18.Text = Text18.Text + CStr(56 + Int(Rnd() * 14))
        Text18.Text = Text18.Text + "CC"
        If yuansudes = "" Then
        Text12.Text = "Ʒ�ʼ��ѵ�"
        Else
        Text12.Text = "����΢��" + yuansudes + "ħ��������"
        End If
    
    
    ElseIf Text29.Text = "ʷʫ" Then
        If Key < 50 Then
        describe = "������"
        Text34(0).Text = "\\""���ﲻ��������������" + aliangci + cname + iname + "�������⣡\\"""
        Text34(1).Text = "��������������������"
        Text34(2).Text = ""
        ElseIf Key < 80 Then
        describe = "���Ƶ�"
        Text34(0).Text = "\\""����" + cname + "�ر�����" + yuansudes + iname + "���ܰ��ɣ�\\"""
        Text34(1).Text = "\\""ʲô�������ã����������㲻���ã�\\"""
        Text34(2).Text = "�������ʼҶ���ʦ������"
        ElseIf Key < 110 Then
        describe = "���ϵ�"
        Text34(0).Text = "\\""��" + aliangci + iname + "�ĵ����Ѿ����ɿ�����\\"""
        Text34(1).Text = "��������ʷѧ�ҡ�������"
        Text34(2).Text = ""
        ElseIf Key < 150 Then
        describe = "��ʷ�ƾõ�"
        Text34(0).Text = "\\""����һ�ι���һ" + aliangci + cname + iname + "���Ƴ���ʷ��\\"""
        Text34(1).Text = "�������ݶ��ʵۡ�������˹"
        Text34(2).Text = ""
        Else
        describe = "���ص�"
        Text34(0).Text = "\\""����һ" + aliangci + "���ص�" + iname + "����\\"""
        Text34(1).Text = "������������ʦ���ڶ�����"
        Text34(2).Text = ""
        End If
        Text36(0).Text = "light_purple"
        Text18.Text = "#"
        Text18.Text = Text18.Text + "FF"
        Text18.Text = Text18.Text + CStr(56 + Int(Rnd() * 44))
        Text18.Text = Text18.Text + "FF"
        If yuansudes = "" Then
        Text12.Text = "�����������"
        Else
        Text12.Text = "�̺�ǿ��" + yuansudes + "ħ����"
        End If
    
    
    ElseIf Text29.Text = "��˵" Then
        If Key < 50 Then
        describe = "�����"
        Text34(0).Text = "\\""����һ�ι���" + iname + "�Ĵ�˵��\\"""
        Text34(1).Text = "�����������ߡ��ݶ�˹"
        ElseIf Key < 100 Then
        describe = "�̺��ֲ�" + yuansudes + "ħ����"
        Text34(0).Text = "\\""ʤ�����������" + iname + "�������㣿�������ɣ�\\"""
        Text34(1).Text = "������ħ���ʵۡ��ն���˿"
        ElseIf Key < 150 Then
        describe = "��ʥ��"
        Text34(0).Text = "\\""ʥ�⽫��ף�����" + iname + "��Ӧ��Ҳ�ᡭ��ף���㣿\\"""
        Text34(1).Text = "����������ʥŮ������"
        Else
        describe = "�����ޡ�" + describe
        For i = 0 To 7
        Text9(i).Text = CStr(Val(Text9(i).Text) + Int((Val(Text9(i).Text) + 2) * poxian * Rnd()))
        Next i
        Text34(0).Text = "\\""ͻ�Ƽ��ްɣ�" + cname + iname + "��\\"""
        Text34(1).Text = "�������������ߡ������"
        End If
        Text36(0).Text = "yellow"
        Text18.Text = "#"
        Text18.Text = Text18.Text + "FF"
        Text18.Text = Text18.Text + "DD"
        Text18.Text = Text18.Text + CStr(46 + Int(Rnd() * 34))
        If yuansudes = "" Then
        Text12.Text = "��������������"
        Else
        Text12.Text = "�̺���һλ" + yuansudes + "���淨ʦħ����"
        End If
    
    ElseIf Text29.Text = "��" Then
        If Key < 50 Then
        describe = "������"
        Text34(0).Text = "\\""�ҵ�ѡ������" + aliangci + iname + "ȥ���ʹ���ɣ�\\"""
        Text34(1).Text = "������������˹"
        ElseIf Key < 100 Then
        describe = "����������"
        Text34(0).Text = "\\""�ҵ�" + cname + iname + "�أ�ȥ���ˣ���\\"""
        Text34(1).Text = "��������֮ʥ�ߡ�����"
        ElseIf Key < 150 Then
        describe = "��͵�"
        Text34(0).Text = "\\""��ѡ�ߣ���" + aliangci + iname + "�������ͽ������ˣ�\\"""
        Text34(1).Text = "����������֮����֯"
        Else
        describe = "��ӡ��������"
        Text34(0).Text = "\\""����������" + iname + "�б�����\\"""
        Text34(1).Text = "����������ʹ���׼Ӷ�"
        End If
        Text36(0).Text = "yellow"
        Text18.Text = "#"
        Text18.Text = Text18.Text + "FF"
        Text18.Text = Text18.Text + CStr(56 + Int(Rnd() * 24))
        Text18.Text = Text18.Text + CStr(0 + Int(Rnd() * 24))
        If yuansudes = "" Then
        Text12.Text = "�ƺ���ĳ�����г�������"
        Else
        Text12.Text = "�̺��ֲ�" + yuansudes + "������"
        End If
    
    ElseIf Text29.Text = "����" Then
        If Key < 50 Then
        describe = "�̺����Ե�"
        Text34(0).Text = "\\""�k�����ڣ�"
        If yuansudesc <> "" Then Text34(0).Text = Text34(0).Text + "�k����" + yuansudesc + yuansunum + "Ԫ�ش������㣡"
        Text34(0).Text = Text34(0).Text + "\\"""
        Text34(1).Text = "�����������˾��������"
        ElseIf Key < 100 Then
        describe = "ب������ǰ��"
        Text34(0).Text = "\\""�k���������£������˹�������\\"""
        Text34(1).Text = "�������Ŵ�ѧ�ߡ�¡���"
        ElseIf Key < 150 Then
        describe = "��������ǰ��"
        Text34(0).Text = "\\""����һ��������Ԫ֮ǰ��һ" + aliangci + iname + "�����Ρ���\\"""
        Text34(1).Text = "��������������˹����"
        ElseIf Key < 180 Then
        describe = "��Ψһ��" + describe
        For i = 0 To 7
        Text9(i).Text = CStr(Val(Text9(i).Text) + Int((Val(Text9(i).Text) + 2) * weiyi * Rnd()))
        Next i
        Text34(0).Text = "\\""ӵ�С�Ψһ����" + cname + iname + "����ӵ���˳�����ʱ��Ŀ��ܣ�\\"""
        Text34(1).Text = "������ʱ��֮����ŵ��"
        Else
        describe = "������"
        Text34(0).Text = "\\""���ﶼ�ḯ�ࡣ�����" + aliangci + cname + iname + "��һ�����⣡\\"""
        Text34(1).Text = "����������֮����ŵ��"
        End If
        Text36(0).Text = "gold"
        Text18.Text = "#"
        Text18.Text = Text18.Text + "FF"
        Text18.Text = Text18.Text + CStr(77 + Int(Rnd() * 23))
        Text18.Text = Text18.Text + CStr(10 + Int(Rnd() * 24))
        If yuansudes = "" Then
        Text12.Text = "���㲻�𡢳��ѹ�����"
        Else
        Text12.Text = "��ӡ��" + yuansudesc + yuansunum + "Ԫ�����Ȩ����"
        End If
    
    
    ElseIf Text29.Text = "����" Then
        If Key < 50 Then
        describe = "ԭ����"
        Text34(0).Text = "\\""���Ǵ������" + iname + "��������ò��ˡ�\\"""
        Text34(1).Text = "������ʱ֮��ʹ������"
        ElseIf Key < 100 Then
        describe = "������"
        Text34(0).Text = "\\""����һ" + aliangci + "��֤�����絮����" + iname + "��\\"""
        Text34(1).Text = "��������Ԩ�������ɵ�"
        ElseIf Key < 150 Then
        describe = "�����ѡ�" + describe
        For i = 0 To 7
        Text9(i).Text = CStr(Val(Text9(i).Text) + Int((Val(Text9(i).Text) + 2) * chaotuo * Rnd()))
        Next i
        Text34(0).Text = "\\""ӵ�С����ѡ���" + iname + "�����п��ִܵ�˰���\\"""
        Text34(1).Text = "�������˰����ߡ�������"
        Else
        describe = "ʼԴ��"
        Text34(0).Text = "\\""�����������ù���" + iname + "��\\"""
        Text34(1).Text = "������ԭ��֮����������"
        End If
        Text36(0).Text = "yellow"
        Text18.Text = "#"
        Text18.Text = Text18.Text + "FF"
        Text18.Text = Text18.Text + CStr(76 + Int(Rnd() * 24))
        Text18.Text = Text18.Text + "FF"
        If Len(yuansudesc) < 2 Then
        Text12.Text = "�Գ�һ��λ�桢�������޿��ܵ�" + yuansudes
        Else
        Text12.Text = yuansudesc + yuansunum + "��Ԫ��������Ϣ������һ��λ���"
        End If
    End If
    
    
ElseIf Mid(laiy, 1, 2) = "����" Then

    Key = Int(100 * Rnd()) '��������
    If Text33.Text = "��������" Then
        ling = "����"
    ElseIf Text33.Text = "��ļ�Ԫ" Or Text33.Text = "�ɹż�Ԫ" Or Text33.Text = "��������ʱ" Then ling = "����"
    ElseIf Text33.Text = "ʮ��ħ��" Then ling = "ħ��"
    ElseIf Key < 20 Then ling = "����"
    ElseIf Key < 35 Then ling = "�����Ѫ"
    ElseIf Key < 45 Then ling = "�������"
    ElseIf Key < 46 Then ling = "ʬ��"
    ElseIf Key < 52 Then ling = "����"
    ElseIf Key < 62 Then ling = "����"
    ElseIf Key < 68 Then ling = "Ԫ��"
    ElseIf Key < 75 Then ling = "Դ��"
    ElseIf Key < 90 Then ling = "����"
    Else: ling = "����"
    End If
    

    If yuansunum = "��" Then yuansudes = "ȫ����"
    
    Key = Int(100 * Rnd()) '��������
    If Text33.Text = "��������" And Key < 20 Then
        zhongzu = "����"
    ElseIf Text33.Text = "��������" And Key < 40 Then zhongzu = "����"
    ElseIf Text33.Text = "��������" And Key < 60 Then zhongzu = "����"
    ElseIf Text33.Text = "��������" And Key < 80 Then zhongzu = "����"
    ElseIf Text33.Text = "��������" And Key < 100 Then zhongzu = "ˮ��"
    ElseIf Text33.Text = "ʮ��ħ��" Then zhongzu = "ħ��"
    ElseIf Key < 80 Then zhongzu = "����"
    ElseIf Key < 85 Then zhongzu = "����"
    ElseIf Key < 87 Then zhongzu = "����"
    ElseIf Key < 89 Then zhongzu = "����"
    ElseIf Key < 91 Then zhongzu = "����"
    ElseIf Key < 93 Then zhongzu = "����"
    ElseIf Key < 95 Then zhongzu = "ħ��"
    ElseIf Key < 96 Then zhongzu = "����"
    ElseIf Key < 97 Then zhongzu = "ˮ��"
    ElseIf Key < 98 Then zhongzu = "����"
    Else: zhongzu = "����"
    End If
    
    Key = Int(80 * Rnd()) '�Ա�
    If Key < 40 Then
        xingbie = "��"
    Else: xingbie = "Ů"
    End If
    
    If xingbie = "Ů" Then
    chenghu = "Ů"
    ElseIf xingbie = "��" Then chenghu = "��"
    End If
    
    Key = Int(70 * Rnd()) '��
    If zhongzu = "����" And Key < 30 Then
        xingshi = "��"
    ElseIf zhongzu = "����" And Key < 50 Then xingshi = "��"
    ElseIf zhongzu = "����" And Key < 80 Then xingshi = "Ϳɽ"
    ElseIf zhongzu = "����" And Key < 60 Then xingshi = "��"
    ElseIf zhongzu = "����" And Key < 60 Then xingshi = "��"
    ElseIf zhongzu = "����" And Key < 60 Then xingshi = "��"
    ElseIf zhongzu = "����" And Key < 60 Then xingshi = "��"
    ElseIf zhongzu = "ħ��" And Key < 20 Then xingshi = "��"
    ElseIf zhongzu = "ħ��" And Key < 35 Then xingshi = "ħ"
    ElseIf zhongzu = "ħ��" And Key < 55 Then xingshi = "Ī"
    ElseIf zhongzu = "ħ��" And Key < 70 Then xingshi = "��"
    ElseIf Key < 3 Then xingshi = "��"
    ElseIf Key < 5 Then xingshi = "�Ϲ�"
    ElseIf Key < 7 Then xingshi = "��"
    ElseIf Key < 10 Then xingshi = "��"
    ElseIf Key < 11 Then xingshi = "����"
    ElseIf Key < 12 Then xingshi = "����"
    ElseIf Key < 13 Then xingshi = "��ľ"
    ElseIf Key < 14 Then xingshi = "ŷ��"
    ElseIf Key < 15 Then xingshi = "����"
    ElseIf Key < 16 Then xingshi = "����"
    ElseIf Key < 17 Then xingshi = "Ľ��"
    ElseIf Key < 18 Then xingshi = "��"
    ElseIf Key < 19 Then xingshi = "�ĺ�"
    ElseIf Key < 20 Then xingshi = "����"
    ElseIf Key < 22 Then xingshi = "Ҷ"
    ElseIf Key < 23 Then xingshi = "���"
    ElseIf Key < 25 Then xingshi = "��"
    ElseIf Key < 27 Then xingshi = "��"
    ElseIf Key < 29 Then xingshi = "��"
    ElseIf Key < 31 Then xingshi = "��"
    ElseIf Key < 33 Then xingshi = "��"
    ElseIf Key < 35 Then xingshi = "��"
    ElseIf Key < 37 Then xingshi = "��"
    ElseIf Key < 39 Then xingshi = "��"
    ElseIf Key < 40 Then xingshi = "��": ElseIf Key < 41 Then xingshi = "֣"
    ElseIf Key < 42 Then xingshi = "��": ElseIf Key < 43 Then xingshi = "��"
    ElseIf Key < 44 Then xingshi = "��": ElseIf Key < 45 Then xingshi = "��"
    ElseIf Key < 46 Then xingshi = "ʩ": ElseIf Key < 47 Then xingshi = "��"
    ElseIf Key < 48 Then xingshi = "��": ElseIf Key < 49 Then xingshi = "��"
    ElseIf Key < 50 Then xingshi = "��": ElseIf Key < 51 Then xingshi = "��"
    ElseIf Key < 52 Then xingshi = "��": ElseIf Key < 53 Then xingshi = "³"
    ElseIf Key < 54 Then xingshi = "��": ElseIf Key < 55 Then xingshi = "��"
    ElseIf Key < 56 Then xingshi = "��": ElseIf Key < 57 Then xingshi = "��"
    ElseIf Key < 58 Then xingshi = "��": ElseIf Key < 59 Then xingshi = "��"
    ElseIf Key < 60 Then xingshi = "��": ElseIf Key < 61 Then xingshi = "ë"
    ElseIf Key < 62 Then xingshi = "��": ElseIf Key < 63 Then xingshi = "��"
    ElseIf Key < 64 Then xingshi = "��": ElseIf Key < 65 Then xingshi = "��"
    ElseIf Key < 66 Then xingshi = "ͯ": ElseIf Key < 67 Then xingshi = "��"
    ElseIf Key < 68 Then xingshi = "��": ElseIf Key < 69 Then xingshi = "��"
    Else: xingshi = "��"
    End If

    Dim mingzi(0 To 2) As String
    mingz = ""
    For i = 0 To 1:
        Key = Int(70 * Rnd()) '����
        If xingbie = "Ů" Then
            If Key < 29 Then
            mingzi(i) = ""
            ElseIf Key < 30 Then mingzi(i) = "��": ElseIf Key < 31 Then mingzi(i) = "��"
            ElseIf Key < 32 Then mingzi(i) = "��": ElseIf Key < 33 Then mingzi(i) = "ٻ"
            ElseIf Key < 34 Then mingzi(i) = "�": ElseIf Key < 35 Then mingzi(i) = "��"
            ElseIf Key < 36 Then mingzi(i) = "��": ElseIf Key < 37 Then mingzi(i) = "ʫ"
            ElseIf Key < 38 Then mingzi(i) = "��": ElseIf Key < 39 Then mingzi(i) = "��"
            ElseIf Key < 40 Then mingzi(i) = "ܿ": ElseIf Key < 41 Then mingzi(i) = "��"
            ElseIf Key < 42 Then mingzi(i) = "��": ElseIf Key < 43 Then mingzi(i) = "��"
            ElseIf Key < 44 Then mingzi(i) = "��": ElseIf Key < 45 Then mingzi(i) = "ܰ"
            ElseIf Key < 46 Then mingzi(i) = "��": ElseIf Key < 47 Then mingzi(i) = "��"
            ElseIf Key < 48 Then mingzi(i) = "��": ElseIf Key < 49 Then mingzi(i) = "��"
            ElseIf Key < 50 Then mingzi(i) = "��": ElseIf Key < 51 Then mingzi(i) = "��"
            ElseIf Key < 52 Then mingzi(i) = "��": ElseIf Key < 53 Then mingzi(i) = "��"
            ElseIf Key < 54 Then mingzi(i) = "ޱ": ElseIf Key < 55 Then mingzi(i) = "��"
            ElseIf Key < 56 Then mingzi(i) = "��": ElseIf Key < 57 Then mingzi(i) = "��"
            ElseIf Key < 58 Then mingzi(i) = "ѩ": ElseIf Key < 59 Then mingzi(i) = "��"
            ElseIf Key < 60 Then mingzi(i) = "��": ElseIf Key < 61 Then mingzi(i) = "��"
            ElseIf Key < 62 Then mingzi(i) = "÷": ElseIf Key < 63 Then mingzi(i) = "��"
            ElseIf Key < 64 Then mingzi(i) = "ݯ": ElseIf Key < 65 Then mingzi(i) = "��"
            ElseIf Key < 66 Then mingzi(i) = "Ө": ElseIf Key < 67 Then mingzi(i) = "��"
            ElseIf Key < 68 Then mingzi(i) = "��": ElseIf Key < 69 Then mingzi(i) = "��"
            Else: mingzi(i) = "��"
            End If
        ElseIf xingbie = "��" Then
            If Key < 29 Then
            mingzi(i) = ""
            ElseIf Key < 30 Then mingzi(i) = "�": ElseIf Key < 31 Then mingzi(i) = "��"
            ElseIf Key < 32 Then mingzi(i) = "��": ElseIf Key < 33 Then mingzi(i) = "ɽ"
            ElseIf Key < 34 Then mingzi(i) = "��": ElseIf Key < 35 Then mingzi(i) = "��"
            ElseIf Key < 36 Then mingzi(i) = "��": ElseIf Key < 37 Then mingzi(i) = "��"
            ElseIf Key < 38 Then mingzi(i) = "��": ElseIf Key < 39 Then mingzi(i) = "��"
            ElseIf Key < 40 Then mingzi(i) = "��": ElseIf Key < 41 Then mingzi(i) = "֪"
            ElseIf Key < 42 Then mingzi(i) = "��": ElseIf Key < 43 Then mingzi(i) = "��"
            ElseIf Key < 44 Then mingzi(i) = "��": ElseIf Key < 45 Then mingzi(i) = "��"
            ElseIf Key < 46 Then mingzi(i) = "��": ElseIf Key < 47 Then mingzi(i) = "��"
            ElseIf Key < 48 Then mingzi(i) = "�": ElseIf Key < 49 Then mingzi(i) = "֮"
            ElseIf Key < 50 Then mingzi(i) = "�": ElseIf Key < 51 Then mingzi(i) = "��"
            ElseIf Key < 52 Then mingzi(i) = "ϧ": ElseIf Key < 53 Then mingzi(i) = "��"
            ElseIf Key < 54 Then mingzi(i) = "Ѱ": ElseIf Key < 55 Then mingzi(i) = "��"
            ElseIf Key < 56 Then mingzi(i) = "��": ElseIf Key < 57 Then mingzi(i) = "֪"
            ElseIf Key < 58 Then mingzi(i) = "��": ElseIf Key < 59 Then mingzi(i) = "��"
            ElseIf Key < 60 Then mingzi(i) = "��": ElseIf Key < 61 Then mingzi(i) = "��"
            ElseIf Key < 62 Then mingzi(i) = "ƽ": ElseIf Key < 63 Then mingzi(i) = "��"
            ElseIf Key < 64 Then mingzi(i) = "͢": ElseIf Key < 65 Then mingzi(i) = "̩"
            ElseIf Key < 66 Then mingzi(i) = "��": ElseIf Key < 67 Then mingzi(i) = "��"
            ElseIf Key < 68 Then mingzi(i) = "��": ElseIf Key < 69 Then mingzi(i) = "��"
            Else: mingzi(i) = "��"
            End If
        End If
    Next i
    
    mingz = mingzi(0) + mingzi(1)
    
    If mingz = "" Then
    Key = 29 + Int(41 * Rnd()) '����
        If xingbie = "Ů" Then
            If Key < 29 Then
            mingz = "���"
            ElseIf Key < 30 Then mingz = "����": ElseIf Key < 31 Then mingz = "��˫"
            ElseIf Key < 32 Then mingz = "����": ElseIf Key < 33 Then mingz = "��ѩ"
            ElseIf Key < 34 Then mingz = "��Ȼ": ElseIf Key < 35 Then mingz = "��ѩ"
            ElseIf Key < 36 Then mingz = "���": ElseIf Key < 37 Then mingz = "����"
            ElseIf Key < 38 Then mingz = "�ĳ�": ElseIf Key < 39 Then mingz = "���"
            ElseIf Key < 40 Then mingz = "ĭ��": ElseIf Key < 41 Then mingz = "ĭѩ"
            ElseIf Key < 42 Then mingz = "��Ө": ElseIf Key < 43 Then mingz = "����"
            ElseIf Key < 44 Then mingz = "���": ElseIf Key < 45 Then mingz = "���"
            ElseIf Key < 46 Then mingz = "����": ElseIf Key < 47 Then mingz = "��ܿ"
            ElseIf Key < 48 Then mingz = "ʱ��": ElseIf Key < 49 Then mingz = "����"
            ElseIf Key < 50 Then mingz = "δ��": ElseIf Key < 51 Then mingz = "����"
            ElseIf Key < 52 Then mingz = "��Ȼ": ElseIf Key < 53 Then mingz = "����"
            ElseIf Key < 54 Then mingz = "����": ElseIf Key < 55 Then mingz = "��ѩ"
            ElseIf Key < 56 Then mingz = "����": ElseIf Key < 57 Then mingz = "����"
            ElseIf Key < 58 Then mingz = "ѩ��": ElseIf Key < 59 Then mingz = "����"
            ElseIf Key < 60 Then mingz = "����": ElseIf Key < 61 Then mingz = "����"
            ElseIf Key < 62 Then mingz = "����": ElseIf Key < 63 Then mingz = "����"
            ElseIf Key < 64 Then mingz = "��Ө": ElseIf Key < 65 Then mingz = "����"
            ElseIf Key < 66 Then mingz = "ĺ��": ElseIf Key < 67 Then mingz = "ܿ��"
            ElseIf Key < 68 Then mingz = "����": ElseIf Key < 69 Then mingz = "����"
            Else: mingz = "����"
            End If
        ElseIf xingbie = "��" Then
            If Key < 29 Then
            mingz = "���"
            ElseIf Key < 30 Then mingz = "���": ElseIf Key < 31 Then mingz = "����"
            ElseIf Key < 32 Then mingz = "�䳾": ElseIf Key < 33 Then mingz = "��ѩ"
            ElseIf Key < 34 Then mingz = "��֮": ElseIf Key < 35 Then mingz = "����"
            ElseIf Key < 36 Then mingz = "�ֹ�": ElseIf Key < 37 Then mingz = "ϧ��"
            ElseIf Key < 38 Then mingz = "����": ElseIf Key < 39 Then mingz = "֪��"
            ElseIf Key < 40 Then mingz = "����": ElseIf Key < 41 Then mingz = "����"
            ElseIf Key < 42 Then mingz = "����": ElseIf Key < 43 Then mingz = "��ɽ"
            ElseIf Key < 44 Then mingz = "�ۺ�": ElseIf Key < 45 Then mingz = "ǧɽ"
            ElseIf Key < 46 Then mingz = "���": ElseIf Key < 47 Then mingz = "��"
            ElseIf Key < 48 Then mingz = "ʱ��": ElseIf Key < 49 Then mingz = "����"
            ElseIf Key < 50 Then mingz = "δ��": ElseIf Key < 51 Then mingz = "ҹ��"
            ElseIf Key < 52 Then mingz = "����": ElseIf Key < 53 Then mingz = "����"
            ElseIf Key < 54 Then mingz = "Զ��": ElseIf Key < 55 Then mingz = "����"
            ElseIf Key < 56 Then mingz = "����": ElseIf Key < 57 Then mingz = "�Ƹ�"
            ElseIf Key < 58 Then mingz = "����": ElseIf Key < 59 Then mingz = "��ҹ"
            ElseIf Key < 60 Then mingz = "ī��": ElseIf Key < 61 Then mingz = "����"
            ElseIf Key < 62 Then mingz = "��֮": ElseIf Key < 63 Then mingz = "��ʱ"
            ElseIf Key < 64 Then mingz = "Զ��": ElseIf Key < 65 Then mingz = "����"
            ElseIf Key < 66 Then mingz = "����": ElseIf Key < 67 Then mingz = "�ƻ�"
            ElseIf Key < 68 Then mingz = "ƽ��": ElseIf Key < 69 Then mingz = "����"
            Else: mingz = "����"
            End If
        End If
    End If


    Key = 29 + Int(41 * Rnd())  '���
    If xingbie = "Ů" Then
        If Key < 29 Then
        zunhao = "���"
        ElseIf Key < 30 Then zunhao = "�Ǻ�": ElseIf Key < 31 Then zunhao = "��˫"
        ElseIf Key < 32 Then zunhao = "����": ElseIf Key < 33 Then zunhao = "��ѩ"
        ElseIf Key < 34 Then zunhao = "����": ElseIf Key < 35 Then zunhao = "����"
        ElseIf Key < 36 Then zunhao = "����": ElseIf Key < 37 Then zunhao = "����"
        ElseIf Key < 38 Then zunhao = "�κ�": ElseIf Key < 39 Then zunhao = "�ѩ"
        ElseIf Key < 40 Then zunhao = "�㺮": ElseIf Key < 41 Then zunhao = "�̷�"
        ElseIf Key < 42 Then zunhao = "��Ө": ElseIf Key < 43 Then zunhao = "�Ǻ�"
        ElseIf Key < 44 Then zunhao = "���": ElseIf Key < 45 Then zunhao = "����"
        ElseIf Key < 46 Then zunhao = "����": ElseIf Key < 47 Then zunhao = "��ܿ"
        ElseIf Key < 48 Then zunhao = "��ѩ": ElseIf Key < 49 Then zunhao = "����"
        ElseIf Key < 50 Then zunhao = "����": ElseIf Key < 51 Then zunhao = "��ϼ"
        ElseIf Key < 52 Then zunhao = "����": ElseIf Key < 53 Then zunhao = "����"
        ElseIf Key < 54 Then zunhao = "���": ElseIf Key < 55 Then zunhao = "��ѩ"
        ElseIf Key < 56 Then zunhao = "��ޱ": ElseIf Key < 57 Then zunhao = "����"
        ElseIf Key < 58 Then zunhao = "�Ǻ�": ElseIf Key < 59 Then zunhao = "����"
        ElseIf Key < 60 Then zunhao = "����": ElseIf Key < 61 Then zunhao = "����"
        ElseIf Key < 62 Then zunhao = "ɽ��": ElseIf Key < 63 Then zunhao = "����"
        ElseIf Key < 64 Then zunhao = "��Ө": ElseIf Key < 65 Then zunhao = "����"
        ElseIf Key < 66 Then zunhao = "ĺ��": ElseIf Key < 67 Then zunhao = "�ƻ�"
        ElseIf Key < 68 Then zunhao = "����": ElseIf Key < 69 Then zunhao = "����"
        Else: zunhao = "����"
        End If
    ElseIf xingbie = "��" Then
        If Key < 29 Then
        zunhao = "���"
        ElseIf Key < 30 Then zunhao = "�˰�": ElseIf Key < 31 Then zunhao = "����"
        ElseIf Key < 32 Then zunhao = "�׺�": ElseIf Key < 33 Then zunhao = "����"
        ElseIf Key < 34 Then zunhao = "��Ϊ": ElseIf Key < 35 Then zunhao = "Ǭ��"
        ElseIf Key < 36 Then zunhao = "����": ElseIf Key < 37 Then zunhao = "�ƻ�"
        ElseIf Key < 38 Then zunhao = "����": ElseIf Key < 39 Then zunhao = "����"
        ElseIf Key < 40 Then zunhao = "��ħ": ElseIf Key < 41 Then zunhao = "ȼ��"
        ElseIf Key < 42 Then zunhao = "��ת": ElseIf Key < 43 Then zunhao = "�춷"
        ElseIf Key < 44 Then zunhao = "��ڤ": ElseIf Key < 45 Then zunhao = "�̺�"
        ElseIf Key < 46 Then zunhao = "����": ElseIf Key < 47 Then zunhao = "�һ�"
        ElseIf Key < 48 Then zunhao = "��Ԫ": ElseIf Key < 49 Then zunhao = "�޼�"
        ElseIf Key < 50 Then zunhao = "����": ElseIf Key < 51 Then zunhao = "����"
        ElseIf Key < 52 Then zunhao = "����": ElseIf Key < 53 Then zunhao = "���"
        ElseIf Key < 54 Then zunhao = "����": ElseIf Key < 55 Then zunhao = "����"
        ElseIf Key < 56 Then zunhao = "����": ElseIf Key < 57 Then zunhao = "�ֻ�"
        ElseIf Key < 58 Then zunhao = "��": ElseIf Key < 59 Then zunhao = "����"
        ElseIf Key < 60 Then zunhao = "̫��": ElseIf Key < 61 Then zunhao = "��Ԫ"
        ElseIf Key < 62 Then zunhao = "����": ElseIf Key < 63 Then zunhao = "��Ȫ"
        ElseIf Key < 64 Then zunhao = "����": ElseIf Key < 65 Then zunhao = "���"
        ElseIf Key < 66 Then zunhao = "��ս": ElseIf Key < 67 Then zunhao = "���"
        ElseIf Key < 68 Then zunhao = "����": ElseIf Key < 69 Then zunhao = "����"
        Else: zunhao = "��ڤ"
        End If
    End If
    
    Key = Int(120 * Rnd()) '�������
    If zhongzu = "����" And Key < 70 Then
    zunhao = "����"
    ElseIf zhongzu = "����" And Key < 40 Then zunhao = "����"
    ElseIf zhongzu = "����" And Key < 60 Then zunhao = "����"
    ElseIf zhongzu = "����" And Key < 100 Then zunhao = "����"
    ElseIf zhongzu = "����" And Key < 80 Then zunhao = "����"
    ElseIf zhongzu = "ˮ��" And Key < 80 Then zunhao = "����"
    ElseIf zhongzu = "����" And Key < 40 Then zunhao = "�׻�"
    ElseIf zhongzu = "����" And Key < 80 Then zunhao = "���"
    End If

    Do While Key > (pinji + 1) * 15 Or Key < (pinji + 1) * 12 - 12
        Key = Int(120 * Rnd()) '�Ƚ�����
    Loop
    
    If ling <> "����" And ling <> "����" And ling <> "ħ��" And ling <> "ʬ��" Then
        If Key < 10 Then
            dengjie = Mid(ling, 1, 1) + "��"
        ElseIf Key < 20 Then dengjie = Mid(ling, 1, 1) + "ʦ"
        ElseIf Key < 30 Then dengjie = Mid(ling, 1, 1) + "��"
        ElseIf Key < 40 Then dengjie = Mid(ling, 1, 1) + "��"
        ElseIf Key < 50 Then dengjie = Mid(ling, 1, 1) + "��"
        ElseIf Key < 60 Then dengjie = Mid(ling, 1, 1) + "��"
        ElseIf Key < 70 Then dengjie = Mid(ling, 1, 1) + "��"
        ElseIf Key < 80 Then dengjie = Mid(ling, 1, 1) + "ʥ"
        ElseIf Key < 90 Then dengjie = Mid(ling, 1, 1) + "��"
        ElseIf Key < 100 Then dengjie = Mid(ling, 1, 1) + "��"
        ElseIf Key < 110 Then dengjie = "����"
        Else: dengjie = "���"
        End If
        
        key1 = Int(20 * Rnd())
        dengjiehou2 = Mid(dengjie, Len(dengjie) - 1, 2)
        
        If Mid(dengjiehou2, 2, 1) = "��" And key1 < 6 And xingbie = "��" Then
        dengjiehou2 = "�ʵ�"
        If Mid(dengjiehou2, 2, 1) = "��" And key1 < 6 And xingbie = "Ů" Then dengjiehou2 = "Ů��"
        
        ElseIf Mid(dengjiehou2, 2, 1) = "��" And key1 < 6 Then dengjiehou2 = "����"
        ElseIf Mid(dengjiehou2, 2, 1) = "ʥ" And key1 < 6 And xingbie = "��" Then dengjiehou2 = "ʥ��"
        ElseIf Mid(dengjiehou2, 2, 1) = "ʥ" And key1 < 6 And xingbie = "Ů" Then dengjiehou2 = "ʥŮ"
        
        ElseIf Mid(dengjiehou2, 2, 1) = "��" And key1 < 5 Then dengjiehou2 = "���"
        ElseIf Mid(dengjiehou2, 2, 1) = "��" And key1 < 10 And xingbie = "Ů" Then dengjiehou2 = "Ů��"
        ElseIf Mid(dengjiehou2, 2, 1) = "��" And key1 < 12 Then dengjiehou2 = "��" + chenghu
        End If
        
        If Key > 50 Then
            If Key Mod 10 = 0 Then
                dengjie = "�벽" + dengjie + "����" + zunhao + dengjiehou2
            Else: dengjie = dengjie + CStr(Key Mod 10) + "��" + "����" + zunhao + dengjiehou2
            End If
        Else
            If Key Mod 10 = 0 Then
                dengjie = "�벽" + dengjie
            Else: dengjie = dengjie + CStr(Key Mod 10) + "��"
            End If
        End If
        
        
    ElseIf ling = "����" Then
        If Key < 10 Then
            dengjie = "��������"
        ElseIf Key < 20 Then dengjie = "��������"
        ElseIf Key < 30 Then dengjie = "������"
        ElseIf Key < 40 Then dengjie = "����ϵ�"
        ElseIf Key < 50 Then dengjie = "½����" + Mid(ling, 1, 1)
        ElseIf Key < 60 Then dengjie = "��" + Mid(ling, 1, 1)
        ElseIf Key < 70 Then dengjie = "��" + Mid(ling, 1, 1)
        ElseIf Key < 80 Then dengjie = "��" + Mid(ling, 1, 1)
        ElseIf Key < 90 Then dengjie = "̫�ҽ�" + Mid(ling, 1, 1)
        ElseIf Key < 100 Then dengjie = "���޽�" + Mid(ling, 1, 1)
        ElseIf Key < 110 Then dengjie = "׼ʥ"
        Else: dengjie = "ʥ��"
        End If
        
        key1 = Int(20 * Rnd())
        dengjiehou2 = Mid(dengjie, Len(dengjie) - 1, 2)
        If key1 < 8 Then
        dengjiehou2 = "��" + chenghu
        ElseIf key1 < 14 And xingbie = "Ů" Then dengjiehou2 = "����"
        ElseIf key1 < 14 And xingbie = "��" Then dengjiehou2 = "ɢ��"
        End If
        
        If Key > 80 Then
            If Key Mod 10 = 0 Then
                dengjie = "�벽" + dengjie + "����" + zunhao + dengjiehou2
            Else: dengjie = dengjie + CStr(Key Mod 10) + "��" + "����" + zunhao + dengjiehou2
            End If
        Else
            If Key Mod 10 = 0 Then
                dengjie = "�벽" + dengjie
            Else: dengjie = dengjie + CStr(Key Mod 10) + "��"
            End If
        End If
        
        
    ElseIf ling = "����" Or ling = "ħ��" Or ling = "ʬ��" Or ling = "����" Then
        If Key < 10 Then
            dengjie = "��" + Mid(ling, 1, 1)
        ElseIf Key < 20 Then dengjie = Mid(ling, 1, 1) + "��"
        ElseIf Key < 30 Then dengjie = Mid(ling, 1, 1) + "��"
        ElseIf Key < 40 Then dengjie = Mid(ling, 1, 1) + "��"
        ElseIf Key < 50 Then dengjie = Mid(ling, 1, 1) + "��"
        ElseIf Key < 60 Then dengjie = Mid(ling, 1, 1) + "��"
        ElseIf Key < 70 Then dengjie = Mid(ling, 1, 1) + "��"
        ElseIf Key < 80 Then dengjie = Mid(ling, 1, 1) + "ʥ"
        ElseIf Key < 90 Then dengjie = Mid(ling, 1, 1) + "��"
        ElseIf Key < 100 Then dengjie = Mid(ling, 1, 1) + "��"
        ElseIf Key < 110 Then dengjie = "����"
        Else: dengjie = "���"
        End If
        
        
        If Key > 60 Then
            If Key Mod 10 = 0 Then
                dengjie = "�벽" + dengjie + "����" + zunhao + dengjie
            Else: dengjie = dengjie + CStr(Key Mod 10) + "��" + "����" + zunhao + dengjie
            End If
        Else
            If Key Mod 10 = 0 Then
                dengjie = "�벽" + dengjie
            Else: dengjie = dengjie + CStr(Key Mod 10) + "��"
            End If
        End If
        
    End If







    Key = Int(200 * Rnd())  '�����

    If Text29.Text = "��ͨ" Then '_____________---------����Ʒ��---------------------------------------------------
        
        If Key < 50 Then
        describe = "���ص�"
        Text34(0).Text = "\\""�����������ƽ������" + aliangci + "" + iname + "�����ǻ�δ�ɳ����ˡ�\\"""
        Text34(1).Text = "\\""�������ھ��죬��������Ƽ֮ĩ��\\"""
        Text34(2).Text = "������" + dengjie + "��" + xingshi + mingz
        ElseIf Key < 100 Then
        describe = "ƽ����"
        Text34(0).Text = "\\""��" + aliangci + iname + "���Ǵ�����õ��ģ�\\"""
        If yuansudes <> "" Then
        Text34(1).Text = "\\""���ƺ�������" + yuansudes + ling + "��\\"""
        Else
        Text34(1).Text = "\\""û���κ��ر�֮����\\"""
        End If
        Text34(2).Text = "������������ʦ��" + xingshi + mingz
        ElseIf Key < 150 Then
        describe = "һ���"
        Text34(0).Text = "\\""��" + aliangci + iname + "�����ʼ�Ļ���\\"""
        Text34(1).Text = "������" + dengjie + "��" + xingshi + mingz
        Text34(2).Text = ""
        Else
        describe = "����Ȼ��"
        Text34(0).Text = "\\""��������һ" + aliangci + "�ܳ�����" + iname + "��\\"""
        Text34(1).Text = "������" + dengjie + "��" + xingshi + mingz
        Text34(2).Text = ""
        End If
        Text36(0).Text = "white"
        
        Text18.Text = "#"
        For i = 1 To 3
        Text18.Text = Text18.Text + CStr(86 + Int(Rnd() * 14))
        Next i
        If yuansudes = "" Then
        Text12.Text = "�洦�ɼ���"
        Else
        Text12.Text = "�����Ѻ۵�" + yuansudes
        describe = "���Ƶ�"
        End If
    
    ElseIf Text29.Text = "����" Then
        If Key < 50 Then
        describe = "������"
        Text34(0).Text = "\\""��" + aliangci + iname + "�Һ�ϲ����\\"""
        Text34(1).Text = "������" + dengjie + "��" + xingshi + mingz
        ElseIf Key < 100 Then
        describe = "���Ƶ�"
        Text34(0).Text = "\\""�����������ŵ���ʽװ����\\"""
        Text34(1).Text = "������" + dengjie + "��" + xingshi + mingz
        ElseIf Key < 150 Then
        describe = "��ĥ����"
        Text34(0).Text = "\\""��" + aliangci + iname + "��������\\"""
        Text34(1).Text = "������" + dengjie + "��" + xingshi + mingz
        Else
        describe = "�����"
        Text34(0).Text = "\\""�����" + iname + "��\\"""
        Text34(1).Text = "������" + dengjie + "��" + xingshi + mingz
        End If
        Text36(0).Text = "dark_green"
        Text18.Text = "#"
        Text18.Text = Text18.Text + CStr(56 + Int(Rnd() * 14))
        Text18.Text = Text18.Text + "EE"
        Text18.Text = Text18.Text + CStr(56 + Int(Rnd() * 14))
        If yuansudes = "" Then
        Text12.Text = "�ǳ������"
        Else
        Text12.Text = "��������" + yuansudes
        End If
        
        ElseIf Text29.Text = "ϡ��" Then
        If Key < 50 Then
        describe = "������"
        Text34(0).Text = "\\""��" + aliangci + iname + "����������һ��\\"""
        Text34(1).Text = "������" + dengjie + "��" + xingshi + mingz
        Text34(2).Text = ""
        ElseIf Key < 100 Then
        describe = "�Ҵ���"
        Text34(0).Text = "\\""�����Ҽ���������������" + cname + iname + "��\\"""
        If yuansudes <> "" Then
        Text34(1).Text = "\\""���ĸ����㣡�����̺���" + yuansudes + "��������\\"""
        Else
        Text34(1).Text = "\\""�ƺ��Ǵ�ĳ�������ż��ϼ�����ģ�\\"""
        End If
        Text34(2).Text = "������" + dengjie + "��" + xingshi + mingz
        ElseIf Key < 150 Then
        describe = "���ѵ�"
        Text34(0).Text = "\\""������" + aliangci + iname + "����̫���ˣ�\\"""
        Text34(1).Text = "\\""�뵽��" + cname + "ȥ��������һ������ţ�\\"""
        Text34(2).Text = "������" + dengjie + "��" + xingshi + mingz
        Else
        describe = "ǿ����"
        Text34(0).Text = "\\""����������" + iname + "����������\\"""
        Text34(1).Text = "������" + dengjie + "��" + xingshi + mingz
        Text34(2).Text = ""
        End If
        Text36(0).Text = "dark_blue"
        Text18.Text = "#"
        Text18.Text = Text18.Text + CStr(7 + Int(Rnd() * 14))
        Text18.Text = Text18.Text + CStr(56 + Int(Rnd() * 14))
        Text18.Text = Text18.Text + "CC"
        If yuansudes = "" Then
        Text12.Text = "Ʒ�ʼ��ѵ�"
        Else
        Text12.Text = "����΢��" + yuansudes + ling + "������"
        End If
    
    
    ElseIf Text29.Text = "ʷʫ" Then
        If Key < 50 Then
        describe = "������"
        Text34(0).Text = "\\""���ﲻ��������������" + aliangci + cname + iname + "�������⣡\\"""
        Text34(1).Text = "������" + zunhao + "��" + chenghu + "��" + xingshi + mingz
        Text34(2).Text = ""
        ElseIf Key < 80 Then
        describe = "���Ƶ�"
        Text34(0).Text = "\\""����" + cname + "�ر�����" + yuansudes + iname + "���ܰ��ɣ�\\"""
        Text34(1).Text = "\\""ʲô�������ã����������㲻���ã�\\"""
        Text34(2).Text = "������" + dengjie + "��" + xingshi + mingz
        ElseIf Key < 110 Then
        describe = "���ϵ�"
        Text34(0).Text = "\\""��" + aliangci + iname + "�ĵ����Ѿ����ɿ�����\\"""
        Text34(1).Text = "������" + zunhao + "��" + chenghu + "��" + xingshi + mingz
        Text34(2).Text = ""
        ElseIf Key < 150 Then
        describe = "��ʷ�ƾõ�"
        Text34(0).Text = "\\""����һ�ι���һ" + aliangci + cname + iname + "���Ƴ���ʷ��\\"""
        Text34(1).Text = "������" + dengjie + "��" + xingshi + mingz
        Text34(2).Text = ""
        Else
        describe = "���ص�"
        Text34(0).Text = "\\""����һ" + aliangci + "���ص�" + iname + "����\\"""
        Text34(1).Text = "������" + zunhao + "��" + chenghu + "��" + xingshi + mingz
        Text34(2).Text = ""
        End If
        Text36(0).Text = "light_purple"
        Text18.Text = "#"
        Text18.Text = Text18.Text + "FF"
        Text18.Text = Text18.Text + CStr(56 + Int(Rnd() * 44))
        Text18.Text = Text18.Text + "FF"
        If yuansudes = "" Then
        Text12.Text = "�����������"
        Else
        Text12.Text = "�̺�ǿ��" + yuansudes + ling + "��"
        End If
    
    
    ElseIf Text29.Text = "��˵" Then
        If Key < 50 Then
        describe = "�����"
        Text34(0).Text = "\\""����һ�ι���" + iname + "�Ĵ�˵��\\"""
        Text34(1).Text = "������" + dengjie + "��" + xingshi + mingz
        ElseIf Key < 100 Then
        describe = "�̺��ֲ�" + yuansudes + ling + "��"
        Text34(0).Text = "\\""ʤ�����������" + iname + "��\\"""
        Text34(1).Text = "������" + dengjie + "��" + xingshi + mingz
        ElseIf Key < 150 Then
        describe = "��ʥ��"
        Text34(0).Text = "\\""ʥ�⽫��ף�����" + iname + "��Ӧ��Ҳ�ᡭ��ף���㣿\\"""
        Text34(1).Text = "������" + dengjie + "��" + xingshi + mingz
        Else
        describe = "�����ޡ�" + describe
        For i = 0 To 7
        Text9(i).Text = CStr(Val(Text9(i).Text) + Int((Val(Text9(i).Text) + 2) * poxian * Rnd()))
        Next i
        Text34(0).Text = "\\""ͻ�Ƽ��ްɣ�" + cname + iname + "��\\"""
        Text34(1).Text = "������" + dengjie + "��" + xingshi + mingz
        End If
        Text36(0).Text = "yellow"
        Text18.Text = "#"
        Text18.Text = Text18.Text + "FF"
        Text18.Text = Text18.Text + "DD"
        Text18.Text = Text18.Text + CStr(46 + Int(Rnd() * 34))
        If yuansudes = "" Then
        Text12.Text = "��������������"
        Else
        Text12.Text = "�̺���һλ����" + yuansudes + ling + "��"
        End If
    
    ElseIf Text29.Text = "��" Then
        If Key < 50 Then
        describe = "������"
        Text34(0).Text = "\\""�����ǣ�����" + aliangci + iname + "ȥ�ػ����Űɣ�\\"""
        Text34(1).Text = "������" + dengjie + "��" + xingshi + mingz
        ElseIf Key < 100 Then
        describe = "����������"
        Text34(0).Text = "\\""�ҵ�" + cname + iname + "�أ�ȥ���ˣ���\\"""
        Text34(1).Text = "������" + dengjie + "��" + xingshi + mingz
        ElseIf Key < 150 Then
        describe = "��͵�"
        Text34(0).Text = "\\""���ǣ���" + aliangci + iname + "�������ͽ������ˣ�\\"""
        Text34(1).Text = "������" + dengjie + "��" + xingshi + mingz
        Else
        describe = "��ӡ�ź���" + ling + "��"
        Text34(0).Text = "\\""" + ling + "������" + iname + "�б�����\\"""
        Text34(1).Text = "������" + dengjie + "��" + xingshi + mingz
        End If
        Text36(0).Text = "yellow"
        Text18.Text = "#"
        Text18.Text = Text18.Text + "FF"
        Text18.Text = Text18.Text + CStr(56 + Int(Rnd() * 24))
        Text18.Text = Text18.Text + CStr(0 + Int(Rnd() * 24))
        If yuansudes = "" Then
        Text12.Text = "�ƺ���ĳ����˵�г�������"
        Else
        Text12.Text = "�̺��ֲ�" + yuansudes + ling + "��������"
        End If
    
    ElseIf Text29.Text = "����" Then
        If Key < 50 Then
        describe = "�����"
        Text34(0).Text = "\\""�k�����ڣ�"
        If yuansudesc <> "" Then Text34(0).Text = Text34(0).Text + "�k����" + yuansudesc + yuansunum + "Ԫ�ش������㣡"
        Text34(0).Text = Text34(0).Text + "\\"""
        Text34(1).Text = "������" + dengjie + "��" + xingshi + mingz
        ElseIf Key < 100 Then
        describe = "ب������ǰ��"
        Text34(0).Text = "\\""�k���������£������˹�������\\"""
        Text34(1).Text = "������" + dengjie + "��" + xingshi + mingz
        ElseIf Key < 150 Then
        describe = "��������ǰ��"
        Text34(0).Text = "\\""����һ��������Ԫ֮ǰ��һ" + aliangci + iname + "�����Ρ���\\"""
        Text34(1).Text = "������" + dengjie + "��" + xingshi + mingz
        ElseIf Key < 180 Then
        describe = "��Ψһ��" + describe
        For i = 0 To 7
        Text9(i).Text = CStr(Val(Text9(i).Text) + Int((Val(Text9(i).Text) + 2) * weiyi * Rnd()))
        Next i
        Text34(0).Text = "\\""ӵ�С�Ψһ����" + cname + iname + "����ӵ���˳�����ʱ��Ŀ��ܣ�\\"""
        Text34(1).Text = "������" + dengjie + "��" + xingshi + mingz
        Else
        describe = "������"
        Text34(0).Text = "\\""���ﶼ�ḯ�ࡣ�����" + aliangci + cname + iname + "��һ�����⣡\\"""
        Text34(1).Text = "������" + dengjie + "��" + xingshi + mingz
        End If
        Text36(0).Text = "gold"
        Text18.Text = "#"
        Text18.Text = Text18.Text + "FF"
        Text18.Text = Text18.Text + CStr(77 + Int(Rnd() * 23))
        Text18.Text = Text18.Text + CStr(10 + Int(Rnd() * 24))
        If yuansudes = "" Then
        Text12.Text = "���㲻�𡢳��ѹ�����"
        Else
        Text12.Text = "��ӡ��" + yuansudes + "�ɵ�Ȩ����"
        End If
    
    
    ElseIf Text29.Text = "����" Then
        If Key < 50 Then
        describe = "ԭ����"
        Text34(0).Text = "\\""����ʥ�˵�" + iname + "��������ò��ˡ�\\"""
        Text34(1).Text = "������" + dengjie + "��" + xingshi + mingz
        ElseIf Key < 100 Then
        describe = "������"
        Text34(0).Text = "\\""����һ" + aliangci + "��֤�����絮����" + iname + "��\\"""
        Text34(1).Text = "������" + dengjie + "��" + xingshi + mingz
        ElseIf Key < 150 Then
        describe = "�����ѡ�" + describe
        For i = 0 To 7
        Text9(i).Text = CStr(Val(Text9(i).Text) + Int((Val(Text9(i).Text) + 2) * chaotuo * Rnd()))
        Next i
        Text34(0).Text = "\\""ӵ�С����ѡ���" + iname + "�����п��ִܵ�˰���\\"""
        Text34(1).Text = "������" + dengjie + "��" + xingshi + mingz
        Else
        describe = "ʼԴ��"
        Text34(0).Text = "\\""����ʥ���ù���" + iname + "��\\"""
        Text34(1).Text = "������" + dengjie + "��" + xingshi + mingz
        End If
        Text36(0).Text = "yellow"
        Text18.Text = "#"
        Text18.Text = Text18.Text + "FF"
        Text18.Text = Text18.Text + CStr(76 + Int(Rnd() * 24))
        Text18.Text = Text18.Text + "FF"
        If Len(yuansudesc) < 2 Then
        Text12.Text = "�Գ�һ�����졢�������޿��ܵ�" + yuansudes
        Else
        Text12.Text = yuansudesc + yuansunum + "������������Ϣ������һ�������"
        End If
    End If

End If '_____________---------------------------------------Ʒ��---------------------



Text36(1).Text = Text36(0).Text
Text36(2).Text = Text36(1).Text
descri = ""
describ = ""
If (Val(Text9(0).Text) > 100 And Val(Text17.Text) = 0) Or (Val(Text9(0).Text) > 5 And Val(Text17.Text) > 0) Then describ = describ + "ľ���ơ�"
If (Val(Text9(1).Text) > 0.8 And Val(Text17.Text) = 0) Or (Val(Text9(1).Text) > 2 And Val(Text17.Text) > 0) Then describ = describ + "��ɽ����"
If (Val(Text9(2).Text) > 1.2 And Val(Text17.Text) = 0) Or (Val(Text9(2).Text) > 3 And Val(Text17.Text) > 0) Then describ = describ + "����ӡ��"
If (Val(Text9(3).Text) > 70 And Val(Text17.Text) = 0) Or (Val(Text9(3).Text) > 4 And Val(Text17.Text) > 0) Then describ = describ + "�����ơ�"
If (Val(Text9(4).Text) > 18 And Val(Text17.Text) = 0) Or (Val(Text9(4).Text) > 2 And Val(Text17.Text) > 0) Then describ = describ + "����ӡ��"
If (Val(Text9(5).Text) > 18 And Val(Text17.Text) = 0) Or (Val(Text9(5).Text) > 2 And Val(Text17.Text) > 0) Then describ = describ + "�����顢"
If (Val(Text9(6).Text) > 10 And Val(Text17.Text) = 0) Or (Val(Text9(6).Text) > 3 And Val(Text17.Text) > 0) Then describ = describ + "���ӡ��"
If (Val(Text9(7).Text) > 40 And Val(Text17.Text) = 0) Or (Val(Text9(7).Text) > 3 And Val(Text17.Text) > 0) Then describ = describ + "���ۡ�"


If describ <> "" Then describ = "������" + describ
If describ <> "" Then describ = Mid(describ, 1, Len(describ) - 1)
If describ <> "" Then describ = describ + "�ȷ��ģ�"

If (Val(Text9(6).Text) < 0) Or (Val(Text9(2).Text) < 0) Then
If Len(iname) > 1 Then
describ = "���صġ�" + describ
Else
descri = descri + "��"
iname = descri + iname
End If
End If


Text10.Text = describe + cname + iname '����
Text35(3).Text = describ    '����ǰ׺
Text35(0).Text = cname + iname '��������


Text10.ForeColor = RGB(sjz(Mid(Text18.Text, 2, 2)), sjz(Mid(Text18.Text, 4, 2)), sjz(Mid(Text18.Text, 6, 2)))


End Sub

Private Sub Command37_Click(Index As Integer)
Text36(Index Mod mo).Text = Command37(Index).Caption
End Sub



Private Sub Command38_Click(Index As Integer)
Text37.Text = Command38(Index).Caption

End Sub

Private Sub Command39_Click(Index As Integer)
Dim total, health, reknock, speed, damage, tough, armor, atspeed, luck, knock As Single
Key = Int(200 * Rnd()) 'ǿ��ָ��������
If Index = 0 Then Text21.Text = """mainhand"""
If Index = 1 Then
    Text21.Text = """offhand"""
    If Text2.Text = "ͷ��" Then Text21.Text = """head"""
    If Text2.Text = "�ؼ�" Then Text21.Text = """chest"""
    If Text2.Text = "����" Then Text21.Text = """legs"""
    If Text2.Text = "ѥ��" Then Text21.Text = """feet"""
End If

If Text29.Text = "��ͨ" Then '_____________---------Ʒ��---------------------------------------------------
    Text17.Text = "0"
    total = Int((5 + Key \ 40)) '5-10��
    extra = Int(total * 0.4 * Rnd())

    If Index = 0 Then
    damage = Int(total * Rnd() + 2)
    speed = Int((total - damage) * Rnd())
    atspeed = Int((total - damage) * Rnd())
    health = Int((total - damage - speed - atspeed - extra))
    ElseIf Index = 1 Then
    health = Int(total * Rnd())
    armor = Int((total - health) * 0.8 * Rnd() + 1)
    tough = Int((total - health) * 0.4 * Rnd() + 1)
    speed = Int((total - health - armor - tough - extra))
    End If


ElseIf Text29.Text = "����" Then
    Text17.Text = "0"
    total = Int((10 + Key \ 20)) '10-20��
    extra = Int(total * 0.3 * Rnd())
    
    If Index = 0 Then
    damage = Int(total * Rnd() + 2)
    speed = Int((total - damage) * Rnd())
    atspeed = Int((total - damage) * Rnd())
    health = Int((total - damage - speed - atspeed - extra))
    ElseIf Index = 1 Then
    health = Int(total * Rnd() + 3)
    armor = Int((total - health) * Rnd() + 1)
    tough = Int((total - health - armor) * Rnd() + 1)
    speed = Int((total - health - armor - tough - extra))
    End If

ElseIf Text29.Text = "ϡ��" Then
    Text17.Text = "0"
    total = Int((20 + Key \ 10)) '20-40��
    extra = Int(total * 0.2 * Rnd())
    
    If Index = 0 Then
    damage = Int(total * Rnd() + 2)
    speed = Int((total - damage) * 0.5 * Rnd())
    atspeed = Int((total - damage) * Rnd())
    health = Int((total - damage - speed - atspeed - extra))
    ElseIf Index = 1 Then
    health = Int(total * Rnd() + 3)
    armor = Int((total - health) * 0.6 * Rnd() + 1)
    tough = Int((total - health) * 0.4 * Rnd() + 1)
    speed = Int((total - health - armor - tough - extra))
    End If

ElseIf Text29.Text = "ʷʫ" Then
    Text17.Text = "0"
    total = Int((40 + Key \ 5)) '40-80��
    extra = Int(total * 0.1 * Rnd())
    
    If Index = 0 Then
    damage = Int(total * 0.7 * Rnd() + 2)
    atspeed = Int((total - damage) * 0.7 * Rnd())
    armor = Int((total - damage - atspeed) * 0.5 * Rnd())
    luck = Int((total - damage - atspeed - armor) * 0.5 * Rnd())
    speed = Int((total - damage - atspeed - luck - armor - extra) * Rnd() * 0.7 * 10) \ 10
    extra = Int((total - damage - atspeed - luck - armor - speed) * 10) \ 10
    ElseIf Index = 1 Then
    health = Int(total * 0.6 * Rnd() + 5)
    armor = Int((total - health) * 0.8 * Rnd() + 1)
    tough = Int((total - health) * 0.4 * Rnd() + 1)
    damage = Int((total - health) * 0.1 * Rnd() + 1)
    luck = Int((total - health) * 0.1 * Rnd() + 1)
    speed = Int((total - health - armor - tough - damage - luck - extra))
    End If

ElseIf Text29.Text = "��˵" Then
Text17.Text = "1"
total = Int((20 + Key / 20)) '20-30��
extra = Int(total * 0.1 * Rnd())

If Index = 0 Then
damage = Int(total * 0.6 * Rnd() * 10) / 10
atspeed = Int((total - damage) * 0.7 * Rnd() * 10) / 10
armor = Int((total - damage - atspeed) * 0.4 * Rnd() * 10) / 10
luck = Int((total - damage - atspeed - armor + 3) * 0.9 * Rnd() * 10) / 10
speed = Int((total - damage - atspeed - luck - armor - extra) * Rnd() * 10) / 10
extra = Int((total - damage - atspeed - luck - armor - speed) * 10) / 10
ElseIf Index = 1 Then
health = Int(total * 0.6 * Rnd() * 10) / 10
armor = Int((total - health) * 0.4 * Rnd() * 10) / 10
tough = Int((total - health) * 0.4 * Rnd() * 10) / 10
damage = Int((total - health) * 0.1 * Rnd() * 10) / 10
luck = Int((total - health) * 0.1 * Rnd() * 10) / 10
speed = Int((total - health - armor - tough - damage - luck - extra) * 10) / 10
End If


ElseIf Text29.Text = "��" Then
Text17.Text = "1"
total = Int((30 + Key / 10)) '30-50��
extra = Int(total * 0.1 * Rnd())

If Index = 0 Then
damage = Int(total * 0.3 * Rnd() * 10) / 10
atspeed = Int((total - damage) * 0.5 * Rnd() * 10) / 10
armor = Int((total - damage - atspeed) * 0.5 * Rnd() * 10) / 10
luck = Int((total - damage - atspeed - armor) * 0.7 * Rnd() * 10) / 10
speed = Int((total - damage - atspeed - luck - armor - extra) * 10) / 10
ElseIf Index = 1 Then
health = Int(total * 0.3 * Rnd() * 10) / 10
armor = Int((total - health) * 0.2 * Rnd() * 10) / 10
tough = Int((total - health) * 0.2 * Rnd() * 10) / 10
damage = Int((total - health) * 0.2 * Rnd() * 10) / 10
atspeed = Int((total - health) * 0.2 * Rnd() * 10) / 10
luck = Int((total - health) * 0.2 * Rnd() * 10) / 10
speed = Int((total - health - armor - tough - damage - atspeed - luck - extra) * 10) / 10
End If

ElseIf Text29.Text = "����" Then
Text17.Text = "1"
total = Int((80 + Key * 0.4)) '80-160��
extra = Int(total * 0.1 * Rnd())

If Index = 0 Then
damage = Int(total * 0.3 * Rnd() * 10) / 10
atspeed = Int((total - damage) * 0.5 * Rnd() * 10) / 10
armor = Int((total - damage - atspeed) * 0.5 * Rnd() * 10) / 10
luck = Int((total - damage - atspeed - armor) * 0.7 * Rnd() * 10) / 10
speed = Int((total - damage - atspeed - luck - armor - extra) * 10) / 10
ElseIf Index = 1 Then
health = Int(total * 0.2 * Rnd() * 10) / 10
armor = Int((total - health) * 0.2 * Rnd() * 10) / 10
tough = Int((total - health) * 0.2 * Rnd() * 10) / 10
damage = Int((total - health) * 0.2 * Rnd() * 10) / 10
atspeed = Int((total - health) * 0.2 * Rnd() * 10) / 10
luck = Int((total - health) * 0.2 * Rnd() * 10) / 10
speed = Int((total - health - armor - tough - damage - atspeed - luck - extra) * 10) / 10
End If


ElseIf Text29.Text = "����" Then

    Text17.Text = "2"
    total = Int((200 + Key * 2)) '200-600��
    
    If Index = 0 Then
    health = Int(total * 0.1 * Rnd() * 10) / 10
    reknock = Int(total * 0.05 * Rnd() * 10) / 10
    speed = Int(total * 0.05 * Rnd() * 10) / 10
    damage = Int(total * 0.2 * Rnd() * 10) / 10
    tough = Int(total * 0.05 * Rnd() * 10) / 10
    armor = Int(total * 0.07 * Rnd() * 10) / 10
    atspeed = Int(total * 0.2 * Rnd() * 10) / 10
    luck = Int(total * 0.2 * Rnd() * 10) / 10
    knock = Int(total * 0.03 * Rnd() * 10) / 10
    
    ElseIf Index = 1 Then
    health = Int(total * 0.2 * Rnd() * 10) / 10
    reknock = Int(total * 0.1 * Rnd() * 10) / 10
    speed = Int(total * 0.05 * Rnd() * 10) / 10
    damage = Int(total * 0.05 * Rnd() * 10) / 10
    tough = Int(total * 0.1 * Rnd() * 10) / 10
    armor = Int(total * 0.1 * Rnd() * 10) / 10
    atspeed = Int(total * 0.05 * Rnd() * 10) / 10
    luck = Int(total * 0.2 * Rnd() * 10) / 10
    knock = Int(total * 0.01 * Rnd() * 10) / 10
    End If
    extra = Int((total - damage - atspeed - luck - armor - speed - tough - reknock - knock - health) * 10) \ 10

ElseIf Text29.Text = pinjiming(8) Then

    Text17.Text = "2"
    total = Int((600 + Key * 3)) '600-1200��
    
    If Index = 0 Then
    health = Int(total * 0.1 * Rnd() * 10) / 10
    reknock = Int(total * 0.05 * Rnd() * 10) / 10
    speed = Int(total * 0.05 * Rnd() * 10) / 10
    damage = Int(total * 0.2 * Rnd() * 10) / 10
    tough = Int(total * 0.05 * Rnd() * 10) / 10
    armor = Int(total * 0.07 * Rnd() * 10) / 10
    atspeed = Int(total * 0.2 * Rnd() * 10) / 10
    luck = Int(total * 0.2 * Rnd() * 10) / 10
    knock = Int(total * 0.03 * Rnd() * 10) / 10
    
    ElseIf Index = 1 Then
    health = Int(total * 0.2 * Rnd() * 10) / 10
    reknock = Int(total * 0.1 * Rnd() * 10) / 10
    speed = Int(total * 0.05 * Rnd() * 10) / 10
    damage = Int(total * 0.05 * Rnd() * 10) / 10
    tough = Int(total * 0.1 * Rnd() * 10) / 10
    armor = Int(total * 0.1 * Rnd() * 10) / 10
    atspeed = Int(total * 0.05 * Rnd() * 10) / 10
    luck = Int(total * 0.2 * Rnd() * 10) / 10
    knock = Int(total * 0.01 * Rnd() * 10) / 10
    End If
    extra = Int((total - damage - atspeed - luck - armor - speed - tough - reknock - knock - health) * 10) \ 10

End If '_____________---------------------------------------Ʒ��---------------------

Label19.Caption = "�ܹ�:" + CStr(total) + vbCrLf + "��ħħ��:" + CStr(extra) + vbCrLf
Label19.Caption = Label19.Caption + "��������:" + CStr(total - extra) + vbCrLf
If health <> 0 Then Label19.Caption = Label19.Caption + "����:" + CStr(health) + vbCrLf
If reknock <> 0 Then Label19.Caption = Label19.Caption + "����:" + CStr(reknock) + vbCrLf
If knock <> 0 Then Label19.Caption = Label19.Caption + "����:" + CStr(knock) + vbCrLf
If speed <> 0 Then Label19.Caption = Label19.Caption + "�ٶ�:" + CStr(speed) + vbCrLf
If damage <> 0 Then Label19.Caption = Label19.Caption + "�˺�:" + CStr(damage) + vbCrLf
If tough <> 0 Then Label19.Caption = Label19.Caption + "����:" + CStr(tough) + vbCrLf
If armor <> 0 Then Label19.Caption = Label19.Caption + "����:" + CStr(armor) + vbCrLf
If atspeed <> 0 Then Label19.Caption = Label19.Caption + "����:" + CStr(atspeed) + vbCrLf
If luck <> 0 Then Label19.Caption = Label19.Caption + "����:" + CStr(luck) + vbCrLf


Text9(0).Text = CStr(health * 3) '��������
Text9(1).Text = CStr(reknock) '���˿���
Text9(2).Text = CStr(speed * 0.15) '�ٶ�
Text9(3).Text = CStr(damage) '�����˺�
Text9(4).Text = CStr(tough) '��������
Text9(5).Text = CStr(armor) '���ױ���
Text9(6).Text = CStr(atspeed * 0.8) '�����ٶ�
Text9(7).Text = CStr(luck)  '����
Text9(8).Text = CStr(knock) '����

For i = 0 To 8
If Text9(i).Text = "0" Then Text9(i).Text = ""
Next i


Call Command4_Click

mult = 1
If Text17.Text = "1" Then mult = 0.4 '��˵��ħ�ȼ�*2.5
If Text17.Text = "2" Then mult = 0.1

Do While extra > 0
If Index = 0 Then '0-2����,3���� 4��ɨ 5���� 18���� 9�;ã�10�޲���11����
Key = Int(10 * Rnd())

If Text2.Text = "��" Then '20-22���޻�26���
Key = Int(14 * Rnd())
If Key <= 5 Then Key = Int(14 * Rnd())
If Key = 10 Then Key = 20
If Key = 11 Then Key = 21
If Key = 12 Then Key = 22
If Key = 13 Then Key = 26

ElseIf Text2.Text = "��" Then '30-32�촩��
Key = Int(13 * Rnd())
If Key <= 5 Then Key = Int(13 * Rnd())
If Key = 10 Then Key = 30
If Key = 11 Then Key = 31
If Key = 12 Then Key = 32

ElseIf Text2.Text = "�����" Then '27-29�Ҽ��ף�34��
Key = Int(14 * Rnd())
If Key <= 5 Then Key = Int(14 * Rnd())
If Key = 10 Then Key = 27
If Key = 11 Then Key = 28
If Key = 12 Then Key = 29
If Key = 13 Then Key = 34

End If
If Key > 5 And Key < 9 Then Key = Key + 3
If Key = 9 Then Key = 18


If Text2.Text = "��" Or Text2.Text = "��" Or Text2.Text = "��" Then  '0-2����,3���� 4��ɨ 5���� 18���� 9�;ã�10�޲���11����  6-8׼ʱЧ
Key = Int(13 * Rnd())
If Key <= 5 Or Key = 12 Then Key = Int(13 * Rnd())
If Key = 12 Then Key = 18
End If






ElseIf Index = 1 Then '9�;ã�10�޲���11����,   12-15������19���䱣�� 16��23��24��38ѥ�� ��17��25ͷ��

Key = Int(8 * Rnd())

If Text2.Text = "ͷ��" Then
Key = Int(12 * Rnd())
If Key = 9 Then Key = 16
If Key = 10 Then Key = 23
If Key = 11 Then Key = 24
If Key = 8 Then Key = 38

ElseIf Text2.Text = "ѥ��" Then
Key = Int(10 * Rnd())
If Key = 9 Then Key = 17
If Key = 8 Then Key = 25

End If

If Key > -1 And Key < 7 Then Key = Key + 9
If Key = 7 Then Key = 19



End If




Call Command3_Click(Int(Key))
lvl = Int((extra * 0.5 * Rnd() + 1) / mult)
If mofaquan(Key) >= 4 Then lvl = 1 '1��2��4��5��6
If lvl >= 256 Then lvl = 255 '1��2��4��5��6
Text3.Text = CStr(lvl)
Call Command2_Click

extra = extra - mofaquan(Key) * lvl * mult



extra = extra - 0.1
Loop



End Sub

Private Sub Command4_Click()
List1.Clear
c(3) = ""
For i = 1 To 100
a(i) = 0
b(i) = 0
e(i) = ""
yuansu(i) = False
Next i
k = 0
List1.AddItem "��ħ��     ���Եȼ�    �ȼ�����   ���ո�ħ�ȼ� "


End Sub

Private Sub Command40_Click()
baoku = baoku + vbCrLf + Text5.Text
baojushu = baojushu + 1
baojushustr = CStr(baojushu)

For i = 0 To 3 - Len(CStr(baojushu))
baojushustr = "0" + baojushustr
Next i

For i = 1 To 4
If Text34(i) = "" Then
    canyuzhe = Text34(i - 1)
    Exit For
End If
Next i

buquan = ""
For i = 0 To 12 - Len(Text10.Text)
buquan = buquan + "  "
Next i

buquan0 = ""
For i = 0 To 4 - Len(Text33.Text)
buquan0 = buquan0 + "  "
Next i

List4.AddItem baojushustr + " " + Text29.Text + " " + Text30.Text + " " + Text10.Text + buquan + Text33.Text + buquan0 + Mid(canyuzhe, 3)
End Sub

Private Sub Command41_Click()
List4.Clear
List4.AddItem "��۱����еı��ߣ�"
List4.AddItem "��� Ʒ�� ���� ����                      ��Դ        ������"
baoku = ""
baojushu = 0
End Sub

Private Sub Command42_Click()
st0 = Mid(baoku, 1, Len(baoku) - 1)
's1 = "D:\VB\temp\" + Text3.Text + ".mcfunction"
s1 = "D:\VB\temp\temp.mcfunction"
If mcfun = True Then
Open s1 For Output As #1
Print #1, st0
Close #1
ElseIf mcfun = False Then
Open Text14.Text & "\" + Text22.Text + "\" + Text13.Text + ".txt" For Output As #1
Print #1, st0
Close #1
End If

Shell "cmd /c D:\VB\temp\change.py"
End Sub

Private Sub Command43_Click()
Timer1.Enabled = True
choujiangcishu = 0
End Sub

Private Sub Command5_Click()
If d(1) = 0 Then d(1) = 1 Else d(1) = 0
If d(1) = 0 Then Command5.Caption = "���Դݻ�"
If d(1) = 1 Then Command5.Caption = "���ɴݻ�"
End Sub

Private Sub Command6_Click(Index As Integer)
Text1.Text = Command6(Index).Caption
Text35(0).Text = Text1.Text + Text2.Text
End Sub




Private Sub Command7_Click()
bukechongfu = Not bukechongfu
If bukechongfu = True Then Command7.Caption = "�����ظ�"
If bukechongfu = False Then Command7.Caption = "�����ظ�"
End Sub

Private Sub Command8_Click(Index As Integer)
Text2.Text = Command8(Index).Caption
Text35(0).Text = Text1.Text + Text2.Text
End Sub

Private Sub Command9_Click()
mcfun = Not mcfun
If mcfun = False Then
Command9.Caption = "���json�ı�"
Text13.Text = "trade"
Text14.Text = "D:\VB\MC������\ʵ�幹��\��ƷĿ¼"
End If
If mcfun = True Then
Command9.Caption = "���mcfunction������ı���"
Text13.Text = "magic"
Text14.Text = "D:\MC\.minecraft\saves\" + Command12(Index).Caption + "\datapacks\ħ������\data\give\functions"
End If
End Sub

Private Sub Form_Load()
pinjiming(0) = "��ͨ": pinjicolor(0) = RGB(150, 150, 150)
pinjiming(1) = "����": pinjicolor(1) = RGB(0, 200, 0)
pinjiming(2) = "ϡ��": pinjicolor(2) = RGB(0, 0, 200)
pinjiming(3) = "ʷʫ": pinjicolor(3) = RGB(200, 0, 250)
pinjiming(4) = "��˵": pinjicolor(4) = RGB(250, 150, 0)
pinjiming(5) = "��": pinjicolor(5) = RGB(200, 30, 0)
pinjiming(6) = "����": pinjicolor(6) = RGB(255, 200, 0)
pinjiming(7) = "����": pinjicolor(7) = RGB(0, 250, 250)
pinjiming(8) = "����": pinjicolor(8) = RGB(250, 230, 0)

For i = 0 To 8
Command32(i).Caption = pinjiming(i)
Next i

laiyuan(0, 0) = "����֮ʱ": laiyuan(0, 1) = "����ʱ��": laiyuan(1, 0) = "��������ʱ": laiyuan(1, 1) = "����ʱ��"
laiyuan(2, 0) = "��һ��Ԫ": laiyuan(2, 1) = "����ʱ��": laiyuan(3, 0) = "��ļ�Ԫ": laiyuan(3, 1) = "����ʱ��"
laiyuan(4, 0) = "�ڶ���Ԫ": laiyuan(4, 1) = "����ʱ��": laiyuan(5, 0) = "�ɹż�Ԫ": laiyuan(5, 1) = "����ʱ��"
laiyuan(6, 0) = "Զ��ʱ��": laiyuan(6, 1) = "����ʱ��": laiyuan(7, 0) = "�Ϲ�֮ʱ": laiyuan(7, 1) = "����ʱ��"
laiyuan(8, 0) = "������ʱ��": laiyuan(6, 1) = "����ʱ��": laiyuan(9, 0) = "��ħ��ʱ��": laiyuan(9, 1) = "����ʱ��"
laiyuan(10, 0) = "����ʱ��": laiyuan(10, 1) = "����ʱ��": laiyuan(11, 0) = "ĩ��ʱ��": laiyuan(11, 1) = "����ʱ��"

laiyuan(12, 0) = "����֮��": laiyuan(12, 1) = "�����ص�": laiyuan(13, 0) = "�ķ��˼�": laiyuan(13, 1) = "�����ص�"
laiyuan(14, 0) = "��Ԩ֮��": laiyuan(14, 1) = "�����ص�": laiyuan(15, 0) = "��������": laiyuan(15, 1) = "�����ص�"
laiyuan(16, 0) = "ңԶȺ��": laiyuan(16, 1) = "�����ص�": laiyuan(17, 0) = "Զ����ͥ": laiyuan(17, 1) = "�����ص�"
laiyuan(18, 0) = "�����": laiyuan(18, 1) = "�����ص�": laiyuan(19, 0) = "ʮ��ħ��": laiyuan(19, 1) = "�����ص�"

For i = 0 To 19
Command35(i).Caption = laiyuan(i, 0)
Next i
Call Command35_Click(13)


Text8(11).Text = """minecraft:instant_health"""

Call HScroll4_Change


Text38.Text = "5"
'��ͨ�弶��1��3����2��1����4�������޲�����5
 '0-2����,3���� 4��ɨ 5���� 18���� 9�;ã�10�޲���11����
mofaquan(0) = 1
mofaquan(1) = 0.7
mofaquan(2) = 0.5
mofaquan(3) = 2
mofaquan(4) = 2
mofaquan(5) = 2

mofaquan(18) = 1
mofaquan(9) = 2
mofaquan(10) = 5
mofaquan(11) = 2

mofaquan(20) = 1 '20-22���޻�26���
mofaquan(21) = 5
mofaquan(22) = 4
mofaquan(26) = 1

mofaquan(30) = 2 '30-32�촩��
mofaquan(31) = 2
mofaquan(32) = 4

mofaquan(27) = 4 '27-29�Ҽ��ף�34��
mofaquan(28) = 5
mofaquan(29) = 5
mofaquan(34) = 1

mofaquan(6) = 4  '6-8׼ʱЧ
mofaquan(7) = 2
mofaquan(8) = 1



'12-15������19���䱣�� 16��23��24��38�����Ǳ��ѥ�� ��17ˮ���٣�25ˮ�º���ͷ��
mofaquan(12) = 1
mofaquan(13) = 1
mofaquan(14) = 1
mofaquan(15) = 1
mofaquan(19) = 1

mofaquan(16) = 2
mofaquan(23) = 2
mofaquan(24) = 2
mofaquan(38) = 2

mofaquan(17) = 4
mofaquan(25) = 2


baoku = ""
Text27.BackColor = RGB((Text15(0).Text), Val(Text15(1).Text), Val(Text15(2).Text))
Text28.Text = "#" + sljz(Val(Text15(0).Text)) + sljz(Val(Text15(1).Text)) + sljz(Val(Text15(2).Text))

Text14.Text = "D:\MC\1.20\.minecraft\saves\��dOlostland 1.0\datapacks\ħ������\data\give\functions"

Text37.Text = ""

ID(0) = "O_"
ID(1) = "M"
ID(2) = "N"
ID(3) = "O01"

Call Command41_Click

mon = 7
For i = 0 To mon * mo
If i \ mo = 0 Then
Command37(i).Caption = "green"
ElseIf i \ mo = 1 Then
Command37(i).Caption = "yellow"
ElseIf i \ mo = 2 Then
Command37(i).Caption = "gold"
ElseIf i \ mo = 3 Then
Command37(i).Caption = "aqua"
ElseIf i \ mo = 4 Then
Command37(i).Caption = "red"
ElseIf i \ mo = 5 Then
Command37(i).Caption = "dark_red"
ElseIf i \ mo = 6 Then
Command37(i).Caption = ""
End If
Next i


For i = 0 To 4
Text34(i).Text = ""
Text36(i).Text = ""
Next i

Text35(1).Text = "����"
Text35(2).Text = "������"

Randomize
Text1.Text = "��"
Text2.Text = "��"
Text3.Text = "3"
Text4.Text = "����"
Label2.Caption = ""
fumomanji = 255
j = 11

sx1(0) = "generic.max_health"
sx1(1) = "generic.knockback_resistance"
sx1(2) = "generic.movement_speed"
sx1(3) = "generic.attack_damage"
sx1(4) = "generic.armor_toughness"
sx1(5) = "generic.armor"
sx1(6) = "generic.attack_speed"
sx1(7) = "generic.luck"
sx1(8) = "generic.attack_knockback"
sx1(9) = "generic.follow_range"
sx1(10) = "horse.jump_strength"
sx1(11) = "zombie.spawn_reinforcements"
sx1(12) = "generic."
sx1(13) = "generic."

sx2(0) = "maxHealth"
sx2(1) = "knockbackResistance"
sx2(2) = "movementSpeed"
sx2(3) = "attackDamage"
sx2(4) = "armorToughness"
sx2(5) = "armor"
sx2(6) = "attackSpeed"
sx2(7) = "luck"
sx2(8) = "attackKnockback"
sx2(9) = "followRange"
sx2(10) = "jumpStrength"
sx2(11) = "spawnReinforcements"
sx2(12) = ""
sx2(13) = ""

d(1) = 1

Text5.Text = ""
mcfun = True
Command9.Caption = "���mcfunction������ı���"
shuxingkaiqi = True

bukechongfu = True
Call Command4_Click
'List1.AddItem "��ħ��     ���Եȼ�    �ȼ�����   ���ո�ħ�ȼ� "
Call Command11_Click
'List2.AddItem "��״       �켣       ��˸       ��ɫ       ����"
Call Command21_Click
'List3.AddItem "ID         ����       ʱ��       ����       ����       ͼ��"
k = 0
End Sub

Private Function bu(x As String) As String '�б���뺯��
bu = ""
buu = 0
d(1) = 1
For i = 1 To Len(x)
If Mid(x, i, 1) > "z" Then buu = buu - 1
Next i

For i = 1 To j - Len(x) + buu
bu = bu + " "
Next i
End Function


Private Function wuzhi(x As String) As String
wuzhi = x
If x = "��" Then wuzhi = "sword"
If x = "��" Then wuzhi = "bow"
If x = "��" Then wuzhi = "crossbow"
If x = "��" Then wuzhi = "hoe"
If x = "�̻����" Then wuzhi = "firework_rocket"
If x = "ͷ��" Then wuzhi = "helmet"
If x = "ñ��" Then wuzhi = "helmet"
If x = "�ؼ�" Then wuzhi = "chestplate"
If x = "����" Then wuzhi = "chestplate"
If x = "����" Then wuzhi = "leggings"
If x = "ѥ��" Then wuzhi = "boots"
If x = "Ь��" Then wuzhi = "boots"
If x = "��" Then wuzhi = "axe"
If x = "��" Then wuzhi = "pickaxe"
If x = "��" Then wuzhi = "shovel"
If x = "�����" Then wuzhi = "trident"
If x = "�ʳ�" Then wuzhi = "elytra"
If x = "�����" Then wuzhi = "fishing_rod"
If x = "����" Then wuzhi = "shield"
If x = "��" Then wuzhi = "tipped_arrow"
If x = "ҩˮ" Then wuzhi = "potion"
If x = "���ɵ�" Then wuzhi = "spawn_egg"
If x = "��" Then wuzhi = "book"
If x = "��" Then wuzhi = "saddle"
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

Public Function sjz(xx) '---------------------------------------ת��ʮ����--------------------------------------
sjz = 0

For i = 0 To 1
x = Mid(xx, i + 1, 1)
If x >= "0" And x <= "9" Then sjz = sjz + Val(x) * 16 ^ (1 - i)
If x >= "A" And x <= "F" Then sjz = sjz + (Asc(x) - Asc("A") + 10) * 16 ^ (1 - i)
Next i


End Function


Private Sub HScroll1_Change(Index As Integer)
Text15(Index).Text = CStr(HScroll1(Index).Value)
Text27.BackColor = RGB(Val(Text15(0).Text), Val(Text15(1).Text), Val(Text15(2).Text))
Text28.Text = "#" + sljz(Val(Text15(0).Text)) + sljz(Val(Text15(1).Text)) + sljz(Val(Text15(2).Text))

mon = 7
For i = 0 To mon * mo
If i \ mo = 6 Then
Command37(i).Caption = Text28.Text
End If
Next i

End Sub

Private Sub HScroll2_Change()
ID(3) = CStr(HScroll2.Value)
For i = Len(ID(3)) To 2
ID(3) = "0" + ID(3)
Next i
Text31.Text = ID(0) + ID(1) + ID(2) + ID(3)

End Sub

Private Sub HScroll3_Change()
Timer1.Interval = HScroll3.Value
End Sub

Private Sub HScroll4_Change()
For i = 0 To 20
If HScroll4.Value = i Then
Label20.Caption = "���Գ�ȡƷ������Ϊ" + pinjiming(i)
Label20.ForeColor = pinjicolor(i)
End If
Next i
End Sub

Private Sub Text15_Change(Index As Integer)
HScroll1(Index).Value = Val(Text15(Index).Text)
Text27.BackColor = RGB((Text15(0).Text), Val(Text15(1).Text), Val(Text15(2).Text))
Text28.Text = "#" + sljz(Val(Text15(0).Text)) + sljz(Val(Text15(1).Text)) + sljz(Val(Text15(2).Text))
End Sub

Private Sub Timer1_Timer()
Dim chou As Integer
'Ʒ��ѡȡ
chou = HScroll4.Value + Int((7 - HScroll4.Value) * Rnd())
Call Command32_Click(chou)

'ʱ���ص�ѡȡ
chou = Int(20 * Rnd())
If chou Mod 2 = 0 Then chou = Int(20 * Rnd())
If chou Mod 2 = 0 Then chou = Int(20 * Rnd())
Call Command35_Click(chou)

'��Ʒѡȡ
chou = Int(20 * Rnd())
Call Command8_Click(chou)

'����ѡȡ
chou = Int(8 * Rnd())
If Text2.Text = "��" Or Text2.Text = "��" Or Text2.Text = "��" Or Text2.Text = "��" Or Text2.Text = "��" Then
Do While chou = 4 Or chou = 5
chou = Int(8 * Rnd())
Loop
ElseIf Text2.Text = "ͷ��" Or Text2.Text = "�ؼ�" Or Text2.Text = "����" Or Text2.Text = "ѥ��" Then
Do While chou = 6 Or chou = 7
chou = Int(8 * Rnd())
Loop
ElseIf Text2.Text = "���ɵ�" Then
chou = 11
Else
chou = 8
End If
Call Command6_Click(chou)

'���ɱ�Դ����
If Text2.Text = "��" Or Text2.Text = "��" Or Text2.Text = "��" Or Text2.Text = "��" Or Text2.Text = "��" Or Text2.Text = "��" Or Text2.Text = "��" Or Text2.Text = "�����" Or Text2.Text = "�����" Or Text2.Text = "����" Then
Call Command39_Click(0)
ElseIf Text2.Text = "ͷ��" Or Text2.Text = "�ؼ�" Or Text2.Text = "����" Or Text2.Text = "ѥ��" Or Text2.Text = "�ʳ�" Or Text2.Text = "��" Or Text2.Text = "�̻����" Or Text2.Text = "��" Then
Call Command39_Click(1)
End If

Call Command36_Click '����

Call Command1_Click '����
Call Command40_Click '���

choujiangcishu = choujiangcishu + 1
If choujiangcishu >= Val(Text38.Text) Then Timer1.Enabled = False
End Sub
