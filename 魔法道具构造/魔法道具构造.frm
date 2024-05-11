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
   StartUpPosition =   3  '窗口缺省
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
      Text            =   "稀有"
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
      Caption         =   "绿宝石"
      Height          =   255
      Index           =   23
      Left            =   9840
      TabIndex        =   307
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "钻石"
      Height          =   255
      Index           =   22
      Left            =   10080
      TabIndex        =   306
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton Command8 
      Caption         =   "锭"
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
      Caption         =   "连抽"
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
      Caption         =   "造物"
      Height          =   375
      Left            =   12240
      TabIndex        =   302
      Top             =   8520
      Width           =   495
   End
   Begin VB.CommandButton Command41 
      Caption         =   "清空"
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
      Caption         =   "入库"
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
      Caption         =   "随机属性：装备"
      Height          =   495
      Index           =   1
      Left            =   4680
      TabIndex        =   296
      Top             =   8400
      Width           =   855
   End
   Begin VB.CommandButton Command39 
      Caption         =   "随机属性：武器"
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
      Caption         =   "随机生成物品历史、信息"
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
      Text            =   "箭"
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
      Caption         =   "「器灵残魂」"
      Height          =   855
      Index           =   7
      Left            =   11880
      TabIndex        =   242
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Command34 
      Caption         =   "「旧日之痕」"
      Height          =   855
      Index           =   6
      Left            =   11640
      TabIndex        =   241
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Command34 
      Caption         =   "「远古回响」"
      Height          =   855
      Index           =   5
      Left            =   11160
      TabIndex        =   240
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Command34 
      Caption         =   "「前生似梦」"
      Height          =   855
      Index           =   4
      Left            =   10440
      TabIndex        =   239
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Command34 
      Caption         =   "「污秽残语」"
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
      Text            =   "大魔法时代"
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton Command34 
      Caption         =   "「符文之语」"
      Height          =   855
      Index           =   2
      Left            =   10920
      TabIndex        =   236
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Command34 
      Caption         =   "「仙纹之语」"
      Height          =   855
      Index           =   1
      Left            =   10200
      TabIndex        =   235
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Command34 
      Caption         =   "「灵纹呓语」"
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
      Text            =   "魔法道具构造.frx":0000
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
      Text            =   $"魔法道具构造.frx":000D
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton Command33 
      Caption         =   "宝石"
      Height          =   495
      Index           =   7
      Left            =   11880
      TabIndex        =   230
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Command33 
      Caption         =   "神器"
      Height          =   495
      Index           =   6
      Left            =   11640
      TabIndex        =   229
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Command33 
      Caption         =   "货币"
      Height          =   495
      Index           =   5
      Left            =   11400
      TabIndex        =   228
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Command33 
      Caption         =   "材料"
      Height          =   495
      Index           =   4
      Left            =   11160
      TabIndex        =   227
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Command33 
      Caption         =   "药水"
      Height          =   495
      Index           =   3
      Left            =   10920
      TabIndex        =   226
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Command33 
      Caption         =   "料理"
      Height          =   495
      Index           =   2
      Left            =   10680
      TabIndex        =   225
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Command33 
      Caption         =   "装备"
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
      Text            =   "材料"
      Top             =   2385
      Width           =   495
   End
   Begin VB.CommandButton Command33 
      Caption         =   "武器"
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
      Caption         =   "不朽"
      Height          =   495
      Index           =   6
      Left            =   11880
      TabIndex        =   220
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Command32 
      BackColor       =   &H8000000E&
      Caption         =   "神话"
      Height          =   495
      Index           =   5
      Left            =   11640
      TabIndex        =   219
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Command32 
      Caption         =   "传说"
      Height          =   495
      Index           =   4
      Left            =   11400
      TabIndex        =   218
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Command32 
      Caption         =   "史诗"
      Height          =   495
      Index           =   3
      Left            =   11160
      TabIndex        =   217
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Command32 
      Caption         =   "稀有"
      Height          =   495
      Index           =   2
      Left            =   10920
      TabIndex        =   216
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Command32 
      Caption         =   "精良"
      Height          =   495
      Index           =   1
      Left            =   10680
      TabIndex        =   215
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Command32 
      BackColor       =   &H8000000D&
      Caption         =   "普通"
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
      Caption         =   "马鞍"
      Height          =   495
      Index           =   20
      Left            =   8160
      TabIndex        =   210
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton Command31 
      Caption         =   "转移"
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
      Caption         =   "书"
      Height          =   255
      Index           =   19
      Left            =   8400
      TabIndex        =   203
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton Command30 
      Caption         =   "录入储存"
      Height          =   495
      Left            =   1800
      TabIndex        =   202
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton Command29 
      Caption         =   "清空"
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
      Caption         =   "保存"
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
      Caption         =   "导出粒子云属性"
      Height          =   495
      Left            =   13080
      TabIndex        =   196
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "北极熊"
      Height          =   255
      Index           =   11
      Left            =   5880
      TabIndex        =   195
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "生成蛋"
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
      Caption         =   "导入"
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
      Caption         =   "新建"
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
      Caption         =   "史诗属性"
      Height          =   495
      Left            =   3600
      TabIndex        =   170
      Top             =   10080
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "锻造"
      Height          =   495
      Left            =   240
      TabIndex        =   169
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton Command20 
      Caption         =   "属性点满"
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
      Text            =   "蕴含着强大力量的"
      Top             =   1920
      Width           =   4335
   End
   Begin VB.CommandButton Command12 
      Caption         =   "新的世界 (1)"
      Height          =   225
      Index           =   1
      Left            =   9600
      TabIndex        =   164
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      Caption         =   "虚空"
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
      Caption         =   "添加"
      Height          =   375
      Left            =   18840
      TabIndex        =   145
      Top             =   6600
      Width           =   735
   End
   Begin VB.CommandButton Command21 
      Caption         =   "清空"
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
      Caption         =   "滞留"
      Height          =   495
      Index           =   10
      Left            =   5280
      TabIndex        =   142
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton Command6 
      Caption         =   "喷溅"
      Height          =   495
      Index           =   9
      Left            =   5520
      TabIndex        =   141
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton Command8 
      Caption         =   "药水"
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
      Caption         =   "开启属性"
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
      Caption         =   "属性清零"
      Height          =   495
      Left            =   3600
      TabIndex        =   120
      Top             =   8400
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "箭"
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
      Caption         =   "终极"
      Height          =   495
      Left            =   1440
      TabIndex        =   116
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton Command8 
      Caption         =   "盾牌"
      Height          =   495
      Index           =   15
      Left            =   7920
      TabIndex        =   115
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton Command17 
      Caption         =   "随机生成UUID"
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
      Text            =   "随机"
      Top             =   7560
      Width           =   2175
   End
   Begin VB.CommandButton Command16 
      Caption         =   "名称"
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
      Caption         =   "颜色代码"
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
      Text            =   "神明遗物"
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
      Caption         =   "烟花火箭"
      Height          =   495
      Index           =   14
      Left            =   9120
      TabIndex        =   89
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton Command8 
      Caption         =   "锄"
      Height          =   255
      Index           =   13
      Left            =   7440
      TabIndex        =   88
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton Command11 
      Caption         =   "清空"
      Height          =   615
      Left            =   19440
      TabIndex        =   87
      Top             =   4560
      Width           =   255
   End
   Begin VB.CommandButton Command10 
      Caption         =   "添加"
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
      Text            =   "随机"
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
      Caption         =   "不可重复"
      Height          =   255
      Left            =   2880
      TabIndex        =   71
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "清空魔咒"
      Height          =   255
      Left            =   1680
      TabIndex        =   70
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "鞘翅"
      Height          =   495
      Index           =   12
      Left            =   7920
      TabIndex        =   69
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton Command8 
      Caption         =   "钓鱼竿"
      Height          =   495
      Index           =   11
      Left            =   8640
      TabIndex        =   68
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton Command8 
      Caption         =   "三叉戟"
      Height          =   495
      Index           =   10
      Left            =   8640
      TabIndex        =   67
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "饵钓"
      Height          =   495
      Index           =   35
      Left            =   5760
      TabIndex        =   66
      Top             =   4920
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "海之眷顾"
      Height          =   495
      Index           =   33
      Left            =   5280
      TabIndex        =   65
      Top             =   4920
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "穿刺"
      Height          =   495
      Index           =   34
      Left            =   4920
      TabIndex        =   64
      Top             =   4920
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "多重射击"
      Height          =   495
      Index           =   32
      Left            =   3360
      TabIndex        =   63
      Top             =   6600
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "穿透"
      Height          =   495
      Index           =   31
      Left            =   3840
      TabIndex        =   62
      Top             =   6600
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "快速装填"
      Height          =   495
      Index           =   30
      Left            =   2880
      TabIndex        =   61
      Top             =   6600
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "引雷"
      Height          =   495
      Index           =   29
      Left            =   4200
      TabIndex        =   60
      Top             =   4920
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "激流"
      Height          =   495
      Index           =   28
      Left            =   4440
      TabIndex        =   59
      Top             =   4920
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "忠诚"
      Height          =   495
      Index           =   27
      Left            =   4680
      TabIndex        =   58
      Top             =   4920
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "冲击"
      Height          =   495
      Index           =   26
      Left            =   4200
      TabIndex        =   57
      Top             =   5400
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "水下呼吸"
      Height          =   495
      Index           =   25
      Left            =   4680
      TabIndex        =   56
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "冰霜行者"
      Height          =   495
      Index           =   24
      Left            =   4920
      TabIndex        =   55
      Top             =   4320
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "灵魂疾行"
      Height          =   495
      Index           =   23
      Left            =   4440
      TabIndex        =   54
      Top             =   4320
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "火矢"
      Height          =   495
      Index           =   22
      Left            =   4920
      TabIndex        =   53
      Top             =   5400
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "无限"
      Height          =   495
      Index           =   21
      Left            =   4680
      TabIndex        =   52
      Top             =   5400
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "力量"
      Height          =   495
      Index           =   20
      Left            =   4440
      TabIndex        =   51
      Top             =   5400
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "弹射物保护"
      Height          =   615
      Index           =   19
      Left            =   2640
      TabIndex        =   50
      Top             =   4920
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "击退"
      Height          =   495
      Index           =   18
      Left            =   2880
      TabIndex        =   49
      Top             =   4320
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "水下速掘"
      Height          =   495
      Index           =   17
      Left            =   4200
      TabIndex        =   48
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "深海探索者"
      Height          =   615
      Index           =   16
      Left            =   3960
      TabIndex        =   47
      Top             =   4320
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "保护"
      Height          =   495
      Index           =   15
      Left            =   3120
      TabIndex        =   46
      Top             =   5400
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "爆炸保护"
      Height          =   495
      Index           =   14
      Left            =   3600
      TabIndex        =   45
      Top             =   5400
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "摔落保护"
      Height          =   495
      Index           =   13
      Left            =   3120
      TabIndex        =   44
      Top             =   4920
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "火焰保护"
      Height          =   495
      Index           =   12
      Left            =   3600
      TabIndex        =   43
      Top             =   4920
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "荆棘"
      Height          =   495
      Index           =   11
      Left            =   3120
      TabIndex        =   42
      Top             =   4320
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "经验修补"
      Height          =   495
      Index           =   10
      Left            =   3360
      TabIndex        =   41
      Top             =   4320
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "耐久"
      Height          =   495
      Index           =   9
      Left            =   2640
      TabIndex        =   40
      Top             =   4320
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "效率"
      Height          =   495
      Index           =   8
      Left            =   3600
      TabIndex        =   39
      Top             =   6000
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "时运"
      Height          =   495
      Index           =   7
      Left            =   3840
      TabIndex        =   38
      Top             =   6000
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "精准采集"
      Height          =   495
      Index           =   6
      Left            =   3120
      TabIndex        =   37
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "石"
      Height          =   255
      Index           =   7
      Left            =   4800
      TabIndex        =   36
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton Command6 
      Caption         =   "木"
      Height          =   255
      Index           =   6
      Left            =   4560
      TabIndex        =   35
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton Command6 
      Caption         =   "皮革"
      Height          =   495
      Index           =   5
      Left            =   4560
      TabIndex        =   34
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton Command6 
      Caption         =   "锁链"
      Height          =   495
      Index           =   4
      Left            =   4800
      TabIndex        =   33
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton Command6 
      Caption         =   "金"
      Height          =   255
      Index           =   3
      Left            =   4800
      TabIndex        =   32
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton Command6 
      Caption         =   "铁"
      Height          =   255
      Index           =   2
      Left            =   4560
      TabIndex        =   31
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton Command8 
      Caption         =   "护腿"
      Height          =   495
      Index           =   9
      Left            =   7440
      TabIndex        =   30
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton Command8 
      Caption         =   "靴子"
      Height          =   495
      Index           =   8
      Left            =   7680
      TabIndex        =   29
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton Command8 
      Caption         =   "头盔"
      Height          =   495
      Index           =   7
      Left            =   6960
      TabIndex        =   28
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton Command8 
      Caption         =   "胸甲"
      Height          =   495
      Index           =   6
      Left            =   7200
      TabIndex        =   27
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton Command8 
      Caption         =   "镐"
      Height          =   255
      Index           =   5
      Left            =   7440
      TabIndex        =   26
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton Command8 
      Caption         =   "锹"
      Height          =   255
      Index           =   4
      Left            =   6960
      TabIndex        =   25
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton Command8 
      Caption         =   "斧"
      Height          =   255
      Index           =   3
      Left            =   7200
      TabIndex        =   24
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton Command8 
      Caption         =   "弩"
      Height          =   255
      Index           =   2
      Left            =   7680
      TabIndex        =   23
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "抢夺"
      Height          =   495
      Index           =   5
      Left            =   2160
      TabIndex        =   22
      Top             =   4800
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "横扫之刃"
      Height          =   495
      Index           =   4
      Left            =   2160
      TabIndex        =   21
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "火焰附加"
      Height          =   495
      Index           =   3
      Left            =   2160
      TabIndex        =   20
      Top             =   4320
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "节肢杀手"
      Height          =   495
      Index           =   2
      Left            =   2160
      TabIndex        =   19
      Top             =   6240
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "亡灵杀手"
      Height          =   495
      Index           =   1
      Left            =   2160
      TabIndex        =   18
      Top             =   5760
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "下界合金"
      Height          =   495
      Index           =   1
      Left            =   5280
      TabIndex        =   17
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton Command8 
      Caption         =   "弓"
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
      Caption         =   "剑"
      Height          =   255
      Index           =   0
      Left            =   6960
      TabIndex        =   13
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton Command6 
      Caption         =   "钻石"
      Height          =   495
      Index           =   0
      Left            =   5040
      TabIndex        =   12
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "不可摧毁"
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
      Caption         =   "锋利"
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
      Caption         =   "翻译并添加"
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
      Caption         =   "弹射物数据"
      Height          =   495
      Index           =   1
      Left            =   7080
      TabIndex        =   200
      Top             =   10200
      Width           =   615
   End
   Begin VB.Label Label18 
      Caption         =   "实体数据"
      Height          =   495
      Index           =   0
      Left            =   7320
      TabIndex        =   190
      Top             =   9720
      Width           =   375
   End
   Begin VB.Label Label9 
      Caption         =   "文件夹"
      Height          =   375
      Index           =   2
      Left            =   5760
      TabIndex        =   188
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "开启栏位"
      Height          =   255
      Left            =   360
      TabIndex        =   179
      Top             =   7800
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "马的跳跃"
      Height          =   375
      Index           =   11
      Left            =   2520
      TabIndex        =   178
      Top             =   9840
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "僵尸召唤"
      Height          =   375
      Index           =   10
      Left            =   2520
      TabIndex        =   177
      Top             =   10320
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "跟踪范围"
      Height          =   375
      Index           =   9
      Left            =   2520
      TabIndex        =   174
      Top             =   9360
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "攻击击退"
      Height          =   495
      Index           =   8
      Left            =   2520
      TabIndex        =   172
      Top             =   8880
      Width           =   375
   End
   Begin VB.Label Label17 
      Caption         =   "损毁程度"
      Height          =   495
      Left            =   360
      TabIndex        =   167
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label Label16 
      Caption         =   $"魔法道具构造.frx":001C
      Height          =   2295
      Left            =   360
      TabIndex        =   162
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label Label15 
      Caption         =   "隐藏属性"
      Height          =   375
      Left            =   6840
      TabIndex        =   161
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "外观"
      Height          =   375
      Index           =   12
      Left            =   17400
      TabIndex        =   159
      Top             =   6600
      Width           =   375
   End
   Begin VB.Label Label14 
      Caption         =   $"魔法道具构造.frx":00AA
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
      Caption         =   "时长"
      Height          =   375
      Index           =   10
      Left            =   13080
      TabIndex        =   156
      Top             =   6240
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "倍率(127/28满级/29弑神)"
      Height          =   735
      Index           =   9
      Left            =   12840
      TabIndex        =   155
      Top             =   5640
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "显示粒子效果"
      Height          =   495
      Index           =   8
      Left            =   14520
      TabIndex        =   154
      Top             =   6600
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "效果图标显示"
      Height          =   495
      Index           =   7
      Left            =   15960
      TabIndex        =   153
      Top             =   6600
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "蓝色光芒"
      Height          =   495
      Index           =   6
      Left            =   13080
      TabIndex        =   152
      Top             =   6600
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "生命上限"
      Height          =   375
      Index           =   0
      Left            =   1320
      TabIndex        =   139
      Top             =   7920
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "击退抗性"
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   138
      Top             =   8400
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "攻击伤害"
      Height          =   375
      Index           =   2
      Left            =   1320
      TabIndex        =   137
      Top             =   9360
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "移动速度"
      Height          =   375
      Index           =   3
      Left            =   1320
      TabIndex        =   136
      Top             =   8880
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "铠甲韧性"
      Height          =   375
      Index           =   4
      Left            =   1320
      TabIndex        =   135
      Top             =   9840
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "铠甲保护"
      Height          =   375
      Index           =   5
      Left            =   1320
      TabIndex        =   134
      Top             =   10320
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "幸运"
      Height          =   375
      Index           =   6
      Left            =   2520
      TabIndex        =   133
      Top             =   8400
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "攻击速度"
      Height          =   375
      Index           =   7
      Left            =   2520
      TabIndex        =   132
      Top             =   7920
      Width           =   495
   End
   Begin VB.Label Label13 
      Caption         =   "操作方式"
      Height          =   495
      Left            =   3600
      TabIndex        =   131
      Top             =   9600
      Width           =   495
   End
   Begin VB.Label Label12 
      Caption         =   "0=小球 1=大球 2=星星 3=苦力怕头   4=爆裂"
      Height          =   1215
      Left            =   14520
      TabIndex        =   108
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label11 
      Caption         =   $"魔法道具构造.frx":0351
      Height          =   3975
      Left            =   6720
      TabIndex        =   107
      Top             =   5880
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "闪烁"
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
         Name            =   "宋体"
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
         Name            =   "宋体"
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
         Name            =   "宋体"
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
      Caption         =   "路径名"
      Height          =   375
      Index           =   1
      Left            =   1800
      TabIndex        =   96
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "文件名"
      Height          =   375
      Index           =   0
      Left            =   1800
      TabIndex        =   95
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label8 
      Caption         =   "描述"
      Height          =   255
      Index           =   1
      Left            =   5160
      TabIndex        =   91
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "名称"
      Height          =   375
      Index           =   0
      Left            =   5400
      TabIndex        =   90
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "重数"
      Height          =   375
      Index           =   4
      Left            =   13320
      TabIndex        =   84
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "颜色"
      Height          =   255
      Index           =   3
      Left            =   13320
      TabIndex        =   82
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "形状"
      Height          =   255
      Index           =   2
      Left            =   13320
      TabIndex        =   80
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "轨迹"
      Height          =   255
      Index           =   1
      Left            =   13320
      TabIndex        =   78
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "射程"
      Height          =   375
      Index           =   0
      Left            =   13320
      TabIndex        =   76
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "数量对象"
      Height          =   495
      Left            =   3000
      TabIndex        =   15
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "附魔等级"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "附魔名"
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
Dim a(1 To 100), fumomanji As Long '第i个附魔属性的最高等级上限
Dim b(1 To 100), n As Long '第i个附魔属性的最终等级
Dim t, bukechongfu As Boolean
Dim c(1 To 100), e(1 To 100), x As String 'c各种字符串，1材质2物质3附魔标签
'e是第i个附魔属性的名称
Dim k, j, buu As Long
Dim d(1 To 10) As Integer '各种数字属性
Dim yanhua, yaoshui As String '跟c差不多
Dim st(0 To 20), st0, sx1(0 To 20), sx2(0 To 20), s1 As String 's1是对应文件路径
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
If x = "木" Then caizhi = "wooden"
If x = "石" Then caizhi = "stone"
If x = "锁链" Then caizhi = "chainmail"
If x = "金" Then caizhi = "golden"
If x = "铁" Then caizhi = "iron"
If x = "钻石" Then caizhi = "diamond"
If x = "皮革" Then caizhi = "leather"
If x = "下界合金" Then caizhi = "netherite"
If x = "滞留" Then caizhi = "lingering"
If x = "喷溅" Then caizhi = "splash"
If x = "北极熊" Then caizhi = "polar_bear"
c(1) = caizhi '材质
If x <> "" Then c(1) = c(1) + "_"
c(2) = wuzhi(Text2.Text) '物质



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

If Text10.Text = "" And Text12.Text = "" Then '——————————————————注释词条——————BEGIN————————————
st(3) = ""
Else
st(3) = "display:{Name:'{""text"":""" + Text10.Text + """,""color"":""" + Text18.Text + """,""italic"":false,""bold"":true}',"
st(3) = st(3) + "Lore:["
'ID,品阶，描述，符文之语
Lore(0) = "'{""text"":""ID:" + Text31.Text + """,""color"":""gold"",""italic"":false,""bold"":false}',"

LoreC(1) = "gray"
If Text29.Text = "普通" Then
LoreC(1) = "white"
ElseIf Text29.Text = "精良" Then
LoreC(1) = "green"
ElseIf Text29.Text = "稀有" Then
LoreC(1) = "blue"
ElseIf Text29.Text = "史诗" Then
LoreC(1) = "dark_purple"
ElseIf Text29.Text = "传说" Then
LoreC(1) = "gold"
ElseIf Text29.Text = "神话" Then
LoreC(1) = "red"
ElseIf Text29.Text = "不朽" Then
LoreC(1) = "yellow"
ElseIf Text29.Text = "永恒" Then
LoreC(1) = "aqua"
End If
Lore(1) = "'{""text"":""" + Text29.Text + " " + Text30.Text + """,""color"":""" + LoreC(1) + """,""italic"":false,""bold"":true}'," '品阶

If Text35(3).Text <> "" Then
LoreC(2) = Text19.Text '描述
Lore(2) = "'{""text"":""" + Text35(3).Text + """,""color"":""" + LoreC(2) + """,""italic"":false,""bold"":true}',"
Else
Lore(2) = ""
End If

LoreC(3) = Text19.Text '描述
Lore(3) = "'{""text"":""" + Text12.Text + Text35(0).Text + "," + Text35(1).Text + Text33.Text + Text35(2).Text + """,""color"":""" + LoreC(3) + """,""italic"":false,""bold"":true}',"


LoreC(4) = "dark_gray"
If Text32.Text = "「仙纹之语」" Then
LoreC(4) = "yellow"
ElseIf Text32.Text = "「前生似梦」" Then
LoreC(4) = "green"
ElseIf Text32.Text = "「灵纹呓语」" Then
LoreC(4) = "blue"
ElseIf Text32.Text = "「符文之语」" Then
LoreC(4) = "aqua"
ElseIf Text32.Text = "「远古回响」" Then
LoreC(4) = "dark_green"
ElseIf Text32.Text = "「污秽残语」" Then
LoreC(4) = "red"
ElseIf Text32.Text = "「旧日之痕」" Then
LoreC(4) = "dark_red"
ElseIf Text32.Text = "「器灵残魂」" Then
LoreC(4) = "white"
End If
Lore(4) = "'{""text"":""" + Text32.Text + "："",""color"":""" + LoreC(4) + """,""italic"":false,""bold"":true}',"

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



End If '——————————————————注释词条—END—————————————————

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
If Text16.Text <> "随机" Then
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

If yaoshui <> "" Then '药水
st(6) = "CustomPotionEffects:[" + yaoshui + "],CustomPotionColor:[I;" + CStr(65536 * Val(Text15(0).Text) + 256 * Val(Text15(1).Text) + Val(Text15(2).Text)) + "],Color:[I;" + CStr(65536 * Val(Text15(0).Text) + 256 * Val(Text15(1).Text) + Val(Text15(2).Text)) + "],Potion:" + Chr(34) + "minecraft:" + Text8(12).Text + Chr(34) + ","
Else
st(6) = ""
End If

If Text23.Text <> "" Then st(7) = "EntityTag:" + Text23.Text + "," Else st(7) = ""

If Text26.Text <> "" And c(2) = "crossbow" Then st(8) = "ChargedProjectiles:[" + Text26.Text + "],Charged:1b," Else st(8) = "" '弹射物


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
st0 = "give " + Text7.Text + " " + c(1) + c(2) + st(0) + " " + Text6.Text 'nbt标签组start
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
If Text8(1).Text <> "随机" Then xingzhuang = Text8(1).Text
If Text8(3).Text <> "随机" Then yanse = Text8(3).Text
For i = 1 To Val(Text8(4).Text)
If Text8(1).Text = "随机" Then xingzhuang = CStr(Int(Rnd() * 5))
If Text8(3).Text = "随机" Then yanse = CStr(Int(Rnd() * 65536 * 256 - 1))
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
List2.AddItem "形状       轨迹       闪烁       颜色       重数"
End Sub


Private Sub Command12_Click(Index As Integer)
Text14.Text = "D:\MC\1.20最新版本\.minecraft\saves\" + Command12(Index).Caption + "\datapacks\魔法构造\data\give\functions"
End Sub


Private Sub Command13_Click()
shuxingkaiqi = Not shuxingkaiqi
If shuxingkaiqi = True Then Command13.Caption = "开启属性" Else Command13.Caption = "关闭属性"
End Sub

Private Sub Command14_Click()
Text8(3).Text = CStr(65536 * Val(Text15(0).Text) + 256 * Val(Text15(1).Text) + Val(Text15(2).Text))
Text27.BackColor = RGB((Text15(0).Text), Val(Text15(1).Text), Val(Text15(2).Text))
End Sub

Private Sub Command15_Click(Index As Integer)
Text13.Text = Command15(Index).Caption
Text22.Text = Command15(Index).Caption
If Index = 0 Then Text14.Text = "D:\MC\.minecraft\saves\" + Command12(Index).Caption + "\datapacks\魔法构造\data\temp\functions"
If Index >= 1 And Index <= 2 Then Text14.Text = "D:\VB\MC—程序\实体构造\物品目录"

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

If Text4.Text = "水下速掘" Then
Label2.Caption = "aqua_affinity"
yuansu(1) = True '水系
a(k) = 1
End If
If Text4.Text = "爆炸保护" Then
Label2.Caption = "blast_protection"
a(k) = 10
End If
If Text4.Text = "深海探索者" Then
Label2.Caption = "depth_strider"
yuansu(1) = True '水系
a(k) = 3
End If
If Text4.Text = "摔落保护" Then
Label2.Caption = "feather_falling"
a(k) = 7
yuansu(6) = True '风系
End If
If Text4.Text = "火焰保护" Then
Label2.Caption = "fire_protection"
a(k) = 10
yuansu(2) = True '火系
End If
If Text4.Text = "弹射物保护" Then
Label2.Caption = "projectile_protection"
yuansu(6) = True '风系
a(k) = 10
End If
If Text4.Text = "冰霜行者" Then
Label2.Caption = "frost_walker"
a(k) = 14
yuansu(3) = True '冰系
End If
If Text4.Text = "经验修补" Then
Label2.Caption = "mending"
a(k) = 1
End If
If Text4.Text = "保护" Then
Label2.Caption = "protection"
a(k) = 20
End If
If Text4.Text = "水下呼吸" Then
Label2.Caption = "respiration"
yuansu(1) = True '水系
End If
If Text4.Text = "灵魂疾行" Then Label2.Caption = "soul_speed"
If Text4.Text = "荆棘" Then
Label2.Caption = "thorns"
yuansu(7) = True '木系
End If
If Text4.Text = "耐久" Then Label2.Caption = "unbreaking"


If Text4.Text = "火焰附加" Then
Label2.Caption = "fire_aspect"
yuansu(2) = True '火系
End If
If Text4.Text = "节肢杀手" Then Label2.Caption = "bane_of_arthropods"
If Text4.Text = "亡灵杀手" Then
Label2.Caption = "smite"
yuansu(4) = True '光系
End If
If Text4.Text = "击退" Then
Label2.Caption = "knockback"
yuansu(6) = True '风系
End If
If Text4.Text = "抢夺" Then Label2.Caption = "looting"
If Text4.Text = "锋利" Then Label2.Caption = "sharpness"
If Text4.Text = "横扫之刃" Then
Label2.Caption = "sweeping"
yuansu(6) = True '风系
End If

If Text4.Text = "火矢" Then
Label2.Caption = "flame"
yuansu(2) = True '火系
a(k) = 1
End If
If Text4.Text = "无限" Then
Label2.Caption = "infinity"
a(k) = 1
End If
If Text4.Text = "力量" Then Label2.Caption = "power"
If Text4.Text = "冲击" Then
Label2.Caption = "punch"
yuansu(6) = True '风系
End If

If Text4.Text = "多重射击" Then
Label2.Caption = "multishot"
a(k) = 1
End If
If Text4.Text = "穿透" Then
Label2.Caption = "piercing"
a(k) = 127
End If
If Text4.Text = "快速装填" Then
Label2.Caption = "quick_charge"
a(k) = 5
End If

If Text4.Text = "引雷" Then
Label2.Caption = "channeling"
yuansu(5) = True '雷系
a(k) = 1
End If
If Text4.Text = "忠诚" Then
Label2.Caption = "loyalty"
a(k) = 127
End If
If Text4.Text = "激流" Then
Label2.Caption = "riptide"
yuansu(1) = True '水系
End If
If Text4.Text = "穿刺" Then Label2.Caption = "impaling"

If Text4.Text = "精准采集" Then
Label2.Caption = "silk_touch"
yuansu(8) = True '土系
a(k) = 1
End If
If Text4.Text = "效率" Then
Label2.Caption = "efficiency"
yuansu(8) = True '土系
End If
If Text4.Text = "时运" Then
Label2.Caption = "fortune"
yuansu(8) = True '土系
End If


If Text4.Text = "海之眷顾" Then
Label2.Caption = "luck_of_the_sea"
yuansu(1) = True '水系
End If
If Text4.Text = "饵钓" Then
Label2.Caption = "lure"
yuansu(1) = True '水系
a(k) = 5
End If
e(k) = Text4.Text
t = False
For i = 1 To k - 1
If e(k) = e(i) Then t = bukechongfu '有重合或失效
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
Text9(2).Text = "0.4" '实际上1024
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
List3.AddItem "ID         倍率       时长       蓝光       粒子       图标"
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
Open "d:\VB\MC—程序\实体构造\实体目录\" + Text24.Text + ".txt" For Input As #1
Do While Not EOF(1)
Line Input #1, shiti
Loop
Close #1
Text23.Text = shiti

End Sub

Private Sub Command27_Click()
If yaoshui <> "" Then '药水
x = "Effects:[" + yaoshui + "],Potion:" + Chr(34) + "minecraft:" + Text8(12).Text + Chr(34) + ""
Else
x = ""
End If

Text5.Text = x
Open "d:\VB\MC—程序\实体构造\物品目录\药水\" + Text25.Text + ".txt" For Output As #1
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

If Command32(Index).Caption = "普通" Then
ID(2) = "N"
ElseIf Command32(Index).Caption = "精良" Then
ID(2) = "S"
ElseIf Command32(Index).Caption = "稀有" Then
ID(2) = "R"
ElseIf Command32(Index).Caption = "史诗" Then
ID(2) = "E"
ElseIf Command32(Index).Caption = "传说" Then
ID(2) = "L"
ElseIf Command32(Index).Caption = "神话" Then
ID(2) = "M"
ElseIf Command32(Index).Caption = "不朽" Then
ID(2) = "I"
ElseIf Command32(Index).Caption = "永恒" Then
ID(2) = "O"
ElseIf Command32(Index).Caption = "无上" Then
ID(2) = "G"
End If
Text31.Text = ID(0) + ID(1) + ID(2) + ID(3)

End Sub

Private Sub Command33_Click(Index As Integer)
Text30.Text = Command33(Index).Caption

If Command33(Index).Caption = "武器" Then
ID(1) = "W"
ElseIf Command33(Index).Caption = "装备" Then
ID(1) = "E"
ElseIf Command33(Index).Caption = "料理" Then
ID(1) = "F"
ElseIf Command33(Index).Caption = "药水" Then
ID(1) = "S"
ElseIf Command33(Index).Caption = "材料" Then
ID(1) = "M"
ElseIf Command33(Index).Caption = "货币" Then
ID(1) = "C"
ElseIf Command33(Index).Caption = "神器" Then
ID(1) = "A"
ElseIf Command33(Index).Caption = "宝石" Then
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

Private Sub Command36_Click() '生成物品信息

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

If Text29.Text = "神话" Or Text29.Text = "不朽" Or Text29.Text = "永恒" Or Text30.Text = "神器" Then
d(1) = 1
Else
d(1) = 0
End If

poxian = 0.2
weiyi = 0.4
chaotuo = 0.8


If d(1) = 0 Then Command5.Caption = "可以摧毁"
If d(1) = 1 Then Command5.Caption = "不可摧毁"

Text35(0).Text = Text1.Text + Text2.Text

iname = Text2.Text

aliangci = ""

Key = Int(200 * Rnd()) '强大指数分类用

If iname = "斧" Then
    
    If Key < 60 Then
    iname = "战斧"
    ElseIf Key < 120 Then iname = "斧"
    End If
    aliangci = "把"

ElseIf iname = "盾牌" Then
    If Key < 60 Then
    iname = "巨盾"
    ElseIf Key < 110 Then iname = "战盾"
    ElseIf Key < 160 Then iname = "盾"
    End If
    aliangci = "个"

ElseIf iname = "鞘翅" Then
    If Key < 50 Then
    iname = "战翼"
    ElseIf Key < 90 Then iname = "滑翔翼"
    ElseIf Key < 130 Then iname = "战翅"
    ElseIf Key < 160 Then iname = "翅膀"
    End If
    aliangci = "件"

ElseIf iname = "三叉戟" Then
    If Key < 60 Then
    iname = "方天画戟"
    ElseIf Key < 110 Then iname = "战戟"
    ElseIf Key < 160 Then iname = "戟"
    End If
    aliangci = "把"

ElseIf iname = "镐" Then
    If Key < 40 Then
    iname = "镰刀"
    ElseIf Key < 80 Then iname = "战镰"
    ElseIf Key < 110 Then iname = "月镰"
    ElseIf Key < 160 Then iname = "镰"
    End If
    aliangci = "把"

ElseIf iname = "书" Then
    If Key < 60 Then
    iname = "魔法书"
    ElseIf Key < 110 Then iname = "神书"
    ElseIf Key < 160 Then iname = "魔法书籍"
    End If
    aliangci = "本"

ElseIf iname = "剑" Then
    If Key < 30 Then
    iname = "长剑"
    ElseIf Key < 60 Then iname = "长刀"
    ElseIf Key < 90 Then iname = "魔剑"
    End If
    aliangci = "把"

ElseIf iname = "胸甲" Then
    If Key < 40 Then
    iname = "铠甲"
    ElseIf Key < 80 Then iname = "盔甲"
    ElseIf Key < 120 Then iname = "战铠"
    ElseIf Key < 160 Then iname = "铠"
    End If
    
    aliangci = "件"

ElseIf iname = "头盔" Then
    If Key < 60 Then
    iname = "盔"
    ElseIf Key < 90 Then iname = "帽子"
    ElseIf Key < 120 Then iname = "盔帽"
    End If
    aliangci = "顶"

ElseIf iname = "靴子" Then
    If Key < 50 Then
    iname = "战靴"
    ElseIf Key < 90 Then iname = "魔法战靴"
    ElseIf Key < 150 Then iname = "靴"
    End If
    aliangci = "双"

ElseIf iname = "弓" Then
    If Key < 60 Then
    iname = "弯弓"
    ElseIf Key < 110 Then iname = "大弓"
    ElseIf Key < 160 Then iname = "长弓"
    End If
    aliangci = "把"

ElseIf iname = "药水" Then
    If Key < 60 Then
    iname = "魔药"
    ElseIf Key < 110 Then iname = "灵药"
    ElseIf Key < 160 Then iname = "魔法药剂"
    End If
    aliangci = "瓶"

ElseIf iname = "烟花火箭" Then
    If Key < 60 Then
    iname = "烟花"
    ElseIf Key < 120 Then iname = "烟火"
    End If
    aliangci = "个"


ElseIf iname = "钓鱼竿" Then
    If Key < 60 Then
    iname = "鱼竿"
    ElseIf Key < 120 Then iname = "钓竿"
    End If
    aliangci = "根"

End If



If Text2.Text = "剑" Or Text2.Text = "锹" Or Text2.Text = "锄" Or Text2.Text = "镐" Or Text2.Text = "斧" Or Text2.Text = "弓" Or Text2.Text = "弩" Or Text2.Text = "三叉戟" Or Text2.Text = "钓鱼竿" Or Text2.Text = "盾牌" Or Text2.Text = "烟花火箭" Or Text2.Text = "箭" Or Text2.Text = "书" Then
Text30.Text = "武器"
If aliangci = "" Then aliangci = "把"
ElseIf Text2.Text = "头盔" Or Text2.Text = "胸甲" Or Text2.Text = "护腿" Or Text2.Text = "靴子" Or Text2.Text = "鞘翅" Or Text2.Text = "马鞍" Then
Text30.Text = "装备"
If aliangci = "" Then aliangci = "件"
End If

If Command33(Index).Caption = "武器" Then
ID(1) = "W"
ElseIf Command33(Index).Caption = "装备" Then ID(1) = "E"
ElseIf Command33(Index).Caption = "料理" Then ID(1) = "F"
ElseIf Command33(Index).Caption = "药水" Then ID(1) = "S"
ElseIf Command33(Index).Caption = "材料" Then ID(1) = "M"
ElseIf Command33(Index).Caption = "货币" Then ID(1) = "C"
ElseIf Command33(Index).Caption = "神器" Then ID(1) = "A"
ElseIf Command33(Index).Caption = "宝石" Then ID(1) = "G"
End If

Text31.Text = ID(0) + ID(1) + ID(2) + ID(3)

whetherdes = False
For i = 0 To 7
If Command32(i).Caption = Text29.Text Then whetherdes = True
Next i

If whetherdes = False Then describe = Text29.Text '品阶

Dim quanjuyuansu(1 To 8) As Boolean
yuansunum = ""
yuansudes = ""
If (Val(Text9(2).Text) > 1.2 And Val(Text17.Text) = 0) Or (Val(Text9(2).Text) > 3 And Val(Text17.Text) > 0) Or (Val(Text9(6).Text) > 10 And Val(Text17.Text) = 0) Or (Val(Text9(6).Text) > 3 And Val(Text17.Text) > 0) Then quanjuyuansu(6) = True
'风元素
If (Val(Text9(0).Text) > 100 And Val(Text17.Text) = 0) Or (Val(Text9(0).Text) > 5 And Val(Text17.Text) > 0) Then quanjuyuansu(7) = True
'木元素
If (Val(Text9(5).Text) > 20 And Val(Text17.Text) = 0) Or (Val(Text9(5).Text) > 5 And Val(Text17.Text) > 0) Then quanjuyuansu(8) = True

For i = 1 To 8: quanjuyuansu(i) = quanjuyuansu(i) Or yuansu(i): Next i

If quanjuyuansu(1) Then yuansudes = yuansudes + "水"
If quanjuyuansu(2) Then yuansudes = yuansudes + "火"
If quanjuyuansu(3) Then yuansudes = yuansudes + "冰"
If quanjuyuansu(4) Then yuansudes = yuansudes + "光"
If quanjuyuansu(5) Then yuansudes = yuansudes + "雷"
If quanjuyuansu(6) Then yuansudes = yuansudes + "风"
If quanjuyuansu(7) Then yuansudes = yuansudes + "木"
If quanjuyuansu(8) Then yuansudes = yuansudes + "土"
yuansudesc = yuansudes

If Len(yuansudes) = 1 Then yuansunum = ""
If Len(yuansudes) = 2 Then yuansunum = "两"
If Len(yuansudes) = 3 Then yuansunum = "三"
If Len(yuansudes) = 4 Then yuansunum = "四"
If Len(yuansudes) = 5 Then yuansunum = "五"
If Len(yuansudes) = 6 Then yuansunum = "六"
If Len(yuansudes) = 7 Then yuansunum = "七"
If Len(yuansudes) = 8 Then yuansunum = "八"
yuansudes = yuansudes + yuansunum

If Len(yuansudes) > 0 Then yuansudes = yuansudes + "系"




Key = Int(200 * Rnd()) '强大指数分类用

cname = Text1.Text

If Text1.Text = "木" Then
If Key < 50 Then
cname = "桃木" '_____________--------------------------------材质----------------------------
ElseIf Key < 100 Then cname = "龙泉木"
ElseIf Key < 120 Then cname = "荆棘"
ElseIf Key < 160 Then cname = "灵木"
Else: cname = "紫檀木"
End If

ElseIf Text1.Text = "石" Then
If Key < 40 Then
cname = "陨石"
ElseIf Key < 80 Then cname = "玄武岩"
ElseIf Key < 110 Then cname = "石岩"
ElseIf Key < 140 Then cname = "花岗岩"
ElseIf Key < 170 Then cname = "玄武石"
Else: cname = "深岩"
End If

ElseIf Text1.Text = "铁" Then
If Key < 40 Then
cname = "钢"
ElseIf Key < 80 Then cname = "精钢"
ElseIf Key < 110 Then cname = "银"
ElseIf Key < 140 Then cname = "秘银"
ElseIf Key < 170 Then cname = "合金"
Else: cname = "陨铁"
End If

ElseIf Text1.Text = "金" Then
If Key < 40 Then
cname = "黄金"
ElseIf Key < 80 Then cname = "精金"
ElseIf Key < 110 Then cname = "耀金"
ElseIf Key < 140 Then cname = "星金"
ElseIf Key < 170 Then cname = "魔金"
Else: cname = "金"
End If

ElseIf Text1.Text = "钻石" Then
If Key < 40 Then
cname = "青石"
ElseIf Key < 80 Then cname = "蓝晶"
ElseIf Key < 110 Then cname = "蓝金"
ElseIf Key < 140 Then cname = "水晶"
ElseIf Key < 170 Then cname = "碧玺石"
Else: cname = "钻石"
End If

ElseIf Text1.Text = "下界合金" Then
If Key < 40 Then
cname = "黑金"
ElseIf Key < 80 Then cname = "黑钢"
ElseIf Key < 110 Then cname = "黑晶"
ElseIf Key < 140 Then cname = "神钢"
ElseIf Key < 170 Then cname = "黑暗合金"
Else: cname = "深渊合金"
End If

ElseIf Text1.Text = "锁链" Then
If Key < 50 Then
cname = "铁链"
ElseIf Key < 100 Then cname = "锁子"
ElseIf Key < 150 Then cname = "银链"
Else: cname = "铁锁链"
End If

ElseIf Text1.Text = "皮革" Then
If Key < 50 Then
cname = "真皮"
ElseIf Key < 100 Then cname = "皮革"
ElseIf Key < 150 Then cname = "棉布"
Else: cname = "尼龙"
End If

End If '_____________---------------------------------------材质---------------------

Key = Int(200 * Rnd()) '分类用
Text34(2).Text = ""


If Mid(laiy, 1, 2) = "西方" Then

    If yuansunum = "八" Then yuansudes = "全元素属性"


    If Text29.Text = "普通" Then '_____________---------品阶---------------------------------------------------
        
        If Key < 50 Then
        describe = "随处可见的"
        Text34(0).Text = "\\""这" + aliangci + iname + "是" + cname + "做的，还是很棒的。\\"""
        Text34(1).Text = "——「矮人工匠」特伦"
        Text34(2).Text = ""
        ElseIf Key < 100 Then
        describe = "粗制的"
        Text34(0).Text = "\\""这" + aliangci + iname + "真的很一般！\\"""
        If yuansudes <> "" Then
        Text34(1).Text = "\\""等下，它似乎还蕴含着" + yuansudes + "的力量！\\"""
        Else
        Text34(1).Text = "\\""太粗糙了！\\"""
        End If
        Text34(2).Text = "——「鉴定师」古尔"
        ElseIf Key < 150 Then
        describe = "劣制的"
        Text34(0).Text = "\\""这" + aliangci + iname + "真的很烂！但也有点用。\\"""
        Text34(1).Text = "——「冒险者」迪卢思"
        Text34(2).Text = ""
        Else
        describe = "毫无长处的"
        Text34(0).Text = "\\""一" + aliangci + "很常见的" + iname + "，但不代表很普通。\\"""
        Text34(1).Text = "——「地精商人」伯恩斯"
        Text34(2).Text = ""
        End If
        Text36(0).Text = "white"
        
        Text18.Text = "#"
        For i = 1 To 3
        Text18.Text = Text18.Text + CStr(86 + Int(Rnd() * 14))
        Next i
        If yuansudes = "" Then
        Text12.Text = "平凡的"
        Else
        Text12.Text = "布满裂痕的" + yuansudes
        describe = "残破的"
        End If
    
    ElseIf Text29.Text = "精良" Then
        If Key < 50 Then
        describe = "优良的"
        Text34(0).Text = "\\""这" + aliangci + iname + "我很喜欢！\\"""
        Text34(1).Text = "——「战士」席勒"
        ElseIf Key < 100 Then
        describe = "精制的"
        Text34(0).Text = "\\""这是瑞马尔斯军团的制式装备。\\"""
        Text34(1).Text = "——「军团长」马里乌斯"
        ElseIf Key < 150 Then
        describe = "打磨过的"
        Text34(0).Text = "\\""这" + aliangci + iname + "质量不错。\\"""
        Text34(1).Text = "——「英雄锻造师」乌斯里"
        Else
        describe = "优秀的"
        Text34(0).Text = "\\""不错的" + iname + "！\\"""
        Text34(1).Text = "——「不败将军」格里西亚斯"
        End If
        Text36(0).Text = "dark_green"
        Text18.Text = "#"
        Text18.Text = Text18.Text + CStr(56 + Int(Rnd() * 14))
        Text18.Text = Text18.Text + "EE"
        Text18.Text = Text18.Text + CStr(56 + Int(Rnd() * 14))
        If yuansudes = "" Then
        Text12.Text = "非常不错的"
        Else
        Text12.Text = "保养过的" + yuansudes
        End If
        
        ElseIf Text29.Text = "稀有" Then
        If Key < 50 Then
        describe = "罕见的"
        Text34(0).Text = "\\""这" + aliangci + iname + "真是万中无一！\\"""
        Text34(1).Text = "——「银色男爵」德沃斯"
        Text34(2).Text = ""
        ElseIf Key < 100 Then
        describe = "家传的"
        Text34(0).Text = "\\""这是我家祖上流传下来的" + cname + iname + "。\\"""
        If yuansudes <> "" Then
        Text34(1).Text = "\\""悄悄告诉你！它还蕴含着" + yuansudes + "的力量！\\"""
        Else
        Text34(1).Text = "\\""似乎是从某个古战场上捡回来的！\\"""
        End If
        Text34(2).Text = "——「黄金骑士」克劳伦斯"
        ElseIf Key < 150 Then
        describe = "极佳的"
        Text34(0).Text = "\\""不错！这" + aliangci + iname + "真是太美了！\\"""
        Text34(1).Text = "\\""想到用" + cname + "去锻造他的一定是天才！\\"""
        Text34(2).Text = "——「战争雕塑家」科斯特"
        Else
        describe = "强力的"
        Text34(0).Text = "\\""力量！有了" + iname + "才有力量！\\"""
        Text34(1).Text = "——「大骑士」姆斯伍德"
        Text34(2).Text = ""
        End If
        Text36(0).Text = "dark_blue"
        Text18.Text = "#"
        Text18.Text = Text18.Text + CStr(7 + Int(Rnd() * 14))
        Text18.Text = Text18.Text + CStr(56 + Int(Rnd() * 14))
        Text18.Text = Text18.Text + "CC"
        If yuansudes = "" Then
        Text12.Text = "品质极佳的"
        Else
        Text12.Text = "有着微弱" + yuansudes + "魔力波动的"
        End If
    
    
    ElseIf Text29.Text = "史诗" Then
        If Key < 50 Then
        describe = "完美的"
        Text34(0).Text = "\\""事物不能完美，但是那" + aliangci + cname + iname + "可以例外！\\"""
        Text34(1).Text = "——「哲人王」道格特"
        Text34(2).Text = ""
        ElseIf Key < 80 Then
        describe = "特制的"
        Text34(0).Text = "\\""我用" + cname + "特别锻造的" + yuansudes + iname + "，很棒吧？\\"""
        Text34(1).Text = "\\""什么？不好用？……那是你不会用！\\"""
        Text34(2).Text = "——「皇家锻造师」苏特"
        ElseIf Key < 110 Then
        describe = "古老的"
        Text34(0).Text = "\\""那" + aliangci + iname + "的诞生已经不可考究。\\"""
        Text34(1).Text = "——「历史学家」福伦萨"
        Text34(2).Text = ""
        ElseIf Key < 150 Then
        describe = "历史悠久的"
        Text34(0).Text = "\\""这是一段关于一" + aliangci + cname + iname + "的悠长历史。\\"""
        Text34(1).Text = "——「凛冬皇帝」佛罗伦斯"
        Text34(2).Text = ""
        Else
        describe = "神秘的"
        Text34(0).Text = "\\""这是一" + aliangci + "神秘的" + iname + "……\\"""
        Text34(1).Text = "——「诡秘术师」乌尔夫伦"
        Text34(2).Text = ""
        End If
        Text36(0).Text = "light_purple"
        Text18.Text = "#"
        Text18.Text = Text18.Text + "FF"
        Text18.Text = Text18.Text + CStr(56 + Int(Rnd() * 44))
        Text18.Text = Text18.Text + "FF"
        If yuansudes = "" Then
        Text12.Text = "完美到极点的"
        Else
        Text12.Text = "蕴含强大" + yuansudes + "魔力的"
        End If
    
    
    ElseIf Text29.Text = "传说" Then
        If Key < 50 Then
        describe = "传奇的"
        Text34(0).Text = "\\""这是一段关于" + iname + "的传说。\\"""
        Text34(1).Text = "——「旅行者」拜恩斯"
        ElseIf Key < 100 Then
        describe = "蕴含恐怖" + yuansudes + "魔力的"
        Text34(0).Text = "\\""胜利将常伴你的" + iname + "！至于你？额……或许吧？\\"""
        Text34(1).Text = "——「魔法皇帝」赫尔萝丝"
        ElseIf Key < 150 Then
        describe = "神圣的"
        Text34(0).Text = "\\""圣光将会祝福你的" + iname + "！应该也会……祝福你？\\"""
        Text34(1).Text = "——「光明圣女」琳妮"
        Else
        describe = "「破限」" + describe
        For i = 0 To 7
        Text9(i).Text = CStr(Val(Text9(i).Text) + Int((Val(Text9(i).Text) + 2) * poxian * Rnd()))
        Next i
        Text34(0).Text = "\\""突破极限吧！" + cname + iname + "！\\"""
        Text34(1).Text = "——「传奇勇者」珀尔伦"
        End If
        Text36(0).Text = "yellow"
        Text18.Text = "#"
        Text18.Text = Text18.Text + "FF"
        Text18.Text = Text18.Text + "DD"
        Text18.Text = Text18.Text + CStr(46 + Int(Rnd() * 34))
        If yuansudes = "" Then
        Text12.Text = "曾经缔造过传奇的"
        Else
        Text12.Text = "蕴含着一位" + yuansudes + "传奇法师魔力的"
        End If
    
    ElseIf Text29.Text = "神话" Then
        If Key < 50 Then
        describe = "神铸的"
        Text34(0).Text = "\\""我的选民，用这" + aliangci + iname + "去完成使命吧！\\"""
        Text34(1).Text = "——「雷神」宙斯"
        ElseIf Key < 100 Then
        describe = "神明遗留的"
        Text34(0).Text = "\\""我的" + cname + iname + "呢？去哪了？？\\"""
        Text34(1).Text = "——「诡之圣者」该隐"
        ElseIf Key < 150 Then
        describe = "天赐的"
        Text34(0).Text = "\\""神选者，这" + aliangci + iname + "和天命就交给你了！\\"""
        Text34(1).Text = "——「命运之神」音织"
        Else
        describe = "封印着神力的"
        Text34(0).Text = "\\""神力，将从" + iname + "中爆发！\\"""
        Text34(1).Text = "——「炽天使」米加尔"
        End If
        Text36(0).Text = "yellow"
        Text18.Text = "#"
        Text18.Text = Text18.Text + "FF"
        Text18.Text = Text18.Text + CStr(56 + Int(Rnd() * 24))
        Text18.Text = Text18.Text + CStr(0 + Int(Rnd() * 24))
        If yuansudes = "" Then
        Text12.Text = "似乎在某个神话中出场过的"
        Else
        Text12.Text = "蕴含恐怖" + yuansudes + "神力的"
        End If
    
    ElseIf Text29.Text = "不朽" Then
        If Key < 50 Then
        describe = "蕴含神性的"
        Text34(0).Text = "\\""祂将永在！"
        If yuansudesc <> "" Then Text34(0).Text = Text34(0).Text + "祂将用" + yuansudesc + yuansunum + "元素创造永恒！"
        Text34(0).Text = Text34(0).Text + "\\"""
        Text34(1).Text = "——「混沌祭司」古拉金"
        ElseIf Key < 100 Then
        describe = "亘古岁月前的"
        Text34(0).Text = "\\""祂看淡了岁月，沉寂了光阴……\\"""
        Text34(1).Text = "——「古代学者」隆格尔"
        ElseIf Key < 150 Then
        describe = "无量光阴前的"
        Text34(0).Text = "\\""这是一段无数纪元之前，一" + aliangci + iname + "做的梦……\\"""
        Text34(1).Text = "——「精灵王」斯潘塞"
        ElseIf Key < 180 Then
        describe = "「唯一」" + describe
        For i = 0 To 7
        Text9(i).Text = CStr(Val(Text9(i).Text) + Int((Val(Text9(i).Text) + 2) * weiyi * Rnd()))
        Next i
        Text34(0).Text = "\\""拥有「唯一」的" + cname + iname + "，就拥有了超脱于时光的可能！\\"""
        Text34(1).Text = "——「时光之王」诺亚"
        Else
        describe = "不坏的"
        Text34(0).Text = "\\""万物都会腐朽。额……那" + aliangci + cname + iname + "是一个例外！\\"""
        Text34(1).Text = "——「日月之主」诺雅"
        End If
        Text36(0).Text = "gold"
        Text18.Text = "#"
        Text18.Text = Text18.Text + "FF"
        Text18.Text = Text18.Text + CStr(77 + Int(Rnd() * 23))
        Text18.Text = Text18.Text + CStr(10 + Int(Rnd() * 24))
        If yuansudes = "" Then
        Text12.Text = "永恒不灭、超脱光阴的"
        Else
        Text12.Text = "封印着" + yuansudesc + yuansunum + "元素神格权柄的"
        End If
    
    
    ElseIf Text29.Text = "永恒" Then
        If Key < 50 Then
        describe = "原初的"
        Text34(0).Text = "\\""这是创世神的" + iname + "，你可能用不了。\\"""
        Text34(1).Text = "——「时之天使」阿蒙"
        ElseIf Key < 100 Then
        describe = "创世的"
        Text34(0).Text = "\\""这是一" + aliangci + "见证了世界诞生的" + iname + "。\\"""
        Text34(1).Text = "——「深渊领主」蒙德"
        ElseIf Key < 150 Then
        describe = "「超脱」" + describe
        For i = 0 To 7
        Text9(i).Text = CStr(Val(Text9(i).Text) + Int((Val(Text9(i).Text) + 2) * chaotuo * Rnd()))
        Next i
        Text34(0).Text = "\\""拥有「超脱」的" + iname + "，才有可能抵达彼岸！\\"""
        Text34(1).Text = "——「彼岸钓者」灵天尊"
        Else
        describe = "始源的"
        Text34(0).Text = "\\""这是造物主用过的" + iname + "。\\"""
        Text34(1).Text = "——「原初之柱」埃伦萨"
        End If
        Text36(0).Text = "yellow"
        Text18.Text = "#"
        Text18.Text = Text18.Text + "FF"
        Text18.Text = Text18.Text + CStr(76 + Int(Rnd() * 24))
        Text18.Text = Text18.Text + "FF"
        If Len(yuansudesc) < 2 Then
        Text12.Text = "自成一方位面、内蕴无限可能的" + yuansudes
        Else
        Text12.Text = yuansudesc + yuansunum + "大元素周流不息，内蕴一方位面的"
        End If
    End If
    
    
ElseIf Mid(laiy, 1, 2) = "东方" Then

    Key = Int(100 * Rnd()) '灵气名称
    If Text33.Text = "八重妖域" Then
        ling = "妖力"
    ElseIf Text33.Text = "洪荒纪元" Or Text33.Text = "仙古纪元" Or Text33.Text = "巫妖量劫时" Then ling = "仙力"
    ElseIf Text33.Text = "十层魔海" Then ling = "魔气"
    ElseIf Key < 20 Then ling = "斗气"
    ElseIf Key < 35 Then ling = "武道气血"
    ElseIf Key < 45 Then ling = "武道真气"
    ElseIf Key < 46 Then ling = "尸气"
    ElseIf Key < 52 Then ling = "巫力"
    ElseIf Key < 62 Then ling = "魂力"
    ElseIf Key < 68 Then ling = "元力"
    ElseIf Key < 75 Then ling = "源力"
    ElseIf Key < 90 Then ling = "灵力"
    Else: ling = "仙力"
    End If
    

    If yuansunum = "八" Then yuansudes = "全属性"
    
    Key = Int(100 * Rnd()) '种族名称
    If Text33.Text = "八重妖域" And Key < 20 Then
        zhongzu = "狐族"
    ElseIf Text33.Text = "八重妖域" And Key < 40 Then zhongzu = "龙族"
    ElseIf Text33.Text = "八重妖域" And Key < 60 Then zhongzu = "蛇族"
    ElseIf Text33.Text = "八重妖域" And Key < 80 Then zhongzu = "猴族"
    ElseIf Text33.Text = "八重妖域" And Key < 100 Then zhongzu = "水族"
    ElseIf Text33.Text = "十层魔海" Then zhongzu = "魔族"
    ElseIf Key < 80 Then zhongzu = "人族"
    ElseIf Key < 85 Then zhongzu = "魂族"
    ElseIf Key < 87 Then zhongzu = "龙族"
    ElseIf Key < 89 Then zhongzu = "狐族"
    ElseIf Key < 91 Then zhongzu = "蛇族"
    ElseIf Key < 93 Then zhongzu = "猴族"
    ElseIf Key < 95 Then zhongzu = "魔族"
    ElseIf Key < 96 Then zhongzu = "凤族"
    ElseIf Key < 97 Then zhongzu = "水族"
    ElseIf Key < 98 Then zhongzu = "虎族"
    Else: zhongzu = "灵族"
    End If
    
    Key = Int(80 * Rnd()) '性别
    If Key < 40 Then
        xingbie = "男"
    Else: xingbie = "女"
    End If
    
    If xingbie = "女" Then
    chenghu = "女"
    ElseIf xingbie = "男" Then chenghu = "君"
    End If
    
    Key = Int(70 * Rnd()) '姓
    If zhongzu = "狐族" And Key < 30 Then
        xingshi = "白"
    ElseIf zhongzu = "狐族" And Key < 50 Then xingshi = "苏"
    ElseIf zhongzu = "狐族" And Key < 80 Then xingshi = "涂山"
    ElseIf zhongzu = "猴族" And Key < 60 Then xingshi = "孙"
    ElseIf zhongzu = "魂族" And Key < 60 Then xingshi = "魂"
    ElseIf zhongzu = "凤族" And Key < 60 Then xingshi = "凤"
    ElseIf zhongzu = "龙族" And Key < 60 Then xingshi = "龙"
    ElseIf zhongzu = "魔族" And Key < 20 Then xingshi = "古"
    ElseIf zhongzu = "魔族" And Key < 35 Then xingshi = "魔"
    ElseIf zhongzu = "魔族" And Key < 55 Then xingshi = "莫"
    ElseIf zhongzu = "魔族" And Key < 70 Then xingshi = "乌"
    ElseIf Key < 3 Then xingshi = "李"
    ElseIf Key < 5 Then xingshi = "南宫"
    ElseIf Key < 7 Then xingshi = "姬"
    ElseIf Key < 10 Then xingshi = "王"
    ElseIf Key < 11 Then xingshi = "东方"
    ElseIf Key < 12 Then xingshi = "独孤"
    ElseIf Key < 13 Then xingshi = "端木"
    ElseIf Key < 14 Then xingshi = "欧阳"
    ElseIf Key < 15 Then xingshi = "钟离"
    ElseIf Key < 16 Then xingshi = "宇文"
    ElseIf Key < 17 Then xingshi = "慕容"
    ElseIf Key < 18 Then xingshi = "朱"
    ElseIf Key < 19 Then xingshi = "夏侯"
    ElseIf Key < 20 Then xingshi = "公孙"
    ElseIf Key < 22 Then xingshi = "叶"
    ElseIf Key < 23 Then xingshi = "令狐"
    ElseIf Key < 25 Then xingshi = "夏"
    ElseIf Key < 27 Then xingshi = "张"
    ElseIf Key < 29 Then xingshi = "杜"
    ElseIf Key < 31 Then xingshi = "顾"
    ElseIf Key < 33 Then xingshi = "陈"
    ElseIf Key < 35 Then xingshi = "秦"
    ElseIf Key < 37 Then xingshi = "杨"
    ElseIf Key < 39 Then xingshi = "周"
    ElseIf Key < 40 Then xingshi = "吴": ElseIf Key < 41 Then xingshi = "郑"
    ElseIf Key < 42 Then xingshi = "沈": ElseIf Key < 43 Then xingshi = "孔"
    ElseIf Key < 44 Then xingshi = "曹": ElseIf Key < 45 Then xingshi = "吕"
    ElseIf Key < 46 Then xingshi = "施": ElseIf Key < 47 Then xingshi = "严"
    ElseIf Key < 48 Then xingshi = "华": ElseIf Key < 49 Then xingshi = "姜"
    ElseIf Key < 50 Then xingshi = "章": ElseIf Key < 51 Then xingshi = "潘"
    ElseIf Key < 52 Then xingshi = "金": ElseIf Key < 53 Then xingshi = "鲁"
    ElseIf Key < 54 Then xingshi = "马": ElseIf Key < 55 Then xingshi = "苗"
    ElseIf Key < 56 Then xingshi = "林": ElseIf Key < 57 Then xingshi = "唐"
    ElseIf Key < 58 Then xingshi = "穆": ElseIf Key < 59 Then xingshi = "罗"
    ElseIf Key < 60 Then xingshi = "孟": ElseIf Key < 61 Then xingshi = "毛"
    ElseIf Key < 62 Then xingshi = "徐": ElseIf Key < 63 Then xingshi = "明"
    ElseIf Key < 64 Then xingshi = "熊": ElseIf Key < 65 Then xingshi = "季"
    ElseIf Key < 66 Then xingshi = "童": ElseIf Key < 67 Then xingshi = "钟"
    ElseIf Key < 68 Then xingshi = "高": ElseIf Key < 69 Then xingshi = "蔡"
    Else: xingshi = "江"
    End If

    Dim mingzi(0 To 2) As String
    mingz = ""
    For i = 0 To 1:
        Key = Int(70 * Rnd()) '名字
        If xingbie = "女" Then
            If Key < 29 Then
            mingzi(i) = ""
            ElseIf Key < 30 Then mingzi(i) = "妍": ElseIf Key < 31 Then mingzi(i) = "嫣"
            ElseIf Key < 32 Then mingzi(i) = "瑶": ElseIf Key < 33 Then mingzi(i) = "倩"
            ElseIf Key < 34 Then mingzi(i) = "绮": ElseIf Key < 35 Then mingzi(i) = "芳"
            ElseIf Key < 36 Then mingzi(i) = "玲": ElseIf Key < 37 Then mingzi(i) = "诗"
            ElseIf Key < 38 Then mingzi(i) = "星": ElseIf Key < 39 Then mingzi(i) = "悦"
            ElseIf Key < 40 Then mingzi(i) = "芸": ElseIf Key < 41 Then mingzi(i) = "雅"
            ElseIf Key < 42 Then mingzi(i) = "萱": ElseIf Key < 43 Then mingzi(i) = "怡"
            ElseIf Key < 44 Then mingzi(i) = "乐": ElseIf Key < 45 Then mingzi(i) = "馨"
            ElseIf Key < 46 Then mingzi(i) = "慧": ElseIf Key < 47 Then mingzi(i) = "涵"
            ElseIf Key < 48 Then mingzi(i) = "琪": ElseIf Key < 49 Then mingzi(i) = "如"
            ElseIf Key < 50 Then mingzi(i) = "雨": ElseIf Key < 51 Then mingzi(i) = "菲"
            ElseIf Key < 52 Then mingzi(i) = "嘉": ElseIf Key < 53 Then mingzi(i) = "宁"
            ElseIf Key < 54 Then mingzi(i) = "薇": ElseIf Key < 55 Then mingzi(i) = "竹"
            ElseIf Key < 56 Then mingzi(i) = "柔": ElseIf Key < 57 Then mingzi(i) = "芷"
            ElseIf Key < 58 Then mingzi(i) = "雪": ElseIf Key < 59 Then mingzi(i) = "月"
            ElseIf Key < 60 Then mingzi(i) = "琴": ElseIf Key < 61 Then mingzi(i) = "芹"
            ElseIf Key < 62 Then mingzi(i) = "梅": ElseIf Key < 63 Then mingzi(i) = "佳"
            ElseIf Key < 64 Then mingzi(i) = "莜": ElseIf Key < 65 Then mingzi(i) = "婉"
            ElseIf Key < 66 Then mingzi(i) = "莹": ElseIf Key < 67 Then mingzi(i) = "香"
            ElseIf Key < 68 Then mingzi(i) = "兰": ElseIf Key < 69 Then mingzi(i) = "梦"
            Else: mingzi(i) = "初"
            End If
        ElseIf xingbie = "男" Then
            If Key < 29 Then
            mingzi(i) = ""
            ElseIf Key < 30 Then mingzi(i) = "琛": ElseIf Key < 31 Then mingzi(i) = "清"
            ElseIf Key < 32 Then mingzi(i) = "晔": ElseIf Key < 33 Then mingzi(i) = "山"
            ElseIf Key < 34 Then mingzi(i) = "海": ElseIf Key < 35 Then mingzi(i) = "天"
            ElseIf Key < 36 Then mingzi(i) = "明": ElseIf Key < 37 Then mingzi(i) = "真"
            ElseIf Key < 38 Then mingzi(i) = "星": ElseIf Key < 39 Then mingzi(i) = "云"
            ElseIf Key < 40 Then mingzi(i) = "文": ElseIf Key < 41 Then mingzi(i) = "知"
            ElseIf Key < 42 Then mingzi(i) = "凡": ElseIf Key < 43 Then mingzi(i) = "羽"
            ElseIf Key < 44 Then mingzi(i) = "浩": ElseIf Key < 45 Then mingzi(i) = "锦"
            ElseIf Key < 46 Then mingzi(i) = "华": ElseIf Key < 47 Then mingzi(i) = "国"
            ElseIf Key < 48 Then mingzi(i) = "皓": ElseIf Key < 49 Then mingzi(i) = "之"
            ElseIf Key < 50 Then mingzi(i) = "瀚": ElseIf Key < 51 Then mingzi(i) = "非"
            ElseIf Key < 52 Then mingzi(i) = "惜": ElseIf Key < 53 Then mingzi(i) = "宁"
            ElseIf Key < 54 Then mingzi(i) = "寻": ElseIf Key < 55 Then mingzi(i) = "竹"
            ElseIf Key < 56 Then mingzi(i) = "刚": ElseIf Key < 57 Then mingzi(i) = "知"
            ElseIf Key < 58 Then mingzi(i) = "阳": ElseIf Key < 59 Then mingzi(i) = "月"
            ElseIf Key < 60 Then mingzi(i) = "琴": ElseIf Key < 61 Then mingzi(i) = "卿"
            ElseIf Key < 62 Then mingzi(i) = "平": ElseIf Key < 63 Then mingzi(i) = "九"
            ElseIf Key < 64 Then mingzi(i) = "廷": ElseIf Key < 65 Then mingzi(i) = "泰"
            ElseIf Key < 66 Then mingzi(i) = "尘": ElseIf Key < 67 Then mingzi(i) = "安"
            ElseIf Key < 68 Then mingzi(i) = "白": ElseIf Key < 69 Then mingzi(i) = "礼"
            Else: mingzi(i) = "闲"
            End If
        End If
    Next i
    
    mingz = mingzi(0) + mingzi(1)
    
    If mingz = "" Then
    Key = 29 + Int(41 * Rnd()) '名字
        If xingbie = "女" Then
            If Key < 29 Then
            mingz = "倾城"
            ElseIf Key < 30 Then mingz = "星言": ElseIf Key < 31 Then mingz = "无双"
            ElseIf Key < 32 Then mingz = "瑶瑶": ElseIf Key < 33 Then mingz = "梦雪"
            ElseIf Key < 34 Then mingz = "嫣然": ElseIf Key < 35 Then mingz = "嫣雪"
            ElseIf Key < 36 Then mingz = "倾国": ElseIf Key < 37 Then mingz = "七月"
            ElseIf Key < 38 Then mingz = "夏初": ElseIf Key < 39 Then mingz = "晴初"
            ElseIf Key < 40 Then mingz = "沫晴": ElseIf Key < 41 Then mingz = "沫雪"
            ElseIf Key < 42 Then mingz = "疏莹": ElseIf Key < 43 Then mingz = "栖月"
            ElseIf Key < 44 Then mingz = "瑾栀": ElseIf Key < 45 Then mingz = "道韫"
            ElseIf Key < 46 Then mingz = "穗言": ElseIf Key < 47 Then mingz = "槿芸"
            ElseIf Key < 48 Then mingz = "时笙": ElseIf Key < 49 Then mingz = "云韶"
            ElseIf Key < 50 Then mingz = "未晞": ElseIf Key < 51 Then mingz = "言柒"
            ElseIf Key < 52 Then mingz = "嘉然": ElseIf Key < 53 Then mingz = "梓绪"
            ElseIf Key < 54 Then mingz = "翩月": ElseIf Key < 55 Then mingz = "沁雪"
            ElseIf Key < 56 Then mingz = "竹心": ElseIf Key < 57 Then mingz = "南栖"
            ElseIf Key < 58 Then mingz = "雪初": ElseIf Key < 59 Then mingz = "秋璃"
            ElseIf Key < 60 Then mingz = "琴语": ElseIf Key < 61 Then mingz = "霁月"
            ElseIf Key < 62 Then mingz = "紫嫣": ElseIf Key < 63 Then mingz = "若兮"
            ElseIf Key < 64 Then mingz = "清莹": ElseIf Key < 65 Then mingz = "婉琳"
            ElseIf Key < 66 Then mingz = "暮雨": ElseIf Key < 67 Then mingz = "芸华"
            ElseIf Key < 68 Then mingz = "半夏": ElseIf Key < 69 Then mingz = "青黛"
            Else: mingz = "霓裳"
            End If
        ElseIf xingbie = "男" Then
            If Key < 29 Then
            mingz = "倾城"
            ElseIf Key < 30 Then mingz = "淳罡": ElseIf Key < 31 Then mingz = "闻笙"
            ElseIf Key < 32 Then mingz = "卿尘": ElseIf Key < 33 Then mingz = "梦雪"
            ElseIf Key < 34 Then mingz = "淮之": ElseIf Key < 35 Then mingz = "无期"
            ElseIf Key < 36 Then mingz = "戎关": ElseIf Key < 37 Then mingz = "惜年"
            ElseIf Key < 38 Then mingz = "谨严": ElseIf Key < 39 Then mingz = "知闲"
            ElseIf Key < 40 Then mingz = "忘忧": ElseIf Key < 41 Then mingz = "少卿"
            ElseIf Key < 42 Then mingz = "弦青": ElseIf Key < 43 Then mingz = "青山"
            ElseIf Key < 44 Then mingz = "观海": ElseIf Key < 45 Then mingz = "千山"
            ElseIf Key < 46 Then mingz = "晴空": ElseIf Key < 47 Then mingz = "今安"
            ElseIf Key < 48 Then mingz = "时笙": ElseIf Key < 49 Then mingz = "云韶"
            ElseIf Key < 50 Then mingz = "未晞": ElseIf Key < 51 Then mingz = "夜明"
            ElseIf Key < 52 Then mingz = "非真": ElseIf Key < 53 Then mingz = "长庚"
            ElseIf Key < 54 Then mingz = "远帆": ElseIf Key < 55 Then mingz = "晏清"
            ElseIf Key < 56 Then mingz = "竹心": ElseIf Key < 57 Then mingz = "云歌"
            ElseIf Key < 58 Then mingz = "云渺": ElseIf Key < 59 Then mingz = "秋夜"
            ElseIf Key < 60 Then mingz = "墨白": ElseIf Key < 61 Then mingz = "霁月"
            ElseIf Key < 62 Then mingz = "羡之": ElseIf Key < 63 Then mingz = "清时"
            ElseIf Key < 64 Then mingz = "远舟": ElseIf Key < 65 Then mingz = "云起"
            ElseIf Key < 66 Then mingz = "星眠": ElseIf Key < 67 Then mingz = "云华"
            ElseIf Key < 68 Then mingz = "平生": ElseIf Key < 69 Then mingz = "无漾"
            Else: mingz = "温言"
            End If
        End If
    End If


    Key = 29 + Int(41 * Rnd())  '尊号
    If xingbie = "女" Then
        If Key < 29 Then
        zunhao = "倾城"
        ElseIf Key < 30 Then zunhao = "星河": ElseIf Key < 31 Then zunhao = "无双"
        ElseIf Key < 32 Then zunhao = "清梦": ElseIf Key < 33 Then zunhao = "梦雪"
        ElseIf Key < 34 Then zunhao = "冰蕊": ElseIf Key < 35 Then zunhao = "玄阴"
        ElseIf Key < 36 Then zunhao = "须弥": ElseIf Key < 37 Then zunhao = "念云"
        ElseIf Key < 38 Then zunhao = "梦寒": ElseIf Key < 39 Then zunhao = "绮雪"
        ElseIf Key < 40 Then zunhao = "广寒": ElseIf Key < 41 Then zunhao = "碧风"
        ElseIf Key < 42 Then zunhao = "疏莹": ElseIf Key < 43 Then zunhao = "星海"
        ElseIf Key < 44 Then zunhao = "瑾栀": ElseIf Key < 45 Then zunhao = "明月"
        ElseIf Key < 46 Then zunhao = "穗言": ElseIf Key < 47 Then zunhao = "槿芸"
        ElseIf Key < 48 Then zunhao = "寒雪": ElseIf Key < 49 Then zunhao = "彩云"
        ElseIf Key < 50 Then zunhao = "日月": ElseIf Key < 51 Then zunhao = "朝霞"
        ElseIf Key < 52 Then zunhao = "晴云": ElseIf Key < 53 Then zunhao = "玲珑"
        ElseIf Key < 54 Then zunhao = "苍穹": ElseIf Key < 55 Then zunhao = "沁雪"
        ElseIf Key < 56 Then zunhao = "紫薇": ElseIf Key < 57 Then zunhao = "青云"
        ElseIf Key < 58 Then zunhao = "星寒": ElseIf Key < 59 Then zunhao = "秋璃"
        ElseIf Key < 60 Then zunhao = "琴语": ElseIf Key < 61 Then zunhao = "霁月"
        ElseIf Key < 62 Then zunhao = "山河": ElseIf Key < 63 Then zunhao = "光阴"
        ElseIf Key < 64 Then zunhao = "清莹": ElseIf Key < 65 Then zunhao = "青玉"
        ElseIf Key < 66 Then zunhao = "暮雨": ElseIf Key < 67 Then zunhao = "云华"
        ElseIf Key < 68 Then zunhao = "命运": ElseIf Key < 69 Then zunhao = "青黛"
        Else: zunhao = "霓裳"
        End If
    ElseIf xingbie = "男" Then
        If Key < 29 Then
        zunhao = "倾城"
        ElseIf Key < 30 Then zunhao = "彼岸": ElseIf Key < 31 Then zunhao = "大日"
        ElseIf Key < 32 Then zunhao = "沧海": ElseIf Key < 33 Then zunhao = "玄清"
        ElseIf Key < 34 Then zunhao = "无为": ElseIf Key < 35 Then zunhao = "乾坤"
        ElseIf Key < 36 Then zunhao = "青莲": ElseIf Key < 37 Then zunhao = "云华"
        ElseIf Key < 38 Then zunhao = "青天": ElseIf Key < 39 Then zunhao = "纯阳"
        ElseIf Key < 40 Then zunhao = "天魔": ElseIf Key < 41 Then zunhao = "燃灯"
        ElseIf Key < 42 Then zunhao = "斗转": ElseIf Key < 43 Then zunhao = "天斗"
        ElseIf Key < 44 Then zunhao = "幽冥": ElseIf Key < 45 Then zunhao = "碧海"
        ElseIf Key < 46 Then zunhao = "玄霄": ElseIf Key < 47 Then zunhao = "烈火"
        ElseIf Key < 48 Then zunhao = "混元": ElseIf Key < 49 Then zunhao = "无极"
        ElseIf Key < 50 Then zunhao = "天雷": ElseIf Key < 51 Then zunhao = "神霄"
        ElseIf Key < 52 Then zunhao = "玉虚": ElseIf Key < 53 Then zunhao = "玉衡"
        ElseIf Key < 54 Then zunhao = "天枢": ElseIf Key < 55 Then zunhao = "长生"
        ElseIf Key < 56 Then zunhao = "永恒": ElseIf Key < 57 Then zunhao = "轮回"
        ElseIf Key < 58 Then zunhao = "神剑": ElseIf Key < 59 Then zunhao = "雷霆"
        ElseIf Key < 60 Then zunhao = "太白": ElseIf Key < 61 Then zunhao = "镇元"
        ElseIf Key < 62 Then zunhao = "碧落": ElseIf Key < 63 Then zunhao = "黄泉"
        ElseIf Key < 64 Then zunhao = "九幽": ElseIf Key < 65 Then zunhao = "昊天"
        ElseIf Key < 66 Then zunhao = "狂战": ElseIf Key < 67 Then zunhao = "苍穹"
        ElseIf Key < 68 Then zunhao = "雷霆": ElseIf Key < 69 Then zunhao = "烈日"
        Else: zunhao = "玄冥"
        End If
    End If
    
    Key = Int(120 * Rnd()) '特殊尊号
    If zhongzu = "狐族" And Key < 70 Then
    zunhao = "青丘"
    ElseIf zhongzu = "龙族" And Key < 40 Then zunhao = "青龙"
    ElseIf zhongzu = "龙族" And Key < 60 Then zunhao = "玄龙"
    ElseIf zhongzu = "龙族" And Key < 100 Then zunhao = "炎龙"
    ElseIf zhongzu = "猴族" And Key < 80 Then zunhao = "齐天"
    ElseIf zhongzu = "水族" And Key < 80 Then zunhao = "玄武"
    ElseIf zhongzu = "凤族" And Key < 40 Then zunhao = "炎凰"
    ElseIf zhongzu = "凤族" And Key < 80 Then zunhao = "神凤"
    End If

    Do While Key > (pinji + 1) * 15 Or Key < (pinji + 1) * 12 - 12
        Key = Int(120 * Rnd()) '等阶名称
    Loop
    
    If ling <> "仙力" And ling <> "妖力" And ling <> "魔气" And ling <> "尸气" Then
        If Key < 10 Then
            dengjie = Mid(ling, 1, 1) + "者"
        ElseIf Key < 20 Then dengjie = Mid(ling, 1, 1) + "师"
        ElseIf Key < 30 Then dengjie = Mid(ling, 1, 1) + "宗"
        ElseIf Key < 40 Then dengjie = Mid(ling, 1, 1) + "君"
        ElseIf Key < 50 Then dengjie = Mid(ling, 1, 1) + "王"
        ElseIf Key < 60 Then dengjie = Mid(ling, 1, 1) + "皇"
        ElseIf Key < 70 Then dengjie = Mid(ling, 1, 1) + "尊"
        ElseIf Key < 80 Then dengjie = Mid(ling, 1, 1) + "圣"
        ElseIf Key < 90 Then dengjie = Mid(ling, 1, 1) + "帝"
        ElseIf Key < 100 Then dengjie = Mid(ling, 1, 1) + "神"
        ElseIf Key < 110 Then dengjie = "神王"
        Else: dengjie = "神帝"
        End If
        
        key1 = Int(20 * Rnd())
        dengjiehou2 = Mid(dengjie, Len(dengjie) - 1, 2)
        
        If Mid(dengjiehou2, 2, 1) = "皇" And key1 < 6 And xingbie = "男" Then
        dengjiehou2 = "皇帝"
        If Mid(dengjiehou2, 2, 1) = "皇" And key1 < 6 And xingbie = "女" Then dengjiehou2 = "女皇"
        
        ElseIf Mid(dengjiehou2, 2, 1) = "尊" And key1 < 6 Then dengjiehou2 = "尊者"
        ElseIf Mid(dengjiehou2, 2, 1) = "圣" And key1 < 6 And xingbie = "男" Then dengjiehou2 = "圣者"
        ElseIf Mid(dengjiehou2, 2, 1) = "圣" And key1 < 6 And xingbie = "女" Then dengjiehou2 = "圣女"
        
        ElseIf Mid(dengjiehou2, 2, 1) = "帝" And key1 < 5 Then dengjiehou2 = "大帝"
        ElseIf Mid(dengjiehou2, 2, 1) = "帝" And key1 < 10 And xingbie = "女" Then dengjiehou2 = "女帝"
        ElseIf Mid(dengjiehou2, 2, 1) = "帝" And key1 < 12 Then dengjiehou2 = "帝" + chenghu
        End If
        
        If Key > 50 Then
            If Key Mod 10 = 0 Then
                dengjie = "半步" + dengjie + "」「" + zunhao + dengjiehou2
            Else: dengjie = dengjie + CStr(Key Mod 10) + "重" + "」「" + zunhao + dengjiehou2
            End If
        Else
            If Key Mod 10 = 0 Then
                dengjie = "半步" + dengjie
            Else: dengjie = dengjie + CStr(Key Mod 10) + "重"
            End If
        End If
        
        
    ElseIf ling = "仙力" Then
        If Key < 10 Then
            dengjie = "炼精化气"
        ElseIf Key < 20 Then dengjie = "炼气化神"
        ElseIf Key < 30 Then dengjie = "炼神返虚"
        ElseIf Key < 40 Then dengjie = "炼虚合道"
        ElseIf Key < 50 Then dengjie = "陆地神" + Mid(ling, 1, 1)
        ElseIf Key < 60 Then dengjie = "真" + Mid(ling, 1, 1)
        ElseIf Key < 70 Then dengjie = "天" + Mid(ling, 1, 1)
        ElseIf Key < 80 Then dengjie = "玄" + Mid(ling, 1, 1)
        ElseIf Key < 90 Then dengjie = "太乙金" + Mid(ling, 1, 1)
        ElseIf Key < 100 Then dengjie = "大罗金" + Mid(ling, 1, 1)
        ElseIf Key < 110 Then dengjie = "准圣"
        Else: dengjie = "圣人"
        End If
        
        key1 = Int(20 * Rnd())
        dengjiehou2 = Mid(dengjie, Len(dengjie) - 1, 2)
        If key1 < 8 Then
        dengjiehou2 = "仙" + chenghu
        ElseIf key1 < 14 And xingbie = "女" Then dengjiehou2 = "仙子"
        ElseIf key1 < 14 And xingbie = "男" Then dengjiehou2 = "散人"
        End If
        
        If Key > 80 Then
            If Key Mod 10 = 0 Then
                dengjie = "半步" + dengjie + "」「" + zunhao + dengjiehou2
            Else: dengjie = dengjie + CStr(Key Mod 10) + "重" + "」「" + zunhao + dengjiehou2
            End If
        Else
            If Key Mod 10 = 0 Then
                dengjie = "半步" + dengjie
            Else: dengjie = dengjie + CStr(Key Mod 10) + "重"
            End If
        End If
        
        
    ElseIf ling = "妖力" Or ling = "魔气" Or ling = "尸气" Or ling = "巫力" Then
        If Key < 10 Then
            dengjie = "大" + Mid(ling, 1, 1)
        ElseIf Key < 20 Then dengjie = Mid(ling, 1, 1) + "将"
        ElseIf Key < 30 Then dengjie = Mid(ling, 1, 1) + "宗"
        ElseIf Key < 40 Then dengjie = Mid(ling, 1, 1) + "君"
        ElseIf Key < 50 Then dengjie = Mid(ling, 1, 1) + "王"
        ElseIf Key < 60 Then dengjie = Mid(ling, 1, 1) + "皇"
        ElseIf Key < 70 Then dengjie = Mid(ling, 1, 1) + "尊"
        ElseIf Key < 80 Then dengjie = Mid(ling, 1, 1) + "圣"
        ElseIf Key < 90 Then dengjie = Mid(ling, 1, 1) + "帝"
        ElseIf Key < 100 Then dengjie = Mid(ling, 1, 1) + "神"
        ElseIf Key < 110 Then dengjie = "神王"
        Else: dengjie = "神帝"
        End If
        
        
        If Key > 60 Then
            If Key Mod 10 = 0 Then
                dengjie = "半步" + dengjie + "」「" + zunhao + dengjie
            Else: dengjie = dengjie + CStr(Key Mod 10) + "重" + "」「" + zunhao + dengjie
            End If
        Else
            If Key Mod 10 = 0 Then
                dengjie = "半步" + dengjie
            Else: dengjie = dengjie + CStr(Key Mod 10) + "重"
            End If
        End If
        
    End If







    Key = Int(200 * Rnd())  '随机数

    If Text29.Text = "普通" Then '_____________---------东方品阶---------------------------------------------------
        
        If Key < 50 Then
        describe = "朴素的"
        Text34(0).Text = "\\""天地万物生来平凡。这" + aliangci + "" + iname + "不过是还未成长罢了。\\"""
        Text34(1).Text = "\\""风虽上于九天，亦起于青萍之末。\\"""
        Text34(2).Text = "——「" + dengjie + "」" + xingshi + mingz
        ElseIf Key < 100 Then
        describe = "平凡的"
        Text34(0).Text = "\\""这" + aliangci + iname + "你是从哪里得到的？\\"""
        If yuansudes <> "" Then
        Text34(1).Text = "\\""它似乎缠绕着" + yuansudes + ling + "！\\"""
        Else
        Text34(1).Text = "\\""没有任何特别之处！\\"""
        End If
        Text34(2).Text = "——「鉴定大师」" + xingshi + mingz
        ElseIf Key < 150 Then
        describe = "一般的"
        Text34(0).Text = "\\""这" + aliangci + iname + "是我最开始的回忆\\"""
        Text34(1).Text = "——「" + dengjie + "」" + xingshi + mingz
        Text34(2).Text = ""
        Else
        describe = "纯天然的"
        Text34(0).Text = "\\""我这里有一" + aliangci + "很常见的" + iname + "。\\"""
        Text34(1).Text = "——「" + dengjie + "」" + xingshi + mingz
        Text34(2).Text = ""
        End If
        Text36(0).Text = "white"
        
        Text18.Text = "#"
        For i = 1 To 3
        Text18.Text = Text18.Text + CStr(86 + Int(Rnd() * 14))
        Next i
        If yuansudes = "" Then
        Text12.Text = "随处可见的"
        Else
        Text12.Text = "布满裂痕的" + yuansudes
        describe = "残破的"
        End If
    
    ElseIf Text29.Text = "精良" Then
        If Key < 50 Then
        describe = "优良的"
        Text34(0).Text = "\\""这" + aliangci + iname + "我很喜欢！\\"""
        Text34(1).Text = "——「" + dengjie + "」" + xingshi + mingz
        ElseIf Key < 100 Then
        describe = "精制的"
        Text34(0).Text = "\\""这是我们宗门的制式装备。\\"""
        Text34(1).Text = "——「" + dengjie + "」" + xingshi + mingz
        ElseIf Key < 150 Then
        describe = "打磨过的"
        Text34(0).Text = "\\""这" + aliangci + iname + "质量不错。\\"""
        Text34(1).Text = "——「" + dengjie + "」" + xingshi + mingz
        Else
        describe = "优秀的"
        Text34(0).Text = "\\""不错的" + iname + "！\\"""
        Text34(1).Text = "——「" + dengjie + "」" + xingshi + mingz
        End If
        Text36(0).Text = "dark_green"
        Text18.Text = "#"
        Text18.Text = Text18.Text + CStr(56 + Int(Rnd() * 14))
        Text18.Text = Text18.Text + "EE"
        Text18.Text = Text18.Text + CStr(56 + Int(Rnd() * 14))
        If yuansudes = "" Then
        Text12.Text = "非常不错的"
        Else
        Text12.Text = "保养过的" + yuansudes
        End If
        
        ElseIf Text29.Text = "稀有" Then
        If Key < 50 Then
        describe = "罕见的"
        Text34(0).Text = "\\""这" + aliangci + iname + "真是万中无一！\\"""
        Text34(1).Text = "——「" + dengjie + "」" + xingshi + mingz
        Text34(2).Text = ""
        ElseIf Key < 100 Then
        describe = "家传的"
        Text34(0).Text = "\\""这是我家祖上流传下来的" + cname + iname + "。\\"""
        If yuansudes <> "" Then
        Text34(1).Text = "\\""悄悄告诉你！它还蕴含着" + yuansudes + "的力量！\\"""
        Else
        Text34(1).Text = "\\""似乎是从某个洞天遗迹上捡回来的！\\"""
        End If
        Text34(2).Text = "——「" + dengjie + "」" + xingshi + mingz
        ElseIf Key < 150 Then
        describe = "极佳的"
        Text34(0).Text = "\\""不错！这" + aliangci + iname + "真是太美了！\\"""
        Text34(1).Text = "\\""想到用" + cname + "去锻造他的一定是天才！\\"""
        Text34(2).Text = "——「" + dengjie + "」" + xingshi + mingz
        Else
        describe = "强力的"
        Text34(0).Text = "\\""力量！有了" + iname + "才有力量！\\"""
        Text34(1).Text = "——「" + dengjie + "」" + xingshi + mingz
        Text34(2).Text = ""
        End If
        Text36(0).Text = "dark_blue"
        Text18.Text = "#"
        Text18.Text = Text18.Text + CStr(7 + Int(Rnd() * 14))
        Text18.Text = Text18.Text + CStr(56 + Int(Rnd() * 14))
        Text18.Text = Text18.Text + "CC"
        If yuansudes = "" Then
        Text12.Text = "品质极佳的"
        Else
        Text12.Text = "有着微弱" + yuansudes + ling + "波动的"
        End If
    
    
    ElseIf Text29.Text = "史诗" Then
        If Key < 50 Then
        describe = "完美的"
        Text34(0).Text = "\\""事物不能完美，但是那" + aliangci + cname + iname + "可以例外！\\"""
        Text34(1).Text = "——「" + zunhao + "仙" + chenghu + "」" + xingshi + mingz
        Text34(2).Text = ""
        ElseIf Key < 80 Then
        describe = "特制的"
        Text34(0).Text = "\\""我用" + cname + "特别锻造的" + yuansudes + iname + "，很棒吧？\\"""
        Text34(1).Text = "\\""什么？不好用？……那是你不会用！\\"""
        Text34(2).Text = "——「" + dengjie + "」" + xingshi + mingz
        ElseIf Key < 110 Then
        describe = "古老的"
        Text34(0).Text = "\\""那" + aliangci + iname + "的诞生已经不可考究。\\"""
        Text34(1).Text = "——「" + zunhao + "仙" + chenghu + "」" + xingshi + mingz
        Text34(2).Text = ""
        ElseIf Key < 150 Then
        describe = "历史悠久的"
        Text34(0).Text = "\\""这是一段关于一" + aliangci + cname + iname + "的悠长历史。\\"""
        Text34(1).Text = "——「" + dengjie + "」" + xingshi + mingz
        Text34(2).Text = ""
        Else
        describe = "神秘的"
        Text34(0).Text = "\\""这是一" + aliangci + "神秘的" + iname + "……\\"""
        Text34(1).Text = "——「" + zunhao + "仙" + chenghu + "」" + xingshi + mingz
        Text34(2).Text = ""
        End If
        Text36(0).Text = "light_purple"
        Text18.Text = "#"
        Text18.Text = Text18.Text + "FF"
        Text18.Text = Text18.Text + CStr(56 + Int(Rnd() * 44))
        Text18.Text = Text18.Text + "FF"
        If yuansudes = "" Then
        Text12.Text = "完美到极点的"
        Else
        Text12.Text = "蕴含强大" + yuansudes + ling + "的"
        End If
    
    
    ElseIf Text29.Text = "传说" Then
        If Key < 50 Then
        describe = "传奇的"
        Text34(0).Text = "\\""这是一段关于" + iname + "的传说。\\"""
        Text34(1).Text = "——「" + dengjie + "」" + xingshi + mingz
        ElseIf Key < 100 Then
        describe = "蕴含恐怖" + yuansudes + ling + "的"
        Text34(0).Text = "\\""胜利将常伴你的" + iname + "！\\"""
        Text34(1).Text = "——「" + dengjie + "」" + xingshi + mingz
        ElseIf Key < 150 Then
        describe = "神圣的"
        Text34(0).Text = "\\""圣光将会祝福你的" + iname + "！应该也会……祝福你？\\"""
        Text34(1).Text = "——「" + dengjie + "」" + xingshi + mingz
        Else
        describe = "「破限」" + describe
        For i = 0 To 7
        Text9(i).Text = CStr(Val(Text9(i).Text) + Int((Val(Text9(i).Text) + 2) * poxian * Rnd()))
        Next i
        Text34(0).Text = "\\""突破极限吧！" + cname + iname + "！\\"""
        Text34(1).Text = "——「" + dengjie + "」" + xingshi + mingz
        End If
        Text36(0).Text = "yellow"
        Text18.Text = "#"
        Text18.Text = Text18.Text + "FF"
        Text18.Text = Text18.Text + "DD"
        Text18.Text = Text18.Text + CStr(46 + Int(Rnd() * 34))
        If yuansudes = "" Then
        Text12.Text = "曾经缔造过传奇的"
        Else
        Text12.Text = "蕴含着一位大能" + yuansudes + ling + "的"
        End If
    
    ElseIf Text29.Text = "神话" Then
        If Key < 50 Then
        describe = "仙铸的"
        Text34(0).Text = "\\""弟子们，用这" + aliangci + iname + "去守护宗门吧！\\"""
        Text34(1).Text = "——「" + dengjie + "」" + xingshi + mingz
        ElseIf Key < 100 Then
        describe = "仙人遗留的"
        Text34(0).Text = "\\""我的" + cname + iname + "呢？去哪了？？\\"""
        Text34(1).Text = "——「" + dengjie + "」" + xingshi + mingz
        ElseIf Key < 150 Then
        describe = "天赐的"
        Text34(0).Text = "\\""后辈们，这" + aliangci + iname + "和天命就交给你了！\\"""
        Text34(1).Text = "——「" + dengjie + "」" + xingshi + mingz
        Else
        describe = "封印着海量" + ling + "的"
        Text34(0).Text = "\\""" + ling + "，将从" + iname + "中爆发！\\"""
        Text34(1).Text = "——「" + dengjie + "」" + xingshi + mingz
        End If
        Text36(0).Text = "yellow"
        Text18.Text = "#"
        Text18.Text = Text18.Text + "FF"
        Text18.Text = Text18.Text + CStr(56 + Int(Rnd() * 24))
        Text18.Text = Text18.Text + CStr(0 + Int(Rnd() * 24))
        If yuansudes = "" Then
        Text12.Text = "似乎在某个传说中出场过的"
        Else
        Text12.Text = "蕴含恐怖" + yuansudes + ling + "与仙力的"
        End If
    
    ElseIf Text29.Text = "不朽" Then
        If Key < 50 Then
        describe = "永恒的"
        Text34(0).Text = "\\""祂将永在！"
        If yuansudesc <> "" Then Text34(0).Text = Text34(0).Text + "祂将用" + yuansudesc + yuansunum + "元素创造永恒！"
        Text34(0).Text = Text34(0).Text + "\\"""
        Text34(1).Text = "——「" + dengjie + "」" + xingshi + mingz
        ElseIf Key < 100 Then
        describe = "亘古岁月前的"
        Text34(0).Text = "\\""祂看淡了岁月，沉寂了光阴……\\"""
        Text34(1).Text = "——「" + dengjie + "」" + xingshi + mingz
        ElseIf Key < 150 Then
        describe = "无量光阴前的"
        Text34(0).Text = "\\""这是一段无数纪元之前，一" + aliangci + iname + "做的梦……\\"""
        Text34(1).Text = "——「" + dengjie + "」" + xingshi + mingz
        ElseIf Key < 180 Then
        describe = "「唯一」" + describe
        For i = 0 To 7
        Text9(i).Text = CStr(Val(Text9(i).Text) + Int((Val(Text9(i).Text) + 2) * weiyi * Rnd()))
        Next i
        Text34(0).Text = "\\""拥有「唯一」的" + cname + iname + "，就拥有了超脱于时光的可能！\\"""
        Text34(1).Text = "——「" + dengjie + "」" + xingshi + mingz
        Else
        describe = "不坏的"
        Text34(0).Text = "\\""万物都会腐朽。额……那" + aliangci + cname + iname + "是一个例外！\\"""
        Text34(1).Text = "——「" + dengjie + "」" + xingshi + mingz
        End If
        Text36(0).Text = "gold"
        Text18.Text = "#"
        Text18.Text = Text18.Text + "FF"
        Text18.Text = Text18.Text + CStr(77 + Int(Rnd() * 23))
        Text18.Text = Text18.Text + CStr(10 + Int(Rnd() * 24))
        If yuansudes = "" Then
        Text12.Text = "永恒不灭、超脱光阴的"
        Else
        Text12.Text = "封印着" + yuansudes + "仙帝权柄的"
        End If
    
    
    ElseIf Text29.Text = "永恒" Then
        If Key < 50 Then
        describe = "原初的"
        Text34(0).Text = "\\""这是圣人的" + iname + "，你可能用不了。\\"""
        Text34(1).Text = "——「" + dengjie + "」" + xingshi + mingz
        ElseIf Key < 100 Then
        describe = "创世的"
        Text34(0).Text = "\\""这是一" + aliangci + "见证了世界诞生的" + iname + "。\\"""
        Text34(1).Text = "——「" + dengjie + "」" + xingshi + mingz
        ElseIf Key < 150 Then
        describe = "「超脱」" + describe
        For i = 0 To 7
        Text9(i).Text = CStr(Val(Text9(i).Text) + Int((Val(Text9(i).Text) + 2) * chaotuo * Rnd()))
        Next i
        Text34(0).Text = "\\""拥有「超脱」的" + iname + "，才有可能抵达彼岸！\\"""
        Text34(1).Text = "——「" + dengjie + "」" + xingshi + mingz
        Else
        describe = "始源的"
        Text34(0).Text = "\\""这是圣人用过的" + iname + "。\\"""
        Text34(1).Text = "——「" + dengjie + "」" + xingshi + mingz
        End If
        Text36(0).Text = "yellow"
        Text18.Text = "#"
        Text18.Text = Text18.Text + "FF"
        Text18.Text = Text18.Text + CStr(76 + Int(Rnd() * 24))
        Text18.Text = Text18.Text + "FF"
        If Len(yuansudesc) < 2 Then
        Text12.Text = "自成一方洞天、内蕴无限可能的" + yuansudes
        Else
        Text12.Text = yuansudesc + yuansunum + "大灵气周流不息，内蕴一方世界的"
        End If
    End If

End If '_____________---------------------------------------品阶---------------------



Text36(1).Text = Text36(0).Text
Text36(2).Text = Text36(1).Text
descri = ""
describ = ""
If (Val(Text9(0).Text) > 100 And Val(Text17.Text) = 0) Or (Val(Text9(0).Text) > 5 And Val(Text17.Text) > 0) Then describ = describ + "木命纹、"
If (Val(Text9(1).Text) > 0.8 And Val(Text17.Text) = 0) Or (Val(Text9(1).Text) > 2 And Val(Text17.Text) > 0) Then describ = describ + "重山符、"
If (Val(Text9(2).Text) > 1.2 And Val(Text17.Text) = 0) Or (Val(Text9(2).Text) > 3 And Val(Text17.Text) > 0) Then describ = describ + "疾风印、"
If (Val(Text9(3).Text) > 70 And Val(Text17.Text) = 0) Or (Val(Text9(3).Text) > 4 And Val(Text17.Text) > 0) Then describ = describ + "湮灭纹、"
If (Val(Text9(4).Text) > 18 And Val(Text17.Text) = 0) Or (Val(Text9(4).Text) > 2 And Val(Text17.Text) > 0) Then describ = describ + "刚韧印、"
If (Val(Text9(5).Text) > 18 And Val(Text17.Text) = 0) Or (Val(Text9(5).Text) > 2 And Val(Text17.Text) > 0) Then describ = describ + "铠岩灵、"
If (Val(Text9(6).Text) > 10 And Val(Text17.Text) = 0) Or (Val(Text9(6).Text) > 3 And Val(Text17.Text) > 0) Then describ = describ + "狂风印、"
If (Val(Text9(7).Text) > 40 And Val(Text17.Text) = 0) Or (Val(Text9(7).Text) > 3 And Val(Text17.Text) > 0) Then describ = describ + "神眷痕、"


If describ <> "" Then describ = "铭刻有" + describ
If describ <> "" Then describ = Mid(describ, 1, Len(describ) - 1)
If describ <> "" Then describ = describ + "等符文，"

If (Val(Text9(6).Text) < 0) Or (Val(Text9(2).Text) < 0) Then
If Len(iname) > 1 Then
describ = "沉重的、" + describ
Else
descri = descri + "重"
iname = descri + iname
End If
End If


Text10.Text = describe + cname + iname '名称
Text35(3).Text = describ    '描述前缀
Text35(0).Text = cname + iname '描述宾语


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
Key = Int(200 * Rnd()) '强大指数分类用
If Index = 0 Then Text21.Text = """mainhand"""
If Index = 1 Then
    Text21.Text = """offhand"""
    If Text2.Text = "头盔" Then Text21.Text = """head"""
    If Text2.Text = "胸甲" Then Text21.Text = """chest"""
    If Text2.Text = "护腿" Then Text21.Text = """legs"""
    If Text2.Text = "靴子" Then Text21.Text = """feet"""
End If

If Text29.Text = "普通" Then '_____________---------品阶---------------------------------------------------
    Text17.Text = "0"
    total = Int((5 + Key \ 40)) '5-10点
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


ElseIf Text29.Text = "精良" Then
    Text17.Text = "0"
    total = Int((10 + Key \ 20)) '10-20点
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

ElseIf Text29.Text = "稀有" Then
    Text17.Text = "0"
    total = Int((20 + Key \ 10)) '20-40点
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

ElseIf Text29.Text = "史诗" Then
    Text17.Text = "0"
    total = Int((40 + Key \ 5)) '40-80点
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

ElseIf Text29.Text = "传说" Then
Text17.Text = "1"
total = Int((20 + Key / 20)) '20-30点
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


ElseIf Text29.Text = "神话" Then
Text17.Text = "1"
total = Int((30 + Key / 10)) '30-50点
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

ElseIf Text29.Text = "不朽" Then
Text17.Text = "1"
total = Int((80 + Key * 0.4)) '80-160点
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


ElseIf Text29.Text = "永恒" Then

    Text17.Text = "2"
    total = Int((200 + Key * 2)) '200-600点
    
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
    total = Int((600 + Key * 3)) '600-1200点
    
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

End If '_____________---------------------------------------品阶---------------------

Label19.Caption = "总共:" + CStr(total) + vbCrLf + "附魔魔力:" + CStr(extra) + vbCrLf
Label19.Caption = Label19.Caption + "利用属性:" + CStr(total - extra) + vbCrLf
If health <> 0 Then Label19.Caption = Label19.Caption + "生命:" + CStr(health) + vbCrLf
If reknock <> 0 Then Label19.Caption = Label19.Caption + "击抗:" + CStr(reknock) + vbCrLf
If knock <> 0 Then Label19.Caption = Label19.Caption + "击退:" + CStr(knock) + vbCrLf
If speed <> 0 Then Label19.Caption = Label19.Caption + "速度:" + CStr(speed) + vbCrLf
If damage <> 0 Then Label19.Caption = Label19.Caption + "伤害:" + CStr(damage) + vbCrLf
If tough <> 0 Then Label19.Caption = Label19.Caption + "韧性:" + CStr(tough) + vbCrLf
If armor <> 0 Then Label19.Caption = Label19.Caption + "保护:" + CStr(armor) + vbCrLf
If atspeed <> 0 Then Label19.Caption = Label19.Caption + "攻速:" + CStr(atspeed) + vbCrLf
If luck <> 0 Then Label19.Caption = Label19.Caption + "幸运:" + CStr(luck) + vbCrLf


Text9(0).Text = CStr(health * 3) '生命提升
Text9(1).Text = CStr(reknock) '击退抗性
Text9(2).Text = CStr(speed * 0.15) '速度
Text9(3).Text = CStr(damage) '攻击伤害
Text9(4).Text = CStr(tough) '铠甲韧性
Text9(5).Text = CStr(armor) '铠甲保护
Text9(6).Text = CStr(atspeed * 0.8) '攻击速度
Text9(7).Text = CStr(luck)  '幸运
Text9(8).Text = CStr(knock) '击退

For i = 0 To 8
If Text9(i).Text = "0" Then Text9(i).Text = ""
Next i


Call Command4_Click

mult = 1
If Text17.Text = "1" Then mult = 0.4 '传说附魔等级*2.5
If Text17.Text = "2" Then mult = 0.1

Do While extra > 0
If Index = 0 Then '0-2锋利,3火焰 4横扫 5抢夺 18击退 9耐久，10修补，11荆棘
Key = Int(10 * Rnd())

If Text2.Text = "弓" Then '20-22力无火，26冲击
Key = Int(14 * Rnd())
If Key <= 5 Then Key = Int(14 * Rnd())
If Key = 10 Then Key = 20
If Key = 11 Then Key = 21
If Key = 12 Then Key = 22
If Key = 13 Then Key = 26

ElseIf Text2.Text = "弩" Then '30-32快穿多
Key = Int(13 * Rnd())
If Key <= 5 Then Key = Int(13 * Rnd())
If Key = 10 Then Key = 30
If Key = 11 Then Key = 31
If Key = 12 Then Key = 32

ElseIf Text2.Text = "三叉戟" Then '27-29忠激雷，34刺
Key = Int(14 * Rnd())
If Key <= 5 Then Key = Int(14 * Rnd())
If Key = 10 Then Key = 27
If Key = 11 Then Key = 28
If Key = 12 Then Key = 29
If Key = 13 Then Key = 34

End If
If Key > 5 And Key < 9 Then Key = Key + 3
If Key = 9 Then Key = 18


If Text2.Text = "镐" Or Text2.Text = "斧" Or Text2.Text = "锄" Then  '0-2锋利,3火焰 4横扫 5抢夺 18击退 9耐久，10修补，11荆棘  6-8准时效
Key = Int(13 * Rnd())
If Key <= 5 Or Key = 12 Then Key = Int(13 * Rnd())
If Key = 12 Then Key = 18
End If






ElseIf Index = 1 Then '9耐久，10修补，11荆棘,   12-15保护，19弹射保护 16，23，24，38靴子 ，17，25头盔

Key = Int(8 * Rnd())

If Text2.Text = "头盔" Then
Key = Int(12 * Rnd())
If Key = 9 Then Key = 16
If Key = 10 Then Key = 23
If Key = 11 Then Key = 24
If Key = 8 Then Key = 38

ElseIf Text2.Text = "靴子" Then
Key = Int(10 * Rnd())
If Key = 9 Then Key = 17
If Key = 8 Then Key = 25

End If

If Key > -1 And Key < 7 Then Key = Key + 9
If Key = 7 Then Key = 19



End If




Call Command3_Click(Int(Key))
lvl = Int((extra * 0.5 * Rnd() + 1) / mult)
If mofaquan(Key) >= 4 Then lvl = 1 '1，2，4，5，6
If lvl >= 256 Then lvl = 255 '1，2，4，5，6
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
List1.AddItem "附魔名     尝试等级    等级上限   最终附魔等级 "


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
List4.AddItem "神帝宝库中的宝具："
List4.AddItem "编号 品阶 种类 名称                      来源        残语者"
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
If d(1) = 0 Then Command5.Caption = "可以摧毁"
If d(1) = 1 Then Command5.Caption = "不可摧毁"
End Sub

Private Sub Command6_Click(Index As Integer)
Text1.Text = Command6(Index).Caption
Text35(0).Text = Text1.Text + Text2.Text
End Sub




Private Sub Command7_Click()
bukechongfu = Not bukechongfu
If bukechongfu = True Then Command7.Caption = "不可重复"
If bukechongfu = False Then Command7.Caption = "可以重复"
End Sub

Private Sub Command8_Click(Index As Integer)
Text2.Text = Command8(Index).Caption
Text35(0).Text = Text1.Text + Text2.Text
End Sub

Private Sub Command9_Click()
mcfun = Not mcfun
If mcfun = False Then
Command9.Caption = "输出json文本"
Text13.Text = "trade"
Text14.Text = "D:\VB\MC—程序\实体构造\物品目录"
End If
If mcfun = True Then
Command9.Caption = "输出mcfunction并输出文本框"
Text13.Text = "magic"
Text14.Text = "D:\MC\.minecraft\saves\" + Command12(Index).Caption + "\datapacks\魔法构造\data\give\functions"
End If
End Sub

Private Sub Form_Load()
pinjiming(0) = "普通": pinjicolor(0) = RGB(150, 150, 150)
pinjiming(1) = "精良": pinjicolor(1) = RGB(0, 200, 0)
pinjiming(2) = "稀有": pinjicolor(2) = RGB(0, 0, 200)
pinjiming(3) = "史诗": pinjicolor(3) = RGB(200, 0, 250)
pinjiming(4) = "传说": pinjicolor(4) = RGB(250, 150, 0)
pinjiming(5) = "神话": pinjicolor(5) = RGB(200, 30, 0)
pinjiming(6) = "不朽": pinjicolor(6) = RGB(255, 200, 0)
pinjiming(7) = "永恒": pinjicolor(7) = RGB(0, 250, 250)
pinjiming(8) = "无上": pinjicolor(8) = RGB(250, 230, 0)

For i = 0 To 8
Command32(i).Caption = pinjiming(i)
Next i

laiyuan(0, 0) = "创世之时": laiyuan(0, 1) = "西方时间": laiyuan(1, 0) = "巫妖量劫时": laiyuan(1, 1) = "东方时间"
laiyuan(2, 0) = "第一纪元": laiyuan(2, 1) = "西方时间": laiyuan(3, 0) = "洪荒纪元": laiyuan(3, 1) = "东方时间"
laiyuan(4, 0) = "第二纪元": laiyuan(4, 1) = "西方时间": laiyuan(5, 0) = "仙古纪元": laiyuan(5, 1) = "东方时间"
laiyuan(6, 0) = "远古时代": laiyuan(6, 1) = "西方时间": laiyuan(7, 0) = "上古之时": laiyuan(7, 1) = "东方时间"
laiyuan(8, 0) = "大文艺时代": laiyuan(6, 1) = "西方时间": laiyuan(9, 0) = "大魔法时代": laiyuan(9, 1) = "西方时间"
laiyuan(10, 0) = "神秘时代": laiyuan(10, 1) = "西方时间": laiyuan(11, 0) = "末法时期": laiyuan(11, 1) = "东方时间"

laiyuan(12, 0) = "神弃之地": laiyuan(12, 1) = "西方地点": laiyuan(13, 0) = "四方人间": laiyuan(13, 1) = "东方地点"
laiyuan(14, 0) = "深渊之底": laiyuan(14, 1) = "西方地点": laiyuan(15, 0) = "八重妖域": laiyuan(15, 1) = "东方地点"
laiyuan(16, 0) = "遥远群星": laiyuan(16, 1) = "西方地点": laiyuan(17, 0) = "远古天庭": laiyuan(17, 1) = "东方地点"
laiyuan(18, 0) = "地狱深处": laiyuan(18, 1) = "西方地点": laiyuan(19, 0) = "十层魔海": laiyuan(19, 1) = "东方地点"

For i = 0 To 19
Command35(i).Caption = laiyuan(i, 0)
Next i
Call Command35_Click(13)


Text8(11).Text = """minecraft:instant_health"""

Call HScroll4_Change


Text38.Text = "5"
'普通五级的1，3级的2，1级的4，经验修补无限5
 '0-2锋利,3火焰 4横扫 5抢夺 18击退 9耐久，10修补，11荆棘
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

mofaquan(20) = 1 '20-22力无火，26冲击
mofaquan(21) = 5
mofaquan(22) = 4
mofaquan(26) = 1

mofaquan(30) = 2 '30-32快穿多
mofaquan(31) = 2
mofaquan(32) = 4

mofaquan(27) = 4 '27-29忠激雷，34刺
mofaquan(28) = 5
mofaquan(29) = 5
mofaquan(34) = 1

mofaquan(6) = 4  '6-8准时效
mofaquan(7) = 2
mofaquan(8) = 1



'12-15保护，19弹射保护 16，23，24，38海灵冰潜，靴子 ，17水下速，25水下呼吸头盔
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

Text14.Text = "D:\MC\1.20\.minecraft\saves\§dOlostland 1.0\datapacks\魔法构造\data\give\functions"

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

Text35(1).Text = "来自"
Text35(2).Text = "的遗物"

Randomize
Text1.Text = "铁"
Text2.Text = "剑"
Text3.Text = "3"
Text4.Text = "荆棘"
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
Command9.Caption = "输出mcfunction并输出文本框"
shuxingkaiqi = True

bukechongfu = True
Call Command4_Click
'List1.AddItem "附魔名     尝试等级    等级上限   最终附魔等级 "
Call Command11_Click
'List2.AddItem "形状       轨迹       闪烁       颜色       重数"
Call Command21_Click
'List3.AddItem "ID         倍率       时长       蓝光       粒子       图标"
k = 0
End Sub

Private Function bu(x As String) As String '列表对齐函数
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
If x = "剑" Then wuzhi = "sword"
If x = "弓" Then wuzhi = "bow"
If x = "弩" Then wuzhi = "crossbow"
If x = "锄" Then wuzhi = "hoe"
If x = "烟花火箭" Then wuzhi = "firework_rocket"
If x = "头盔" Then wuzhi = "helmet"
If x = "帽子" Then wuzhi = "helmet"
If x = "胸甲" Then wuzhi = "chestplate"
If x = "铠甲" Then wuzhi = "chestplate"
If x = "护腿" Then wuzhi = "leggings"
If x = "靴子" Then wuzhi = "boots"
If x = "鞋子" Then wuzhi = "boots"
If x = "斧" Then wuzhi = "axe"
If x = "镐" Then wuzhi = "pickaxe"
If x = "锹" Then wuzhi = "shovel"
If x = "三叉戟" Then wuzhi = "trident"
If x = "鞘翅" Then wuzhi = "elytra"
If x = "钓鱼竿" Then wuzhi = "fishing_rod"
If x = "盾牌" Then wuzhi = "shield"
If x = "箭" Then wuzhi = "tipped_arrow"
If x = "药水" Then wuzhi = "potion"
If x = "生成蛋" Then wuzhi = "spawn_egg"
If x = "书" Then wuzhi = "book"
If x = "马鞍" Then wuzhi = "saddle"
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

Public Function sjz(xx) '---------------------------------------转成十进制--------------------------------------
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
Label20.Caption = "尝试抽取品阶下限为" + pinjiming(i)
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
'品阶选取
chou = HScroll4.Value + Int((7 - HScroll4.Value) * Rnd())
Call Command32_Click(chou)

'时代地点选取
chou = Int(20 * Rnd())
If chou Mod 2 = 0 Then chou = Int(20 * Rnd())
If chou Mod 2 = 0 Then chou = Int(20 * Rnd())
Call Command35_Click(chou)

'物品选取
chou = Int(20 * Rnd())
Call Command8_Click(chou)

'材质选取
chou = Int(8 * Rnd())
If Text2.Text = "剑" Or Text2.Text = "锹" Or Text2.Text = "锄" Or Text2.Text = "镐" Or Text2.Text = "斧" Then
Do While chou = 4 Or chou = 5
chou = Int(8 * Rnd())
Loop
ElseIf Text2.Text = "头盔" Or Text2.Text = "胸甲" Or Text2.Text = "护腿" Or Text2.Text = "靴子" Then
Do While chou = 6 Or chou = 7
chou = Int(8 * Rnd())
Loop
ElseIf Text2.Text = "生成蛋" Then
chou = 11
Else
chou = 8
End If
Call Command6_Click(chou)

'生成本源属性
If Text2.Text = "剑" Or Text2.Text = "锹" Or Text2.Text = "锄" Or Text2.Text = "镐" Or Text2.Text = "斧" Or Text2.Text = "弓" Or Text2.Text = "弩" Or Text2.Text = "三叉戟" Or Text2.Text = "钓鱼竿" Or Text2.Text = "盾牌" Then
Call Command39_Click(0)
ElseIf Text2.Text = "头盔" Or Text2.Text = "胸甲" Or Text2.Text = "护腿" Or Text2.Text = "靴子" Or Text2.Text = "鞘翅" Or Text2.Text = "马鞍" Or Text2.Text = "烟花火箭" Or Text2.Text = "箭" Then
Call Command39_Click(1)
End If

Call Command36_Click '命名

Call Command1_Click '锻造
Call Command40_Click '入库

choujiangcishu = choujiangcishu + 1
If choujiangcishu >= Val(Text38.Text) Then Timer1.Enabled = False
End Sub
