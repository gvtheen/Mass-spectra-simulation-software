VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form ElementTable 
   AutoRedraw      =   -1  'True
   Caption         =   "IOP Mass Simulation"
   ClientHeight    =   6315
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7590
   FillColor       =   &H00C0C0FF&
   ForeColor       =   &H8000000B&
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   7590
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox SavPicture 
      Height          =   375
      Left            =   6720
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   118
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Element_Table"
      Height          =   5895
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin MSCommLib.MSComm MSComm1 
         Left            =   240
         Top             =   5400
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0FF&
         Height          =   495
         Index           =   110
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   136
         Top             =   4800
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Height          =   495
         Index           =   109
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   135
         Top             =   4320
         Width           =   375
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H80000009&
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   134
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         Caption         =   "放射性元素"
         Height          =   375
         Index           =   6
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   117
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "碱土金属"
         Height          =   375
         Index           =   5
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   116
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         Caption         =   "非金属"
         Height          =   375
         Index           =   4
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   115
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   "主族金属"
         Height          =   375
         Index           =   3
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   114
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "碱金属"
         Height          =   375
         Index           =   2
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   113
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "过渡金属"
         Height          =   375
         Index           =   1
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   112
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "稀有气体"
         Height          =   375
         Index           =   0
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   111
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0FF&
         Height          =   495
         Index           =   108
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   108
         Top             =   3600
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0FF&
         Height          =   495
         Index           =   89
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   4800
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0FF&
         Height          =   495
         Index           =   107
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   109
         Top             =   3600
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0FF&
         Height          =   495
         Index           =   106
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   107
         Top             =   3600
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0FF&
         Height          =   495
         Index           =   105
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   106
         Top             =   3600
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0FF&
         Height          =   495
         Index           =   104
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   105
         Top             =   3600
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0FF&
         Height          =   495
         Index           =   103
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   104
         Top             =   3600
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0FF&
         Height          =   495
         Index           =   102
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   103
         Top             =   4800
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0FF&
         Height          =   495
         Index           =   101
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   102
         Top             =   4800
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0FF&
         Height          =   495
         Index           =   100
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   101
         Top             =   4800
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0FF&
         Height          =   495
         Index           =   99
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   100
         Top             =   4800
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0FF&
         Height          =   495
         Index           =   98
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   4800
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0FF&
         Height          =   495
         Index           =   97
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   4800
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0FF&
         Height          =   495
         Index           =   96
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   97
         Top             =   4800
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0FF&
         Height          =   495
         Index           =   95
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   96
         Top             =   4800
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0FF&
         Height          =   495
         Index           =   94
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   95
         Top             =   4800
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0FF&
         Height          =   495
         Index           =   93
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   4800
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0FF&
         Height          =   495
         Index           =   92
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   4800
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0FF&
         Height          =   495
         Index           =   91
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   4800
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0FF&
         Height          =   495
         Index           =   90
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   4800
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0FF&
         Height          =   495
         Index           =   88
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   3600
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0FF&
         Height          =   495
         Index           =   87
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   3600
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0FF&
         Height          =   495
         Index           =   86
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   3600
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0FF&
         Height          =   495
         Index           =   85
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   3120
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0FF&
         Height          =   495
         Index           =   84
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   3120
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0FF&
         Height          =   495
         Index           =   83
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   3120
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080C0FF&
         Height          =   495
         Index           =   82
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   3120
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080C0FF&
         Height          =   495
         Index           =   81
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   3120
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080C0FF&
         Height          =   495
         Index           =   80
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   3120
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Height          =   495
         Index           =   79
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   3120
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Height          =   495
         Index           =   78
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   3120
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Height          =   495
         Index           =   77
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   3120
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Height          =   495
         Index           =   76
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   3120
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Height          =   495
         Index           =   75
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   3120
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Height          =   495
         Index           =   74
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   3120
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Height          =   495
         Index           =   73
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   3120
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Height          =   495
         Index           =   72
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   3120
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Height          =   495
         Index           =   71
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   3120
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Height          =   495
         Index           =   70
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   4320
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Height          =   495
         Index           =   69
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   4320
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Height          =   495
         Index           =   68
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   4320
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Height          =   495
         Index           =   67
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   4320
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Height          =   495
         Index           =   66
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   4320
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Height          =   495
         Index           =   65
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   4320
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Height          =   495
         Index           =   64
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   4320
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Height          =   495
         Index           =   63
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   4320
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Height          =   495
         Index           =   62
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   4320
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0FF&
         Height          =   495
         Index           =   61
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   4320
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Height          =   495
         Index           =   60
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   4320
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Height          =   495
         Index           =   59
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   4320
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Height          =   495
         Index           =   58
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   4320
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Height          =   495
         Index           =   57
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   4320
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Height          =   495
         Index           =   56
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   3120
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFC0&
         Height          =   495
         Index           =   55
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   3120
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF8080&
         Height          =   495
         Index           =   54
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   3120
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00808080&
         Height          =   495
         Index           =   53
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   2640
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Height          =   495
         Index           =   52
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   2640
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Height          =   495
         Index           =   51
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   2640
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080C0FF&
         Height          =   495
         Index           =   50
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   2640
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080C0FF&
         Height          =   495
         Index           =   49
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   2640
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080C0FF&
         Height          =   495
         Index           =   48
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   2640
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Height          =   495
         Index           =   47
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   2640
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Height          =   495
         Index           =   46
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   2640
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Height          =   495
         Index           =   45
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   2640
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Height          =   495
         Index           =   44
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   2640
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Height          =   495
         Index           =   43
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   2640
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0FF&
         Height          =   495
         Index           =   42
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   2640
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Height          =   495
         Index           =   41
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   2640
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Height          =   495
         Index           =   40
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   2640
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Height          =   495
         Index           =   39
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   2640
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Height          =   495
         Index           =   38
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   2640
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFC0&
         Height          =   495
         Index           =   37
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   2640
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF8080&
         Height          =   495
         Index           =   36
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   2640
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00808080&
         Height          =   495
         Index           =   35
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   2160
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Height          =   495
         Index           =   34
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   2160
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Height          =   495
         Index           =   33
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   2160
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Height          =   495
         Index           =   32
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   2160
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080C0FF&
         Height          =   495
         Index           =   31
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   2160
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080C0FF&
         Height          =   495
         Index           =   30
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   2160
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Height          =   495
         Index           =   29
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   2160
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Height          =   495
         Index           =   28
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2160
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Height          =   495
         Index           =   27
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   2160
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Height          =   495
         Index           =   26
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   2160
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Height          =   495
         Index           =   25
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2160
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Height          =   495
         Index           =   24
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   2160
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Height          =   495
         Index           =   23
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   2160
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Height          =   495
         Index           =   22
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   2160
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Height          =   495
         Index           =   21
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   2160
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Height          =   495
         Index           =   20
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2160
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFC0&
         Height          =   495
         Index           =   19
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2160
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF8080&
         Height          =   495
         Index           =   18
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2160
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00808080&
         Height          =   495
         Index           =   17
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1680
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Height          =   495
         Index           =   16
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1680
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Height          =   495
         Index           =   15
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1680
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Height          =   495
         Index           =   14
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1680
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Height          =   495
         Index           =   13
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1680
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080C0FF&
         Height          =   495
         Index           =   12
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1680
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFC0&
         Height          =   495
         Index           =   11
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1680
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF8080&
         Height          =   495
         Index           =   10
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1680
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00808080&
         Height          =   495
         Index           =   9
         Left            =   6240
         MaskColor       =   &H00C0C0FF&
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1200
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Height          =   495
         Index           =   8
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1200
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Height          =   495
         Index           =   7
         Left            =   5520
         MaskColor       =   &H0080C0FF&
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1200
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Height          =   495
         Index           =   6
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1200
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Height          =   495
         Index           =   5
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1200
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Height          =   495
         Index           =   4
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1200
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFC0&
         Height          =   495
         Index           =   3
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1200
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF8080&
         Height          =   495
         Index           =   2
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1200
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00808080&
         Height          =   495
         Index           =   1
         Left            =   6240
         MaskColor       =   &H000000C0&
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Height          =   495
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Caption         =   " 锕 系"
         Height          =   495
         Index           =   17
         Left            =   480
         TabIndex        =   139
         Top             =   4800
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   " 镧 系"
         Height          =   495
         Index           =   16
         Left            =   480
         TabIndex        =   138
         Top             =   4320
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "  IIA"
         Height          =   495
         Index           =   1
         Left            =   480
         TabIndex        =   137
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "    O"
         Height          =   495
         Index           =   15
         Left            =   6240
         TabIndex        =   133
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "   VIIA"
         Height          =   495
         Index           =   14
         Left            =   5880
         TabIndex        =   132
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "   VIA"
         Height          =   495
         Index           =   13
         Left            =   5520
         TabIndex        =   131
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "   VA"
         Height          =   495
         Index           =   12
         Left            =   5160
         TabIndex        =   130
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "   IVA"
         Height          =   495
         Index           =   11
         Left            =   4800
         TabIndex        =   129
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "   III"
         Height          =   495
         Index           =   10
         Left            =   4440
         TabIndex        =   128
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "  IIB"
         Height          =   495
         Index           =   9
         Left            =   4080
         TabIndex        =   127
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "   IB"
         Height          =   495
         Index           =   8
         Left            =   3720
         TabIndex        =   126
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Caption         =   "          IVB"
         Height          =   495
         Index           =   7
         Left            =   2640
         TabIndex        =   125
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "  VIIB"
         Height          =   495
         Index           =   6
         Left            =   2280
         TabIndex        =   124
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "  VIB"
         Height          =   495
         Index           =   5
         Left            =   1920
         TabIndex        =   123
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "   VB"
         Height          =   495
         Index           =   4
         Left            =   1560
         TabIndex        =   122
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "  IVB"
         Height          =   495
         Index           =   3
         Left            =   1200
         TabIndex        =   121
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   " IIIB"
         Height          =   495
         Index           =   2
         Left            =   840
         TabIndex        =   120
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "   IA"
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   119
         Top             =   240
         Width           =   375
      End
      Begin VB.Label ShowChemForm 
         BackColor       =   &H80000009&
         Height          =   495
         Left            =   840
         TabIndex        =   110
         Top             =   1200
         Width           =   3615
      End
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Save 
         Caption         =   "SavePeriod"
         Shortcut        =   ^S
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "Edit"
      Begin VB.Menu Formula_Edit 
         Caption         =   "Formula.."
         Shortcut        =   ^F
      End
      Begin VB.Menu Sound 
         Caption         =   "Sound"
         Begin VB.Menu TurnOn 
            Caption         =   "Turn On  "
         End
         Begin VB.Menu TurnOff 
            Caption         =   "Turn Off   Y"
         End
      End
      Begin VB.Menu Delete 
         Caption         =   "Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu Chart 
         Caption         =   "FindMaxAbun"
      End
   End
   Begin VB.Menu Calculation 
      Caption         =   "Calcul.."
      Begin VB.Menu DisPlay 
         Caption         =   "DisPlay"
         Shortcut        =   ^D
      End
      Begin VB.Menu Distribution 
         Caption         =   "Distri.."
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu GUANYU 
      Caption         =   "About"
      Begin VB.Menu About 
         Caption         =   "About"
         Shortcut        =   ^A
      End
      Begin VB.Menu Help 
         Caption         =   "Help"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "ElementTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Showbool As Boolean
Dim SoundBool As Boolean


  Private Type PALETTEENTRY
        peRed As Byte
        peGreen As Byte
        peBlue As Byte
        peFlags As Byte
    End Type
   
    Private Type LOGPALETTE
        palVersion As Integer
        palNumEntries As Integer
        palPalEntry(255) As PALETTEENTRY
    End Type

    Private Type GUID
      Data1 As Long
      Data2 As Integer
      Data3 As Integer
      Data4(7) As Byte
    End Type

    Private Const RASTERCAPS As Long = 38
    Private Const RC_PALETTE As Long = &H100
    Private Const SIZEPALETTE As Long = 104

    Private Type RECT
       Left As Long
       Top As Long
       Right As Long
       Bottom As Long
    End Type
    
    Private Type PicBmp
      Size As Long
      Type As Long
      hBmp As Long
      hPal As Long
      Reserved As Long
    End Type
    
 Private Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hDC As Long) As Long
 
 Private Declare Function CreateCompatibleBitmap Lib "GDI32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
 
 Private Declare Function GetDeviceCaps Lib "GDI32" (ByVal hDC As Long, ByVal iCapabilitiy As Long) As Long
 
 Private Declare Function GetSystemPaletteEntries Lib "GDI32" (ByVal hDC As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
 
 Private Declare Function CreatePalette Lib "GDI32" (lpLogPalette As LOGPALETTE) As Long

 Private Declare Function SelectObject Lib "GDI32" (ByVal hDC As Long, ByVal hObject As Long) As Long

 Private Declare Function BitBlt Lib "GDI32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long

 Private Declare Function DeleteDC Lib "GDI32" (ByVal hDC As Long) As Long

 Private Declare Function GetForegroundWindow Lib "USER32" () As Long

 Private Declare Function SelectPalette Lib "GDI32" (ByVal hDC As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long

 Private Declare Function RealizePalette Lib "GDI32" (ByVal hDC As Long) As Long

 Private Declare Function GetWindowDC Lib "USER32" (ByVal hWnd As Long) As Long

 Private Declare Function GetDC Lib "USER32" (ByVal hWnd As Long) As Long

 Private Declare Function GetWindowRect Lib "USER32" (ByVal hWnd As Long, lpRect As RECT) As Long
 
 Private Declare Function ReleaseDC Lib "USER32" (ByVal hWnd As Long, ByVal hDC As Long) As Long

 Private Declare Function GetDesktopWindow Lib "USER32" () As Long
 
 Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture

    Dim r As Long

    Dim Pic As PicBmp

    Dim IPic As IPicture

    Dim IID_IDispatch As GUID

    '填充IDispatch界面

    With IID_IDispatch

    .Data1 = &H20400

    .Data4(0) = &HC0

    .Data4(7) = &H46

    End With

    '填充Pic

    With Pic

    .Size = Len(Pic)

    ' Pic结构长度

    .Type = vbPicTypeBitmap

    ' 图像类型

    .hBmp = hBmp

    ' 位图句柄

    .hPal = hPal

    ' 调色板句柄

    End With

    '建立Picture图像

    r = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)

    '返回Picture对象

    Set CreateBitmapPicture = IPic

    End Function
Private Function CaptureWindow(ByVal hWndSrc As Long, ByVal Client As Boolean, ByVal LeftSrc As Long, ByVal TopSrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long) As Picture

    Dim hDCMemory As Long

    Dim hBmp As Long

    Dim hBmpPrev As Long

    Dim r As Long

    Dim hDCSrc As Long

    Dim hPal As Long

    Dim hPalPrev As Long

    Dim RasterCapsScrn As Long

    Dim HasPaletteScrn As Long

    Dim PaletteSizeScrn As Long

    Dim LogPal As LOGPALETTE

    If Client Then

    hDCSrc = GetDC(hWndSrc)

    Else

    hDCSrc = GetWindowDC(hWndSrc)

    End If

    hDCMemory = CreateCompatibleDC(hDCSrc)

    hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)

    hBmpPrev = SelectObject(hDCMemory, hBmp)

    '获得屏幕属性

    RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS)

    HasPaletteScrn = RasterCapsScrn And RC_PALETTE

    PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE)

    '如果屏幕对象有调色板则获得屏幕调色板

    If HasPaletteScrn And (PaletteSizeScrn = 256) Then

    '建立屏幕调色板的拷贝

    LogPal.palVersion = &H300

    LogPal.palNumEntries = 256

    r = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))

    hPal = CreatePalette(LogPal)

    '将新建立的调色板选如建立的内存绘图句柄中

    hPalPrev = SelectPalette(hDCMemory, hPal, 0)

    r = RealizePalette(hDCMemory)

    End If

    '拷贝图像

    r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, vbSrcCopy)

    hBmp = SelectObject(hDCMemory, hBmpPrev)

    If HasPaletteScrn And (PaletteSizeScrn = 256) Then

    hPal = SelectPalette(hDCMemory, hPalPrev, 0)

    End If

    '释放资源

    r = DeleteDC(hDCMemory)
    r = ReleaseDC(hWndSrc, hDCSrc)
    Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
    End Function
    'capturescreen函数捕捉整个屏幕图像
    Public Function CaptureScreen() As Picture
    Dim hWndScreen As Long
    '获得桌面的窗口句柄
    
    Set CaptureScreen = CaptureWindow(ElementTable.hWnd, True, 0, 0, ElementTable.Width \ Screen.TwipsPerPixelX, ElementTable.Height \ Screen.TwipsPerPixelY)
    End Function
Private Function CaptureActiveWindow() As Picture
    Dim hWndActive As Long
    Dim r As Long
    Dim RectActive As RECT
    hWndActive = GetForegroundWindow()
    r = GetWindowRect(hWndActive, RectActive)
    Set CaptureActiveWindow = CaptureWindow(hWndActive, False, 0, 0, _
    RectActive.Right - RectActive.Left, RectActive.Bottom - RectActive.Top)
    End Function
Private Sub About_Click()
AboutForm.Show
End Sub

Private Sub Chart_Click()
Caculate.Show
End Sub

Private Sub Command1_Click(Index As Integer)
    On Error GoTo H1
    If SoundBool Then
       Dim SoundFile As String
       Static AuPlayer1 As Object
       Select Case Index
         Case 109
             SoundFile = App.Path & "\ElementSound" & "\" & Trim(Command1(55).Caption) & ".au"
         Case 110
             SoundFile = App.Path & "\ElementSound" & "\" & Trim(Command1(87).Caption) & ".au"
         Case Else
             SoundFile = App.Path & "\ElementSound" & "\" & Trim(Command1(Index).Caption) & ".au"
       End Select
       Set AuPlayer1 = New FilgraphManager
       AuPlayer1.RenderFile SoundFile
       AuPlayer1.Run
    End If
    
    InputForm.Show
    If ShowChemFormBool Then ShowChemForm.Caption = "": ShowChemFormBool = False: Call ResetElemTable
    If Index = 109 Then Call InputForm.SetIOPData(55): Exit Sub
    If Index = 110 Then Call InputForm.SetIOPData(87): Exit Sub
    Call InputForm.SetIOPData(Index)
    Exit Sub
    
H1:     MsgBox Error.Description
End Sub
Private Sub Command3_Click()
    On Error GoTo h2
    Dim pat As String
    pat = App.Path & "\HELP.CHM"
    Shell "hh.exe " & pat, vbNormalFocus
    Exit Sub
h2:     MsgBox Error.Description
End Sub
Private Sub Delete_Click()
ShowChemForm.Caption = ""
Call ResetElemTable
End Sub
Public Sub DisPlay_Click()
Dim str As String

str = Trim(ShowChemForm.Caption)
MocularForm = str
ShowChemFormBool = True
If str <> "" And StrComp(str, SaveLast_Chem) <> 0 Then Call DealInput(str): SaveLast_Chem = str: ElementTable.Hide: Exit Sub

If StrComp(str, SaveLast_Chem) = 0 And str <> "" Then isDealWithData = False: DisForm.Show: DisForm.Repeat_Paint: ElementTable.Hide: Exit Sub
End Sub
Private Sub Initiate()
FormAtomNum = 0
IOPValueBool = False
Showbool = True
isDealWithData = True
End Sub
Public Sub Exit_Click()
Unload Me
End Sub
Private Sub Form_Initialize()
On Error GoTo H1
Save.Enabled = False
'Set SavPicture.Picture = CaptureScreen()
SoundBool = False
H1: Exit Sub
End Sub
Private Sub Form_Load()
    On Error GoTo h3
    Dim fil As String, filnum As Integer, i As Integer, EleName As String, num1 As Integer, striName As String, filpath As String, num As Integer
    num1 = 0: num = 0
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    filpath = App.Path & "\TABLE.DAT"
    If Not fso.FileExists(filpath) Then MsgBox "Sorry,缺少TABLE.DAT文件，请在软件包中将其拷到安装目录下", vbInformation + vbOKOnly, "消息提示":  Unload Me: Exit Sub
    filpath = App.Path & "\ElementSound"
    If Not fso.FolderExists(filpath) Then MsgBox "Sorry,缺少ElementSound文件夹，请在软件包中将其拷到安装目录下", vbInformation + vbOKOnly, "消息提示":  Unload Me: Exit Sub
    filnum = FreeFile
    filpath = App.Path & "\TABLE.DAT"
    Open filpath For Input As #filnum
      Input #filnum, num
      ReDim TotleElemTable(num)
      Do Until EOF(filnum)
           Input #filnum, striName
           Command1(num1).Caption = striName
           If num1 = 56 Then Command1(109).Caption = striName
           If num1 = 88 Then Command1(110).Caption = striName
           TotleElemTable(num1).ElementShorName = striName
           Input #filnum, TotleElemTable(num1).ElementName
           Input #filnum, TotleElemTable(num1).IopNum
           ReDim TotleElemTable(num1).ElementIopData(TotleElemTable(num1).IopNum, 2)
           For i = 0 To TotleElemTable(num1).IopNum - 1
               Input #filnum, TotleElemTable(num1).ElementIopData(i, 0), TotleElemTable(num1).ElementIopData(i, 1)
           Next i
           num1 = num1 + 1
      Loop
    Close #filnum
    
    filpath = App.Path & "\IOPMassSimulation Data"
    If Not fso.FolderExists(filpath) Then fso.CreateFolder (filpath)
    
    
    Distribution.Enabled = False
    
    
    Showbool = True
    ShowChemFormBool = False
    Exit Sub
h3:     MsgBox Error.Description
End Sub
Public Sub ShowElement(nn As Integer, mm As Integer)

ShowChemForm.Caption = ShowChemForm.Caption & " " & Command1(mm).Caption & nn

End Sub
Private Sub ResetElemTable()
Dim fil As String, filnum As Integer, i As Integer, EleName As String, num1 As Integer, striName As String, filpath As String, num As Integer
num1 = 0: num = 0
filpath = App.Path & "\TABLE.DAT"
filnum = FreeFile

Open filpath For Input As #filnum
  Input #filnum, num
  Do Until EOF(filnum)
       Input #filnum, striName
       Command1(num1).Caption = striName
       TotleElemTable(num1).ElementShorName = striName
       Input #filnum, TotleElemTable(num1).ElementName
       Input #filnum, TotleElemTable(num1).IopNum
       ReDim TotleElemTable(num1).ElementIopData(TotleElemTable(num1).IopNum, 2)
       For i = 0 To TotleElemTable(num1).IopNum - 1
           Input #filnum, TotleElemTable(num1).ElementIopData(i, 0), TotleElemTable(num1).ElementIopData(i, 1)
       Next i
       num1 = num1 + 1
  Loop
Close #filnum

End Sub
Public Sub DealInput(mystr As String)
Dim i As Integer, length As Integer, GetStr As String, sum As Integer

If mystr = "" Then Exit Sub
length = Len(Trim(mystr))

Call Initiate

For i = 1 To length
    If Mid(mystr, i, 1) <> " " Then sum = sum + 1
Next i
sum = sum + 1

ReDim ProcDataElem(sum) As ProcDataType             '输入的分子式中原子个数的计算

For i = 1 To length
    If Mid(mystr, i, 1) = " " Then Call procStr(Trim(GetStr)): GetStr = ""
       GetStr = GetStr & Mid(mystr, i, 1)
    If i = length Then Call procStr(Trim(GetStr))
Next i

If Showbool Then Call DisForm.Show:  DisForm.WindowState = 0

End Sub
Private Sub procStr(mm As String)
    Dim leng As Integer, i As Integer, nn As String, str As String
    mm = Trim(mm)
    leng = Len(mm): nn = ""
    For i = 1 To leng
        If Asc(Mid(mm, i, 1)) >= 48 And Asc(Mid(mm, i, 1)) <= 57 Then Exit For
        nn = nn & Mid(mm, i, 1)
    Next i
    str = Mid(mm, Len(Trim(nn)) + 1, leng)
    
    If FindElement(nn, Val(str)) = False Then ShowChemForm.Caption = "": MsgBox "The Wrong Formula:" & nn, vbInformation + vbCritical, "Message": Showbool = False: ShutFormuFormBool = False
           
End Sub
Private Function FindElement(ss As String, mm1 As Integer) As Boolean
Dim i As Integer, j As Integer
For i = 0 To 108
    If StrComp(ss, Trim(Command1(i).Caption), 1) = 0 Then
       FindElement = True
       ProcDataElem(FormAtomNum).PointNum = i
       ProcDataElem(FormAtomNum).num = mm1
       FormAtomNum = FormAtomNum + 1
       Exit Function
    End If
Next i
FindElement = False
End Function
Private Sub Formula_Edit_Click()
FormuForm.Show
End Sub

Private Sub Help_Click()
Dim pat As String
pat = App.Path & "\ElementPeriodTable.CHM"
Shell "hh.exe " & pat, vbNormalFocus
End Sub


Private Sub Save_Click()
Set SavPicture.Picture = CaptureScreen()
    CommonDialog1.DefaultExt = ".BMP"
    CommonDialog1.Filter = "Bitmap Image (＊.bmp)|＊.bmp"
    CommonDialog1.InitDir = App.Path
    CommonDialog1.ShowSave
    If CommonDialog1.FileName <> "" Then
        SavePicture SavPicture.Picture, CommonDialog1.FileName
    End If
Set SavPicture.Picture = Nothing    '清楚内存空间
End Sub
Private Sub TurnOff_Click()
SoundBool = False
TurnOn.Caption = "Turn On"
TurnOff.Caption = "Turn Off Y"
End Sub

Private Sub TurnOn_Click()
SoundBool = True
TurnOn.Caption = "Turn On Y"
TurnOff.Caption = "Turn Off"

End Sub
