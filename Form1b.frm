VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Matrix Inverse and Transpose Calculator"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   60
      TabIndex        =   114
      Top             =   1440
      Width           =   1455
      Begin VB.OptionButton optMMULT 
         Caption         =   "[A]*[A-1] = [I]"
         Height          =   255
         Left            =   90
         TabIndex        =   117
         Top             =   690
         Width           =   1305
      End
      Begin VB.OptionButton optINVERSE 
         Caption         =   "Inverse"
         Height          =   255
         Left            =   90
         TabIndex        =   116
         Top             =   210
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.OptionButton optTRANSPOSE 
         Caption         =   "Transpose"
         Height          =   255
         Left            =   90
         TabIndex        =   115
         Top             =   450
         Width           =   1155
      End
   End
   Begin VB.TextBox Text3 
      Height          =   1845
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   113
      Text            =   "Form1b.frx":0000
      Top             =   3930
      Width           =   7485
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   450
      Style           =   2  'Dropdown List
      TabIndex        =   111
      Top             =   630
      Width           =   645
   End
   Begin VB.CommandButton cmdSOLVE 
      Caption         =   "Calculate -  Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   60
      TabIndex        =   110
      Top             =   2610
      Width           =   1425
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   99
      Left            =   7020
      TabIndex        =   99
      Top             =   3510
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   98
      Left            =   6420
      TabIndex        =   98
      Top             =   3510
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   97
      Left            =   5820
      TabIndex        =   97
      Top             =   3510
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   96
      Left            =   5220
      TabIndex        =   96
      Top             =   3510
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   95
      Left            =   4620
      TabIndex        =   95
      Top             =   3510
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   94
      Left            =   4020
      TabIndex        =   94
      Top             =   3510
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   93
      Left            =   3420
      TabIndex        =   93
      Top             =   3510
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   92
      Left            =   2820
      TabIndex        =   92
      Top             =   3510
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   91
      Left            =   2220
      TabIndex        =   91
      Top             =   3510
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   90
      Left            =   1620
      TabIndex        =   90
      Top             =   3510
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   89
      Left            =   7020
      TabIndex        =   89
      Top             =   3150
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   88
      Left            =   6420
      TabIndex        =   88
      Top             =   3150
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   87
      Left            =   5820
      TabIndex        =   87
      Top             =   3150
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   86
      Left            =   5220
      TabIndex        =   86
      Top             =   3150
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   85
      Left            =   4620
      TabIndex        =   85
      Top             =   3150
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   84
      Left            =   4020
      TabIndex        =   84
      Top             =   3150
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   83
      Left            =   3420
      TabIndex        =   83
      Top             =   3150
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   82
      Left            =   2820
      TabIndex        =   82
      Top             =   3150
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   81
      Left            =   2220
      TabIndex        =   81
      Top             =   3150
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   80
      Left            =   1620
      TabIndex        =   80
      Top             =   3150
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   79
      Left            =   7020
      TabIndex        =   79
      Top             =   2790
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   78
      Left            =   6420
      TabIndex        =   78
      Top             =   2790
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   77
      Left            =   5820
      TabIndex        =   77
      Top             =   2790
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   76
      Left            =   5220
      TabIndex        =   76
      Top             =   2790
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   75
      Left            =   4620
      TabIndex        =   75
      Top             =   2790
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   74
      Left            =   4020
      TabIndex        =   74
      Top             =   2790
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   73
      Left            =   3420
      TabIndex        =   73
      Top             =   2790
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   72
      Left            =   2820
      TabIndex        =   72
      Top             =   2790
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   71
      Left            =   2220
      TabIndex        =   71
      Top             =   2790
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   70
      Left            =   1620
      TabIndex        =   70
      Top             =   2790
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   69
      Left            =   7020
      TabIndex        =   69
      Top             =   2430
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   68
      Left            =   6420
      TabIndex        =   68
      Top             =   2430
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   67
      Left            =   5820
      TabIndex        =   67
      Top             =   2430
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   66
      Left            =   5220
      TabIndex        =   66
      Top             =   2430
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   65
      Left            =   4620
      TabIndex        =   65
      Top             =   2430
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   64
      Left            =   4020
      TabIndex        =   64
      Top             =   2430
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   63
      Left            =   3420
      TabIndex        =   63
      Top             =   2430
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   62
      Left            =   2820
      TabIndex        =   62
      Top             =   2430
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   61
      Left            =   2220
      TabIndex        =   61
      Top             =   2430
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   60
      Left            =   1620
      TabIndex        =   60
      Top             =   2430
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   59
      Left            =   7020
      TabIndex        =   59
      Top             =   2070
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   58
      Left            =   6420
      TabIndex        =   58
      Top             =   2070
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   57
      Left            =   5820
      TabIndex        =   57
      Top             =   2070
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   56
      Left            =   5220
      TabIndex        =   56
      Top             =   2070
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   55
      Left            =   4620
      TabIndex        =   55
      Top             =   2070
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   54
      Left            =   4020
      TabIndex        =   54
      Top             =   2070
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   53
      Left            =   3420
      TabIndex        =   53
      Top             =   2070
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   52
      Left            =   2820
      TabIndex        =   52
      Top             =   2070
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   51
      Left            =   2220
      TabIndex        =   51
      Top             =   2070
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   50
      Left            =   1620
      TabIndex        =   50
      Top             =   2070
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   49
      Left            =   7020
      TabIndex        =   49
      Top             =   1710
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   48
      Left            =   6420
      TabIndex        =   48
      Top             =   1710
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   47
      Left            =   5820
      TabIndex        =   47
      Top             =   1710
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   46
      Left            =   5220
      TabIndex        =   46
      Top             =   1710
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   45
      Left            =   4620
      TabIndex        =   45
      Top             =   1710
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   44
      Left            =   4020
      TabIndex        =   44
      Top             =   1710
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   43
      Left            =   3420
      TabIndex        =   43
      Top             =   1710
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   42
      Left            =   2820
      TabIndex        =   42
      Top             =   1710
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   41
      Left            =   2220
      TabIndex        =   41
      Top             =   1710
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   40
      Left            =   1620
      TabIndex        =   40
      Top             =   1710
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   39
      Left            =   7020
      TabIndex        =   39
      Top             =   1350
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   38
      Left            =   6420
      TabIndex        =   38
      Top             =   1350
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   37
      Left            =   5820
      TabIndex        =   37
      Top             =   1350
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   36
      Left            =   5220
      TabIndex        =   36
      Top             =   1350
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   35
      Left            =   4620
      TabIndex        =   35
      Top             =   1350
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   34
      Left            =   4020
      TabIndex        =   34
      Top             =   1350
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   33
      Left            =   3420
      TabIndex        =   33
      Text            =   "5"
      Top             =   1350
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   32
      Left            =   2820
      TabIndex        =   32
      Text            =   "1"
      Top             =   1350
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   31
      Left            =   2220
      TabIndex        =   31
      Text            =   "4"
      Top             =   1350
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   30
      Left            =   1620
      TabIndex        =   30
      Text            =   "2"
      Top             =   1350
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   29
      Left            =   7020
      TabIndex        =   29
      Top             =   990
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   28
      Left            =   6420
      TabIndex        =   28
      Top             =   990
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   27
      Left            =   5820
      TabIndex        =   27
      Top             =   990
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   26
      Left            =   5220
      TabIndex        =   26
      Top             =   990
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   25
      Left            =   4620
      TabIndex        =   25
      Top             =   990
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   24
      Left            =   4020
      TabIndex        =   24
      Top             =   990
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   23
      Left            =   3420
      TabIndex        =   23
      Text            =   "6"
      Top             =   990
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   22
      Left            =   2820
      TabIndex        =   22
      Text            =   "2"
      Top             =   990
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   21
      Left            =   2220
      TabIndex        =   21
      Text            =   "2"
      Top             =   990
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   20
      Left            =   1620
      TabIndex        =   20
      Text            =   "4"
      Top             =   990
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   19
      Left            =   7020
      TabIndex        =   19
      Top             =   630
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   18
      Left            =   6420
      TabIndex        =   18
      Top             =   630
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   17
      Left            =   5820
      TabIndex        =   17
      Top             =   630
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   16
      Left            =   5220
      TabIndex        =   16
      Top             =   630
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   15
      Left            =   4620
      TabIndex        =   15
      Top             =   630
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   14
      Left            =   4020
      TabIndex        =   14
      Top             =   630
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   13
      Left            =   3420
      TabIndex        =   13
      Text            =   "4"
      Top             =   630
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   12
      Left            =   2820
      TabIndex        =   12
      Text            =   "1"
      Top             =   630
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   11
      Left            =   2220
      TabIndex        =   11
      Text            =   "4"
      Top             =   630
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   10
      Left            =   1620
      TabIndex        =   10
      Text            =   "1"
      Top             =   630
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   9
      Left            =   7020
      TabIndex        =   9
      Top             =   270
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   8
      Left            =   6420
      TabIndex        =   8
      Top             =   270
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   7
      Left            =   5820
      TabIndex        =   7
      Top             =   270
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   5220
      TabIndex        =   6
      Top             =   270
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   4620
      TabIndex        =   5
      Top             =   270
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   4020
      TabIndex        =   4
      Top             =   270
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   3420
      TabIndex        =   3
      Text            =   "2"
      Top             =   270
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   2820
      TabIndex        =   2
      Text            =   "4"
      Top             =   270
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   2220
      TabIndex        =   1
      Text            =   "2"
      Top             =   270
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   1620
      TabIndex        =   0
      Text            =   "1"
      Top             =   270
      Width           =   555
   End
   Begin VB.Label Label3 
      Caption         =   "Matrix dimension:"
      Height          =   255
      Left            =   60
      TabIndex        =   112
      Top             =   210
      Width           =   1485
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "x10"
      Height          =   195
      Index           =   9
      Left            =   7020
      TabIndex        =   109
      Top             =   30
      Width           =   555
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "x9"
      Height          =   195
      Index           =   8
      Left            =   6420
      TabIndex        =   108
      Top             =   30
      Width           =   555
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "x8"
      Height          =   195
      Index           =   7
      Left            =   5820
      TabIndex        =   107
      Top             =   30
      Width           =   555
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "x7"
      Height          =   195
      Index           =   6
      Left            =   5220
      TabIndex        =   106
      Top             =   30
      Width           =   555
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "x6"
      Height          =   195
      Index           =   5
      Left            =   4620
      TabIndex        =   105
      Top             =   30
      Width           =   555
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "x5"
      Height          =   195
      Index           =   4
      Left            =   4020
      TabIndex        =   104
      Top             =   30
      Width           =   555
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "x4"
      Height          =   195
      Index           =   3
      Left            =   3420
      TabIndex        =   103
      Top             =   30
      Width           =   555
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "x3"
      Height          =   195
      Index           =   2
      Left            =   2820
      TabIndex        =   102
      Top             =   30
      Width           =   555
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "x2"
      Height          =   195
      Index           =   1
      Left            =   2220
      TabIndex        =   101
      Top             =   30
      Width           =   555
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "x1"
      Height          =   195
      Index           =   0
      Left            =   1620
      TabIndex        =   100
      Top             =   30
      Width           =   555
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Calculates the Inverse of a Rectangular Matrix [A] (Dimensions N x N) using the Gauss elimination method,
'the product [A]*[A-1] for verification purposes (must be always equal to Singular Matrix [I])
'and also the transpose of Matrix [A]. The interface is limited to 10x10 dimensions, but the solver itself can
'be used to calculate the Inverse of any Rectangular Matrix, provided the determinant of it is non-zero.

'v1.10: Fixed a bug in the Matrix Inverse calculation routine (which occured only in some special cases),
'where line k is changed with line_1

'(c) 2003, Vagelis Plevris, Greece
'mailto: vplevris@tee.gr

Private Sub Form_Load()
    'Give numbers 1-10 to combo1
    For N = 1 To 10
        Combo1.AddItem N
    Next N
    'Default = 4 (Matrix dimensions 4X4)
    Combo1.Text = 4
    Call Hide_Show_Textboxes
End Sub

Private Sub Combo1_Click()
    Call Hide_Show_Textboxes
End Sub

Private Sub cmdSOLVE_Click()
    Call Build_Matrix
    Call Calculate_Inverse
    Call Calculate_Transpose
    Call Calculate_MatrixMult
    Call Type_Result
End Sub

Sub Hide_Show_Textboxes()

System_DIM = Val(Combo1.Text) 'Matrix [A] dimensions
'Hide all textboxes
For N = 0 To 99 '0-99 for Matrix [A]
    Text1(N).Visible = False
Next N
'Show the appropriate textboxes for the Matrix dimensions
For N = 0 To System_DIM - 1
    For k = 0 To 10 * (System_DIM - 1) Step 10
        Text1(N + k).Visible = True 'matrix [A]
    Next k
Next N

End Sub

Sub Build_Matrix()
'Builds the [A] Matrix assigning values from the textboxes
'Matrix_A dimensions are set to max 10x10 for the interface needs, but can be increased to whatever

'Build Matrix_A
For N = 1 To System_DIM
    For m = 1 To System_DIM
        Matrix_A(N, m) = Val(Text1(m - 1 + (N - 1) * 10))
    Next m
Next N

End Sub

Sub Calculate_Inverse()
'Uses Gauss elimination method in order to calculate the inverse matrix [A]-1
'Method: Puts matrix [A] at the left and the singular matrix [I] at the right:
'[ a11 a12 a13 | 1 0 0 ]
'[ a21 a22 a23 | 0 1 0 ]
'[ a31 a32 a33 | 0 0 1 ]
'Then using line operations, we try to build the singular matrix [I] at the left.
'After we have finished, the inverse matrix [A]-1 (bij) will be at the right:
'[ 1 0 0 | b11 b12 b13 ]
'[ 0 1 0 | b21 b22 b23 ]
'[ 0 0 1 | b31 b32 b33 ]

On Error GoTo errhandler 'In case the inverse cannot be found (Determinant = 0)

Solution_Problem = False

'Assign values from matrix [A] at the left
For N = 1 To System_DIM
    For m = 1 To System_DIM
        Operations_Matrix(m, N) = Matrix_A(m, N)
    Next
Next

'Assign values from singular matrix [I] at the right
For N = 1 To System_DIM
    For m = 1 To System_DIM
        If N = m Then
            Operations_Matrix(m, N + System_DIM) = 1
        Else
            Operations_Matrix(m, N + System_DIM) = 0
        End If
    Next
Next

'Build the Singular matrix [I] at the left
For k = 1 To System_DIM
   'Bring a non-zero element first by changes lines if necessary
   If Operations_Matrix(k, k) = 0 Then
      For N = k To System_DIM
        If Operations_Matrix(N, k) <> 0 Then line_1 = N: Exit For 'Finds line_1 with non-zero element
      Next N
      'Change line k with line_1
      For m = k To System_DIM * 2
         temporary_1 = Operations_Matrix(k, m)
         Operations_Matrix(k, m) = Operations_Matrix(line_1, m)
         Operations_Matrix(line_1, m) = temporary_1
      Next m
   End If
   
    elem1 = Operations_Matrix(k, k)
   For N = k To 2 * System_DIM
    Operations_Matrix(k, N) = Operations_Matrix(k, N) / elem1
   Next N
   
   'For other lines, make a zero element by using:
   'Ai1=Aij-A11*(Aij/A11)
   'and change all the line using the same formula for other elements
   For N = 1 To System_DIM
        If N = k And N = MAX_DIM Then Exit For 'Finished
        If N = k And N < MAX_DIM Then N = N + 1 'Do not change that element (already equals to 1), go for next one
      If Operations_Matrix(N, k) <> 0 Then 'if it is zero, stays as it is
         multiplier_1 = Operations_Matrix(N, k) / Operations_Matrix(k, k)
         For m = k To 2 * System_DIM
            Operations_Matrix(N, m) = Operations_Matrix(N, m) - Operations_Matrix(k, m) * multiplier_1
         Next m
      End If
   Next N
Next k

'Assign the right part to the Inverse_Matrix
For N = 1 To System_DIM
    For k = 1 To System_DIM
        Inverse_Matrix(N, k) = Operations_Matrix(N, System_DIM + k)
    Next k
Next N

Exit Sub

errhandler:
message$ = "An error occured during the calculation process. Determinant of Matrix [A] is probably equal to zero."
response = MsgBox(message$, vbCritical)
Solution_Problem = True

End Sub

Sub Calculate_Transpose()
'Calculates the transpose of matrix [A]
For N = 1 To System_DIM
    For k = 1 To System_DIM
        Transpose_Matrix(N, k) = Matrix_A(k, N)
    Next k
Next N

End Sub

Sub Calculate_MatrixMult()
'Calculates the product [A]*[A-1] which must be always equal to the Singular Matrix [I]
'The same result must also come up for the product: [A-1]*[A]=[I]
'You can try it using: Matrix_Mult(k, m) = Matrix_Mult(k, m) + Inverse_Matrix(L, m) * Matrix_A(k, L)

Erase Matrix_Mult

For k = 1 To System_DIM
    For m = 1 To System_DIM
        For L = 1 To System_DIM
            Matrix_Mult(k, m) = Matrix_Mult(k, m) + Matrix_A(k, L) * Inverse_Matrix(L, m)
        Next L
    Next m
Next k

End Sub

Sub Type_Result()
'Types the result to the multi-line textbox
If Solution_Problem = False Then 'The inverse has been found
    result_1$ = "Result:" + vbCrLf
    For N = 1 To System_DIM
        line_1$ = ""
        For k = 1 To System_DIM
            'Type the corresponding result based on user's selection
            If optINVERSE = True Then line_1$ = line_1$ + Str(Inverse_Matrix(N, k)) 'Fill the row
            If optTRANSPOSE = True Then line_1$ = line_1$ + Str(Transpose_Matrix(N, k))
            If optMMULT = True Then line_1$ = line_1$ + Str(Matrix_Mult(N, k))
            If k < System_DIM Then line_1$ = line_1$ + ", " 'Do not print comma ',' after last column
        Next k
            result_1$ = result_1$ + vbCrLf + line_1$
    Next N
Else
    result_1$ = "Error!" 'Inverse could not be found, determinant probably equal to zero
End If

Text3.Text = result_1$

End Sub


