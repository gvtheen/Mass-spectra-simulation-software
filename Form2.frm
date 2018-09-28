VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   6900
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8535
   LinkTopic       =   "Form2"
   ScaleHeight     =   6900
   ScaleWidth      =   8535
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "Set"
      Height          =   735
      Left            =   5640
      TabIndex        =   1
      Top             =   5160
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "丰度设置"
      Height          =   6255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7935
      Begin VB.Frame Frame6 
         Caption         =   "Br"
         Height          =   1815
         Left            =   360
         TabIndex        =   39
         Top             =   4320
         Width           =   2055
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   15
            Left            =   960
            TabIndex        =   41
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   16
            Left            =   960
            TabIndex        =   40
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "78.92"
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   43
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label5 
            Caption         =   "80.92"
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   42
            Top             =   1080
            Width           =   495
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "O"
         Height          =   1815
         Index           =   2
         Left            =   5400
         TabIndex        =   14
         Top             =   2400
         Width           =   2175
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   14
            Left            =   840
            TabIndex        =   38
            Top             =   1200
            Width           =   855
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   13
            Left            =   840
            TabIndex        =   37
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   12
            Left            =   840
            TabIndex        =   36
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "17.99"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   25
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label7 
            Caption         =   "16.99"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   24
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label7 
            Caption         =   "15.99"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   23
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "N"
         Height          =   1815
         Index           =   1
         Left            =   5400
         TabIndex        =   13
         Top             =   360
         Width           =   2175
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   5
            Left            =   840
            TabIndex        =   29
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   4
            Left            =   840
            TabIndex        =   28
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "15.00"
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   16
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label4 
            Caption         =   "14.00"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   15
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "S"
         Height          =   1815
         Index           =   0
         Left            =   2880
         TabIndex        =   6
         Top             =   2400
         Width           =   2175
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   11
            Left            =   840
            TabIndex        =   35
            Top             =   1320
            Width           =   855
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   10
            Left            =   840
            TabIndex        =   34
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   9
            Left            =   840
            TabIndex        =   33
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   8
            Left            =   840
            TabIndex        =   32
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label6 
            Caption         =   "35.97"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   22
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label6 
            Caption         =   "33.97"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   21
            Top             =   960
            Width           =   495
         End
         Begin VB.Label Label6 
            Caption         =   "32.97"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   20
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label6 
            Caption         =   "31.97"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   19
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Cl"
         Height          =   1815
         Left            =   360
         TabIndex        =   5
         Top             =   2400
         Width           =   2055
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   7
            Left            =   960
            TabIndex        =   31
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   6
            Left            =   960
            TabIndex        =   30
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "36.96"
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   18
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label5 
            Caption         =   "34.96"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   17
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "C"
         Height          =   1815
         Left            =   2880
         TabIndex        =   4
         Top             =   360
         Width           =   2175
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   3
            Left            =   840
            TabIndex        =   27
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   2
            Left            =   840
            TabIndex        =   26
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "13.00"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Top             =   1200
            Width           =   495
         End
         Begin VB.Label Label3 
            Caption         =   "12.00"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   480
            Width           =   495
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "H"
         Height          =   1815
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   2055
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   1
            Left            =   960
            TabIndex        =   8
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   0
            Left            =   960
            TabIndex        =   7
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "2.01"
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   10
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "1.00"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   9
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   90
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim i As Integer
For i = 0 To 16
    If Text1(i).Text <> "" Then ElementMass(i, 1) = Val(Trim(Text1(i).Text))
Next i
SetBool = False
Form2.Hide
End Sub
