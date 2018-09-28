VERSION 5.00
Begin VB.Form AboutForm 
   Caption         =   "About"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5355
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":058A
   ScaleHeight     =   6240
   ScaleWidth      =   5355
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "本程序由BlackHawk编写,发现使用的问题及时向我提出！"
      Height          =   420
      Left            =   480
      TabIndex        =   1
      Top             =   5160
      Width           =   4500
   End
   Begin VB.Label Label1 
      Caption         =   "欢迎使用IOP MASS Simulation."
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   4800
      Width           =   3255
   End
End
Attribute VB_Name = "AboutForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
