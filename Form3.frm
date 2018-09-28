VERSION 5.00
Begin VB.Form FormuForm 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edit Formula"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4530
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Caption         =   "InputFormula"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2040
         TabIndex        =   4
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ok"
         Default         =   -1  'True
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox FormText 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   2
         ToolTipText     =   "不要用空格隔开"
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Formula"
         Height          =   180
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   630
      End
   End
End
Attribute VB_Name = "FormuForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
    Dim i As Integer, str As String, leng As Integer, str1 As String
    ShutFormuFormBool = True
    str = Trim(FormText.Text)
    leng = Len(str)
    For i = 1 To leng
        str1 = str1 & Mid(str, i, 1)
        If i <= leng - 1 And i > 1 Then
          If (Asc(Mid(str, i, 1)) >= 48 And Asc(Mid(str, i, 1)) <= 57) And ((Asc(Mid(str, i + 1, 1)) >= 65 And Asc(Mid(str, i + 1, 1)) <= 90) Or (Asc(Mid(str, i + 1, 1)) >= 97 And Asc(Mid(str, i + 1, 1)) <= 122)) Then
           str1 = str1 & " "
          End If
        End If
     
    Next i
    MocularForm = str1
    ElementTable.ShowChemForm.Caption = str1
    Call ElementTable.DealInput(str1)
            If ShutFormuFormBool = False Then FormText.Text = "": Exit Sub
    Unload Me
End Sub
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Form_Load()
FormText.Text = MocularForm
End Sub
