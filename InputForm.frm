VERSION 5.00
Begin VB.Form InputForm 
   Caption         =   "SettingTheData"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5700
   Icon            =   "InputForm.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   5700
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Height          =   4695
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.TextBox AtomNumber 
         Height          =   270
         Index           =   9
         Left            =   2520
         TabIndex        =   36
         Top             =   3960
         Width           =   735
      End
      Begin VB.TextBox AtomNumber 
         Height          =   270
         Index           =   8
         Left            =   2520
         TabIndex        =   35
         Top             =   3600
         Width           =   735
      End
      Begin VB.TextBox AtomNumber 
         Height          =   270
         Index           =   7
         Left            =   2520
         TabIndex        =   34
         Top             =   3240
         Width           =   735
      End
      Begin VB.TextBox AtomNumber 
         Height          =   270
         Index           =   6
         Left            =   2520
         TabIndex        =   33
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox AtomNumber 
         Height          =   270
         Index           =   5
         Left            =   2520
         TabIndex        =   32
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox AtomNumber 
         Height          =   270
         Index           =   4
         Left            =   2520
         TabIndex        =   31
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox AtomNumber 
         Height          =   270
         Index           =   3
         Left            =   2520
         TabIndex        =   30
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox AtomNumber 
         Height          =   270
         Index           =   2
         Left            =   2520
         TabIndex        =   29
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox AtomNumber 
         Height          =   270
         Index           =   1
         Left            =   2520
         TabIndex        =   28
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox AtomNumber 
         Height          =   270
         Index           =   0
         Left            =   2520
         TabIndex        =   27
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox AtomIOPNum 
         Height          =   375
         Left            =   1440
         TabIndex        =   26
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox FengDu 
         Height          =   270
         Index           =   9
         Left            =   1440
         TabIndex        =   23
         Top             =   3960
         Width           =   975
      End
      Begin VB.TextBox FengDu 
         Height          =   270
         Index           =   8
         Left            =   1440
         TabIndex        =   19
         Top             =   3600
         Width           =   975
      End
      Begin VB.TextBox FengDu 
         Height          =   270
         Index           =   7
         Left            =   1440
         TabIndex        =   18
         Top             =   3240
         Width           =   975
      End
      Begin VB.TextBox FengDu 
         Height          =   270
         Index           =   6
         Left            =   1440
         TabIndex        =   17
         Top             =   2880
         Width           =   975
      End
      Begin VB.TextBox FengDu 
         Height          =   270
         Index           =   5
         Left            =   1440
         TabIndex        =   14
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox FengDu 
         Height          =   270
         Index           =   4
         Left            =   1440
         TabIndex        =   13
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox FengDu 
         Height          =   270
         Index           =   3
         Left            =   1440
         TabIndex        =   11
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox FengDu 
         Height          =   270
         Index           =   2
         Left            =   1440
         TabIndex        =   9
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox FengDu 
         Height          =   270
         Index           =   1
         Left            =   1440
         TabIndex        =   8
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox FengDu 
         Height          =   270
         Index           =   0
         Left            =   1440
         TabIndex        =   7
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   3600
         TabIndex        =   2
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CommandButton OK 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   3600
         TabIndex        =   1
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "AtomNumber:"
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label IopValue 
         Height          =   375
         Index           =   9
         Left            =   360
         TabIndex        =   24
         Top             =   3960
         Width           =   975
      End
      Begin VB.Label IopValue 
         Height          =   375
         Index           =   8
         Left            =   360
         TabIndex        =   22
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label IopValue 
         Height          =   375
         Index           =   7
         Left            =   240
         TabIndex        =   21
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label IopValue 
         Height          =   375
         Index           =   6
         Left            =   240
         TabIndex        =   20
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label IopValue 
         Height          =   375
         Index           =   5
         Left            =   360
         TabIndex        =   16
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label IopValue 
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   15
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label AtomNum 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3600
         TabIndex        =   12
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label IopValue 
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   10
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label IopValue 
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label IopValue 
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label IopValue 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   975
      End
      Begin VB.Label EleName 
         Alignment       =   2  'Center
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
         Left            =   3480
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "InputForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NameStrNum As Integer
Dim ChangeAtomNumBool As Boolean
Public Sub SetIOPData(mm As Integer)
Dim i As Integer
EleName = TotleElemTable(mm).ElementName

NameStrNum = mm
AtomNum = mm + 1
AtomIOPNum.ToolTipText = "此文本填写原子数，必填"
For i = 0 To 9
   IopValue(i).Visible = True: AtomNumber(i).Visible = True
   FengDu(i).Visible = True
   IopValue(i).Enabled = True: FengDu(i).Enabled = True: AtomNumber(i).Enabled = True
 Next i
For i = 0 To TotleElemTable(mm).IopNum - 1
      IopValue(i).Caption = TotleElemTable(mm).ElementIopData(i, 0)
      FengDu(i).Text = FormatNumber(TotleElemTable(mm).ElementIopData(i, 1), 5, vbTrue)
      FengDu(i).ToolTipText = "此文本填写同位素丰度"
      AtomNumber(i).ToolTipText = "此文本填写同位素原子数"
Next i
For i = TotleElemTable(mm).IopNum To 9
     IopValue(i).Visible = False: IopValue(i).Enabled = False
     FengDu(i).Visible = False: FengDu(i).Enabled = False
     AtomNumber(i).Visible = False: AtomNumber(i).Enabled = False
 Next i
End Sub
Private Sub AtomNumber_Change(Index As Integer)
Dim i As Integer, sum As Single

If AtomIOPNum.Text = "" Then Exit Sub

If Val(AtomNumber(Index).Text) > AtomIOPNum.Text Then AtomNumber(Index).Text = AtomIOPNum.Text
sum = 0
If AtomIOPNum.Text <> "" And Index = TotleElemTable(NameStrNum).IopNum - 2 Then
     For i = 0 To Index
         sum = sum + Val(AtomNumber(i).Text)
     Next i
     AtomNumber(Index + 1).Text = Val(AtomIOPNum.Text) - sum
    
End If
End Sub
Private Sub OK_Click()
Dim i As Integer, j As Integer, sum As Double, SetElemData() As Double, num1 As Double, num2 As Double, num3 As Integer, AtomMulti() As Double
If AtomIOPNum.Text = " " Then
    For i = 0 To TotleElemTable(NameStrNum).IopNum - 1
        AtomIOPNum.Text = Val(AtomIOPNum.Text) + Val(AtomNumber(i).Text)
    Next i
End If
If Val(AtomIOPNum.Text) <= 0 Then AtomIOPNum.Text = "": MsgBox "Pleast input The atom numbers in the AtomNumber blank", vbInformation + vbYes, "Message": Exit Sub

Call ElementTable.ShowElement(Val(AtomIOPNum.Text), NameStrNum)
For i = 0 To TotleElemTable(NameStrNum).IopNum - 1
    sum = sum + Val(FengDu(i).Text)
Next i

For i = 0 To TotleElemTable(NameStrNum).IopNum - 1      '检查输入各个同位素原子数是否等于输入原子总数
     num3 = num3 + Val(Trim(AtomNumber(i).Text))
Next i
If num3 = Val(Trim(AtomIOPNum.Text)) Then ChangeAtomNumBool = True

If Abs(sum - 100) > 0.5 And ChangeAtomNumBool = False Then       '直接输入丰度时检测处理部分
   MsgBox "丰度输入有误，请重输", vbInformation + vbOKCancel, "Message"
     For i = 0 To TotleElemTable(NameStrNum).IopNum - 1
      IopValue(i).Caption = TotleElemTable(NameStrNum).ElementIopData(i, 0)
      FengDu(i).Text = TotleElemTable(NameStrNum).ElementIopData(i, 1)
    Next i
    Exit Sub
 Else
  If IOPValueBool Then                                        '判断是否改变丰度
    ReDim SetElemData(TotleElemTable(NameStrNum).IopNum, 2)
    For i = 0 To TotleElemTable(NameStrNum).IopNum - 1
         SetElemData(i, 1) = Val(Trim(FengDu(i).Text)): If SetElemData(i, 1) = 0 Then SetElemData(i, 1) = SetElemData(i, 1) + 1E-20
         SetElemData(i, 0) = Val(Trim(IopValue(i).Caption))
    Next i
    For i = 0 To TotleElemTable(NameStrNum).IopNum - 1
        For j = i + 1 To TotleElemTable(NameStrNum).IopNum - 1
            If SetElemData(i, 1) < SetElemData(j, 1) Then
                num1 = SetElemData(i, 1): SetElemData(i, 1) = SetElemData(j, 1): SetElemData(j, 1) = num1
                num2 = SetElemData(i, 0): SetElemData(i, 0) = SetElemData(j, 0): SetElemData(j, 0) = num2
            End If
    Next j, i
    For i = 0 To TotleElemTable(NameStrNum).IopNum - 1
           TotleElemTable(NameStrNum).ElementIopData(i, 0) = SetElemData(i, 0)
           TotleElemTable(NameStrNum).ElementIopData(i, 1) = SetElemData(i, 1)
    Next i
 End If
End If
If ChangeAtomNumBool Then
    ReDim AtomMulti(TotleElemTable(NameStrNum).IopNum, 2)
    For i = 0 To TotleElemTable(NameStrNum).IopNum - 1
        AtomMulti(i, 0) = Val(Trim(IopValue(i).Caption))
        AtomMulti(i, 1) = 100 * Val(Trim(AtomNumber(i).Text)) / Val(Trim(AtomIOPNum.Text))
        If AtomMulti(i, 1) = 0 Then AtomMulti(i, 1) = AtomMulti(i, 1) + 1E-30
    Next i
    For i = 0 To TotleElemTable(NameStrNum).IopNum - 1
       For j = i + 1 To TotleElemTable(NameStrNum).IopNum - 1
            If AtomMulti(i, 1) < AtomMulti(j, 1) Then
               num1 = AtomMulti(i, 1)
               AtomMulti(i, 1) = AtomMulti(j, 1)
               AtomMulti(j, 1) = num1
               num2 = AtomMulti(i, 0)
               AtomMulti(i, 0) = AtomMulti(j, 0)
               AtomMulti(j, 0) = num2
            End If
    Next j, i
    For i = 0 To TotleElemTable(NameStrNum).IopNum - 1
           TotleElemTable(NameStrNum).ElementIopData(i, 0) = AtomMulti(i, 0)
           TotleElemTable(NameStrNum).ElementIopData(i, 1) = AtomMulti(i, 1)
    Next i
    ChangeAtomNumBool = False
End If
Unload Me
End Sub
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub FengDu_Change(Index As Integer)
Dim i As Integer, sum As Double
If Val(FengDu(Index).Text) > 100 Then FengDu(Index).Text = 100
If Index = TotleElemTable(NameStrNum).IopNum - 2 Then
     For i = 0 To Index
         sum = Val(FengDu(i).Text)
     Next i
     FengDu(Index + 1).Text = 100 - sum
End If
IOPValueBool = True
End Sub
Private Sub Form_Activate()
AtomIOPNum.SetFocus
IOPValueBool = False
ChangeAtomNumBool = False
End Sub

Private Sub Form_Load()
Dim i As Integer
For i = 0 To 9
   IopValue(i).Left = 240
   IopValue(i).Top = 720 + i * 360
Next i
End Sub
