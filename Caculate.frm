VERSION 5.00
Begin VB.Form Caculate 
   Caption         =   "FindMaxAbundance"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   5160
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame FindMaxAbundance 
      Caption         =   "FindMaxAbundance"
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4815
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   2400
         TabIndex        =   15
         Top             =   5160
         Width           =   975
      End
      Begin VB.CommandButton OK 
         Caption         =   "OK"
         Height          =   495
         Left            =   1080
         TabIndex        =   14
         Top             =   5160
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   5
         Left            =   2160
         TabIndex        =   12
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   4
         Left            =   2160
         TabIndex        =   10
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   3
         Left            =   2160
         TabIndex        =   9
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   2
         Left            =   2160
         TabIndex        =   6
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   1
         Left            =   2160
         TabIndex        =   4
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   0
         Left            =   2160
         TabIndex        =   3
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "For_Mass"
         Height          =   375
         Index           =   5
         Left            =   960
         TabIndex        =   13
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Thi_Mass"
         Height          =   375
         Index           =   4
         Left            =   960
         TabIndex        =   11
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Sec_Mass"
         Height          =   375
         Index           =   3
         Left            =   960
         TabIndex        =   8
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label ShowAbundance 
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   4200
         Width           =   4335
      End
      Begin VB.Label Label1 
         Caption         =   "Fir_Mass"
         Height          =   375
         Index           =   2
         Left            =   960
         TabIndex        =   5
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Cl_Number"
         Height          =   375
         Index           =   1
         Left            =   960
         TabIndex        =   2
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "C_Number"
         Height          =   375
         Index           =   0
         Left            =   960
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Caculate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim C12NumData() As Integer, GroupNum As Integer, C_Num As Integer
Private Sub OK_Click()

Dim i As Integer, j As Integer, k As Integer
Dim MaxMinData() As Double, DataNum As Integer
Dim maxData As Double, minData As Double, DataMaxMin() As Double
Call GetC12Num
C_Num = Val(Text1(0).Text)
ReDim DataMaxMin(GroupNum, 2)
ReDim MaxMinData(GroupNum, C_Num - 1)
For i = 0 To GroupNum - 1
     DataNum = 0
     
     For j = 1 To C_Num
         For k = 0 To GroupNum - 1
           If C12NumData(k) <> j Then MaxMinData(i, DataNum) = CompareData(C12NumData(i), j): DataNum = DataNum + 1
     Next k, j
     
     For k = GroupNum - 1 To 0 Step -1
        If k > i Then MaxMinData(i, DataNum) = CompareData(C12NumData(i), C12NumData(k)): DataNum = DataNum + 1
        If k < i Then MaxMinData(i, DataNum) = CompareData(C12NumData(k), C12NumData(i)): DataNum = DataNum + 1
     Next k
Next i
For i = 0 To GroupNum - 1
     maxData = 1#: minData = MaxMinData(i, 0)
     For j = DataNum - 1 To DataNum - 1 - i Step -1
          If maxData > MaxMinData(j, 0) Then maxData = MaxMinData(j, 0)
     Next j
     For j = 0 To DataNum - 1 - i
          If minData < MaxMinData(j, 0) Then minData = MaxMinData(j, 0)
     Next j
     DataMaxMin(i, 0) = minData: DataMaxMin(i, 1) = maxData
Next i
minData = DataMaxMin(i, 0): maxData = DataMaxMin(i, 1)
For i = 0 To GroupNum - 1
    If minData < DataMaxMin(i, 0) Then minData = DataMaxMin(i, 0)
    If maxData > DataMaxMin(i, 1) Then maxData = DataMaxMin(i, 1)
Next i
ShowAbundance.Caption = minData & " < " & "Aboun" & " < " & maxData
'    C13_Num = C_Num - C12_Num
'    C12_Abun_Max = 1 - (C_Num - C12_Num + 1) / (C_Num + 1)
'    C12_Abun_Min = 1 - (C_Num - C12_Num) / (C_Num + 1)
'    ShowAbundance = C_Num * FormatNumber((1 - C12_Abun_Max), 6, vbTrue) & " < " & " C13 Of Abun " & "<" & C_Num * FormatNumber((1 - C12_Abun_Min), 6, vbTrue)
'    TotleElemTable(11).ElementIopData(0, 1) = (C12_Abun_Min + C12_Abun_Max) / 2
'    TotleElemTable(11).ElementIopData(1, 1) = 1 - (C12_Abun_Min + C12_Abun_Max) / 2
'    ElementTable.ShowChemForm.Caption = "C" & C_Num & " " & "Cl" & Cl_Num
 


End Sub
Private Function CompareData(n1 As Integer, n2 As Integer) As Double
Dim nn1 As Integer, nn2 As Integer
nn1 = n1: nn2 = n2
CompareData = 1 - 1 / (1 + (ProSecData(C_Num, n1) / ProSecData(C_Num, n2)) ^ (1 / (n1 - n2)))
End Function
Private Sub GetC12Num()
Dim i As Integer, Cl_Num As Integer
Dim Totle_Mass As Single, C_MASS As Double
GroupNum = 0
Cl_Num = Val(Text1(1).Text)

ReDim C12NumData(4)
For i = 2 To 5
    If Text1(i).Text <> "" Then
      
      Totle_Mass = Val(Text1(i).Text)
      C_MASS = Totle_Mass - TotleElemTable(16).ElementIopData(0, 0) * Cl_Num
      C12NumData(GroupNum) = (C_MASS - C_Num * TotleElemTable(5).ElementIopData(1, 0)) / (TotleElemTable(5).ElementIopData(0, 0) - TotleElemTable(5).ElementIopData(1, 0))
      GroupNum = GroupNum + 1
      
    End If
Next i
ReDim Preserve C12NumData(GroupNum)
End Sub
Private Function ProSecData(n As Integer, m As Integer) As Double '二项式系数处理
Dim bool As Integer
On Error GoTo H1
Dim i As Integer, aa As Double, bb As Double, cc As Double, sum As Double
aa = n: bb = m: cc = n - m
sum = 1


If m = 0 Or n = m Then ProSecData = 1: Exit Function
For i = 0 To n - 1
    
    aa = n - i
    If bb = 1# Then
      bb = 1#
    Else
      bb = m - i
    End If
    If cc = 1# Then
      cc = 1#
    Else
      cc = n - m - i
    End If
    sum = sum * aa / (bb * cc)
    
Next i

ProSecData = sum
Exit Function
H1: Exit Function
Exit Function
End Function
Private Sub Command2_Click()
Unload Me
End Sub

