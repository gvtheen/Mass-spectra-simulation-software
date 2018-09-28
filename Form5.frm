VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Distribution 
   Caption         =   "Form1"
   ClientHeight    =   4905
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   7110
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4335
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   7646
      _Version        =   393216
      BackColorSel    =   -2147483643
      BackColorBkg    =   16777215
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu SaveData 
         Caption         =   "SaveData"
      End
      Begin VB.Menu SaveData_As 
         Caption         =   "SaveData As"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Distribution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call GridInitate(4, 5)
End Sub

