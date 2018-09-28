VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form DisForm 
   AutoRedraw      =   -1  'True
   Caption         =   "Mass Spectrum Display"
   ClientHeight    =   6525
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10200
   Icon            =   "PaintPict.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   10200
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Distr_Frame 
      Caption         =   "Destribution"
      Height          =   5655
      Left            =   600
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   7215
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   5295
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   9340
         _Version        =   393216
         BackColorBkg    =   16777215
         AllowUserResizing=   3
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture2 
      Height          =   495
      Left            =   480
      ScaleHeight     =   435
      ScaleWidth      =   1275
      TabIndex        =   9
      Top             =   5880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "信息"
      Height          =   5535
      Left            =   8160
      TabIndex        =   1
      Top             =   240
      Width           =   1575
      Begin VB.CommandButton LabelCommand 
         Caption         =   "标记谱峰"
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label ShowNum 
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "显示数目："
         Height          =   180
         Left            =   120
         TabIndex        =   6
         Top             =   2160
         Width           =   915
      End
      Begin VB.Label TotalNum 
         Height          =   300
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   825
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "谱线总数："
         Height          =   180
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   915
      End
      Begin VB.Label ChemForm 
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "化学式："
         Height          =   180
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   720
      End
   End
   Begin VB.PictureBox MassPicture 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   600
      ScaleHeight     =   5655
      ScaleWidth      =   7215
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      Begin VB.Line Line4 
         BorderStyle     =   3  'Dot
         Visible         =   0   'False
         X1              =   720
         X2              =   3360
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line3 
         BorderStyle     =   3  'Dot
         Visible         =   0   'False
         X1              =   720
         X2              =   3360
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line2 
         BorderStyle     =   3  'Dot
         Visible         =   0   'False
         X1              =   720
         X2              =   3360
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  'Dot
         Visible         =   0   'False
         X1              =   840
         X2              =   3480
         Y1              =   1560
         Y2              =   1560
      End
   End
   Begin VB.Menu PopManu 
      Caption         =   "PopManu"
      Visible         =   0   'False
      Begin VB.Menu Zoomlarge 
         Caption         =   "Zoom in"
      End
      Begin VB.Menu ZoomSmall 
         Caption         =   "Zoom Out"
      End
      Begin VB.Menu Back 
         Caption         =   "Back"
      End
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Save 
         Caption         =   "Save"
         Begin VB.Menu Save_Data 
            Caption         =   "SaveData"
         End
         Begin VB.Menu Save_Picture 
            Caption         =   "SaveSpectrun"
            Shortcut        =   ^S
         End
      End
      Begin VB.Menu Printer1 
         Caption         =   "PrintSpectrum"
         Shortcut        =   ^P
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu Tool 
      Caption         =   "Tool"
      Begin VB.Menu Taskmgr 
         Caption         =   "Taskmgr"
         Shortcut        =   ^T
      End
      Begin VB.Menu Calculator 
         Caption         =   "Calculator"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu Display 
      Caption         =   "Display"
      Begin VB.Menu Spectrum 
         Caption         =   "Spectrum"
      End
      Begin VB.Menu Distribution 
         Caption         =   "Distribution"
      End
   End
   Begin VB.Menu Setting 
      Caption         =   "Setting"
      WindowList      =   -1  'True
      Begin VB.Menu Resolv 
         Caption         =   "Resolution"
         Begin VB.Menu Resolv_1 
            Caption         =   "1.0"
         End
         Begin VB.Menu Resolv_05 
            Caption         =   "0.5"
         End
         Begin VB.Menu Resolv_03 
            Caption         =   "0.3 Y"
         End
         Begin VB.Menu Resolv_003 
            Caption         =   "0.03"
         End
         Begin VB.Menu Resolv_0003 
            Caption         =   "0.003"
         End
         Begin VB.Menu Resolv_00003 
            Caption         =   "0.0003"
         End
         Begin VB.Menu Resolv_Ide 
            Caption         =   "Ideal"
         End
         Begin VB.Menu User_Defined 
            Caption         =   "User_Defined"
         End
      End
      Begin VB.Menu BIAO 
         Caption         =   "LabelRange"
         Begin VB.Menu More_1 
            Caption         =   ">1 "
         End
         Begin VB.Menu More_10 
            Caption         =   ">10 Y"
         End
         Begin VB.Menu More_20 
            Caption         =   ">20"
         End
         Begin VB.Menu More_30 
            Caption         =   ">30"
         End
         Begin VB.Menu More_40 
            Caption         =   ">40"
         End
         Begin VB.Menu More_50 
            Caption         =   ">50"
         End
         Begin VB.Menu More_60 
            Caption         =   ">60"
         End
         Begin VB.Menu More_70 
            Caption         =   ">70"
         End
         Begin VB.Menu More_80 
            Caption         =   ">80"
         End
         Begin VB.Menu More_90 
            Caption         =   ">90"
         End
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "DisForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Zoom As Boolean, X1 As Double, X2 As Double, Y1 As Double, Y2 As Double, ZoomX As Double, ZoomY As Double
'the varibles used in the adjust coordinate by mouse
Dim CoorWidth As Double, CoorHeight As Double, MinXcor As Double, MaxXcor As Double, MinYcor As Double, MaxYcor As Double
'the varibles used in the print the coordinate,which stand for the true the value of data in the coordinate
Dim IniXmin As Double, IniXmax As Double
Dim LabelBool As Boolean, HeightPicture As Long, maxData As Double, minData As Double
Dim EleFurProData() As DealDataType, SaveData() As Double, PointDealNum As Integer, ProceDataSave() As DataType, NumData As Long, LastNum As Long
Dim PaintNum As Integer
Dim DieCycleBool As Boolean, SaveResValue As Single
Dim Label_Value As Single
Dim AbunDataContant As Double

   Private Type Cooordinate
      xMin As Double
      Xmax As Double
      yMin As Double
      Ymax As Double
   End Type
   
   Private PictCoor As Cooordinate

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

    Private Type RECT                                '对象的大小长，宽，顶，底部（自定义）
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

                  hDCSrc = GetDC(hWndSrc)                                                         '获取指定窗口的设备场景

    Else

                   hDCSrc = GetWindowDC(hWndSrc)                                                  '获取整个窗口（包括边框、滚动条、标题栏、菜单等）的设备场景

    End If                                                                                        '创建一个与特定设备场景一致的内存设备场景

              hDCMemory = CreateCompatibleDC(hDCSrc)

              hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)             '创建一幅与指定大小和句柄的位图

              hBmpPrev = SelectObject(hDCMemory, hBmp)                             '每个设备场景都可能有选入其中的图形对象

    '获得屏幕属性

              RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS)     ' RASTERCAPS 关键词                            '根据指定设备场景代表的设备的功能返回信息

              HasPaletteScrn = RasterCapsScrn And RC_PALETTE           'RC_PALETTE：设备基于调色板

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

    r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, vbSrcCopy)                '将一幅位图从一个设备场景复制到另一个。源和目标DC相互间必须兼容

    hBmp = SelectObject(hDCMemory, hBmpPrev)

    If HasPaletteScrn And (PaletteSizeScrn = 256) Then

                  hPal = SelectPalette(hDCMemory, hPalPrev, 0)

    End If

    '释放资源

    r = DeleteDC(hDCMemory)
    r = ReleaseDC(hWndSrc, hDCSrc)
    Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
    End Function
                                                                            'capturescreen函数捕捉MassPicture图像
    Public Function CaptureScreen() As Picture
    
        Set CaptureScreen = CaptureWindow(MassPicture.hWnd, True, 0, 0, MassPicture.Width \ Screen.TwipsPerPixelX, MassPicture.Height \ Screen.TwipsPerPixelY)
        
    End Function
'Private Function CaptureActiveWindow() As Picture
 '   Dim hWndActive As Long
 '   Dim r As Long
 '   Dim RectActive As RECT
 '   hWndActive = GetForegroundWindow()
    
  '  r = GetWindowRect(hWndActive, RectActive)
 '   Set CaptureActiveWindow = CaptureWindow(hWndActive, False, 0, 0, RectActive.Right - RectActive.Left, RectActive.Bottom - RectActive.Top)
 '   End Function
 Private Sub PrintPictureToFitPage(Prn As Printer, Pic As Picture)
    Const vbHiMetric As Integer = 8
    Dim PicRatio As Double
    Dim PrnWidth As Double
    Dim PrnHeight As Double
    Dim PrnRatio As Double
    Dim PrnPicWidth As Double
    Dim PrnPicHeight As Double
    If Pic.Height >= Pic.Width Then
    Prn.Orientation = vbPRORPortrait
    Else
    Prn.Orientation = vbPRORLandscape
    End If
    PicRatio = Pic.Width / Pic.Height
    PrnWidth = Prn.ScaleX(Prn.ScaleWidth, Prn.ScaleMode, vbHiMetric)
    PrnHeight = Prn.ScaleY(Prn.ScaleHeight, Prn.ScaleMode, vbHiMetric)
    PrnRatio = PrnWidth / PrnHeight
    If PicRatio >= PrnRatio Then
    PrnPicWidth = Prn.ScaleX(PrnWidth, vbHiMetric, Prn.ScaleMode)
    PrnPicHeight = Prn.ScaleY(PrnWidth / PicRatio, vbHiMetric, Prn.ScaleMode)
    Else
    PrnPicHeight = Prn.ScaleY(PrnHeight, vbHiMetric, Prn.ScaleMode)
    PrnPicWidth = Prn.ScaleX(PrnHeight * PicRatio, vbHiMetric, Prn.ScaleMode)
    End If
    Prn.PaintPicture Pic, 0, 0, PrnPicWidth, PrnPicHeight
    End Sub
Private Sub Back_Click()
MinXcor = IniXmin: MaxXcor = IniXmax

Call Coordinate
Call PrinPic
End Sub
Private Sub Initiate()
PaintNum = 0
PointDealNum = 0
NumData = 0
LastNum = 0
LabelBool = False
Label_Value = 10
DieCycleBool = False


PictCoor.Xmax = MassPicture.Width - 450
 PictCoor.xMin = 450
 PictCoor.Ymax = MassPicture.Height - 500
 PictCoor.yMin = 800
 End Sub
Public Sub ProcData()
'On Error GoTo h1
Dim i As Integer, j As Integer, k As Integer, SaveDataNum As Long, datapoint() As Integer, num1 As Double, num2 As Double
ReDim EleFurProData(FormAtomNum) As DealDataType

  MinYcor = 0:  MaxYcor = 100

For i = 0 To FormAtomNum - 1
     Select Case TotleElemTable(ProcDataElem(i).PointNum).IopNum
       Case 1
            Call OneDealIOP(ProcDataElem(i).PointNum, ProcDataElem(i).num)
       Case 2
            Call TwoDealIOP(ProcDataElem(i).PointNum, ProcDataElem(i).num)
       Case 3
            Call ThrDealIOP(ProcDataElem(i).PointNum, ProcDataElem(i).num)
       Case 4
            Call FurDealIOP(ProcDataElem(i).PointNum, ProcDataElem(i).num)
       Case 5
            Call FivDealIOP(ProcDataElem(i).PointNum, ProcDataElem(i).num)
       Case 6
            Call SixDealIOP(ProcDataElem(i).PointNum, ProcDataElem(i).num)
       Case Is >= 7
            Call SevDealIOP(ProcDataElem(i).PointNum, ProcDataElem(i).num)
      
       Case Else
     End Select
Next i

SaveDataNum = 1
For i = 0 To PointDealNum - 1
    SaveDataNum = SaveDataNum * EleFurProData(i).num
Next i

NumData = SaveDataNum
ReDim SaveData(SaveDataNum, 2)
ReDim datapoint(PointDealNum)
For i = 0 To SaveDataNum - 1
   SaveData(i, 0) = 0
   SaveData(i, 1) = 1
Next i

For i = 0 To PointDealNum - 1
    datapoint(i) = 0
Next i

SaveDataNum = 0
Do                                                'sum1存储数据的个数，databool(sum1)存储是数据的位置
     For j = 0 To PointDealNum - 1
         SaveData(SaveDataNum, 0) = SaveData(SaveDataNum, 0) + EleFurProData(j).MassData(datapoint(j))
         SaveData(SaveDataNum, 1) = SaveData(SaveDataNum, 1) * EleFurProData(j).AbunData(datapoint(j))
     Next j
     datapoint(PointDealNum - 1) = datapoint(PointDealNum - 1) + 1
     SaveDataNum = SaveDataNum + 1
     k = PointDealNum - 1
     While datapoint(k) >= EleFurProData(k).num
         
           datapoint(k) = 0
           k = k - 1
           If k < 0 Then Exit Do
           datapoint(k) = datapoint(k) + 1
           If datapoint(0) > EleFurProData(0).num Then Exit Do
     Wend
     
Loop

 For i = 0 To NumData - 1                                                           '对得出的数据进行从小到大排序
      For j = i To NumData - 1
          If SaveData(i, 0) > SaveData(j, 0) Then
              num1 = SaveData(i, 0): SaveData(i, 0) = SaveData(j, 0): SaveData(j, 0) = num1
              num2 = SaveData(i, 1): SaveData(i, 1) = SaveData(j, 1): SaveData(j, 1) = num2
          End If
 Next j, i
 
 Call ProcDataFur(0.3)
 Call AdjustCoor
 Call PrinPic
Exit Sub

'h1: MsgBox "对不起！你输入的分子存在同位素峰太多了，程序无法计算", vbInformation + vbOKCancel, "消息提示": Exit Sub
End Sub
Private Sub ProcDataFur(ResValue As Double)
Dim i As Integer, j As Integer, Max_Abun As Double

LastNum = 0: i = 0: j = 0
SaveResValue = ResValue
For i = 0 To NumData - 1         '对于SaveData(i,2)标记进行初始化
   SaveData(i, 2) = 0
Next i

ReDim ProceDataSave(NumData)
'While i < NumData
'     ProceDataSave(LastNum).Abundance = SaveData(i, 1): ProceDataSave(LastNum).MassData = SaveData(i, 0)
'
'     While SaveData(j, 0) <= SaveData(i, 0) + ResValue And j < NumData
'           If SaveData(j, 1) > SaveData(i, 1) Then ProceDataSave(LastNum).MassData = SaveData(j, 0)
'           ProceDataSave(LastNum).Abundance = ProceDataSave(LastNum).Abundance + SaveData(j, 1)
'           j = j + 1
'     Wend
'     i = j + 1
'    LastNum = LastNum + 1
'Wend
For i = 0 To NumData - 1
    If SaveData(i, 2) <> 1# Then ProceDataSave(LastNum).Abundance = SaveData(i, 1): ProceDataSave(LastNum).MassData = SaveData(i, 0): ProceDataSave(LastNum).Labelpoint = 1
    For j = i + 1 To NumData - 1
       If Abs(SaveData(i, 0) - SaveData(j, 0)) < ResValue And SaveData(j, 2) <> 1# Then
          ProceDataSave(LastNum).Abundance = ProceDataSave(LastNum).Abundance + SaveData(j, 1): ProceDataSave(LastNum).MassData = (SaveData(i, 0) + SaveData(j, 0)) / 2
            SaveData(i, 2) = 1#: SaveData(j, 2) = 1
            ProceDataSave(LastNum).Labelpoint = ProceDataSave(LastNum).Labelpoint + 1
        End If
    Next j
LastNum = LastNum + 1

Next i

   DataMax_Abundance = ProceDataSave(0).Abundance                                                 '寻找最大值
   For i = 0 To LastNum - 1
       If DataMax_Abundance < ProceDataSave(i).Abundance Then DataMax_Abundance = ProceDataSave(i).Abundance
   Next i
   
   For i = 0 To LastNum - 1
       ProceDataSave(i).Abundance = ProceDataSave(i).Abundance / DataMax_Abundance
   Next i
   ReDim PaintData(LastNum)
   
   PaintNum = 0
 
   For i = 0 To LastNum - 1
      If ProceDataSave(i).Abundance >= 0.00000001 Then PaintData(PaintNum).MassData = ProceDataSave(i).MassData: PaintData(PaintNum).Abundance = ProceDataSave(i).Abundance: PaintData(PaintNum).Labelpoint = ProceDataSave(i).Labelpoint: PaintNum = PaintNum + 1
   Next i
   Call GridInitate(PaintNum + 1, DataMax_Abundance)
   
   End Sub
Private Sub OneDealIOP(EleLabel As Integer, AtomNum As Integer)       '一个同位素处理
    EleFurProData(PointDealNum).num = 1
    ReDim EleFurProData(PointDealNum).MassData(1)
    ReDim EleFurProData(PointDealNum).AbunData(1)
    EleFurProData(PointDealNum).MassData(0) = AtomNum * TotleElemTable(EleLabel).ElementIopData(0, 0)
    EleFurProData(PointDealNum).AbunData(0) = TotleElemTable(EleLabel).ElementIopData(0, 1) / 100
    PointDealNum = PointDealNum + 1
End Sub
Private Sub TwoDealIOP(EleLabel As Integer, AtomNum As Integer)        '二个同位素处理
    Dim i As Integer, ValueSav As Double, Sum1 As Integer
    
    Sum1 = 0
    ReDim EleFurProData(PointDealNum).MassData(AtomNum + 1)
    ReDim EleFurProData(PointDealNum).AbunData(AtomNum + 1)
    For i = 0 To AtomNum
        If DieCycleBool Then Exit Sub
            ValueSav = ProSecData(AtomNum, i) * (TotleElemTable(EleLabel).ElementIopData(0, 1) / 100) ^ i * (TotleElemTable(EleLabel).ElementIopData(1, 1) / 100) ^ (AtomNum - i)
            If ValueSav > AbunDataContant Then
               EleFurProData(PointDealNum).MassData(Sum1) = i * TotleElemTable(EleLabel).ElementIopData(0, 0) + (AtomNum - i) * TotleElemTable(EleLabel).ElementIopData(1, 0)
               EleFurProData(PointDealNum).AbunData(Sum1) = ValueSav
               Sum1 = Sum1 + 1
            End If
    Next i
    EleFurProData(PointDealNum).num = Sum1
    ReDim Preserve EleFurProData(PointDealNum).MassData(Sum1): ReDim Preserve EleFurProData(PointDealNum).AbunData(Sum1)
    PointDealNum = PointDealNum + 1
End Sub
Private Sub ThrDealIOP(EleLabel As Integer, AtomNum As Integer)              '三个同位素处理
    Dim i As Integer, sum As Integer, j As Integer, Sum1 As Integer, ValueSav As Double
    sum = AtomNum + 1
    Sum1 = 0
    
    
    ReDim EleFurProData(PointDealNum).MassData(sum)
    ReDim EleFurProData(PointDealNum).AbunData(sum)
    
    
    For i = 0 To AtomNum
        For j = 0 To AtomNum - i
           If DieCycleBool Then Exit Sub
           If Sum1 > sum - 1 Then sum = sum + 50: ReDim Preserve EleFurProData(PointDealNum).MassData(sum): ReDim Preserve EleFurProData(PointDealNum).AbunData(sum)
           ValueSav = ProSecData(AtomNum, i) * (TotleElemTable(EleLabel).ElementIopData(0, 1) / 100) ^ i * ProSecData(AtomNum - i, j) * (TotleElemTable(EleLabel).ElementIopData(1, 1) / 100) ^ j * (TotleElemTable(EleLabel).ElementIopData(2, 1) / 100) ^ (AtomNum - i - j)
           If ValueSav > AbunDataContant Then
                EleFurProData(PointDealNum).MassData(Sum1) = i * TotleElemTable(EleLabel).ElementIopData(0, 0) + j * TotleElemTable(EleLabel).ElementIopData(1, 0) + (AtomNum - i - j) * TotleElemTable(EleLabel).ElementIopData(2, 0)
                EleFurProData(PointDealNum).AbunData(Sum1) = ValueSav
                Sum1 = Sum1 + 1
           End If
    Next j, i
    
    EleFurProData(PointDealNum).num = Sum1
    ReDim Preserve EleFurProData(PointDealNum).MassData(Sum1): ReDim Preserve EleFurProData(PointDealNum).AbunData(Sum1)
    PointDealNum = PointDealNum + 1
End Sub
Private Sub FurDealIOP(EleLabel As Integer, AtomNum As Integer)                 '多于四个同位素的处理
Dim i As Integer, j As Integer, k As Integer, sum As Double, Sum1 As Double, ValueSav As Double
    sum = AtomNum + 50
    Sum1 = 0
    
    ReDim EleFurProData(PointDealNum).MassData(sum)
    ReDim EleFurProData(PointDealNum).AbunData(sum)
    
    sum = 0
    For i = 0 To AtomNum
       For j = 0 To AtomNum - i
         For k = 0 To AtomNum - i - j
            If DieCycleBool Then Exit Sub
            If Sum1 > sum - 1 Then sum = sum + 50: ReDim Preserve EleFurProData(PointDealNum).MassData(sum): ReDim Preserve EleFurProData(PointDealNum).AbunData(sum)
            ValueSav = ProSecData(AtomNum, i) * (TotleElemTable(EleLabel).ElementIopData(0, 1) / 100) ^ i * ProSecData(AtomNum - i, j) * (TotleElemTable(EleLabel).ElementIopData(1, 1) / 100) ^ j * ProSecData(AtomNum - i - j, k) * (TotleElemTable(EleLabel).ElementIopData(2, 1) / 100) ^ k * (TotleElemTable(EleLabel).ElementIopData(3, 1) / 100) ^ (AtomNum - i - j - k)
            If ValueSav > AbunDataContant Then
                EleFurProData(PointDealNum).MassData(Sum1) = i * TotleElemTable(EleLabel).ElementIopData(0, 0) + j * TotleElemTable(EleLabel).ElementIopData(1, 0) + k * TotleElemTable(EleLabel).ElementIopData(2, 0) + (AtomNum - i - j - k) * TotleElemTable(EleLabel).ElementIopData(3, 0)
                EleFurProData(PointDealNum).AbunData(Sum1) = ValueSav
                Sum1 = Sum1 + 1
            End If
    Next k, j, i
    
    EleFurProData(PointDealNum).num = Sum1
    ReDim Preserve EleFurProData(PointDealNum).MassData(Sum1): ReDim Preserve EleFurProData(PointDealNum).AbunData(Sum1)
    PointDealNum = PointDealNum + 1
End Sub
Private Sub FivDealIOP(EleLabel As Integer, AtomNum As Integer)
   Dim i As Integer, j As Integer, k As Integer, g As Integer, sum As Integer, Sum1 As Double, ValueSav As Double
    Sum1 = 0
    If TotleElemTable(EleLabel).ElementIopData(TotleElemTable(EleLabel).IopNum - 1, 1) < 1 Then Call FurDealIOP(EleLabel, AtomNum): Exit Sub
    
    sum = AtomNum + 100
    ReDim EleFurProData(PointDealNum).MassData(sum)
    ReDim EleFurProData(PointDealNum).AbunData(sum)
    For i = 0 To AtomNum
       For j = 0 To AtomNum - i
         For k = 0 To AtomNum - i - j
            For g = 0 To AtomNum - i - j - k - g
              If DieCycleBool Then Exit Sub
              If Sum1 > sum - 1 Then sum = sum + 50: ReDim Preserve EleFurProData(PointDealNum).MassData(sum): ReDim Preserve EleFurProData(PointDealNum).AbunData(sum)
              ValueSav = ProSecData(AtomNum, i) * (TotleElemTable(EleLabel).ElementIopData(0, 1) / 100) ^ i * ProSecData(AtomNum - i, j) * (TotleElemTable(EleLabel).ElementIopData(1, 1) / 100) ^ j * ProSecData(AtomNum - i - j, k) * (TotleElemTable(EleLabel).ElementIopData(2, 1) / 100) ^ k * ProSecData(AtomNum - i - j - k, g) * (TotleElemTable(EleLabel).ElementIopData(3, 1) / 100) ^ g * (TotleElemTable(EleLabel).ElementIopData(4, 1) / 100) ^ (AtomNum - i - j - k - g)
              If ValueSav > AbunDataContant Then
                EleFurProData(PointDealNum).MassData(Sum1) = i * TotleElemTable(EleLabel).ElementIopData(0, 0) + j * TotleElemTable(EleLabel).ElementIopData(1, 0) + k * TotleElemTable(EleLabel).ElementIopData(2, 0) + g * TotleElemTable(EleLabel).ElementIopData(3, 0) + (AtomNum - i - j - k - g) * TotleElemTable(EleLabel).ElementIopData(4, 0)
                EleFurProData(PointDealNum).AbunData(Sum1) = ValueSav
                Sum1 = Sum1 + 1
              End If
    Next g, k, j, i
    
     EleFurProData(PointDealNum).num = Sum1
     ReDim Preserve EleFurProData(PointDealNum).MassData(Sum1): ReDim Preserve EleFurProData(PointDealNum).AbunData(Sum1)
     PointDealNum = PointDealNum + 1
End Sub
Private Sub SixDealIOP(EleLabel As Integer, AtomNum As Integer)
Dim i As Integer, j As Integer, k As Integer, g As Integer, m As Integer, sum As Integer, Sum1 As Integer, ValueSav As Double
    Sum1 = 0
    If TotleElemTable(EleLabel).ElementIopData(TotleElemTable(EleLabel).IopNum - 1, 1) < 1 Then Call FivDealIOP(EleLabel, AtomNum): Exit Sub
    sum = AtomNum + 100
    ReDim EleFurProData(PointDealNum).MassData(sum)
    ReDim EleFurProData(PointDealNum).AbunData(sum)
    For i = 0 To AtomNum
       For j = 0 To AtomNum - i
         For k = 0 To AtomNum - i - j
            For g = 0 To AtomNum - i - j - k - g
                 For m = 0 To AtomNum - i - j - k - g - m
                    If DieCycleBool Then Exit Sub
                    If Sum1 > sum - 1 Then sum = sum + 50: ReDim Preserve EleFurProData(PointDealNum).MassData(sum): ReDim Preserve EleFurProData(PointDealNum).AbunData(sum)
                    ValueSav = ProSecData(AtomNum, i) * (TotleElemTable(EleLabel).ElementIopData(0, 1) / 100) ^ i * ProSecData(AtomNum - i, j) * (TotleElemTable(EleLabel).ElementIopData(1, 1) / 100) ^ j * ProSecData(AtomNum - i - j, k) * (TotleElemTable(EleLabel).ElementIopData(2, 1) / 100) ^ k * ProSecData(AtomNum - i - j - k, g) * (TotleElemTable(EleLabel).ElementIopData(3, 1) / 100) ^ g * ProSecData(AtomNum - i - j - k - g, m) * (TotleElemTable(EleLabel).ElementIopData(4, 1) / 100) ^ m * (TotleElemTable(EleLabel).ElementIopData(5, 1) / 100) ^ (AtomNum - i - j - k - g - m)
                    If ValueSav > AbunDataContant Then
                        EleFurProData(PointDealNum).MassData(Sum1) = i * TotleElemTable(EleLabel).ElementIopData(0, 0) + j * TotleElemTable(EleLabel).ElementIopData(1, 0) + k * TotleElemTable(EleLabel).ElementIopData(2, 0) + g * TotleElemTable(EleLabel).ElementIopData(3, 0) + m * TotleElemTable(EleLabel).ElementIopData(4, 0) + (AtomNum - i - j - k - g) * TotleElemTable(EleLabel).ElementIopData(5, 0)
                        EleFurProData(PointDealNum).AbunData(Sum1) = ValueSav
                        Sum1 = Sum1 + 1
                    End If
    Next m, g, k, j, i
    
     EleFurProData(PointDealNum).num = Sum1
     ReDim Preserve EleFurProData(PointDealNum).MassData(Sum1): ReDim Preserve EleFurProData(PointDealNum).AbunData(Sum1)
     PointDealNum = PointDealNum + 1
End Sub
Private Sub SevDealIOP(EleLabel As Integer, AtomNum As Integer)
Dim i As Integer, j As Integer, k As Integer, g As Integer, m As Integer, n As Integer, sum As Integer, Sum1 As Integer, ValueSav As Double
    Sum1 = 0
    If TotleElemTable(EleLabel).ElementIopData(TotleElemTable(EleLabel).IopNum - 1, 1) < 1 Then Call SixDealIOP(EleLabel, AtomNum): Exit Sub
    sum = AtomNum + 100
    ReDim EleFurProData(PointDealNum).MassData(sum)
    ReDim EleFurProData(PointDealNum).AbunData(sum)
    
    For i = 0 To AtomNum
       For j = 0 To AtomNum - i
         For k = 0 To AtomNum - i - j
            For g = 0 To AtomNum - i - j - k - g
                 For m = 0 To AtomNum - i - j - k - g - m
                    For n = 0 To AtomNum - i - j - k - g - m - n
                        If DieCycleBool Then Exit Sub
                        If Sum1 > sum - 1 Then sum = sum + 50: ReDim Preserve EleFurProData(PointDealNum).MassData(sum): ReDim Preserve EleFurProData(PointDealNum).AbunData(sum)
                        ValueSav = ProSecData(AtomNum, i) * (TotleElemTable(EleLabel).ElementIopData(0, 1) / 100) ^ i * ProSecData(AtomNum - i, j) * (TotleElemTable(EleLabel).ElementIopData(1, 1) / 100) ^ j * ProSecData(AtomNum - i - j, k) * (TotleElemTable(EleLabel).ElementIopData(2, 1) / 100) ^ k * ProSecData(AtomNum - i - j - k, g) * (TotleElemTable(EleLabel).ElementIopData(3, 1) / 100) ^ g * ProSecData(AtomNum - i - j - k - g, m) * (TotleElemTable(EleLabel).ElementIopData(4, 1) / 100) ^ m * ProSecData(AtomNum - i - j - k - g - m, n) * (TotleElemTable(EleLabel).ElementIopData(5, 1) / 100) ^ n * (TotleElemTable(EleLabel).ElementIopData(6, 1) / 100) ^ (AtomNum - i - j - k - g - m - n)
                        If ValueSav > AbunDataContant Then
                            EleFurProData(PointDealNum).MassData(Sum1) = i * TotleElemTable(EleLabel).ElementIopData(0, 0) + j * TotleElemTable(EleLabel).ElementIopData(1, 0) + k * TotleElemTable(EleLabel).ElementIopData(2, 0) + g * TotleElemTable(EleLabel).ElementIopData(3, 0) + m * TotleElemTable(EleLabel).ElementIopData(4, 0) + n * TotleElemTable(EleLabel).ElementIopData(5, 0) + (AtomNum - i - j - k - g - m - n) * TotleElemTable(EleLabel).ElementIopData(6, 0)
                            EleFurProData(PointDealNum).AbunData(Sum1) = ValueSav
                            Sum1 = Sum1 + 1
                       End If
    Next n, m, g, k, j, i
    EleFurProData(PointDealNum).num = Sum1
     
     ReDim Preserve EleFurProData(PointDealNum).MassData(Sum1): ReDim Preserve EleFurProData(PointDealNum).AbunData(Sum1)
     PointDealNum = PointDealNum + 1
     
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
H1:  bool = MsgBox("对不起！你输入的分子存在同位素峰太多了，程序无法计算", vbInformation + vbOKCancel, "消息提示")
     DieCycleBool = True
    If bool Then
        Unload Me
    Else
        Unload Me
    End If
    Call ElementTable.Exit_Click
End Function

Private Sub Calculator_Click()
    If Shell("C:\WINDOWS\system32\calc.exe", 1) Then
      Exit Sub
    Else
       MsgBox "无法与计算器进行关联", vbInformation + vbExclamation
    End If
End Sub

Private Sub Distribution_Click()
MassPicture.Visible = False
Distr_Frame.Visible = True
LabelCommand.Enabled = False
Save_Picture.Enabled = False
End Sub

Private Sub Exit_Click()
    Unload Me
End Sub
Private Sub File_Click()
    Set Picture2.Picture = CaptureScreen()
    
End Sub
Private Sub Form_Activate()
    MinXcor = 0:  MaxXcor = 100:  MinYcor = 0:  MaxYcor = 100
    CoorWidth = MaxXcor - MinXcor
    CoorHeight = MaxYcor - MinYcor
    ZoomX = MassPicture.Width / CoorWidth
    ZoomY = MassPicture.Height / CoorHeight
    'Printer1.Enabled = False
    PictCoor.Xmax = MassPicture.Width - 450
    PictCoor.xMin = 450
    PictCoor.Ymax = MassPicture.Height - 500
    PictCoor.yMin = 800
    If isDealWithData Then
       Call Initiate
       Call ProcData
    End If
 End Sub
Private Sub LabelPic(rev As Single)
    Dim i As Integer, m As Single
    m = rev / 100
    
    If LabelBool = False Then
      Call PrinPic
      For i = 0 To PaintNum - 1
         If PaintData(i).Abundance >= m And PaintData(i).MassData > MinXcor And PaintData(i).MassData < MaxXcor Then
          MassPicture.CurrentX = (PaintData(i).MassData - MinXcor) * ZoomX + PictCoor.xMin
          MassPicture.CurrentY = PictCoor.Ymax - PaintData(i).Abundance * 100 * ZoomY
          MassPicture.Print Round(PaintData(i).MassData, 2)
        End If
      Next i
      Set Picture2.Picture = Nothing
      Set Picture2.Picture = CaptureScreen()
      LabelBool = True
      LabelCommand.Caption = "除去标记": Exit Sub
    Else
       Call PrinPic
       LabelCommand.Caption = "标记谱图"
       LabelBool = False
    End If
End Sub
Public Sub Repeat_Paint()
Call AdjustCoor
Call PrinPic
End Sub
Private Sub Coordinate()      '画坐标和刻度
    Dim StepX As Double, StepY As Double, i As Integer, Ymax As Long, Xmax As Long, j As Integer, currex As Double, currey As Double
    Dim Xcoor(10) As Double, Ycoor(10) As Double
    
    CoorWidth = MaxXcor - MinXcor                           '计算刻度
    CoorHeight = MaxYcor - MinYcor
    For i = 0 To 9
         Xcoor(i) = ((MaxXcor - MinXcor) / 10) * (i + 1) + MinXcor
         Ycoor(i) = ((MaxYcor - MinYcor) / 10) * (i + 1) + MinYcor
    Next
    
    StepX = (PictCoor.Xmax - PictCoor.xMin) / 10
    StepY = (PictCoor.Ymax - PictCoor.yMin) / 10
    Xmax = PictCoor.Xmax
    Ymax = PictCoor.Ymax
    MassPicture.Line (PictCoor.xMin, PictCoor.Ymax)-(PictCoor.Xmax, PictCoor.Ymax), vbBlue                          '横坐标轴
    MassPicture.Line (PictCoor.xMin, PictCoor.yMin)-(PictCoor.xMin, PictCoor.Ymax), vbRed             '众坐标轴
    For i = 0 To 9
        MassPicture.Line (PictCoor.xMin, Ymax - (i + 1) * StepY)-(PictCoor.xMin - 120, Ymax - (i + 1) * StepY), vbRed
        currex = MassPicture.CurrentX: currey = MassPicture.CurrentY
          For j = 1 To 9                                               '纵坐标刻度
             If j <> 5 Then
                MassPicture.Line (PictCoor.xMin, Ymax - i * StepY - j * StepY / 10)-(PictCoor.xMin - 80, Ymax - i * StepY - j * StepY / 10), vbRed
             Else
                MassPicture.Line (PictCoor.xMin, Ymax - i * StepY - j * StepY / 10)-(PictCoor.xMin - 100, Ymax - i * StepY - j * StepY / 10), vbRed
             End If
          Next j
        MassPicture.CurrentX = currex - 300: MassPicture.CurrentY = currey - 70
        MassPicture.Print Ycoor(i)
        MassPicture.Line (PictCoor.xMin + StepX * (i + 1), Ymax + 20)-(PictCoor.xMin + StepX * (i + 1), Ymax + 20 + 100), vbBlue   '横坐标刻度
        currex = MassPicture.CurrentX: currey = MassPicture.CurrentY
          For j = 1 To 9
             If j <> 5 Then
                MassPicture.Line (PictCoor.xMin + StepX * i + j * StepX / 10, Ymax + 20)-(PictCoor.xMin + StepX * i + j * StepX / 10, Ymax + 20 + 50), vbBlue
             Else
                MassPicture.Line (PictCoor.xMin + StepX * i + j * StepX / 10, Ymax + 20)-(PictCoor.xMin + StepX * i + j * StepX / 10, Ymax + 20 + 80), vbBlue
             End If
          Next j
        MassPicture.CurrentX = currex - 180: MassPicture.CurrentY = currey + 20
        MassPicture.Print Format(Xcoor(i), "Fixed")
    Next i
End Sub
Private Sub PrinPic()
                         '画图
    On Error GoTo H1
    Dim i As Integer, show_Num As Long
    show_Num = 0
    MassPicture.Cls
    Call Coordinate
    
    ZoomX = (PictCoor.Xmax - PictCoor.xMin) / CoorWidth
    ZoomY = (PictCoor.Ymax - PictCoor.yMin) / CoorHeight
    For i = 0 To PaintNum - 1
         If PaintData(i).MassData > MinXcor And PaintData(i).MassData < MaxXcor Then
            MassPicture.Line ((PaintData(i).MassData - MinXcor) * ZoomX + PictCoor.xMin, PictCoor.Ymax)-((PaintData(i).MassData - MinXcor) * ZoomX + PictCoor.xMin, PictCoor.Ymax - PaintData(i).Abundance * 100 * ZoomY), vbRed
            show_Num = show_Num + 1
         End If
    Next i
    
    Call ShowMessage(show_Num)
    MassPicture.CurrentX = 1000: MassPicture.CurrentY = 300
    If SaveResValue <> 0 Then
      MassPicture.Print "化学式：" & " " & ChemForm.Caption & "   " & "分辨率：" & " " & Format(SaveResValue, "Scientific")
    Else
      MassPicture.Print "化学式：" & " " & ChemForm.Caption & "   " & "分辨率：" & " " & "Idel"
    End If
    'Set Picture2.Picture = Nothing
    'Set Picture2.Picture = CaptureScreen()             'Save Picture
H1:   Exit Sub
End Sub
Private Sub ShowMessage(show_Num As Long)
    Dim i As Integer, str As String
    
    ChemForm.Caption = ""
    TotalNum.Caption = LastNum
    ShowNum.Caption = show_Num
    For i = 0 To FormAtomNum - 1
        str = TotleElemTable(ProcDataElem(i).PointNum).ElementShorName & ProcDataElem(i).num
        ChemForm.Caption = ChemForm.Caption & str
    Next i

End Sub
Private Sub GridInitate(rowNum As Integer, max_data As Double)
    Dim i As Integer, j As Integer, row_Num As Integer
    row_Num = 13
    If rowNum >= 13 Then row_Num = rowNum
    
        MSFlexGrid1.Cols = 4: MSFlexGrid1.Rows = row_Num
        MSFlexGrid1.Col = 0
        MSFlexGrid1.Row = 0
        MSFlexGrid1.Text = "      m/z"
        MSFlexGrid1.Col = 1
        MSFlexGrid1.Row = 0
        MSFlexGrid1.Text = "    Abundance"
        MSFlexGrid1.Col = 2
        MSFlexGrid1.Row = 0
        MSFlexGrid1.Text = "    Spread"
        MSFlexGrid1.Col = 3
        MSFlexGrid1.Row = 0
        MSFlexGrid1.Text = "  Multiplicity"
    
    For i = 0 To 3
      MSFlexGrid1.ColWidth(i) = 1550
    Next i
    
    For i = 0 To row_Num - 1
      MSFlexGrid1.RowHeight(i) = 400
    Next i
    
    For i = 1 To rowNum - 1
          MSFlexGrid1.Col = 0: MSFlexGrid1.Row = i
          MSFlexGrid1.Text = "   " & FormatNumber(PaintData(i - 1).MassData, 5, vbTrue) & "   "     '第一列数据
          MSFlexGrid1.Col = 1: MSFlexGrid1.Row = i
          MSFlexGrid1.Text = FormatNumber(PaintData(i - 1).Abundance * max_data, 8, vbTrue)     '第二列数据
          MSFlexGrid1.Col = 2: MSFlexGrid1.Row = i
          MSFlexGrid1.Text = FormatNumber(PaintData(i - 1).Abundance, 8, vbTrue)                '第三列数据
          MSFlexGrid1.Col = 3: MSFlexGrid1.Row = i
          MSFlexGrid1.Text = PaintData(i - 1).Labelpoint
    Next i
End Sub

Private Sub Form_Deactivate()
'isDealWithData = False
End Sub
Private Sub Form_Paint()
isDealWithData = False
End Sub

Private Sub Form_Resize()
    ElementTable.Show
    isDealWithData = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ElementTable.Show
End Sub
Private Sub Help_Click()
On Error GoTo H1
  Dim pat As String
  pat = App.Path & "\HELP.CHM"
  Shell "hh.exe " & pat, vbNormalFocus
  Exit Sub
H1: MsgBox Err.Description
End Sub

Private Sub LabelCommand_Click()
Call LabelPic(Label_Value)
End Sub

Private Sub More_1_Click()
    LabelBool = False
    Label_Value = 1
    More_1.Caption = ">1 Y"
    More_10.Caption = ">10"
    More_20.Caption = ">20"
    More_30.Caption = ">30"
    More_40.Caption = ">40"
    More_50.Caption = ">50"
    More_60.Caption = ">60"
    More_70.Caption = ">70"
    More_80.Caption = ">80"
    More_90.Caption = ">90"
    
    Call LabelPic(1)
End Sub
Private Sub More_10_Click()
    LabelBool = False
    Label_Value = 10
    More_1.Caption = ">1 "
    More_10.Caption = ">10 Y"
    More_20.Caption = ">20"
    More_30.Caption = ">30"
    More_40.Caption = ">40"
    More_50.Caption = ">50"
    More_60.Caption = ">60"
    More_70.Caption = ">70"
    More_80.Caption = ">80"
    More_90.Caption = ">90"
    
    Call LabelPic(10)
End Sub
Private Sub More_20_Click()
    LabelBool = False
    Label_Value = 20
    More_1.Caption = ">1 "
    More_10.Caption = ">10"
    More_20.Caption = ">20 Y"
    More_30.Caption = ">30"
    More_40.Caption = ">40"
    More_50.Caption = ">50"
    More_60.Caption = ">60"
    More_70.Caption = ">70"
    More_80.Caption = ">80"
    More_90.Caption = ">90"
    Call LabelPic(20)
End Sub

Private Sub More_30_Click()
    LabelBool = False
    Label_Value = 30
    More_1.Caption = ">1 "
    More_10.Caption = ">10"
    More_20.Caption = ">20"
    More_30.Caption = ">30 Y"
    More_40.Caption = ">40"
    More_50.Caption = ">50"
    More_60.Caption = ">60"
    More_70.Caption = ">70"
    More_80.Caption = ">80"
    More_90.Caption = ">90"
    Call LabelPic(30)
End Sub

Private Sub More_40_Click()
    LabelBool = False
    Label_Value = 40
    More_1.Caption = ">1 "
    More_10.Caption = ">10"
    More_20.Caption = ">20"
    More_30.Caption = ">30"
    More_40.Caption = ">40 Y"
    More_50.Caption = ">50"
    More_60.Caption = ">60"
    More_70.Caption = ">70"
    More_80.Caption = ">80"
    More_90.Caption = ">90"
    Call LabelPic(40)
End Sub

Private Sub More_50_Click()
    LabelBool = False
    Label_Value = 50
    More_1.Caption = ">1 "
    More_10.Caption = ">10"
    More_20.Caption = ">20"
    More_30.Caption = ">30"
    More_40.Caption = ">40"
    More_50.Caption = ">50 Y"
    More_60.Caption = ">60"
    More_70.Caption = ">70"
    More_80.Caption = ">80"
    More_90.Caption = ">90"
    Call LabelPic(50)
End Sub

Private Sub More_60_Click()
    LabelBool = False
    Label_Value = 60
    More_1.Caption = ">1 "
    More_10.Caption = ">10"
    More_20.Caption = ">20"
    More_30.Caption = ">30"
    More_40.Caption = ">40"
    More_50.Caption = ">50"
    More_60.Caption = ">60 Y"
    More_70.Caption = ">70"
    More_80.Caption = ">80"
    More_90.Caption = ">90"
    Call LabelPic(60)
End Sub

Private Sub More_70_Click()
    LabelBool = False
    Label_Value = 70
    More_1.Caption = ">1 "
    More_10.Caption = ">10"
    More_20.Caption = ">20"
    More_30.Caption = ">30"
    More_40.Caption = ">40"
    More_50.Caption = ">50"
    More_60.Caption = ">60"
    More_70.Caption = ">70 Y"
    More_80.Caption = ">80"
    More_90.Caption = ">90"
    Call LabelPic(70)
End Sub

Private Sub More_80_Click()
    LabelBool = False
    Label_Value = 80
    More_1.Caption = ">1 "
    More_10.Caption = ">10"
    More_20.Caption = ">20"
    More_30.Caption = ">30"
    More_40.Caption = ">40"
    More_50.Caption = ">50"
    More_60.Caption = ">60"
    More_70.Caption = ">70"
    More_80.Caption = ">80 Y"
    More_90.Caption = ">90"
    Call LabelPic(80)
End Sub
Private Sub More_90_Click()
    LabelBool = False
    Label_Value = 90
    More_1.Caption = ">1 "
    More_10.Caption = ">10"
    More_20.Caption = ">20"
    More_30.Caption = ">30"
    More_40.Caption = ">40"
    More_50.Caption = ">50"
    More_60.Caption = ">60"
    More_70.Caption = ">70"
    More_80.Caption = ">80"
    More_90.Caption = ">90 Y"
    Call LabelPic(90)
End Sub
Private Sub MassPicture_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
       PopupMenu PopManu
    End If
End Sub

Private Sub Printer1_Click()
'On Error GoTo h1

    Set Picture2.Picture = CaptureScreen()
    
    PrintPictureToFitPage Printer, Picture2.Picture
        Printer.EndDoc
    Set Picture2 = Nothing
'h1: Exit Sub
End Sub

Private Sub Resolv_00003_Click()
    Resolv_003.Caption = "0.03"
    Resolv_0003.Caption = "0.003"
    Resolv_03.Caption = "0.3"
    Resolv_1.Caption = "1.0"
    Resolv_00003.Caption = "0.0003" & "  " & "Y"
    Resolv_05.Caption = "0.5"
    Resolv_Ide.Caption = "Ideal"
    
    Call ProcDataFur(0.0003)
     Call AdjustCoor
     Call PrinPic
End Sub

Private Sub Resolv_0003_Click()
    Resolv_003.Caption = "0.03"
    Resolv_0003.Caption = "0.003" & "  " & "Y"
    Resolv_03.Caption = "0.3"
    Resolv_1.Caption = "1.0"
    Resolv_00003.Caption = "0.0003"
    Resolv_05.Caption = "0.5"
    Resolv_Ide.Caption = "Ideal"
    
     Call ProcDataFur(0.003)
     Call AdjustCoor
     Call PrinPic
End Sub

Private Sub Resolv_003_Click()
    Resolv_003.Caption = "0.03" & "  " & "Y"
    Resolv_0003.Caption = "0.003"
    Resolv_03.Caption = "0.3"
    Resolv_1.Caption = "1.0"
    Resolv_00003.Caption = "0.0003"
    Resolv_05.Caption = "0.5"
    Resolv_Ide.Caption = "Ideal"
    
     Call ProcDataFur(0.03)
     Call AdjustCoor
     Call PrinPic
End Sub

Private Sub Resolv_03_Click()
    Resolv_003.Caption = "0.03"
    Resolv_0003.Caption = "0.003"
    Resolv_03.Caption = "0.3" & "  " & "Y"
    Resolv_1.Caption = "1.0"
    Resolv_00003.Caption = "0.0003"
    Resolv_05.Caption = "0.5"
    Resolv_Ide.Caption = "Ideal"
    
     Call ProcDataFur(0.3)
     Call AdjustCoor
     Call PrinPic
End Sub

Private Sub Resolv_05_Click()
    Resolv_003.Caption = "0.03"
    Resolv_0003.Caption = "0.003"
    Resolv_03.Caption = "0.3"
    Resolv_1.Caption = "1.0"
    Resolv_00003.Caption = "0.0003"
    Resolv_05.Caption = "0.5" & "  " & "Y"
    Resolv_Ide.Caption = "Ideal"
     Call ProcDataFur(0.5)
     Call AdjustCoor
     Call PrinPic
End Sub

Private Sub Resolv_1_Click()
    Resolv_003.Caption = "0.03"
    Resolv_0003.Caption = "0.003"
    Resolv_03.Caption = "0.3"
    Resolv_1.Caption = "1.0" & "  " & "Y"
    Resolv_00003.Caption = "0.0003"
    Resolv_05.Caption = "0.5"
    Resolv_Ide.Caption = "Ideal"
     Call ProcDataFur(1)
     Call AdjustCoor
     Call PrinPic
End Sub
Private Sub Resolv_Ide_Click()
    Resolv_003.Caption = "0.03"
    Resolv_0003.Caption = "0.003"
    Resolv_03.Caption = "0.3"
    Resolv_1.Caption = "1.0"
    Resolv_00003.Caption = "0.0003"
    Resolv_05.Caption = "0.5"
    Resolv_Ide.Caption = "Ideal" & "  " & "Y"

    Call ProcDataFur(0)
    Call AdjustCoor
    Call PrinPic
End Sub
Private Sub Save_Data_Click()
On Error GoTo h10
    Dim i As Integer, file_num As Integer, file_path As String, j As Integer
        
        CommonDialog1.DefaultExt = ".txt"
        CommonDialog1.Filter = "txt (＊.txt)|＊.txt"
        CommonDialog1.InitDir = App.Path & "\IOPMassSimulation Data"
        CommonDialog1.ShowSave
        
        
        If CommonDialog1.FileName <> "" Then
           file_path = CommonDialog1.FileName: file_num = FreeFile()
           Open file_path For Output As file_num
                Print #file_num, "The IOP Mass Simulation Data Of " & ChemForm.Caption
                Print #file_num, "**********************************************************************"
                Print #file_num, "The abundance value of all kinds of IOP elements in the different atom is as follow"
                For i = 0 To FormAtomNum - 1
                     Print #file_num, TotleElemTable(ProcDataElem(i).PointNum).ElementShorName
                      For j = 0 To TotleElemTable(ProcDataElem(i).PointNum).IopNum - 1
                           Print #file_num, FormatNumber(TotleElemTable(ProcDataElem(i).PointNum).ElementIopData(j, 0), 6, vbTrue), FormatNumber(TotleElemTable(ProcDataElem(i).PointNum).ElementIopData(j, 1) / 100, 6, vbTrue)
                           
                      Next j
                Next i
                Print #file_num, "**********************************************************************"
                Print #file_num, "The distribution of abundance in " & ChemForm.Caption & " is as follow"
                Print #file_num, "  M / Z", " Abundance", "   Spread", "   Multiplicity"
                For i = 0 To PaintNum - 1
                    Print #file_num, FormatNumber(PaintData(i).MassData, 6, vbTrue), FormatNumber(PaintData(i).Abundance * DataMax_Abundance, 8, vbTrue), FormatNumber(PaintData(i).Abundance, 8, vbTrue), "       " & PaintData(i).Labelpoint
                Next i
            Close #file_num
        End If
        
      Exit Sub
      
h10: MsgBox Error.Description
End Sub

Private Sub Save_Picture_Click()
   
    CommonDialog1.DefaultExt = ".BMP"
    CommonDialog1.Filter = "Bitmap Image (＊.bmp)|＊.bmp"
    CommonDialog1.InitDir = App.Path
    CommonDialog1.ShowSave
    If CommonDialog1.FileName <> "" Then
        SavePicture Picture2.Picture, CommonDialog1.FileName
    End If
Set Picture2.Picture = Nothing    '清楚内存空间
End Sub
Private Sub Spectrum_Click()
MassPicture.Visible = True
Distr_Frame.Visible = False
LabelCommand.Enabled = True
Save_Picture.Enabled = True
End Sub
Private Sub Taskmgr_Click()
    If Shell("C:\WINDOWS\system32\taskmgr.exe", 1) Then
       Exit Sub
    Else
       MsgBox "无法与任务管理器进行关联", vbInformation + vbExclamation
    End If
End Sub
Private Sub User_Defined_Click()
On Error GoTo hh1
Dim value As Double, aa As String
aa = InputBox("Input the value of Resolution" & " " & ChemForm.Caption, "Set the value of Resolution", FormatNumber(0.3, 2, vbTrue))

If aa = "" Then Exit Sub

value = Val(aa)
Select Case value
       Case 1
          Call Resolv_1_Click
       Case 0.5
          Call Resolv_05_Click
       Case 0.3
          Call Resolv_03_Click
       Case 0.03
          Call Resolv_003_Click
       Case 0.003
          Call Resolv_1_Click
       Case 0.0003
          Call Resolv_00003_Click
       Case Else
            Resolv_003.Caption = "0.03"
            Resolv_0003.Caption = "0.003"
            Resolv_03.Caption = "0.3"
            Resolv_1.Caption = "1.0"
            Resolv_00003.Caption = "0.0003"
            Resolv_05.Caption = "0.5"
            Resolv_Ide.Caption = "Ideal"
            Call ProcDataFur(value)
            Call AdjustCoor
            Call PrinPic
End Select
Exit Sub

hh1: MsgBox "Please input the value"
End Sub

Private Sub ZoomSmall_Click()
    MinXcor = MinXcor - CoorWidth / 2: MaxXcor = MaxXcor + CoorWidth / 2
    
    Call Coordinate
    Call PrinPic

End Sub
Private Sub ZoomLarge_Click()

MinXcor = Fix(minData): MaxXcor = Fix(maxData)

Call Coordinate
Call PrinPic
End Sub
Private Sub AdjustCoor()
    Dim i As Integer, mmi As Double
    
    
    maxData = PaintData(0).MassData
    minData = PaintData(0).MassData
    For i = 0 To PaintNum - 1
        If maxData <= PaintData(i).MassData And PaintData(i).Abundance > 0.001 Then maxData = PaintData(i).MassData
        If minData >= PaintData(i).MassData And PaintData(i).Abundance > 0.001 Then minData = PaintData(i).MassData
    Next i
    
    
    MinXcor = minData - 1: IniXmin = minData - 1
    
    MaxXcor = maxData + 1: IniXmax = maxData + 1
    
    Call Coordinate
End Sub
Private Sub MassPicture_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo H1
Dim i As Double
    If Button = 1 Then
       If Zoom = False Then
          Zoom = True
          X1 = X
          Y1 = Y
          
          Line1.X1 = X
          Line1.Y1 = Y
          Line1.X2 = X
          Line1.Y2 = Y
          Line1.Visible = True
          
          Line2.X1 = X
          Line2.Y1 = Y
          Line2.X2 = X
          Line2.Y2 = Y
          Line2.Visible = True
          
          Line3.X1 = X
          Line3.Y1 = Y
          Line3.X2 = X
          Line3.Y2 = Y
          Line3.Visible = True
          
          Line4.X1 = X
          Line4.Y1 = Y
          Line4.X2 = X
          Line4.Y2 = Y
          Line4.Visible = True
       Else
          X2 = X1
          Y2 = Y1
          X1 = X
          Y1 = Y
          If X1 > X2 Then
             i = X1
             X1 = X2
             X2 = i
          End If
          If X1 = X2 Then X2 = X2 + 1
          If Y1 > Y2 Then
             i = Y1
             Y1 = Y2
             Y2 = i
          End If
          If Y1 = Y2 Then Y2 = Y2 + 1
          Line4.Visible = False
          Line3.Visible = False
          Line2.Visible = False
          Line1.Visible = False
          Zoom = False
          Call JustCoor(X1, X2, Y1, Y2)
          
        End If
    End If
  Exit Sub
H1: Call Coordinate
   Exit Sub
End Sub
Private Sub JustCoor(XX1 As Double, XX2 As Double, YY1 As Double, YY2 As Double)
    If XX1 = XX2 Then GoTo H1
    
        Dim xMin As Double, yMin As Double     'reserve the startponting of last Coordinate
        xMin = MinXcor: yMin = MinYcor
        MinXcor = (XX1 - PictCoor.xMin) / ZoomX + xMin
        MaxXcor = (XX2 - PictCoor.xMin) / ZoomX + xMin
        MinYcor = 0
        MaxYcor = 100
        MassPicture.Cls
              Call Coordinate
              Call PrinPic
H1:     Exit Sub
End Sub
Private Sub MassPicture_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Zoom Then
       Line1.Y2 = Y
       
       Line2.X2 = X
       
       Line3.X1 = X
       Line3.X2 = X
       Line3.Y1 = Y1
       Line3.Y2 = Y
       
       Line4.Y1 = Y
       Line4.Y2 = Y
       Line4.X1 = X1
       Line4.X2 = X
    End If
End Sub


