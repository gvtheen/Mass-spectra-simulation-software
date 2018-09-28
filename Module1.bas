Attribute VB_Name = "Module1"

Public Type ElementType
    ElementName As String
    ElementShorName As String
    IopNum As Integer
    ElementIopData() As Double
End Type

Public Type ProcDataType
   PointNum As Integer
   num As Integer
End Type

Public Type DealDataType
   num As Long
   MassData() As Double
   AbunData() As Double
End Type

Public Type DataType
  MassData As Double
  Abundance As Double
  Labelpoint As Integer
End Type

Public ProcDataElem() As ProcDataType                              '存储化学式中各个原子可能存在的分布数和分布情况
Public TotleElemTable() As ElementType                             '元素周期表
Public FormAtomNum As Integer                                      '存储化学式原子个数
Public IOPValueBool As Boolean                                     '判断是否改变元素的丰度
Public ShowChemFormBool As Boolean
Public ShutFormuFormBool As Boolean
Public PaintData() As DataType
Public SaveLast_Chem As String
Public isDealWithData As Boolean
Public DataMax_Abundance As Double, MocularForm As String



