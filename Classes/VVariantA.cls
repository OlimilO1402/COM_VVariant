VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VVariantPtr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements VVariant
'https://docs.microsoft.com/en-us/windows/win32/api/propidlbase/ns-propidlbase-propvariant
'https://devblogs.microsoft.com/oldnewthing/20171221-00/?p=97625

Private Type TVar
    Value As Variant
End Type

Private Type TVariantPtr
    VarVal As TVar
    Ptr()  As TVar
    udtptr As TUDTPtr
End Type

Private m_Var As TVariantPtr

Private Declare Sub RtlMoveMemory Lib "kernel32" (pDst As Any, pSrc As Any, ByVal BytLen As Long)

Private Sub Class_Initialize()
    MPtr.New_UDTPtr m_Var.udtptr, FADF_RECORD Or FADF_FIXEDSIZE, MVariant.Variant_Size
    ResetPtr
    SAPtr(ArrPtr(m_Var.Ptr)) = m_Var.udtptr.pSA
End Sub

Private Sub Class_Terminate()
    ZeroSAPtr ArrPtr(m_Var.Ptr)
End Sub

Friend Sub New_(aValue)
    m_Var.Ptr(0).Value = aValue
End Sub

Friend Sub NewVt(Optional ByVal vt As EVbVarType = vbIllegal, Optional aValue)
    New_ aValue
    Me.VarType = vt
End Sub

Friend Property Let Ptr(ByVal Value As LongPtr)
    m_Var.udtptr.pvData = Value
End Property
Friend Property Get Ptr() As LongPtr
    Ptr = m_Var.udtptr.pvData
End Property
Public Sub ResetPtr()
    m_Var.udtptr.pvData = VarPtr(m_Var)
End Sub

Public Property Let VarType(Value As EVbVarType)
    RtlMoveMemory ByVal Me.Ptr, Value, 2
End Property
Public Property Get VarType() As EVbVarType
    RtlMoveMemory VarType, ByVal Me.Ptr, 2
End Property
Public Function VarTypeToStr() As String
    VarTypeToStr = MEVbVarType.VarType2_ToStr(m_Var.Ptr(0).Value)
End Function

Public Function GetByte() As Byte
    GetMem1 ByVal Me.Ptr + 8, GetByte
End Function

Public Function GetInteger() As Integer
    GetMem2 ByVal Me.Ptr + 8, GetInteger
End Function

Public Function GetLong() As Long
    GetMem4 ByVal Me.Ptr + 8, GetLong
End Function

Public Function GetSingle() As Single
    GetMem4 ByVal Me.Ptr + 8, GetSingle
End Function

Public Function GetDouble() As Double
    GetMem8 ByVal Me.Ptr + 8, GetDouble
End Function

Public Function GetCurrency() As Currency
    GetMem8 ByVal Me.Ptr + 8, GetCurrency
End Function

Public Function GetBoolean() As Boolean
    GetMem2 ByVal Me.Ptr + 8, GetBoolean
End Function

Public Function GetDate() As Date
    GetMem8 ByVal Me.Ptr + 8, GetDate
End Function

#If VBA7 Then
Public Function GetLongLong() As LongLong
    GetMem8 ByVal Me.Ptr + 8, GetLongLong
End Function
#End If

Public Function GetLongPtr() As LongPtr
    RtlMoveMemory GetLongPtr, ByVal Me.Ptr + 8, MPtr.LongPtr_Size
End Function

Public Function GetString() As String
    GetString = m_Var.Ptr(0).Value
End Function

Public Function GetObject() As Object
    Set GetObject = m_Var.Ptr(0).Value
End Function

Public Function GetDecimal() As Variant
    RtlMoveMemory GetDecimal, ByVal Me.Ptr, MVariant.Variant_Size
End Function

Public Function ToStr() As String
    Select Case Me.VarType
    Case EVbVarType.vbByte, EVbVarType.vbInteger, EVbVarType.vbLong, EVbVarType.vbSingle, EVbVarType.vbDouble, EVbVarType.vbCurrency, EVbVarType.vbDate, EVbVarType.vbString
        ToStr = CStr(Me.Value)
    Case EVbVarType.vbObject:
    Case Else:
    End Select
End Function

Public Property Get Value() As Variant
    Select Case Me.VarType
    Case EVbVarType.vbByte:       Value = Me.GetByte
    Case EVbVarType.vbInteger:    Value = Me.GetInteger
    Case EVbVarType.vbLong:       Value = Me.GetLong
    Case EVbVarType.vbSingle:     Value = Me.GetSingle
    Case EVbVarType.vbDouble:     Value = Me.GetDouble
    Case EVbVarType.vbCurrency:   Value = Me.GetCurrency
    Case EVbVarType.vbDate:       Value = Me.GetDate
    Case EVbVarType.vbString:     Value = Me.GetString
    Case EVbVarType.vbObject: Set Value = Me.GetObject
    Case Else:
    End Select
End Property
Public Property Let Value(aValue)
    m_Var.Ptr(0).Value = aValue
End Property
Public Property Set Value(aValue)
    Set m_Var.Ptr(0).Value = aValue
End Property

Public Function UDTPtrToStr() As String
    UDTPtrToStr = MPtr.UDTPtrToString(m_Var.udtptr)
End Function




' ############################## '    Implements VVariant    ' ############################## '
Private Property Get VVariant_Ptr() As LongPtr
    VVariant_Ptr = Me.Ptr
End Property

Private Property Let VVariant_VarType(ByVal vt As EVbVarType)
    Me.VarType = vt
End Property
Public Property Get VVariant_VarType() As EVbVarType
    VVariant_VarType = Me.VarType
End Property

Private Function VVariant_VarTypeToStr() As String
    VVariant_VarTypeToStr = Me.VarTypeToStr
End Function

'all GetData-functions:
Private Function VVariant_GetByte() As Byte
    VVariant_GetByte = Me.GetByte
End Function

Private Function VVariant_GetInteger() As Integer
    VVariant_GetInteger = Me.GetInteger
End Function

Private Function VVariant_GetLong() As Long
    VVariant_GetLong = Me.GetLong
End Function

#If VBA7 Then
Private Function VVariant_GetLongLong() As LongLong
    VVariant_GetLongLong = Me.GetLongLong
End Function
#End If

Private Function VVariant_GetLongPtr() As LongPtr
    VVariant_GetLongPtr = Me.GetLongPtr
End Function

Private Function VVariant_GetSingle() As Single
    VVariant_GetSingle = Me.GetSingle
End Function

Private Function VVariant_GetDouble() As Double
    VVariant_GetDouble = Me.GetDouble
End Function

Private Function VVariant_GetCurrency() As Currency
    VVariant_GetCurrency = Me.GetCurrency
End Function

Private Function VVariant_GetBoolean() As Boolean
    VVariant_GetBoolean = Me.GetBoolean
End Function

Private Function VVariant_GetDate() As Date
    VVariant_GetDate = Me.GetDate
End Function

Private Function VVariant_GetString() As String
    VVariant_GetString = Me.GetString
End Function

Private Function VVariant_GetObject() As Object
    Set VVariant_GetObject = Me.GetObject
End Function

Private Function VVariant_GetDecimal() As Variant
    VVariant_GetDecimal = Me.GetDecimal
End Function

Private Function VVariant_ToStr() As String
    VVariant_ToStr = Me.ToStr
End Function

Private Property Get VVariant_Value() As Variant
    If Me.VarType = EVbVarType.vbObject Then
        Set VVariant_Value = Me.Value
    Else
        VVariant_Value = Me.Value
    End If
End Property
Private Property Let VVariant_Value(aValue)
    Me.Value = aValue
End Property
Private Property Set VVariant_Value(aValue)
    Set Me.Value = aValue
End Property
