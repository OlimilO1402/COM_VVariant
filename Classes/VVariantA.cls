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
'https://docs.microsoft.com/en-us/windows/win32/api/propidlbase/ns-propidlbase-propvariant
'https://devblogs.microsoft.com/oldnewthing/20171221-00/?p=97625
Implements VVariant

#If Win64 Then
    Private Const Variant_Size As Long = 24
#Else
    Private Const Variant_Size As Long = 16
#End If

Private Const S_OK As Long = 0

#If VBA7 Then
    'all Variant functions of Oleaut32.dll
    'https://docs.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-varabs
    Private Declare PtrSafe Function VarAbs Lib "oleaut32" (ByRef pvarIn As Any, ByRef pvarResult As Any) As Long
    Private Declare PtrSafe Function VarAdd Lib "oleaut32" (ByRef pvarLeft As Any, ByRef pvarRight As Any, ByRef pvarResult As Any) As Long
    Private Declare PtrSafe Function VarAnd Lib "oleaut32" (ByRef pvarLeft As Any, ByRef pvarRight As Any, ByRef pvarResult As Any) As Long
    Private Declare PtrSafe Function VarDiv Lib "oleaut32" (ByRef pvarLeft As Any, ByRef pvarRight As Any, ByRef pvarResult As Any) As Long
    Private Declare PtrSafe Function VarEqv Lib "oleaut32" (ByRef pvarLeft As Any, ByRef pvarRight As Any, ByRef pvarResult As Any) As Long
    Private Declare PtrSafe Function VarFix Lib "oleaut32" (ByRef pvarIn As Any, ByRef pvarResult As Any) As Long
    Private Declare PtrSafe Function VarIdiv Lib "oleaut32" (ByRef pvarLeft As Any, ByRef pvarRight As Any, ByRef pvarResult As Any) As Long
    Private Declare PtrSafe Function VarImp Lib "oleaut32" (ByRef pvarLeft As Any, ByRef pvarRight As Any, ByRef pvarResult As Any) As Long
    Private Declare PtrSafe Function VarInt Lib "oleaut32" (ByRef pvarIn As Any, ByRef pvarResult As Any) As Long
    Private Declare PtrSafe Function VarMod Lib "oleaut32" (ByRef pvarLeft As Any, ByRef pvarRight As Any, ByRef pvarResult As Any) As Long
    Private Declare PtrSafe Function VarMul Lib "oleaut32" (ByRef pvarLeft As Any, ByRef pvarRight As Any, ByRef pvarResult As Any) As Long
    Private Declare PtrSafe Function VarNeg Lib "oleaut32" (ByRef pvarIn As Any, ByRef pvarResult As Any) As Long
    Private Declare PtrSafe Function VarNot Lib "oleaut32" (ByRef pvarIn As Any, ByRef pvarResult As Any) As Long
    Private Declare PtrSafe Function VarOr Lib "oleaut32" (ByRef pvarLeft As Any, ByRef pvarRight As Any, ByRef pvarResult As Any) As Long
    Private Declare PtrSafe Function VarPow Lib "oleaut32" (ByRef pvarLeft As Any, ByRef pvarRight As Any, ByRef pvarResult As Any) As Long
    Private Declare PtrSafe Function VarSub Lib "oleaut32" (ByRef pvarLeft As Any, ByRef pvarRight As Any, ByRef pvarResult As Any) As Long
    Private Declare PtrSafe Function VarXor Lib "oleaut32" (ByRef pvarLeft As Any, ByRef pvarRight As Any, ByRef pvarResult As Any) As Long
    Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (pDst As Any, pSrc As Any, ByVal BytLen As Long)
#Else
    'all Variant functions of Oleaut32.dll
    'https://docs.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-varabs
    Private Declare Function VarAbs Lib "oleaut32" (ByRef pvarIn As Any, ByRef pvarResult As Any) As Long
    Private Declare Function VarAdd Lib "oleaut32" (ByRef pvarLeft As Any, ByRef pvarRight As Any, ByRef pvarResult As Any) As Long
    Private Declare Function VarAnd Lib "oleaut32" (ByRef pvarLeft As Any, ByRef pvarRight As Any, ByRef pvarResult As Any) As Long
    Private Declare Function VarDiv Lib "oleaut32" (ByRef pvarLeft As Any, ByRef pvarRight As Any, ByRef pvarResult As Any) As Long
    Private Declare Function VarEqv Lib "oleaut32" (ByRef pvarLeft As Any, ByRef pvarRight As Any, ByRef pvarResult As Any) As Long
    Private Declare Function VarFix Lib "oleaut32" (ByRef pvarIn As Any, ByRef pvarResult As Any) As Long
    Private Declare Function VarIdiv Lib "oleaut32" (ByRef pvarLeft As Any, ByRef pvarRight As Any, ByRef pvarResult As Any) As Long
    Private Declare Function VarImp Lib "oleaut32" (ByRef pvarLeft As Any, ByRef pvarRight As Any, ByRef pvarResult As Any) As Long
    Private Declare Function VarInt Lib "oleaut32" (ByRef pvarIn As Any, ByRef pvarResult As Any) As Long
    Private Declare Function VarMod Lib "oleaut32" (ByRef pvarLeft As Any, ByRef pvarRight As Any, ByRef pvarResult As Any) As Long
    Private Declare Function VarMul Lib "oleaut32" (ByRef pvarLeft As Any, ByRef pvarRight As Any, ByRef pvarResult As Any) As Long
    Private Declare Function VarNeg Lib "oleaut32" (ByRef pvarIn As Any, ByRef pvarResult As Any) As Long
    Private Declare Function VarNot Lib "oleaut32" (ByRef pvarIn As Any, ByRef pvarResult As Any) As Long
    Private Declare Function VarOr Lib "oleaut32" (ByRef pvarLeft As Any, ByRef pvarRight As Any, ByRef pvarResult As Any) As Long
    Private Declare Function VarPow Lib "oleaut32" (ByRef pvarLeft As Any, ByRef pvarRight As Any, ByRef pvarResult As Any) As Long
    Private Declare Function VarSub Lib "oleaut32" (ByRef pvarLeft As Any, ByRef pvarRight As Any, ByRef pvarResult As Any) As Long
    Private Declare Function VarXor Lib "oleaut32" (ByRef pvarLeft As Any, ByRef pvarRight As Any, ByRef pvarResult As Any) As Long
    Private Declare Sub RtlMoveMemory Lib "kernel32" (pDst As Any, pSrc As Any, ByVal BytLen As Long)
#End If

Private Type TVar
    Value As Variant
End Type

Private Type TVariantPtr
    VarVal As TVar
    Ptr()  As TVar
    udtptr As TUDTPtr
End Type

Private m_Var As TVariantPtr

Private Sub Class_Initialize()
    MPtr.New_UDTPtr m_Var.udtptr, FADF_RECORD Or FADF_FIXEDSIZE, Variant_Size
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

Public Property Get SizeInBytes() As Integer
    SizeInBytes = LenB(m_Var)
End Property
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
    RtlMoveMemory GetDecimal, ByVal Me.Ptr, Variant_Size
End Function

Public Function ToStr() As String
    Select Case Me.VarType
    'Case EVbVarType.vbByte, EVbVarType.vbInteger, EVbVarType.vbLong, EVbVarType.vbSingle, EVbVarType.vbDouble, EVbVarType.vbCurrency, EVbVarType.vbDate, EVbVarType.vbString
    '    ToStr = CStr(Me.Value)
    Case EVbVarType.vbObject:
        Dim obj As Object: Set obj = Me.GetObject
        ToStr = obj.ToStr 'every obj must have a Function ToStr As String
    Case Else
        ToStr = CStr(Me.Value)
    End Select
End Function

Public Property Get Value() As Variant
    Select Case Me.VarType
    Case EVbVarType.vbObject: Set Value = m_Var.Ptr(0).Value
    Case Else:                    Value = m_Var.Ptr(0).Value
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

' ############################## '    Arithmetic and Boolean Operations    ' ############################## '
Public Function VAbs() As VVariant
    'Returns the absolute value of a variant.
    Dim vRes: If VarAbs(m_Var.Ptr(0).Value, vRes) = S_OK Then m_Var.Ptr(0).Value = vRes
    Set VAbs = Me
End Function
Public Function VAdd(Value) As VVariant
    'Returns the sum of two variants.
    Dim vRes: If VarAdd(m_Var.Ptr(0).Value, Value, vRes) = S_OK Then m_Var.Ptr(0).Value = vRes
    Set VAdd = Me
End Function
Public Function VAnd(Value) As VVariant
    'Performs a bitwise And operation between two variants of any integral type.
    Dim vRes: If VarAnd(m_Var.Ptr(0).Value, Value, vRes) = S_OK Then m_Var.Ptr(0).Value = vRes
    Set VAnd = Me
End Function
Public Function VDiv(Value) As VVariant
    'Returns the result from dividing two variants.
    Dim vRes: If VarDiv(m_Var.Ptr(0).Value, Value, vRes) = S_OK Then m_Var.Ptr(0).Value = vRes
    Set VDiv = Me
End Function
Public Function VEqv(Value) As VVariant
    'Performs a bitwise equivalence on two variants.
    Dim vRes: If VarEqv(m_Var.Ptr(0).Value, Value, vRes) = S_OK Then m_Var.Ptr(0).Value = vRes
    Set VEqv = Me
End Function
Public Function VFix() As VVariant
     'Returns the integer portion of a variant.
    Dim vRes: If VarFix(m_Var.Ptr(0).Value, vRes) = S_OK Then m_Var.Ptr(0).Value = vRes
    Set VFix = Me
End Function
Public Function VIdiv(Value) As VVariant
    'Converts two variants of any type to integers then returns the result from dividing them.
    Dim vRes: If VarIdiv(m_Var.Ptr(0).Value, Value, vRes) = S_OK Then m_Var.Ptr(0).Value = vRes
    Set VIdiv = Me
End Function
Public Function VImp(Value) As VVariant
    'Performs a bitwise implication on two variants.
    Dim vRes: If VarImp(m_Var.Ptr(0).Value, Value, vRes) = S_OK Then m_Var.Ptr(0).Value = vRes
    Set VImp = Me
End Function
Public Function VInt() As VVariant
    'Returns the integer portion of a variant.
    'similar to VFix
    Dim vRes: If VarInt(m_Var.Ptr(0).Value, vRes) = S_OK Then m_Var.Ptr(0).Value = vRes
    Set VInt = Me
End Function
Public Function VMod(Value) As VVariant
    'Divides two variants and returns only the remainder.
    Dim vRes: If VarMod(m_Var.Ptr(0).Value, Value, vRes) = S_OK Then m_Var.Ptr(0).Value = vRes
    Set VMod = Me
End Function
Public Function VMul(Value) As VVariant
    'Returns the result from multiplying two variants.
    Dim vRes: If VarMul(m_Var.Ptr(0).Value, Value, vRes) = S_OK Then m_Var.Ptr(0).Value = vRes
    Set VMul = Me
End Function
Public Function VNeg() As VVariant
    'Performs logical negation on a variant.
    Dim vRes: If VarNeg(m_Var.Ptr(0).Value, vRes) = S_OK Then m_Var.Ptr(0).Value = vRes
    Set VNeg = Me
End Function
Public Function VNot() As VVariant
    'Performs the bitwise not negation operation on a variant.
    Dim vRes: If VarNot(m_Var.Ptr(0).Value, vRes) = S_OK Then m_Var.Ptr(0).Value = vRes
    Set VNot = Me
End Function
Public Function VOr(Value) As VVariant
    'Performs a logical disjunction on two variants.
    Dim vRes: If VarOr(m_Var.Ptr(0).Value, Value, vRes) = S_OK Then m_Var.Ptr(0).Value = vRes
    Set VOr = Me
End Function
Public Function VPow(Value) As VVariant
    'Returns the result of performing the power function with two variants.
    Dim vRes: If VarPow(m_Var.Ptr(0).Value, Value, vRes) = S_OK Then m_Var.Ptr(0).Value = vRes
    Set VPow = Me
End Function
Public Function VSub(Value) As VVariant
    'Subtracts two variants.
    Dim vRes: If VarSub(m_Var.Ptr(0).Value, Value, vRes) = S_OK Then m_Var.Ptr(0).Value = vRes
    Set VSub = Me
End Function
Public Function VXor(Value) As VVariant
    'Performs a logical exclusion on two variants.
    Dim vRes: If VarXor(m_Var.Ptr(0).Value, Value, vRes) = S_OK Then m_Var.Ptr(0).Value = vRes
    Set VXor = Me
End Function

' ############################## '    Implements VVariant    ' ############################## '
Private Property Get VVariant_Ptr() As LongPtr
    VVariant_Ptr = Me.Ptr
End Property

Private Property Get VVariant_SizeInBytes() As Integer
    VVariant_SizeInBytes = Me.SizeInBytes
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

' ############################## '           Implements VVariant           ' ############################## '
' ############################## '    Arithmetic and Boolean Operations    ' ############################## '

Public Function VVariant_VAbs() As VVariant
    Set VVariant_VAbs = Me.VAbs()
End Function
Public Function VVariant_VAdd(Value) As VVariant
    Set VVariant_VAdd = Me.VAdd(Value)
End Function
Public Function VVariant_VAnd(Value) As VVariant
    Set VVariant_VAnd = Me.VAnd(Value)
End Function
Public Function VVariant_VDiv(Value) As VVariant
    Set VVariant_VDiv = Me.VDiv(Value)
End Function
Public Function VVariant_VEqv(Value) As VVariant
    Set VVariant_VEqv = Me.VEqv(Value)
End Function
Public Function VVariant_VFix() As VVariant
    Set VVariant_VFix = Me.VFix()
End Function
Public Function VVariant_VIdiv(Value) As VVariant
    Set VVariant_VIdiv = Me.VIdiv(Value)
End Function
Public Function VVariant_VImp(Value) As VVariant
    Set VVariant_VImp = Me.VImp(Value)
End Function
Public Function VVariant_VInt() As VVariant
    Set VVariant_VInt = Me.VInt()
End Function
Public Function VVariant_VMod(Value) As VVariant
    Set VVariant_VMod = Me.VMod(Value)
End Function
Public Function VVariant_VMul(Value) As VVariant
    Set VVariant_VMul = Me.VMul(Value)
End Function
Public Function VVariant_VNeg() As VVariant
    Set VVariant_VNeg = Me.VNeg()
End Function
Public Function VVariant_VNot() As VVariant
    Set VVariant_VNot = Me.VNot()
End Function
Public Function VVariant_VOr(Value) As VVariant
    Set VVariant_VOr = Me.VOr(Value)
End Function
Public Function VVariant_VPow(Value) As VVariant
    Set VVariant_VPow = Me.VPow(Value)
End Function
Public Function VVariant_VSub(Value) As VVariant
    Set VVariant_VSub = Me.VSub(Value)
End Function
Public Function VVariant_VXor(Value) As VVariant
    Set VVariant_VXor = Me.VXor(Value)
End Function
