VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VVariant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'https://docs.microsoft.com/en-us/windows/win32/api/propidlbase/ns-propidlbase-propvariant
'https://devblogs.microsoft.com/oldnewthing/20171221-00/?p=97625

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

Private m_Var As TVar


Private Sub Class_Initialize()
    '
End Sub

Private Sub Class_Terminate()
    'delete the modified vt if the variant does not hold anything
    Me.VarType = vbEmpty
End Sub

Friend Sub New_(Value)
    If IsObject(Value) Then Set m_Var.Value = Value Else m_Var.Value = Value
End Sub

Friend Sub NewVt(Optional ByVal vt As EVbVarType = EVbVarType.vbIllegal, Optional Value)
    Me.New_ Value
    If vt <> EVbVarType.vbIllegal Then Me.VarType = vt
End Sub

'Public Function VVariant(aValue) As VVariant
'    'Create a VVariant-object with a Variant
'    Set VVariant = New VVariant: VVariant.New_ aValue
'End Function
'Public Function VVariantVt(vt As EVbVarType, aValue) As VVariant
'    'Create a VVariant-object with a Variant and set the vartype yourself,
'    'like e.g. give a signed Long and set vt to unsigned Long
'    Set VVariantVt = New VVariant: VVariantVt.NewVt vt, aValue
'End Function

Friend Sub NewC(other As VVariant)
    New_ other.Value
End Sub

Friend Function Clone() As VVariant
    Set Clone = New VVariant: Clone.NewC Me
End Function

Public Property Get Ptr() As LongPtr
    Ptr = VarPtr(m_Var)
End Property

Public Property Let VarType(ByVal vt As EVbVarType)
    GetMem2 vt, ByVal Me.Ptr
End Property
Public Property Get VarType() As EVbVarType
    GetMem2 ByVal Me.Ptr, VarType
End Property

Public Function VarTypeToStr() As String
    VarTypeToStr = MEVbVarType.VarType2_ToStr(m_Var.Value)
End Function

Public Property Let Slot(ByVal Index As Byte, ByVal Value As Long)
    Dim maxIndex As Byte: maxIndex = MEVbVarType.SizeOf_Variant / 4 - 1
    If Index > maxIndex Then Index = maxIndex
    Dim pDst As LongPtr: pDst = Me.Ptr + Index * 4
    GetMem4 Value, ByVal pDst
End Property
Public Property Get Slot(ByVal Index As Byte) As Long
    Dim maxIndex As Byte: maxIndex = MEVbVarType.SizeOf_Variant / 4 - 1
    If Index > maxIndex Then Index = maxIndex
    Dim pSrc As LongPtr: pSrc = Me.Ptr + Index * 4
    GetMem4 ByVal pSrc, Slot
End Property

Public Function HexDump() As String
    Dim i As Long, s As String ': s = ""
    For i = 0 To MEVbVarType.SizeOf_Variant / 4 - 1
        s = s & Hex$(Me.Slot(i)) & vbCrLf
    Next
    HexDump = s
End Function

Public Property Get SizeInBytes() As Integer
    SizeInBytes = MEVbVarType.SizeOf_Variant
End Property

'all GetData-functions:
Public Function GetByte() As Byte
    GetMem1 ByVal Me.Ptr + 8, GetByte
End Function

Public Function GetInteger() As Integer
    GetMem2 ByVal Me.Ptr + 8, GetInteger
End Function

Public Function GetLong() As Long
    GetMem4 ByVal Me.Ptr + 8, GetLong
End Function

#If VBA7 Then
Public Function GetLongLong() As LongLong
    GetMem8 ByVal Me.Ptr + 8, GetLongLong
End Function
#End If

Public Function GetLongPtr() As LongPtr
    RtlMoveMemory GetLongPtr, ByVal Me.Ptr + 8, MPtr.SizeOf_LongPtr
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

Public Function GetString() As String
    GetString = m_Var.Value
End Function

Public Function GetObject() As Object
    Set GetObject = m_Var.Value
End Function

Public Function GetDecimal() As Variant
    RtlMoveMemory GetDecimal, ByVal Me.Ptr, MEVbVarType.SizeOf_Variant
End Function

Public Function ToStr() As String
    'What shall we present here?
    'the value, the VarType, the Ptr, what else
    Select Case Me.VarType
    Case EVbVarType.vbObject
        Dim obj As Object: Set obj = m_Var.Value
        ToStr = obj.ToStr
    Case Else
        ToStr = CStr(Me.Value)
    End Select
End Function

Public Property Get Value() As Variant
    If Me.VarType = EVbVarType.vbObject Then
        Set Value = m_Var.Value
    Else
        Value = m_Var.Value
    End If
End Property

Public Property Let Value(aValue)
    m_Var.Value = aValue
End Property

Public Property Set Value(aValue)
    Set m_Var.Value = aValue
End Property

' ############################## '    Arithmetic and Boolean Operations    ' ############################## '
Public Function VAbs() As VVariant
    'Returns the absolute value of a variant.
    Dim vRes: If VarAbs(m_Var.Value, vRes) = S_OK Then m_Var.Value = vRes
    Set VAbs = Me
End Function
Public Function VAdd(Value) As VVariant
    'Returns the sum of two variants.
    Dim vRes, hr As Long
    
    hr = VarAdd(m_Var.Value, Value, ByVal VarPtr(vRes))
    If hr = S_OK Then
        m_Var.Value = vRes
    Else
        Debug.Print Hex(hr)
        Debug.Print WinApiError_ToStr(hr)
    End If
    Set VAdd = Me
End Function
Public Function VAnd(Value) As VVariant
    'Performs a bitwise And operation between two variants of any integral type.
    Dim vRes: If VarAnd(m_Var.Value, Value, vRes) = S_OK Then m_Var.Value = vRes
    Set VAnd = Me
End Function
Public Function VDiv(Value) As VVariant
    'Returns the result from dividing two variants.
    Dim vRes: If VarDiv(m_Var.Value, Value, vRes) = S_OK Then m_Var.Value = vRes
    Set VDiv = Me
End Function
Public Function VEqv(Value) As VVariant
    'Performs a bitwise equivalence on two variants.
    Dim vRes, hr As Long
    hr = VarEqv(m_Var.Value, Value, vRes)
    If hr = S_OK Then
        m_Var.Value = vRes
    Else
        Debug.Print Hex(hr) & " " & MErr.WinApiError_ToStr(hr)
    End If
    Set VEqv = Me
End Function
Public Function VFix() As VVariant
     'Returns the integer portion of a variant.
    Dim vRes: If VarFix(m_Var.Value, vRes) = S_OK Then m_Var.Value = vRes
    Set VFix = Me
End Function
Public Function VIdiv(Value) As VVariant
    'Converts two variants of any type to integers then returns the result from dividing them.
    Dim vRes: If VarIdiv(m_Var.Value, Value, vRes) = S_OK Then m_Var.Value = vRes
    Set VIdiv = Me
End Function
Public Function VImp(Value) As VVariant
    'Performs a bitwise implication on two variants.
    Dim vRes: If VarImp(m_Var.Value, Value, vRes) = S_OK Then m_Var.Value = vRes
    Set VImp = Me
End Function
Public Function VInt() As VVariant
    'Returns the integer portion of a variant.
    'similar to VFix
    Dim vRes: If VarInt(m_Var.Value, vRes) = S_OK Then m_Var.Value = vRes
    Set VInt = Me
End Function
Public Function VMod(Value) As VVariant
    'Divides two variants and returns only the remainder.
    Dim vRes: If VarMod(m_Var.Value, Value, vRes) = S_OK Then m_Var.Value = vRes
    Set VMod = Me
End Function
Public Function VMul(Value) As VVariant
    'Returns the result from multiplying two variants.
    Dim vRes: If VarMul(m_Var.Value, Value, vRes) = S_OK Then m_Var.Value = vRes
    Set VMul = Me
End Function
Public Function VNeg() As VVariant
    'Performs logical negation on a variant.
    Dim vRes: If VarNeg(m_Var.Value, vRes) = S_OK Then m_Var.Value = vRes
    Set VNeg = Me
End Function
Public Function VNot() As VVariant
    'Performs the bitwise not negation operation on a variant.
    Dim vRes: If VarNot(m_Var.Value, vRes) = S_OK Then m_Var.Value = vRes
    Set VNot = Me
End Function
Public Function VOr(Value) As VVariant
    'Performs a logical disjunction on two variants.
    Dim vRes: If VarOr(m_Var.Value, Value, vRes) = S_OK Then m_Var.Value = vRes
    Set VOr = Me
End Function
Public Function VPow(Value) As VVariant
    'Returns the result of performing the power function with two variants.
    Dim vRes: If VarPow(m_Var.Value, Value, vRes) = S_OK Then m_Var.Value = vRes
    Set VPow = Me
End Function
Public Function VSub(Value) As VVariant
    'Subtracts two variants.
    Dim vRes: If VarSub(m_Var.Value, Value, vRes) = S_OK Then m_Var.Value = vRes
    Set VSub = Me
End Function
Public Function VXor(Value) As VVariant
    'Performs a logical exclusion on two variants.
    Dim vRes: If VarXor(m_Var.Value, Value, vRes) = S_OK Then m_Var.Value = vRes
    Set VXor = Me
End Function
