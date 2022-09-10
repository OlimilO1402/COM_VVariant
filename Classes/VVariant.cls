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

Private Type TVar
    Value As Variant
End Type

Private m_Var As TVar

Private Declare Sub RtlMoveMemory Lib "kernel32" (pDst As Any, pSrc As Any, ByVal BytLen As Long)

Friend Sub New_(Value)
    If IsObject(Value) Then Set m_Var.Value = Value Else m_Var.Value = Value
End Sub

Friend Sub NewVt(Optional ByVal vt As EVbVarType = EVbVarType.vbIllegal, Optional Value)
    If vt = EVbVarType.vbIllegal Then Me.New_ Value: Exit Sub
    Me.VarType = vt
End Sub

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
    RtlMoveMemory ByVal Me.Ptr, vt, 2
End Property
Public Property Get VarType() As EVbVarType
    RtlMoveMemory VarType, ByVal Me.Ptr, 2
End Property

Public Function VarTypeToStr() As String
    VarTypeToStr = MEVbVarType.VarType2_ToStr(m_Var.Value)
End Function

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
    RtlMoveMemory GetLongPtr, ByVal Me.Ptr + 8, MPtr.LongPtr_Size
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
    RtlMoveMemory GetDecimal, ByVal Me.Ptr, MVariant.Variant_Size
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

