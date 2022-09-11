Attribute VB_Name = "MNew"
Option Explicit

Public Function VVariant(aValue) As VVariant
    'Create a VVariant-object with a Variant
    Set VVariant = New VVariant: VVariant.New_ aValue
End Function
Public Function VVariantVt(vt As EVbVarType, aValue) As VVariant
    'Create a VVariant-object with a Variant and set the vartype yourself,
    'like e.g. give a signed Long and set vt to unsigned Long
    Set VVariantVt = New VVariant: VVariantVt.NewVt vt, aValue
End Function

Public Function VVariantPtr(aValue) As VVariantPtr
    'Create a VVariantPtr-object with a Variant
    Set VVariantPtr = New VVariantPtr: VVariantPtr.New_ aValue
End Function
Public Function VVariantPtrVt(vt As EVbVarType, aValue) As VVariantPtr
    'Create a VVariant-object with a Variant and set the vartype yourself,
    'like e.g. give a signed Long and set vt to unsigned Long
    Set VVariantPtrVt = New VVariantPtr: VVariantPtrVt.NewVt vt, aValue
End Function

Public Function TestDummy(aName As String) As TestDummy
    Set TestDummy = New TestDummy: TestDummy.New_ aName
End Function
