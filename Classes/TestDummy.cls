VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestDummy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name As String

Friend Sub New_(aName As String)
    Name = aName
End Sub
'Public Function TestDummy(aName As String) As TestDummy
'    Set TestDummy = New TestDummy: TestDummy.New_ aName
'End Function

Public Function ToStr() As String
    ToStr = Name
End Function
