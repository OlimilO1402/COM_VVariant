VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "Form1"
   ClientHeight    =   12975
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   ScaleHeight     =   12975
   ScaleWidth      =   6255
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnTest 
      Caption         =   "Test VVariant && VVariantPtr"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   12375
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   0
      Top             =   600
      Width           =   6255
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BtnTest_Click()
    
    DebugPrint "Create a VVariant-object v1: "
    Dim v1 As VVariant: Set v1 = MNew.VVariant("Oliver")
    DebugPrint "v1 contains: "
    DebugPrint VVariant_ToDebugStr(v1)
    
    DebugPrint "Create a VVariantPtr-object v2:"
    Dim v2 As VVariantPtr: Set v2 = MNew.VVariantPtr("Frank")
    DebugPrint "v2 contains: "
    DebugPrint VVariant_ToDebugStr(v2)
    
    DebugPrint "Create a VVariant-object v3 by cloning v1: "
    Dim v3 As VVariant: Set v3 = v1.Clone
    DebugPrint "v3 contains: "
    DebugPrint VVariant_ToDebugStr(v3)
    
    DebugPrint "Assigning the v3.Ptr to v2.Ptr: "
    v2.Ptr = v3.Ptr
    DebugPrint "Now v2 contains: "
    DebugPrint VVariant_ToDebugStr(v2)
    
    DebugPrint "Assigning another value to v2. (""Chris"")"
    DebugPrint "The value of v3 will also be changed, because "
    DebugPrint "v2 points to the same memory-location as v3: "
    DebugPrint ""
    v2.Value = "Chris"
    
    DebugPrint "Now show me v2: "
    DebugPrint VVariant_ToDebugStr(v2)
    
    DebugPrint "Now show me v3: "
    DebugPrint VVariant_ToDebugStr(v3)
    
    DebugPrint "Resetting the pointer of v2: "
    v2.ResetPtr
    DebugPrint "Now v2 contains again: "
    DebugPrint VVariant_ToDebugStr(v2)
End Sub

Sub DebugPrint(Value As String)
    Text1.Text = Text1.Text & Value & vbCrLf
End Sub

Function VVariant_ToDebugStr(v As VVariant) As String
    Dim s As String
    s = "VarType:    " & v.VarType & "; Value: " & v.Value & vbCrLf & _
        "GetLongPtr: " & v.GetLongPtr & vbCrLf & _
        "Ptr:        " & v.Ptr & vbCrLf & _
        "ToStr:      " & v.ToStr & vbCrLf
    VVariant_ToDebugStr = s
End Function

