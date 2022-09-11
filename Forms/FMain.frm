VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "Form1"
   ClientHeight    =   12975
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   ScaleHeight     =   12975
   ScaleWidth      =   9330
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnTestArithmetic 
      Caption         =   "Test Arithmetic Functions"
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
      Left            =   2880
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
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
      Width           =   2655
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

Private Sub Form_Resize()
    Dim L As Single
    Dim T As Single: T = Text1.Top
    Dim W As Single: W = Me.ScaleWidth
    Dim H As Single: H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then Text1.Move L, T, W, H
End Sub

Private Sub BtnTest_Click()
    
    DebugClear
    DebugPrint "Testing basic functions of VVariant and VVariantPtr"
    DebugPrint "==================================================="
    
    DebugPrint "Create a VVariant-object v1: "
    Dim v1 As VVariant: Set v1 = MNew.VVariant("Oliver")
    DebugPrint "v1 contains: "
    DebugPrint VVariant_ToDebugStr1(v1)
    
    DebugPrint "Create a VVariantPtr-object v2:"
    Dim v2 As VVariantPtr: Set v2 = MNew.VVariantPtr("Frank")
    DebugPrint "v2 contains: "
    DebugPrint VVariant_ToDebugStr1(v2)
    
    DebugPrint "Create a VVariant-object v3 by cloning v1: "
    Dim v3 As VVariant: Set v3 = v1.Clone
    DebugPrint "v3 contains: "
    DebugPrint VVariant_ToDebugStr1(v3)
    
    DebugPrint "Assigning the v3.Ptr to v2.Ptr: "
    v2.Ptr = v3.Ptr
    DebugPrint "Now v2 contains: "
    DebugPrint VVariant_ToDebugStr1(v2)
    
    DebugPrint "Assigning another value to v2. (""Chris"")"
    DebugPrint "The value of v3 will also be changed, because "
    DebugPrint "v2 points to the same memory-location as v3: "
    DebugPrint ""
    v2.Value = "Chris"
    
    DebugPrint "Now show me v2: "
    DebugPrint VVariant_ToDebugStr1(v2)
    
    DebugPrint "Now show me v3: "
    DebugPrint VVariant_ToDebugStr1(v3)
    
    DebugPrint "Resetting the pointer of v2: "
    v2.ResetPtr
    DebugPrint "Now v2 contains again: "
    DebugPrint VVariant_ToDebugStr1(v2)
    
    DebugPrint "Create a VVariant-object v4: "
    Dim v4 As VVariant: Set v4 = MNew.VVariant(123)
    DebugPrint "v4 contains: "
    DebugPrint VVariant_ToDebugStr1(v4)
    
End Sub

Sub DebugClear()
    Text1.Text = ""
End Sub

Sub DebugPrint(Value As String)
    Text1.Text = Text1.Text & Value & vbCrLf
End Sub

Function VVariant_ToDebugStr1(v As VVariant) As String
    Dim s As String
    s = "VarType:    " & v.VarTypeToStr & vbCrLf & _
        "Value:      " & v.Value & vbCrLf & _
        "GetLongPtr: " & v.GetLongPtr & vbCrLf & _
        "Ptr:        " & v.Ptr & vbCrLf & _
        "ToStr:      " & v.ToStr & vbCrLf
    VVariant_ToDebugStr1 = s
End Function

Function VVariant_ToDebugStr2(v As VVariant) As String
    Dim s As String
    s = "VarType:    " & v.VarTypeToStr & vbCrLf & _
        "Value:      " & v.Value & vbCrLf
    VVariant_ToDebugStr2 = s
End Function

Private Sub BtnTestArithmetic_Click()
    
    DebugClear
    DebugPrint "Testing arithmetic operations"
    DebugPrint "============================="
    
    Dim v: v = -123456
    DebugPrint "Create a VVariant-object v1: "
    Dim v1 As VVariant: Set v1 = MNew.VVariant(v)
    DebugPrint VVariant_ToDebugStr2(v1)
    
    DebugPrint "Clone v1 to v2: "
    Dim v2 As VVariant: Set v2 = v1.Clone
    DebugPrint VVariant_ToDebugStr2(v2)
    
    DebugPrint "Absolute value of v2: "
    v2.VAbs
    DebugPrint VVariant_ToDebugStr2(v2)
    
    v = 456789
    DebugPrint "Adding " & v & " to v2: "
    v2.VAdd v
    DebugPrint VVariant_ToDebugStr2(v2)
    
    v = 12345
    DebugPrint "And-operation with " & v & " on v2: "
    v2.VAnd v
    DebugPrint VVariant_ToDebugStr2(v2)
    
    v = 213
    DebugPrint "Dividing v2 by " & v & ": "
    v2.VDiv 213
    DebugPrint VVariant_ToDebugStr2(v2)
    
    DebugPrint "Only whole part of v2:"
    v2.VFix
    DebugPrint VVariant_ToDebugStr2(v2)
    
    v = 2
    DebugPrint "Equivalenting v2 by " & v & ": "
    v2.VEqv v
    DebugPrint VVariant_ToDebugStr2(v2)
    
End Sub
