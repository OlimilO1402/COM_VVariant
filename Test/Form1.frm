VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type TVariant
    Value As Variant
End Type

'Private m_Var As TVariant

Private Declare Sub GetMem2 Lib "msvbvm60" (ByRef Src As Any, ByRef Dst As Any)


Private Sub Command1_Click()
    
    Dim aVar As TVariant
    
    VVariant1 aVar, vbArray Or vbString, Empty
    
End Sub

Private Sub Command2_Click()
    
    Dim aVar As TVariant
    
    VVariant2 aVar, vbArray Or vbString

End Sub

Private Sub Command3_Click()
    
    Dim aVar As TVariant
    
    VVariant3 aVar, vbArray Or vbString

End Sub

Private Sub VVariant1(this As TVariant, vt As VbVarType, ByVal aVal As Variant)
    
    this.Value = aVal
    GetMem2 ByVal vt, ByVal VarPtr(this)
    
End Sub

Private Sub VVariant2(this As TVariant, vt As VbVarType, Optional ByVal aVal As Variant = Empty)
    
    this.Value = aVal
    GetMem2 ByVal vt, ByVal VarPtr(this)
    
End Sub
Private Sub VVariant3(this As TVariant, vt As VbVarType, Optional aVal As Variant)
        
    this.Value = aVal
    GetMem2 ByVal vt, ByVal VarPtr(this)
    
End Sub
