VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Values:"
   ClientHeight    =   1020
   ClientLeft      =   3270
   ClientTop       =   2805
   ClientWidth     =   3765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1020
   ScaleWidth      =   3765
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lbl2 
      BackStyle       =   0  'Transparent
      Caption         =   "2nd Value: "
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "1st Value: "
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function VALUE_1() As String 'this function returns the first value
    Dim tmpValue1 As String
    tmpValue1 = "VALUE1$first value<END>                       "  'this is what the other exe is changing (first value)
    VALUE_1 = Mid(tmpValue1, 8, InStr(tmpValue1, "<END>") - 8) 'returns the string without VALUE1$ and <END>
End Function

Function VALUE_2() As String 'this function returns the second value
    Dim tmpValue2 As String
    tmpValue2 = "VALUE2$second value<END>                       " 'this is what the other exe is changing (first value)
    VALUE_2 = Mid(tmpValue2, 8, InStr(tmpValue2, "<END>") - 8) 'returns the string without VALUE2$ and <END>
End Function

Private Sub Form_Load()
    lbl1.Caption = lbl1.Caption & VALUE_1 'lbl1 = first value
    lbl2.Caption = lbl2.Caption & VALUE_2 'lbl2 = second value
End Sub
