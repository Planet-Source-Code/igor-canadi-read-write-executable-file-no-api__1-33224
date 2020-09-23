VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Read-write"
   ClientHeight    =   1905
   ClientLeft      =   7665
   ClientTop       =   3450
   ClientWidth     =   3705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   3705
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.TextBox txtValue 
      Height          =   285
      Index           =   1
      Left            =   1080
      TabIndex        =   2
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox txtValue 
      Height          =   285
      Index           =   0
      Left            =   1080
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton cmdWrite 
      Caption         =   "Write"
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "Read"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "File:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "2nd Value:"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "1st Value:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'I didn't used a module because it makes the project a little bit slower. It's only 3 functions that could be in a module.
Option Explicit
Private Const BUFF_LEN = 30 'maximum lenght of values

Function FindOffset(ByVal wFile As String, ByVal wMarker As String) As Long 'finds position of a marker
    On Error GoTo er
    Dim FF As Long 'free file
    Dim rByte As Byte
    Dim Junk As Byte 'not used bytes
    Dim CmpStr As String
    FF = FreeFile
    Open wFile For Binary As FF 'let's open the file
        Do Until EOF(FF)
            If Not EOF(FF) Then Get FF, , rByte 'reading
            If Not EOF(FF) Then Get FF, , Junk 'not used bytes
            If Len(CmpStr) < Len(wMarker) Then
                CmpStr = CmpStr + Chr(rByte) ' make the string longer
            Else
                CmpStr = Mid$(CmpStr, 2) + Chr(rByte) 'move the string, so it could recognize the marker
            End If
                        
            If CmpStr = wMarker Then 'I've got the marker!!!
                FindOffset = Loc(FF) + 1 'location of the marker
                Exit Do
            End If
        Loop
    Close #1
    Exit Function
er:
    FindOffset = 0 'error
End Function

Function WriteToFile(ByVal wFile As String, ByVal wValue As String, ByVal wMarker As String) As Boolean
    On Error GoTo er
    Dim Base As Long 'position of the marker
    Dim i As Byte
    Dim bVal As Byte
    Base = FindOffset(wFile, wMarker)
    wValue = wValue + "<END>"
    If Base = 0 Then GoTo er
    Open wFile For Binary As #1 'opens the file
        For i = 1 To Len(wValue)
            bVal = Asc(Mid(wValue, i, 1)) 'char witch will go to the exe
            Put #1, Base + (i - 1) * 2, bVal 'write to file (every second byte)
        Next i
    Close #1
    
    WriteToFile = True 'everything went fine
    Exit Function
er:
    WriteToFile = False 'something didn't work correctly
End Function

Function ReadFromFile(ByVal wFile As String, ByVal wMarker As String) As String 'to read from file
    On Error Resume Next
    Dim Base As Long 'position of the marker
    Dim tStr As String 'temporary string
    Dim i As Byte
    Dim bVal As Byte
    Base = FindOffset(wFile, wMarker) 'finds position of the marker
    If Base <> 0 Then 'if there is no error then...
        Open wFile For Binary As #1 'opens the file
            Do Until (InStr(1, tStr, "<END>") <> 0) Or (i = BUFF_LEN * 2) 'stop if there is the end of the string, or the string is on the max length
                Get #1, Base + i * 2, bVal 'get byte (every second byte)
                i = i + 1 'move the position
                tStr = tStr + Chr(bVal)
            Loop
        Close #1
    
        ReadFromFile = Left$(tStr, Len(tStr) - 5) 'returns string witch we read
    Else
        MsgBox "Something's wrong. Check if you spelled file name correctly", vbCritical 'error
    End If
End Function

Private Sub cmdRead_Click()
    Dim file As String 'file path
    file = txtFile.Text
    txtValue(0).Text = ReadFromFile(file, "VALUE1$") 'read 1st value
    txtValue(1).Text = ReadFromFile(file, "VALUE2$") 'read 2nd value
End Sub

Private Sub cmdWrite_Click()
    Dim file As String 'file path
    file = txtFile.Text
    If WriteToFile(file, txtValue(0).Text, "VALUE1$") = False Then MsgBox "Something's wrong. Check if you spelled file name correctly", vbCritical: Exit Sub 'write 1st value and trap the error
    If WriteToFile(file, txtValue(1).Text, "VALUE2$") = False Then MsgBox "Something's wrong. Check if you spelled file name correctly", vbCritical: Exit Sub 'write 1st value and trap the error
    txtValue(0).Text = "" 'reset value
    txtValue(1).Text = "" 'reset value
    Shell file 'let user see what has happend to the exe file
End Sub

Private Sub Form_Load()
    txtFile.Text = App.Path & "\test_exe.exe" 'this makes writing path easier
End Sub
