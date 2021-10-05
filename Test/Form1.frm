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
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim s As String

Private Sub Command1_Click()
    Dim bInIde          As Boolean
    
    s = ""
    Debug.Assert SetTrue(bInIde)
    If bInIde Then
        ' do stuff
        s = s & "Yeah-1;  "
    End If
    ' do other stuff
    s = s & "Nope-1;"
End Sub

Public Function SetTrue(bValue As Boolean) As Boolean
    bValue = True
    SetTrue = True
End Function


Private Sub Command2_Click()
    Dim bInIde          As Boolean
    
    s = ""
    Debug.Assert SetTrue(bInIde)
    If bInIde Then
        ' do stuff
        s = s & "Yeah-2;  "
    End If
    ' do other stuff
    s = s & "Nope-2;"
End Sub

Private Sub Command3_Click()
    MsgBox s
End Sub
