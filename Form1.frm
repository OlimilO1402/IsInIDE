VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "The Question ""Is in IDE"" Answered in 6 different  ways"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows-Standard
   Begin Projekt1.UserControl1 UserControl11 
      Height          =   1575
      Left            =   2880
      TabIndex        =   12
      Top             =   240
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   2778
   End
   Begin VB.Label LblIsInIDEQ1 
      AutoSize        =   -1  'True
      Caption         =   "IsInIDE1?"
      Height          =   195
      Left            =   240
      TabIndex        =   14
      Top             =   600
      Width           =   705
   End
   Begin VB.Label LblIsInIDEA1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   1560
      TabIndex        =   13
      Top             =   600
      Width           =   480
   End
   Begin VB.Label LblIsInIDEA 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   1560
      TabIndex        =   11
      Top             =   240
      Width           =   480
   End
   Begin VB.Label LblIsInIDEQ 
      AutoSize        =   -1  'True
      Caption         =   "IsInIDE?"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   240
      Width           =   615
   End
   Begin VB.Label LblIsInIDEA6 
      Height          =   195
      Left            =   1560
      TabIndex        =   9
      Top             =   2400
      Width           =   480
   End
   Begin VB.Label LblIsInIDEQ6 
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   2400
      Width           =   480
   End
   Begin VB.Label LblIsInIDEA5 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   1560
      TabIndex        =   7
      Top             =   2040
      Width           =   480
   End
   Begin VB.Label LblIsInIDEQ5 
      AutoSize        =   -1  'True
      Caption         =   "IsInIDE5?"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   705
   End
   Begin VB.Label LblIsInIDEA4 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   1560
      TabIndex        =   5
      Top             =   1680
      Width           =   480
   End
   Begin VB.Label LblIsInIDEQ4 
      AutoSize        =   -1  'True
      Caption         =   "IsInIDE4?"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   705
   End
   Begin VB.Label LblIsInIDEQ2 
      AutoSize        =   -1  'True
      Caption         =   "IsInIDE2?"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   705
   End
   Begin VB.Label LblIsInIDEQ3 
      AutoSize        =   -1  'True
      Caption         =   "IsInIDE3?"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   705
   End
   Begin VB.Label LblIsInIDEA3 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   1560
      TabIndex        =   1
      Top             =   1320
      Width           =   480
   End
   Begin VB.Label LblIsInIDEA2 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   1560
      TabIndex        =   0
      Top             =   960
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'http://www.activevb.de/tipps/vb6tipps/tipp0347.html  ' 15.06.2003
'http://www.activevb.de/rubriken/faq/faq0151.html     ' 16.12.2006

'4 different answers to the question:
'Is the program running in the IDE during debugging, or does it run as compiled exe on the user-computer?
'when to use which trick if you have
'standard.exe, in VB-IDE during debugging, in User-Control

'for all tricks that work with the debug-object, you must activate:
'"Extras"->"Optionen"->"Allgemein"->"Unterbrechen bei Fehlern"->"Bei nicht verarbeiteten Fehlern"
Private Sub Form_Load()
    LblIsInIDEA.Caption = MIsInIDE.CheckIDEEnum_ToStr(MIsInIDE.IsInIDE)
    LblIsInIDEA1.Caption = MIsInIDE.CheckIDEEnum_ToStr(MIsInIDE.IsInIDE1(UserControl11.hWnd))
    LblIsInIDEA2.Caption = MIsInIDE.IsInIDE2
    LblIsInIDEA3.Caption = MIsInIDE.IsInIDE3
    LblIsInIDEA4.Caption = MIsInIDE.IsInIDE4
    LblIsInIDEA5.Caption = MIsInIDE.IsInIDE5
    'LblIsInIDEA6.Caption = MIsInIDE.IsInIDE6
End Sub

