VERSION 5.00
Begin VB.UserControl UserControl1 
   ClientHeight    =   1665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4005
   ScaleHeight     =   1665
   ScaleWidth      =   4005
   Begin VB.Frame Frame1 
      Caption         =   "UserControl for testing IsInIDE"
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      Begin VB.CommandButton Command2 
         Caption         =   "IsInIDE"
         Height          =   375
         Left            =   2640
         TabIndex        =   6
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "IsInIDE"
         Height          =   375
         Left            =   2640
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label LblIsInIDEQ1 
         AutoSize        =   -1  'True
         Caption         =   "Is In IDE?"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   705
      End
      Begin VB.Label LblIsInIDEA1 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         Height          =   195
         Left            =   1560
         TabIndex        =   4
         Top             =   720
         Width           =   480
      End
      Begin VB.Label LblIsInIDEA 
         AutoSize        =   -1  'True
         Caption         =   "Label"
         Height          =   195
         Left            =   1560
         TabIndex        =   2
         Top             =   360
         Width           =   390
      End
      Begin VB.Label LblIsInIDEQ 
         AutoSize        =   -1  'True
         Caption         =   "Is In IDE?"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   705
      End
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    LblIsInIDEA.Caption = MIsInIDE.CheckIDEEnum_ToStr(MIsInIDE.IsInIDE(UserControl.hWnd))
End Sub

Private Sub Command2_Click()
    LblIsInIDEA1.Caption = MIsInIDE.CheckIDEEnum_ToStr(MIsInIDE.IsInIDE1(UserControl.hWnd))
End Sub

Public Property Get hWnd() As LongPtr
    hWnd = UserControl.hWnd
End Property

Private Sub UserControl_Initialize()
    LblIsInIDEA.Caption = MIsInIDE.CheckIDEEnum_ToStr(MIsInIDE.IsInIDE(UserControl.hWnd))
    LblIsInIDEA1.Caption = MIsInIDE.CheckIDEEnum_ToStr(MIsInIDE.IsInIDE1(UserControl.hWnd))
End Sub
