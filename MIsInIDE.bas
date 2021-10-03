Attribute VB_Name = "MIsInIDE"
Option Explicit
Private Const GW_OWNER As Long = &H4&

Public Enum CheckIDEEnum
    ciIDEDesign = -2 ' designing a user control in the IDE
    ciIDERun = -1    ' using the user control designing a form
    ciCompiled = 0   ' the control used in compiled exe program
End Enum
#If VBA7 Then
    Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Private Declare PtrSafe Function GetParent Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal wCmd As Long) As Long
#Else
    Public Enum LongPtr
        [None]
    End Enum
    Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Private Declare Function GetParent Lib "user32" (ByVal hWnd As LongPtr) As Long
    Private Declare Function GetWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal wCmd As Long) As Long
#End If

'works in every situation
Public Function IsInIDE(Optional ByVal hWnd As LongPtr = 0) As CheckIDEEnum
Try: On Error GoTo Catch
    If hWnd = 0 Then
        'Runtimer-Error 76? do the following:
        '"Extras"->"Options"->"Allgemein"->"Unterbrechen bei Fehlern"->"Bei nicht verarbeiteten Fehlern"
        Debug.Print 1 / 0
    Else
        Dim PhWnd As LongPtr
        Do While Not hWnd = 0
            PhWnd = hWnd
            hWnd = GetParent(PhWnd)
        Loop
        Dim Buffer As String: Buffer = Space$(128)
        hWnd = GetWindow(PhWnd, GW_OWNER)
        GetClassName PhWnd, Buffer, Len(Buffer)
        Buffer = UCase(Left(Buffer, 8))
        Select Case Buffer
        Case "IDEOWNER": IsInIDE = ciIDEDesign
        Case "THUNDERM": IsInIDE = ciIDERun
        Case Else:       IsInIDE = ciCompiled
        End Select
    End If
    Exit Function
Catch: IsInIDE = ciIDERun
End Function

Public Function IsInIDE1(ByVal hWnd As LongPtr) As CheckIDEEnum
    'aka the "idiv"-trick
    'works for UserControls
    'could also work for forms
    'Dim Hwnd As Long
    Dim PhWnd As Long
    Dim Buffer As String
    Dim Result As Long

    'Hwnd = Form1.UserControl11.Hwnd
    Do While Not hWnd = 0
        PhWnd = hWnd
        hWnd = GetParent(PhWnd)
    Loop

    Buffer = Space$(128)
    'PHwnd = GetWindow(PHwnd, GW_OWNER)
    Result = GetClassName(PhWnd, Buffer, Len(Buffer))
    If Left$(Buffer, 8) = "IDEOwner" Then
        'MsgBox "IDE"
        IsInIDE1 = CheckIDEEnum.ciIDEDesign
    ElseIf Left$(Buffer, 11) = "ThunderMain" Then
        'MsgBox "IDE running"
        IsInIDE1 = CheckIDEEnum.ciIDERun
    Else
        'MsgBox "compiled"
        IsInIDE1 = CheckIDEEnum.ciCompiled
    End If
End Function

Public Function CheckIDEEnum_ToStr(e As CheckIDEEnum) As String
    Dim s As String
    Select Case e
    Case CheckIDEEnum.ciIDEDesign: s = "ciIDEDesign" '"IDEOwner"
    Case CheckIDEEnum.ciIDERun:    s = "ciIDERun"    '"ThunderMain"
    Case CheckIDEEnum.ciCompiled:  s = "ciCompiled"  'compiled exe
    End Select
    CheckIDEEnum_ToStr = s
End Function

Public Function IsInIDE2() As Boolean
    'similar to idiv-trick but for windows/Forms only
    Dim hWndParent As Long:  hWndParent = GetWindow(Form1.hWnd, GW_OWNER)
    Dim Buffer     As String: Buffer = Space$(128)
    GetClassName hWndParent, Buffer, Len(Buffer)
    
    If Left(Buffer, 11) = "ThunderMain" Then
        IsInIDE2 = True
    Else
        IsInIDE2 = False
    End If
End Function

Public Function IsInIDE3() As Boolean
    'aka the "Klaus-Langbein"-trick,
    'aka the "C.-Schlegel"-trick,
    'simplest way, works in VB-Classic only
    'for standard-exe only
    IsInIDE3 = App.LogMode = vbLogAuto '0
End Function

'Public Function LogModeConstants_ToStr(e As LogModeConstants) As String
'    Dim s As String
'    Select Case e
'    Case LogModeConstants.vbLogAuto:      s = "vbLogAuto"      ' 0
'    Case LogModeConstants.vbLogOff:       s = "vbLogOff"       ' 1
'    Case LogModeConstants.vbLogToFile:    s = "vbLogToFile"    ' 2
'    Case LogModeConstants.vbLogToNT:      s = "vbLogToNT"      ' 3
'    Case LogModeConstants.vbLogOverwrite: s = "vbLogOverwrite" '16 &H10
'    Case LogModeConstants.vbLogThreadID:  s = "vbLogThreadID"  '32 &H20
'    Case Else: s = CStr(e)
'    End Select
'    LogModeConstants_ToStr = s
'End Function

'Tricks with the debug-object are working in VBC awa in VBA7
'
Public Function IsInIDE4() As Boolean
    'aka the dirty-trick
    'aka the "Dominic-Hoffmann"-trick
    On Error GoTo Fehler

    '   Fehler produzieren
    Debug.Print 1 / 0

    '   Hier angekommen wurde kein Fehler gemeldet
    IsInIDE4 = False

Ende:
    Exit Function

Fehler:
    '   Hier angekommen wurde ein Fehler gemeldet
    IsInIDE4 = True

    Resume Ende
End Function

Public Function IsInIDE5( _
                Optional ByVal bVal As Boolean = False _
                ) As Boolean
    'aka the "Konrad-LM-Rudolph"-trick
    Static bComp__ As Boolean

    If bVal Then
        bComp__ = False
    Else
        bComp__ = True
        Debug.Assert IsInIDE5(True)
        IsInIDE5 = Not bComp__
    End If
End Function

'Kommentar von JTK-One am 02.07.2004 um 16:25
'
'Es gibt auch die Eigenschaft:
'Ambient.UserMode
'Damit kann man bei OCXen prima zwischen Design- und Runtime unterscheide.

'Public Function IsInIDE6() As Boolean
'    'if you see semoetihing like this, don't trust, it does not work
'    Debug.Assert (IsInIDE6 = True)
'End Function
'

Public Sub CheckIsInIDE6()
    Dim bInIde          As Boolean
    
    Debug.Assert SetTrue(bInIde)
    If bInIde Then
        ' do stuff
        MsgBox "Yeah in IDE"
    End If
    ' do other stuff
    MsgBox "Nope not in IDE"
End Sub

Public Function SetTrue(bValue As Boolean) As Boolean
    bValue = True
    SetTrue = True
End Function

