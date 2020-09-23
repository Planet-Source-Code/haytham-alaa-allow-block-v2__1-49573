Attribute VB_Name = "PassStringsBetweenFiles"

' #^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#
'
' Hi, Kasparov's back
'  This module is to enable you to pass strings between VB prgrams
'  Simple idea:
'  1-(Target Project): Just add a "Hook Me" command in the main form Load Procedure
'       Notice: The recived data will be saved into the global variable "inMess"
'         I advice you to make a timer to check if this value is changed
'  2-(Send Project): use "SendMess" function to send data, it has
'    those variables SendMess FormName, MessageToBeSent, Target-TitleBar'sName
'
'  Sorry for litle comments
'
'  Notice:
'     You are free to use this code in your applications, Just
'     leave a comment saying that you take it from
'     kasparov03@hotmail.com , Haytham Alaa , Egypt , Cairo
'
' #^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#^#

Type COPYDATASTRUCT
   dwData As Long
   cbData As Long
   lpData As Long
End Type

Private Declare Function FindWindow Lib "user32" Alias _
   "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName _
   As String) As Long

Private Declare Function SendMessage Lib "user32" Alias _
   "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal _
   wParam As Long, lParam As Any) As Long

'Copies a block of memory from one location to another.
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
   (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Public Const GWL_WNDPROC = (-4)
Public Const WM_COPYDATA = &H4A
Global lpPrevWndProc As Long
Global inMess


Declare Function CallWindowProc Lib "user32" Alias _
   "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As _
   Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As _
   Long) As Long

Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
   (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As _
   Long) As Long

Public Sub Hook(Frm As Form)
    lpPrevWndProc = SetWindowLong(Frm.hwnd, GWL_WNDPROC, _
    AddressOf WindowProc)
End Sub

Public Sub Unhook(Frm As Form)
    Dim temp As Long
    temp = SetWindowLong(Frm.hwnd, GWL_WNDPROC, lpPrevWndProc)
End Sub

Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, _
   ByVal wParam As Long, ByVal lParam As Long) As Long
    If uMsg = WM_COPYDATA Then
        Call mySub(lParam)
    End If
    WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, _
       lParam)
End Function

Sub mySub(lParam As Long)
    Dim cds As COPYDATASTRUCT
    Dim buf(1 To 255) As Byte

    Call CopyMemory(cds, ByVal lParam, Len(cds))

    Select Case cds.dwData
     Case 1
        Debug.Print "got a 1"
     Case 2
        Debug.Print "got a 2"
     Case 3
        Call CopyMemory(buf(1), ByVal cds.lpData, cds.cbData)
        a$ = StrConv(buf, vbUnicode)
        a$ = Left$(a$, InStr(1, a$, Chr$(0)) - 1)
        inMess = a$
        inReceived
    End Select
End Sub

Public Function SendMess(Frm As Form, ByVal Mess As String, ByVal Wind)
    Dim cds As COPYDATASTRUCT
    Dim ThWnd As Long
    Dim buf(1 To 255) As Byte

' Get the hWnd of the target application
    ThWnd = FindWindow(vbNullString, Wind)
    a$ = Mess
' Copy the string into a byte array, converting it to ASCII
    Call CopyMemory(buf(1), ByVal a$, Len(a$))
    cds.dwData = 3
    cds.cbData = Len(a$) + 1
    cds.lpData = VarPtr(buf(1))
    i = SendMessage(ThWnd, WM_COPYDATA, Frm.hwnd, cds)
End Function

Public Function inReceived()
'   MsgBox "Enter the code that will be executed each time the program receives a new Message" & vbCrLf & "Go to a function called inReceived"
      sMess = Split(inMess, ":")
      For i = 0 To frmMain.lstBlock.ListCount - 1
         If frmMain.lstBlock.List(i) = sMess(0) Then Exit Function
         If frmMain.lstBlk.List(i) = sMess(1) Then Exit Function
      Next
      frmMain.lstBlock.AddItem sMess(1)
      frmMain.lstBlk.AddItem sMess(0)
End Function
