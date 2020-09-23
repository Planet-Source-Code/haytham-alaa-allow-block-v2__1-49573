VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Allow / Block"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   2040
      Top             =   840
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Password"
      Height          =   495
      Left            =   1920
      TabIndex        =   11
      Top             =   1680
      Width           =   495
   End
   Begin VB.ListBox lstAll 
      Height          =   2010
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   1695
   End
   Begin VB.ListBox lstBlk 
      Height          =   2010
      Left            =   2520
      TabIndex        =   9
      Top             =   1080
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog dlg1 
      Left            =   1920
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdDelAll 
      Caption         =   "Delete"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add ....."
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdDelBlk 
      Caption         =   "Delete"
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   3120
      Width           =   855
   End
   Begin VB.ListBox lstBlock 
      Height          =   2010
      Left            =   2520
      TabIndex        =   1
      Top             =   3600
      Width           =   1695
   End
   Begin VB.ListBox lstAllow 
      Height          =   2010
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   $"frmMain.frx":0000
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label2 
      Caption         =   "Blocked List :-"
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Allowed List :-"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label lblAllow 
      Caption         =   "<-------"
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label lblBlock 
      Caption         =   "-------->"
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   1320
      Width           =   495
   End
   Begin VB.Menu mnusystray 
      Caption         =   "SysTray"
      Visible         =   0   'False
      Begin VB.Menu mnuSysOpen 
         Caption         =   "Open Window"
      End
      Begin VB.Menu mnuSyMin 
         Caption         =   "Minimize"
      End
      Begin VB.Menu mnuSysExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fl(), Pass
Dim result As Integer
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Sub Timer2_Timer()
   'THIS IS ACTUAL KEY FINDING
   ' SOME ASCII OR USED FOR WINDOWS FUNCTION KEYS, TO STORE TEXT IN TEXTBOX
   For i = 1 To 255
      result = 0
      result = GetAsyncKeyState(i)
      If result = -32767 Then
         ' HIDES THE APPLICATION WHEN YOU PRESS "F10" ON KEYBOARD
         If i = 121 Then Me.Visible = False
         ' THIS WILL BE RAISED WHEN YOU PRESS "F11" ON KEYBOARD.
         If i = 122 Then
            strs = InputBox("What is the Password?", "PassWord Request", "0000")
            If Not LCase(strs) = EnDe(Pass, "25notbroken") Then
               MsgBox "Invalid Password"
               Me.Hide
               Exit Sub
            Else
               Me.Show
            End If
         End If
      End If
   Next i
End Sub

Private Sub cmdAdd_Click()
   With dlg1
      .DialogTitle = "Add a file to Block-List"
      .Filter = "All Files *.*|*.*"
      .ShowOpen
      If .FileName = "" Then Exit Sub
      For i = 0 To lstBlock.ListCount - 1
         If .FileName = lstBlock.List(i) Then
            MsgBox "File is Blocked"
            Exit Sub
         End If
         If .FileName = lstAllow.List(i) Then
            MsgBox "File is Allowed"
            Exit Sub
         End If
      Next
      fFile = Split(.FileName, "\")
      lstBlk.AddItem fFile(UBound(fFile))
      lstBlock.AddItem .FileName
   End With
   ChkChange
End Sub

Private Sub cmdDelAll_Click()
   If lstAll.ListIndex = -1 Then Exit Sub
   lstAllow.RemoveItem lstAll.ListIndex
   lstAll.RemoveItem lstAll.ListIndex
   ChkChange
End Sub

Private Sub cmdDelBlk_Click()
   If lstBlk.ListIndex = -1 Then Exit Sub
   lstBlock.RemoveItem lstBlk.ListIndex
   lstBlk.RemoveItem lstBlk.ListIndex
   ChkChange
End Sub

Private Sub Command1_Click()
   ss = InputBox("Enter previous password:")
   If ss = "" Then Exit Sub
   If ss = EnDe(Pass, "25notbroken") Then
      s2 = InputBox("Enter the new password:")
      Pass = EnDe(s2, "25notbroken")
      Kill App.Path & "\pass.alb"
      Open App.Path & "\pass.ALB" For Output As #1
         Write #1, Pass
      Close #1
   End If
End Sub

Private Sub Form_Load()
   On Error Resume Next
   Hook Me
   Open App.Path & "\pass.ALB" For Input As #1
      Input #1, Pass
   Close #1
'   If Check = "no" Then Me.Hide
   Open App.Path & "\list.ALB" For Input As #1
      Input #1, s
      lst1 = Split(s, "<blk>")
      For i = 1 To UBound(lst1) - 1
         lstBlock.AddItem lst1(i)
         sFile = Split(lst1(i), "\")
         lstBlk.AddItem sFile(UBound(sFile))
      Next
      lst2 = Split(s, "<all>")
      For i = 1 To UBound(lst2) - 1
         lstAllow.AddItem lst2(i)
         sFile2 = Split(lst2(i), "\")
         lstAll.AddItem sFile2(UBound(sFile2))
      Next
   Close #1
   ChkChange

   ' This part to add the program in the run list so that it will be executed everytime the computer starts.
   SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "AllBlk", App.Path & "\" & App.EXEName & ".exe", REG_SZ

   'CHECKS FIRST WEATHER APPLICATION ALREADY RUNNING
   If App.PrevInstance = True Then End
   'HIDES THE APPLICATION FROM 'CTRL-ALT-DEL ' TASK LIST .
   App.TaskVisible = False
   ' HIDES THE APPLICATION
   Me.Visible = False

End Sub

Private Sub Form_LostFocus()
   Me.Hide
   'Me.WindowState = vbMinimized
End Sub

Private Sub Form_Resize()
  If (Me.WindowState = vbMinimized) Then Me.Hide

End Sub

Private Sub lblAllow_Click()
   If lstBlk.ListIndex = -1 Then Exit Sub
   lstAllow.AddItem lstBlock.List(lstBlk.ListIndex)
   lstBlock.RemoveItem lstBlk.ListIndex
   lstAll.AddItem lstBlk.List(lstBlk.ListIndex)
   lstBlk.RemoveItem lstBlk.ListIndex
   ChkChange
End Sub

Private Sub lblBlock_Click()
   If lstAll.ListIndex = -1 Then Exit Sub
   lstBlock.AddItem lstAllow.List(lstAll.ListIndex)
   lstAllow.RemoveItem lstAll.ListIndex
   lstBlk.AddItem lstAll.List(lstAll.ListIndex)
   lstAll.RemoveItem lstAll.ListIndex
   ChkChange
End Sub

Private Function ChkChange()
   ReDim Preserve fl(lstBlock.ListCount + 3)
   For ii = 0 To lstBlock.ListCount - 1
      Close fl(ii + 2)
      fl(ii + 2) = FreeFile
      Open lstBlock.List(ii) For Random As fl(ii + 2)
   Next

   Kill App.Path & "\list.ALB"
   Close fl(lstBlock.ListCount + 2)
   fl(lstBlock.ListCount + 2) = FreeFile
   Open App.Path & "\list.ALB" For Output As fl(lstBlock.ListCount + 2)
      lst1 = "<blk>"
      For i = 0 To lstBlock.ListCount - 1
         lst1 = lst1 & lstBlock.List(i) & "<blk>"
      Next

      lst2 = "<all>"
      For i = 0 To lstAllow.ListCount - 1
         lst2 = lst2 & lstAllow.List(i) & "<all>"
      Next

      Write #fl(lstBlock.ListCount + 2), lst1 & lst2
   Close fl(lstBlock.ListCount + 2)
End Function

' This function i used to encrypt the length of the password to make it so hard to be breaked
Function EnDe(ByVal Secret$, Password$)
    L = Len(Password$)
    For X = 1 To Len(Secret$)
        Char = Asc(Mid$(Password$, (X Mod L) - L * ((X Mod L) = 0), 1))
        Mid$(Secret$, X, 1) = Chr$(Asc(Mid$(Secret$, X, 1)) Xor Char)
    Next
    EnDe = Secret$
End Function
