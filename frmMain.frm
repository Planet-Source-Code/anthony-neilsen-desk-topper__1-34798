VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6795
   ClientLeft      =   1860
   ClientTop       =   2415
   ClientWidth     =   9885
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   9885
   Begin VB.CommandButton Command1 
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   360
      TabIndex        =   5
      Top             =   360
      Width           =   180
   End
   Begin VB.Timer BumpTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   1110
      Top             =   975
   End
   Begin VB.CommandButton cmdBump 
      Height          =   240
      Index           =   3
      Left            =   585
      TabIndex        =   4
      Top             =   345
      Width           =   225
   End
   Begin VB.CommandButton cmdBump 
      Height          =   240
      Index           =   2
      Left            =   90
      TabIndex        =   3
      Top             =   330
      Width           =   225
   End
   Begin VB.CommandButton cmdBump 
      Height          =   240
      Index           =   1
      Left            =   345
      TabIndex        =   2
      Top             =   600
      Width           =   225
   End
   Begin VB.CommandButton cmdBump 
      Height          =   240
      Index           =   0
      Left            =   345
      TabIndex        =   1
      Top             =   90
      Width           =   225
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   7755
      TabIndex        =   0
      Top             =   9375
      Width           =   1140
   End
   Begin DeskTopper.TrayIcon TrayIcon1 
      Left            =   8055
      Top             =   495
      _ExtentX        =   1746
      _ExtentY        =   635
   End
   Begin VB.Image TrayPic 
      Height          =   240
      Left            =   8460
      Picture         =   "frmMain.frx":0E42
      Top             =   135
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==================
'R E A D    T H I S
'==================
'
'ABOUT THE PROGRAM
'When this program is first run, it will make a copy
'of your desktop and put it in the app's path in
'the folder \Original
'However you desktop looked when you first ran the
'program is how the program restores your desktop
'to original from the popupmenu
'If you want to restart the original copy, then
'delete the file "Desktopper.ini"
'and the folder "MyShortcuts" in the app's path
'
'WARNING!!!!!!!!!!
'YOU REALLY SHOULD BACK UP YOUR DESKTOP ICONS
'BEFORE RUNNING THIS PROGRAM, AS IT DOES A LOT
'OF MANIPULATING OF THE DESKTOP ICONS
'JUST COPY ALL THE FILES FROM C:\Windows\Desktop
'TO A FOLDER OF YOUR CHOICE
'YOU'VE BEEN WARNED.
'
'IF YOU DON'T BACKUP AND YOU LOSE YOUR ICONS, THEN
'I DON'T WANT TO KNOW ABOUT IT!
'
'
'IN THE RARE CASE THAT THIS PROGRAM CRASHES AND
'LEAVES YOUR DESKTOP INCOMPLETE, YOU CAN RESTORE
'IT BY COPYING THE FILES BACK TO C:\Windows\Desktop
'
'ONE MORE THING - WHEN USING THE BUMP BUTTONS
'YOU CAN INCREASE THE BUMP AMOUNT BY HOLDING DOWN
'THE <Ctrl>, <Shift> and/or <Alt> KEYS
'
'
'HAVE FUN AND IF YOU HAVE ANY PROBLEMS YOU CAN
'CONTACT ME AT kleena@optushome.com.au
'













Option Explicit
Private mSHIFT As Integer
Private BumpCnt As Integer
Private BumpDir As Integer
Private BumpInc As Integer
Private puMenu As Long
Private tWIN As New tpnWIN
Private BumpX As Integer
Private BumpY As Integer
Private MaxBumpIcons As Integer
Private MyShortcutsFolder As String

Private CurDesktop As String

Private Desktops() As String
Private MaxDesktops As Integer

Dim dtLinks() As String
Dim dtFolds() As String

Dim SubFolds() As String
Private MaxLinks As Integer
Private MaxFolds As Integer
Private MaxIcons As Integer

Function CreateXYstring() As String
Dim str7 As String
Dim i7 As Integer
Dim max As Integer
max = FindIcons

For i7 = 1 To max
   str7 = str7 & Right("00000" & Mid(Str(ixPOS(i7)), 2), 5) & Right("00000" & Mid(Str(iyPOS(i7)), 2), 5)
Next i7
CreateXYstring = str7


End Function

Sub SaveThisDesktop(fold As String)
Dim bln As Boolean
Dim i7 As Integer
Dim xyStr As String
Dim fso




xyStr = CreateXYstring


Call KillThisDesktop(MyShortcutsFolder & "\" & fold)

Set fso = CreateObject("Scripting.FileSystemObject")
 
If Not fso.FolderExists(MyShortcutsFolder & "\" & fold) Then
   MkDir MyShortcutsFolder & "\" & fold
End If



MaxFolds = 0
MaxLinks = 0
bln = RecurseFolderList("C:\Windows\Desktop")


For i7 = 1 To MaxFolds
   If Not fso.FolderExists(MyShortcutsFolder & "\" & fold & Mid(dtFolds(i7), 19)) Then
      MkDir MyShortcutsFolder & "\" & fold & Mid(dtFolds(i7), 19)
   End If
Next i7

For i7 = 1 To MaxLinks
   'List2.AddItem Mid(dtLinks(i7), 19)
   'List2.AddItem dtLinks(i7)
   SetAttr dtLinks(i7), 0
   FileCopy dtLinks(i7), MyShortcutsFolder & "\" & fold & Mid(dtLinks(i7), 19)
Next i7



Set fso = Nothing

Close #1: Open MyShortcutsFolder & "\" & fold & "\xy.dat" For Output As #1
Print #1, xyStr
Close #1


End Sub

Public Function RecurseFolderList(FolderName As String) As Boolean

On Error Resume Next
Dim fso, f, fc, fj, f1
 
Set fso = CreateObject("Scripting.FileSystemObject")
 
If Err.Number > 0 Then
    RecurseFolderList = False
    Exit Function
End If
 
On Error GoTo 0
If fso.FolderExists(FolderName) Then

   Set f = fso.GetFolder(FolderName)
   Set fc = f.Subfolders
   Set fj = f.Files
   For Each f1 In fc
      MaxFolds = MaxFolds + 1
      ReDim Preserve dtFolds(MaxFolds)
      dtFolds(MaxFolds) = f1
      RecurseFolderList (f1)
   Next
   
   'For each folder check for any files
   For Each f1 In fj
      If Right(f1, 6) <> "xy.dat" Then
         MaxLinks = MaxLinks + 1
         ReDim Preserve dtLinks(MaxLinks)
         dtLinks(MaxLinks) = f1
      End If
   Next
   
   Set f = Nothing
   Set fc = Nothing
   Set fj = Nothing
   Set f1 = Nothing

Else
   RecurseFolderList = False
End If

Set fso = Nothing

End Function

Sub SaveDesktops()
Dim i7 As Integer
Close #1: Open App.Path & "\DeskTopper.ini" For Output As #1
Print #1, CurDesktop

For i7 = 1 To MaxDesktops
   Print #1, Desktops(i7)
Next i7



End Sub


Sub InitDesktops()
Close #1: Open App.Path & "\DeskTopper.ini" For Random As #1
If LOF(1) = 0 Then
   Close #1
   Open App.Path & "\DeskTopper.ini" For Output As #1
   Print #1, "Original"
   Close #1
   Call SaveThisDesktop("Original")
End If

End Sub




Private Sub BumpTimer_Timer()
If BumpCnt < 5 Then
   BumpCnt = BumpCnt + 1
Else
   Call LoopBump
End If

End Sub

Sub LoopBump()
Select Case mSHIFT
Case 0: BumpInc = 1
Case 1: BumpInc = 5
Case 2: BumpInc = 10
Case 3: BumpInc = 15
Case 4: BumpInc = 20
Case 5: BumpInc = 30
Case 6: BumpInc = 50
Case 7: BumpInc = 100
Case Else
BumpInc = 1
End Select

Select Case BumpDir
Case 0: Call BumpUP
Case 1: Call BumpDOWN
Case 2: Call BumpLEFT
Case 3: Call BumpRIGHT
Case Else
End Select

End Sub

Sub DoBump()
Dim i7 As Integer
Dim lng As Long
Dim xPos As Integer
Dim yPos As Integer
Dim ret As Long
For i7 = 0 To MaxBumpIcons - 1

   xPos = BumpOrgX(i7 + 1) + BumpX
   yPos = BumpOrgY(i7 + 1) + BumpY
   
   lng = GetDesktopWindow
   ret = SendMessage(lng, LVM_SETITEMPOSITION, i7, ByVal CLng(xPos + yPos * &H10000))
   
Next i7

End Sub
Sub BumpUP()
BumpY = BumpY - BumpInc
Call DoBump

End Sub
Sub BumpDOWN()
BumpY = BumpY + BumpInc
Call DoBump

End Sub
Sub BumpLEFT()
BumpX = BumpX - BumpInc
Call DoBump

End Sub
Sub BumpRIGHT()
BumpX = BumpX + BumpInc
Call DoBump

End Sub

Private Sub cmdBump_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
mSHIFT = Shift
BumpCnt = 0
BumpDir = Index
Call LoopBump
BumpTimer.Enabled = True

End Sub

Private Sub cmdBump_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
mSHIFT = Shift
BumpTimer.Enabled = False

End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmMain.Hide
End Sub

Private Sub Form_Load()
Dim i7 As Integer
Set TrayIcon1.Picture = TrayPic.Picture
TrayIcon1.Add

frmMain.Width = 960
frmMain.Height = 960
frmMain.Top = Screen.Height - frmMain.Height - 390
frmMain.Left = Screen.Width - frmMain.Width - 300

MyShortcutsFolder = App.Path & "\MyShortcuts"


Call CreateShortcutsFolder
Call InitDesktops


Call LoadDesktops
Call ShowThisDesktop(CurDesktop)



End Sub


Sub LoadDesktops()
Dim str7 As String
On Error GoTo ErrH

MaxDesktops = 0

Close #1: Open App.Path & "\DeskTopper.ini" For Input As #1
Input #1, CurDesktop
While Not EOF(1)
   MaxDesktops = MaxDesktops + 1
   ReDim Preserve Desktops(MaxDesktops)
   Input #1, Desktops(MaxDesktops)
Wend

Close #1



Exit Sub
ErrH:
If Err.Number = 53 Then 'file not found
   Close #1: Open App.Path & "\DeskTopper.ini" For Output As #1
   Resume
End If


End Sub




Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
   tWIN.DragOBJ frmMain
End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then frmMain.Hide

End Sub

Private Sub Form_Unload(Cancel As Integer)
If CurDesktop <> "" And CurDesktop <> "Original" Then
   Call SaveThisDesktop(CurDesktop)
End If
Call SaveDesktops
Set frmMain = Nothing

End Sub



Private Sub TrayIcon1_MouseUp(Button As Integer, Shift As Integer)
Dim lng As Long
Dim Pot As POINTAPI


GetCursorPos Pot


Call CreatePuMenu
lng = TrackPopupMenu(puMenu, TPM_RETURNCMD, Pot.X, Pot.Y, 0, Me.hwnd, ByVal 0&)


Select Case lng
Case 1000     'new desktop
   Call DoNewDesktop
Case 2000     'delete desktop
   Call DoDeleteDesktop
Case 3000     'restore original
   If CurDesktop <> "" And CurDesktop <> "Original" Then
      Call SaveThisDesktop(CurDesktop)
   End If
   Call ShowThisDesktop("Original")
   CurDesktop = "Original"
Case 4000  'save desktop
   Call SaveThisDesktop(CurDesktop)
   
   
Case 8000      'Bump
   Call InitBumpOrgArray
   frmMain.Show
   
    
Case 9000     'end prog
   Unload frmMain
   End

Case Else

   If lng = 0 Then Exit Sub
   If CurDesktop <> "" And CurDesktop <> "Original" Then
      Call SaveThisDesktop(CurDesktop)
   End If
   Call ShowThisDesktop(Desktops(lng))
   CurDesktop = Desktops(lng)

End Select



End Sub
Sub DoDeleteDesktop()
Dim i7 As Integer
Dim s7 As Integer
Call KillThisDesktop(CurDesktop)
For i7 = 1 To MaxDesktops
   If Desktops(i7) = "CurDesktop" Then Exit For
Next i7

CurDesktop = "Original"

For s7 = i7 To MaxDesktops - 1
   Desktops(s7) = Desktops(s7 + 1)
Next s7
MaxDesktops = MaxDesktops - 1
ReDim Preserve Desktops(MaxDesktops)

Call ShowThisDesktop(CurDesktop)


End Sub




Sub KillThisDesktop(fold As String)
Dim i7 As Integer
Dim bln As Boolean
On Error GoTo ErrH
MaxLinks = 0
MaxFolds = 0
bln = RecurseFolderList(fold)


For i7 = 1 To MaxLinks
   Kill dtLinks(i7)
Next i7


For i7 = MaxFolds To 1 Step -1
   RmDir dtFolds(i7)
Next i7



Exit Sub
ErrH:

Resume Next

End Sub
Sub ShowThisDesktop(fold)
Dim bln As Boolean
Dim i7 As Integer
Dim fso
Dim wDSK As String
Dim s As Integer
Dim xyStr As String

Call KillThisDesktop("C:\Windows\Desktop")



wDSK = "C:\Windows\Desktop"
Set fso = CreateObject("Scripting.FileSystemObject")
 


MaxLinks = 0
MaxFolds = 0
bln = RecurseFolderList(MyShortcutsFolder & "\" & fold)


s = InStr(dtLinks(1), fold) + Len(fold)



For i7 = 1 To MaxFolds
   If Not fso.FolderExists(wDSK & Mid(dtFolds(i7), s)) Then
      MkDir wDSK & Mid(dtFolds(i7), s)
   End If
Next i7

s = InStr(dtLinks(1), fold) + Len(fold)
For i7 = 1 To MaxLinks
   FileCopy dtLinks(i7), wDSK & Mid(dtLinks(i7), s)
Next i7


Set fso = Nothing


xyStr = CreateXYstring
Close #1: Open MyShortcutsFolder & "\" & fold & "\xy.dat" For Input As #1
Line Input #1, xyStr
Close #1
Call ArrangeIcons(xyStr)


End Sub

Sub ArrangeIcons(str7)
Dim xPos As Integer
Dim yPos As Integer
Dim i7 As Integer
Dim c As Integer
Dim lng As Long
Dim ret As Long
For i7 = 1 To Len(str7) Step 10

   xPos = Val(Mid(str7, i7, 5))
   yPos = Val(Mid(str7, i7 + 5, 5))
   
   lng = GetDesktopWindow
   ret = SendMessage(lng, LVM_SETITEMPOSITION, c, ByVal CLng(xPos + yPos * &H10000))
   c = c + 1
Next i7


End Sub


Sub CreatePuMenu()
Dim L7 As Long
Dim tStr As String
puMenu = CreatePopupMenu()



AppendMenu puMenu, MF_STRING, ByVal 1000&, "New Desktop..."
If MaxDesktops > 0 Then
   AppendMenu puMenu, MF_STRING, ByVal 4000&, "Save Current Desktop"
   If CurDesktop = "Original" Then
      AppendMenu puMenu, MF_STRING Or MF_DISABLED Or MF_GRAYED, ByVal 0&, "Delete Current Desktop"
   Else
      AppendMenu puMenu, MF_STRING, ByVal 2000&, "Delete Current Desktop"
   End If
End If
AppendMenu puMenu, MF_SEPARATOR, ByVal 0&, ByVal 0&

If CurDesktop = "Original" Then
   AppendMenu puMenu, MF_STRING Or MF_CHECKED, ByVal 3000&, "Restore Original"
Else
   AppendMenu puMenu, MF_STRING, ByVal 3000&, "Restore Original"
End If

If MaxDesktops > 0 Then
   AppendMenu puMenu, MF_SEPARATOR, ByVal 0&, ByVal 0&
   AppendMenu puMenu, MF_STRING, ByVal 0&, "Switch to..."
   For L7 = 1 To MaxDesktops
      tStr = Desktops(L7)
      If tStr = CurDesktop Then
      AppendMenu puMenu, MF_STRING Or MF_CHECKED, ByVal L7&, tStr
      Else
      AppendMenu puMenu, MF_STRING, ByVal L7&, tStr
      End If
   Next L7
End If

AppendMenu puMenu, MF_SEPARATOR, ByVal 0&, ByVal 0&
AppendMenu puMenu, MF_STRING, ByVal 8000&, "Bump..."
AppendMenu puMenu, MF_SEPARATOR, ByVal 0&, ByVal 0&
AppendMenu puMenu, MF_STRING, ByVal 9000&, "Exit"


End Sub
Function FindNextDesktop()
Dim i7 As Integer
Dim str7 As String
Dim n As Integer
Dim maxn As Integer
If MaxDesktops = 0 Then
   FindNextDesktop = "My Desktop"
Else
   str7 = "My Desktop"
   For i7 = 1 To MaxDesktops
      If Left(Desktops(i7), 10) = "My Desktop" Then
         n = Val(Mid(Desktops(i7) & " ", 11))
         If n > maxn Then maxn = n
      End If
   Next i7
   FindNextDesktop = "My Desktop" & Str(maxn + 1)
End If


End Function
Sub DoNewDesktop()
Dim str7 As String
Dim xPos As Integer
Dim yPos As Integer
Dim prompt As String

prompt = FindNextDesktop


xPos = Screen.Width - 5550
yPos = Screen.Height - (MaxDesktops * 270) - 3000
str7 = InputBox("Desk Topper", "Create New Desktop", prompt, xPos, yPos)
If UCase(str7) = "ORIGINAL" Then
   MsgBox "This name is reserved!", vbCritical
   Exit Sub
End If
If str7 <> "" Then
   Call AddNewDesktop(str7)
End If



End Sub

Sub AddNewDesktop(nam As String)

MaxDesktops = MaxDesktops + 1
ReDim Preserve Desktops(MaxDesktops)
Desktops(MaxDesktops) = nam

CurDesktop = nam
Call SaveThisDesktop(CurDesktop)


End Sub

Sub CreateShortcutsFolder()
Dim fso, f, fc, fj, f1
 
Set fso = CreateObject("Scripting.FileSystemObject")
 
 
If Not fso.FolderExists(MyShortcutsFolder) Then
   MkDir MyShortcutsFolder
End If


End Sub
Sub InitBumpOrgArray()
MaxBumpIcons = FindIcons
BumpX = 0
BumpY = 0



End Sub
