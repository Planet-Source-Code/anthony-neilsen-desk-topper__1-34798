VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tpnWIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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

Public Sub DragOBJ(frm)

ReleaseCapture
SendMessage frm.hwnd, &HA1, 2, 0&

End Sub
Sub SizeOBJ(f, act)
Dim act2 As Integer
Select Case UCase(act)
Case "T": act2 = 12
Case "B": act2 = 15
Case "L": act2 = 10
Case "R": act2 = 11
Case "TL": act2 = 13
Case "TR": act2 = 14
Case "BL": act2 = 16
Case "BR": act2 = 17
Case Else
Exit Sub
End Select

Call ReleaseCapture
Call SendMessage(f.hwnd, &HA1, act2, 0)

End Sub


Sub FormOnTop(f, act)
Select Case act

Case 1
   SetWindowPos f.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOSIZE Or SWP_NOMOVE
Case 0
   SetWindowPos f.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOSIZE Or SWP_NOMOVE
Case Else

End Select

End Sub


Sub SetRegion(ff, cc As Long)
Dim hRgn As Long
hRgn = GetBitmapRegion(ff.Picture, cc)
SetWindowRgn ff.hwnd, hRgn, True

End Sub

