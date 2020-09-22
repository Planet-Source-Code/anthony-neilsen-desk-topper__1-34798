Attribute VB_Name = "tpnWINmod"
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

Declare Function ReleaseCapture Lib "user32" () As Long

Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Global Const HWND_NOTOPMOST = -2
Global Const SWP_NOACTIVATE = &H10
Global Const SWP_SHOWWINDOW = &H40
Global Const HWND_BOTTOM = 1
Global Const HWND_TOP = 0
Global Const HWND_TOPMOST = -1
Global Const SWP_NOOWNERZORDER = &H200
Global Const SWP_NOSIZE = &H1
Global Const SWP_NOZORDER = &H4
Global Const SWP_NOMOVE = &H2


Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type


Public Function GetBitmapRegion(cPicture As StdPicture, cTransparent As Long)
Dim hRgn As Long, tRgn As Long
Dim X As Integer, Y As Integer, X0 As Integer
Dim hDC As Long, BM As BITMAP
hDC = CreateCompatibleDC(0)
If hDC Then
   SelectObject hDC, cPicture
   GetObject cPicture, Len(BM), BM
   hRgn = CreateRectRgn(0, 0, BM.bmWidth, BM.bmHeight)
        
   For Y = 0 To BM.bmHeight
       For X = 0 To BM.bmWidth
           While X <= BM.bmWidth And GetPixel(hDC, X, Y) <> cTransparent
               X = X + 1
           Wend
           X0 = X
           While X <= BM.bmWidth And GetPixel(hDC, X, Y) = cTransparent
               X = X + 1
           Wend
           If X0 < X Then
               tRgn = CreateRectRgn(X0, Y, X, Y + 1)
               CombineRgn hRgn, hRgn, tRgn, 4
               DeleteObject tRgn
           End If
       Next X
   Next Y
   GetBitmapRegion = hRgn
   DeleteObject SelectObject(hDC, cPicture)
End If
DeleteDC hDC
End Function


