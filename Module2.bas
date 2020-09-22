Attribute VB_Name = "Module2"
'==================
'R E A D    T H I S
'==================
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















' THANKS TO THE GUY BELOW FOR THIS AWESOME CODE !!!


':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' Locate Desktop Icons Source Code VB6
' Created 11 April 1999
' By Paul Pavlic
' abuse and advice to    pepp@cyberdude.com
' Feel free to use anyway you can
' FreeWare
' Works on Win95
'
' Special thanks to Bruce McKinney
' Author of Hardcore Visual Basic
' A wonderful book with a great many insights
'
' Why?
' I looked all over the Internet, and never once saw any working
' code to do this, thought I would make a contribution
' Why call it Paulies Pet?
' Was the name of my Desktop Pet in VB3 rewriting it now to VB6
'
' This code resolves Explorer page faults when trying to send
' LVM_GETITEMPOSITION to the Desktop Listview
'
'*************************************







Option Explicit

'constants
Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const OPEN_ALWAYS = 4
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const SECTION_MAP_WRITE = &H2
Public Const FILE_MAP_WRITE = SECTION_MAP_WRITE

'NOT documented in Win32api.txt
Public Const PAGE_READWRITE As Long = &H4

Public Const LVM_GETTITEMCOUNT& = (&H1000 + 4)
Public Const LVM_SETITEMPOSITION& = (&H1000 + 15)
Public Const LVM_FIRST = &H1000
Public Const LVM_GETITEMPOSITION = (LVM_FIRST + 16)
Public Const LVM_GETITEMTEXT = LVM_FIRST + 45

Public Const GW_CHILD = 5

'damn  hell of a lot of declares
'copymemory *3 avoid byval in code - bug? works this way




Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
(hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Sub CopyMemoryOne Lib "kernel32" Alias "RtlMoveMemory" _
(ByVal hpvDest&, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Sub CopyMemoryTwo Lib "kernel32" Alias "RtlMoveMemory" _
(hpvDest As Any, ByVal hpvSource&, ByVal cbCopy As Long)

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam _
As Any) As Long
Declare Function SendMessageByLong& Lib "user32" Alias "SendMessageA" _
(ByVal hwnd&, ByVal wMsg&, ByVal wParam&, ByVal lParam&)

Private Declare Function FindWindow& Lib "user32" Alias "FindWindowA" _
(ByVal lpClassName As String, ByVal lpWindowName As String)
Private Declare Function FindWindowEx& Lib "user32" Alias _
"FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter _
As Long, ByVal lpClassName As String, ByVal lpWindowName As String)

'declares for printing to the desktop or other window for debug purposes
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal _
hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As _
String, ByVal nCount As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
ByVal hDC As Long) As Long
Public Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, _
lpRect As Any, ByVal bErase As Long) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'declares for memory-mapped files
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" _
(ByVal lpFileName As String, ByVal dwDesiredAccess As Long, _
ByVal dwShareMode As Long, lpSecurityAttributes As Any, _
ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As _
Long, ByVal hTemplateFile As Long) As Long
' changed lpFileMappigAttributes to Any, makes life much easier
Public Declare Function CreateFileMappingTwo Lib "kernel32" Alias _
"CreateFileMappingA" (ByVal hFile As Long, lpFileMappigAttributes _
As Any, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, _
ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
Public Declare Function MapViewOfFile Lib "kernel32" (ByVal _
hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal _
dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal _
dwNumberOfBytesToMap As Long) As Long
Public Declare Function UnmapViewOfFile Lib "kernel32" (lpBaseAddress _
As Any) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject _
As Long) As Long
Public Declare Function FlushViewOfFile Lib "kernel32" (ByVal lpBaseAddress _
As Long, ByVal dwNumberOfBytesToFlush As Long) As Long

'type declarations
Public Type LV_ITEM ' might need this if we ever figure out
    mask As Long    ' how to retrieve the text
    iItem As Long
    iSubItem As Long
    State As Long
    stateMask As Long
    pszText As Long
    cchTextMax As Long
    iImage As Long
    lParam As Long  ' I think we might need a second
    iIndent As Long ' memory mapped file
End Type

Public Type POINTAPI
        X As Long
        Y As Long
End Type


'dimension some variables
Dim pNull As Long
Dim MyValue As Long
Dim MyValue2 As Long
Dim sFileName As String
Dim CurrentDirectory As String
Dim hdesk As Long
Dim i As Long
Public ixPOS() As Integer
Public iyPOS() As Integer




Function FindIcons() As Integer
Dim hFile As Long
Dim hFileMap As Long
Dim pFileMap As Long
Dim c As POINTAPI
Dim iCount As Long

pNull = 0
hdesk = FindWindow("progman", vbNullString)
hdesk = FindWindowEx(hdesk, 0, "shelldll_defview", vbNullString)
hdesk = FindWindowEx(hdesk, 0, "syslistview32", vbNullString)
'hdesk is the handle of the Desktop's listview
iCount = SendMessageByLong(hdesk, LVM_GETTITEMCOUNT, 0, 0&)



'create a memory-mapped file /////
CurrentDirectory = App.Path
If Right(CurrentDirectory, 1) <> "\" Then
   CurrentDirectory = CurrentDirectory & "\"
End If
sFileName = CurrentDirectory & "TEMPPPPP.PPP"
' Open file
hFile = CreateFile(sFileName, GENERIC_READ Or GENERIC_WRITE, 0, ByVal pNull, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, pNull)
' get handle
hFileMap = CreateFileMappingTwo(hFile, ByVal pNull, PAGE_READWRITE, 0, 16, "MyMapping")
' Get pointer to memory representing file
pFileMap = MapViewOfFile(hFileMap, FILE_MAP_WRITE, 0, 0, 0)

ReDim Preserve ixPOS(iCount)
ReDim Preserve iyPOS(iCount)
ReDim Preserve BumpOrgX(iCount)
ReDim Preserve BumpOrgY(iCount)


For i = 0 To iCount - 1
   'lparam is mem-map file Pointer
   Call SendMessageByLong(hdesk, LVM_GETITEMPOSITION, i, pFileMap)
   'copy returned to our POINTAPI (c.x,c.y)
   CopyMemoryTwo c, pFileMap, 8
   'show me where the icons are
   ixPOS(i + 1) = c.X
   iyPOS(i + 1) = c.Y
   BumpOrgX(i + 1) = c.X
   BumpOrgY(i + 1) = c.Y
   'put value in our arrays
   'IconPosition(i) = c
   'back up array for swapping later
   'IconPosition2(i) = c
Next i

'Release resources back to windows
FlushViewOfFile pFileMap, 8
UnmapViewOfFile pFileMap
CloseHandle hFileMap
CloseHandle hFile



FindIcons = iCount

End Function





Public Function GetDesktopWindow() As Long
Dim lng As Long
lng = FindWindow("Progman", vbNullString)
lng = GetWindow(lng, GW_CHILD)
GetDesktopWindow = GetWindow(lng, GW_CHILD)

End Function
