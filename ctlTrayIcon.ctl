VERSION 5.00
Begin VB.UserControl TrayIcon 
   Appearance      =   0  'Flat
   ClientHeight    =   705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1770
   EditAtDesignTime=   -1  'True
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   705
   ScaleWidth      =   1770
   ToolboxBitmap   =   "ctlTrayIcon.ctx":0000
   Begin VB.Image imgIcon 
      Height          =   120
      Left            =   855
      Top             =   495
      Width           =   345
   End
   Begin VB.Label lblText 
      Caption         =   "JKTray Icon Control"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   45
      TabIndex        =   0
      Top             =   15
      Width           =   1740
   End
End
Attribute VB_Name = "TrayIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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















'========================================================
'THIS TRAY ICON USER CONTROL WAS WRITTEN BY SOMEBODY ELSE
' SORRY - I CAN'T REMMEBER WHOIT WAS AT THIS POINT!!!
'========================================================


Option Explicit
Private bIconShown As Boolean

Public Event MouseMove(Button As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer)
Public Event MouseUp(Button As Integer, Shift As Integer)
Public Event LeftDblClick()
Public Event RightDblClick()

'this is the shell function that does the work
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
'this is the type that the "shell_notifyicon" needs
Private Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type
'NotifyIconData(uFlags) constants
Private Const NIM_ADD = &H0
Private Const NIM_DELETE = &H2
Private Const NIM_MODIFY = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIF_MESSAGE = &H1
Private Const NIF_DOALL = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
'Mouse movement constants
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Sub SetSysTray(bUpdate As Boolean)
Dim ReturnVal As Long
Dim TrayIcon As NOTIFYICONDATA
TrayIcon.cbSize = Len(TrayIcon)
TrayIcon.hwnd = UserControl.hwnd
TrayIcon.uID = 1&
TrayIcon.uFlags = NIF_DOALL
TrayIcon.uCallbackMessage = WM_MOUSEMOVE
TrayIcon.hIcon = imgIcon.Picture
TrayIcon.szTip = imgIcon.ToolTipText + Chr$(0)
If bUpdate Then
    ReturnVal = Shell_NotifyIcon(NIM_MODIFY, TrayIcon)
Else
    ReturnVal = Shell_NotifyIcon(NIM_ADD, TrayIcon)
End If
bIconShown = True

End Sub

Public Sub Add()
SetSysTray False

End Sub

Public Sub Remove()
Dim ReturnVal As Long
Dim TrayIcon As NOTIFYICONDATA
TrayIcon.cbSize = Len(TrayIcon)
TrayIcon.hwnd = UserControl.hwnd
TrayIcon.uID = 1&
ReturnVal = Shell_NotifyIcon(NIM_DELETE, TrayIcon)
bIconShown = False

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Result As Long
Dim lMsg As Long
Dim iShift As Integer


SetActiveWindow UserControl.Parent.hwnd

If UserControl.Parent.Visible = False Then SetForegroundWindow UserControl.Parent.hwnd

'check for "shift"
If (GetAsyncKeyState(16) And &H8001) <> 0 Then iShift = iShift + 1
'check for "ctrl"
If (GetAsyncKeyState(17) And &H8001) <> 0 Then iShift = iShift + 2
'check for "alt"
If (GetAsyncKeyState(18) And &H8001) <> 0 Then iShift = iShift + 4


'the value of X will vary depending upon the scalemode setting
If ScaleMode = vbPixels Then
   lMsg = X
Else
   lMsg = X / Screen.TwipsPerPixelX
End If

Select Case lMsg
   Case WM_LBUTTONUP
      RaiseEvent MouseUp(vbLeftButton, iShift)
   Case WM_RBUTTONUP
      RaiseEvent MouseUp(vbRightButton, iShift)
   Case WM_LBUTTONDOWN
      RaiseEvent MouseDown(vbLeftButton, iShift)
   Case WM_RBUTTONDOWN
      RaiseEvent MouseDown(vbRightButton, iShift)
   Case WM_LBUTTONDBLCLK
      RaiseEvent LeftDblClick
   Case WM_RBUTTONDBLCLK
      RaiseEvent RightDblClick
   Case WM_MOUSEMOVE
      RaiseEvent MouseMove(Button, iShift)
End Select

End Sub

Public Sub Restore()
If Not TypeOf UserControl.Parent Is Form Then
    Err.Raise vbObjectError + 3000, "TrayIcon", "Invalid Form object"
Else
    UserControl.Parent.WindowState = vbNormal
    SetForegroundWindow UserControl.Parent.hwnd
    UserControl.Parent.Show
End If

End Sub

Private Sub UserControl_Terminate()
Remove

End Sub

Public Property Get Picture() As Picture
Set Picture = imgIcon.Picture

End Property

Public Property Set Picture(ByVal New_Picture As Picture)
Set imgIcon.Picture = New_Picture
PropertyChanged "Picture"
Refresh

End Property

Public Sub Refresh()
UserControl.Refresh
If Not bIconShown Then Exit Sub
SetSysTray True

End Sub

Public Property Get ToolTipText() As String
ToolTipText = imgIcon.ToolTipText

End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
imgIcon.ToolTipText() = New_ToolTipText
PropertyChanged "ToolTipText"
Refresh

End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Set Picture = PropBag.ReadProperty("Picture", Nothing)
imgIcon.ToolTipText = PropBag.ReadProperty("ToolTipText", "")

End Sub

Private Sub UserControl_Resize()
'With imgIcon
'    .Top = UserControl.ScaleTop
'    .Left = UserControl.ScaleLeft
'    .Height = UserControl.ScaleHeight
'    .Width = UserControl.ScaleWidth
'End With

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("Picture", Picture, Nothing)
Call PropBag.WriteProperty("ToolTipText", imgIcon.ToolTipText, "")

End Sub

