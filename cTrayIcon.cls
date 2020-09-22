VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTrayIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'JASON:
'a few notes...
'compile mode changed to P-Code
'native code has little performance benefit
'and much bigger files...
'vbaccelerator.com agrees with me on that one!
'new control is only 24KB, compared to 44KB
'dll base address changed from default value

'** This component is NOT compatible with your original control **
'aboutbox removed, event declarations different


'// VBWeb 19-Jul-00
'// Header added

'JKTryIcn Control
'
'Easily add a tray icon to your program, using the
'methods and events of the control.  Simply insert the
'control, name it, program the events you want to
'capture, and set the picture and tip text at run time,
'then invoke the .show method
'
'Created by Jason K
'jason_k_3431@yahoo.com


'this is whether or not an icon is shown
Private bIconShown As Boolean

'To fire these events, use RaiseEvent with the following syntax:
'RaiseEvent MouseMove[(arg1, arg2, ... , argn)]

'// VBWeb 19-Jul-00
'// Events changed, note to X/Y args

Public Event MouseMove(Button As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer)
Public Event MouseUp(Button As Integer, Shift As Integer)
Public Event LeftDblClick()
Public Event RightDblClick()
Private m_Picture As IPictureDisp
Private m_sToolTip As String
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
'// VB Web 19-Jul-00
'// API Added
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Implements isubclass
'// VBWeb 19-Jul-00
'// new sub, to avoid repeated
'// code in refresh & add
'// set bUpdate to true to update
Private Sub SetSysTray(bUpdate As Boolean)
    Dim ReturnVal As Long
    Dim TrayIcon As NOTIFYICONDATA
    TrayIcon.cbSize = Len(TrayIcon)
    TrayIcon.hwnd = picIcon.hwnd
    TrayIcon.uID = 1&
    TrayIcon.uFlags = NIF_DOALL
    TrayIcon.uCallbackMessage = WM_MOUSEMOVE
    TrayIcon.hIcon = m_Picture.Handle
    TrayIcon.szTip = picIcon.ToolTipText + Chr$(0)
    If bUpdate Then
        ReturnVal = Shell_NotifyIcon(NIM_MODIFY, TrayIcon)
    Else
        ReturnVal = Shell_NotifyIcon(NIM_ADD, TrayIcon)
    End If
    bIconShown = True
End Sub
'// VBWeb 19-Jul-00
'// sub changed
Public Sub Add()
    SetSysTray False
End Sub
Public Sub Remove()
    Dim ReturnVal As Long
    Dim TrayIcon As NOTIFYICONDATA
    TrayIcon.cbSize = Len(TrayIcon)
    TrayIcon.hwnd = picIcon.hwnd
    TrayIcon.uID = 1&
    ReturnVal = Shell_NotifyIcon(NIM_DELETE, TrayIcon)
    bIconShown = False
End Sub
'// VB Web 19-Jul-00
'// About Box Removed... sorry!
'Public Sub AboutShow()
'    'show the "frmAbout" form
'    On Error Resume Next
'    frmAbout.Show vbModal
'    On Error GoTo 0
'End Sub
'// VB Web 19-Jul-00
'// Procedure re-written
Private Sub picIcon_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'this procedure receives the callbacks from the System Tray icon.
    Dim Result As Long
    Dim lMsg As Long
    'the value of X will vary depending upon the scalemode setting
    If ScaleMode = vbPixels Then
        lMsg = X
    Else
        lMsg = X / Screen.TwipsPerPixelX
    End If
    SetForegroundWindow hwnd 'need this, so that any popup menu wil remain in the proper place
    Select Case lMsg
        Case WM_LBUTTONUP
            RaiseEvent MouseUp(vbLeftButton, Shift)
        Case WM_RBUTTONUP
            RaiseEvent MouseUp(vbRightButton, Shift)
        Case WM_LBUTTONDOWN
            RaiseEvent MouseDown(vbLeftButton, Shift)
        Case WM_RBUTTONDOWN
            RaiseEvent MouseDown(vbRightButton, Shift)
        Case WM_LBUTTONDBLCLK
            RaiseEvent LeftDblClick
        Case WM_RBUTTONDBLCLK
            RaiseEvent RightDblClick
        Case WM_MOUSEMOVE
            RaiseEvent MouseMove(Button, Shift)
    End Select
    'UserControl.Width = 0
End Sub
'// VB Web 19-Jul-00
'// Procedure Added
Public Sub Restore(frmForm As Object)
    If Not TypeOf frmForm Is Form Then
        Err.Raise vbObjectError + 3000, "TrayIcon", "Invalid Form object"
    Else
        frmForm.WindowState = vbNormal
        SetForegroundWindow (frmForm.hwnd)
        frmForm.Show
    End If
End Sub

Private Sub UserControl_Show()
    If Ambient.UserMode = False Then Label1.Visible = False
End Sub

'// VB Web 19-Jul-00
'// Automatically remove icon when exiting
Private Sub UserControl_Terminate()
    Remove
End Sub
Public Property Get Picture() As IPictureDisp
    Set Picture = m_Picture
End Property
Public Property Set Picture(ByVal New_Picture As IPictureDisp)
    Set m_Picture = New_Picture
    PropertyChanged "Picture"
    Refresh
End Property
'// VB Web 19-Jul-00
'// sub now uses SetSysTray sub
Public Sub Refresh()
'    UserControl.Refresh
    If Not bIconShown Then Exit Sub
    SetSysTray True
End Sub
Public Property Get ToolTipText() As String
    ToolTipText = m_sToolTip
End Property
Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    m_sToolTip = New_ToolTipText
    PropertyChanged "ToolTipText"
    Refresh
End Property
'Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'    Set Picture = PropBag.ReadProperty("Picture", Nothing)
'    picIcon.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
'End Sub
'Private Sub UserControl_Resize()
'    With picIcon
'        .Top = UserControl.ScaleTop
'        .Left = UserControl.ScaleLeft
'        .Height = UserControl.ScaleHeight
'        .Width = UserControl.ScaleWidth
'    End With
'End Sub
'Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'    Call PropBag.WriteProperty("Picture", Picture, Nothing)
'    Call PropBag.WriteProperty("ToolTipText", picIcon.ToolTipText, "")
'End Sub

