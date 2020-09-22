VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TrayIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarToolTipText As String 'local copy
Private mvarIcon As PictureBox 'local copy
Private mvarHandle As Long 'local copy
'this is whether or not an icon is shown
Public bIconShown As Boolean

'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent Click[(arg1, arg2, ... , argn)]
Public Event Click(Button As Integer, Keyboard As Integer)

'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent MouseMove[(arg1, arg2, ... , argn)]
Public Event MouseMove(Keyboard As Integer)

'this is the shell function that does the work
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

'this is the type that the "shell_notifyicon" needs
Private Type NOTIFYICONDATA
        cbSize As Long
        hWnd As Long
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

Public Sub Show()
    Dim ReturnVal As Long
    Dim TrayIcon As NOTIFYICONDATA
    TrayIcon.cbSize = Len(TrayIcon)
    TrayIcon.hWnd = mvarHandle
    TrayIcon.uID = 1&
    TrayIcon.uFlags = NIF_DOALL
    TrayIcon.uCallbackMessage = WM_MOUSEMOVE
    TrayIcon.hIcon = mvarIcon
    TrayIcon.szTip = mvarToolTipText + Chr$(0)
    ReturnVal = Shell_NotifyIcon(NIM_ADD, TrayIcon)
    bIconShown = True
End Sub

Public Sub Delete()
    Dim ReturnVal As Long
    Dim TrayIcon As NOTIFYICONDATA
    TrayIcon.cbSize = Len(TrayIcon)
    TrayIcon.hWnd = mvarHandle
    TrayIcon.uID = 1&
    ReturnVal = Shell_NotifyIcon(NIM_DELETE, TrayIcon)
    bIconShown = False
End Sub

Public Property Let Icon(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Icon = Form1
    Set mvarIcon = vData
    IconUpdate
End Property

Public Property Get Icon() As Object
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Icon
    Set Icon = mvarIcon
End Property
Public Property Let hWnd(Handle As Long)
    If bIconShown Then Exit Property
    mvarHandle = Handle
End Property
Public Property Let ToolTipText(vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ToolTipText = 5
    mvarToolTipText = vData
    IconUpdate
End Property

Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_UserMemId = 0
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ToolTipText
    ToolTipText = mvarToolTipText
End Property

Private Sub IconUpdate()
    If Not bIconShown Then Exit Sub
    Dim ReturnVal As Long
    Dim TrayIcon As NOTIFYICONDATA
    TrayIcon.cbSize = Len(TrayIcon)
    TrayIcon.uID = 1&
    TrayIcon.uFlags = NIF_DOALL
    TrayIcon.hIcon = mvarIcon
    TrayIcon.hWnd = mvarHandle
    TrayIcon.szTip = mvarToolTipText + Chr$(0)
    TrayIcon.uCallbackMessage = WM_MOUSEMOVE
    ReturnVal = Shell_NotifyIcon(NIM_MODIFY, TrayIcon)
    Debug.Print ReturnVal
End Sub
