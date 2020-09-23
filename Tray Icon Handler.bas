Attribute VB_Name = "modTrayIcon"
Option Explicit

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal lngHandle As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function RegisterWindowMessage Lib "user32.dll" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal lngHandle As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" (ByVal dwMessage As Long, lpData As TrayIconData) As Long

Public Type TrayIconData
    lngSize As Long
    lngHandler As Long
    lngID As Long
    lngFlags As Long
    lngMessage As Long
    lngIcon As Long
    strToolTip As String * 64
End Type

Public lngOldWindowProcess As Long
Public lngTaskBarCreated As Long
Public udtIconData As TrayIconData

Public Sub SubClass(lngHandle As Long)

    lngOldWindowProcess = SetWindowLong(lngHandle, -4, AddressOf WindowProc)

End Sub

Public Sub UnSubClass(lngHandle As Long)

    SetWindowLong lngHandle, -4, lngOldWindowProcess

End Sub

Public Function WindowProc(ByVal lngHandle As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    If lngHandle = frmMain.hWnd Then
        If uMsg = lngTaskBarCreated Then
            frmMain.Show
            With udtIconData
                .lngSize = Len(udtIconData)
                .lngIcon = frmMain.Icon
                .lngHandler = frmMain.hWnd
                .strToolTip = "Time Synchronizer" & Chr$(0)
                .lngMessage = 512
                .lngFlags = 2 Or 4 Or 1
                .lngID = vbNull
            End With
            If frmMain.WindowState = 1 Then
                Call Shell_NotifyIcon(0, udtIconData)
                frmMain.Hide
            End If
          Else
            WindowProc = CallWindowProc(lngOldWindowProcess, lngHandle, uMsg, wParam, lParam)
        End If
      Else
        WindowProc = CallWindowProc(lngOldWindowProcess, lngHandle, uMsg, wParam, lParam)
    End If

End Function
