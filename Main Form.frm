VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Time Synchronizer"
   ClientHeight    =   1065
   ClientLeft      =   4785
   ClientTop       =   2595
   ClientWidth     =   2400
   Icon            =   "Main Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   71
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   160
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMinimize 
      Caption         =   "    _"
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      ToolTipText     =   "Send To Tray"
      Top             =   480
      Width           =   495
   End
   Begin VB.Timer tmrTimer 
      Interval        =   60000
      Left            =   120
      Top             =   1200
   End
   Begin MSWinsockLib.Winsock WinSock 
      Left            =   1320
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   37
   End
   Begin VB.CommandButton cmdSynchronizeNow 
      Caption         =   "&Synchronize Now"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Synchronize Now"
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Status"
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type SystemTime
    intYear As Integer
    intMonth As Integer
    intWeekDay As Integer
    intDay As Integer
    intHour As Integer
    intMinute As Integer
    intSecond As Integer
    intMillisecond As Integer
End Type

Private Declare Function SetSystemTime Lib "kernel32" (lpSystemTime As SystemTime) As Long

Private intCounter As Integer
Private intCloseMethod As Integer
Private sngDelay As Single
Private strTime As String

Private Sub cmdMinimize_Click()

    WindowState = 1

End Sub

Private Sub cmdSynchronizeNow_Click()

  Dim intFileNumber As Integer
  Dim strServer As String

    intFileNumber = FreeFile
    strServer = "ntps1-0.cs.tu-berlin.de"
    On Error Resume Next
        If Right$(App.Path, 1) = "\" Then
            Open App.Path & "settings.txt" For Input Lock Write As #intFileNumber
            Line Input #intFileNumber, strServer
            Close #intFileNumber
          Else
            Open App.Path & "\settings.txt" For Input Lock Write As #intFileNumber
            Line Input #intFileNumber, strServer
            Close #intFileNumber
        End If
    On Error GoTo 0
    cmdSynchronizeNow.Enabled = False
    lblStatus.Caption = "Synchronizing..."
    WinSock.Close
    strTime = Empty
    WinSock.RemoteHost = strServer
    WinSock.Connect

End Sub

Private Sub Form_Load()

    If App.PrevInstance Then
        intCloseMethod = 1
        Unload Me
    End If
    lngTaskBarCreated = RegisterWindowMessage("TaskbarCreated")
    SubClass frmMain.hWnd
    With udtIconData
        .lngSize = Len(udtIconData)
        .lngIcon = Icon
        .lngHandler = hWnd
        .strToolTip = "Time Synchronizer" & Chr$(0)
        .lngMessage = 512
        .lngFlags = 2 Or 4 Or 1
        .lngID = vbNull
    End With
    cmdSynchronizeNow_Click

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  Dim lngMessage As Long

    lngMessage = X
    If lngMessage = 515 Then
        WindowState = vbNormal
        Shell_NotifyIcon 2, udtIconData
        Show
      ElseIf lngMessage = 518 Then
        Unload Me
    End If

End Sub

Private Sub Form_Resize()

    If WindowState = 1 Then
        Call Shell_NotifyIcon(0, udtIconData)
        Hide
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If intCloseMethod = 0 Then
        UnSubClass frmMain.hWnd
        Shell_NotifyIcon 2, udtIconData
    End If
    Set frmMain = Nothing

End Sub

Private Sub SynchronizeClock(strTimeString As String)

  Dim datTime As Date
  Dim dblTime As Double
  Dim lngTime As Long
  Dim udtSystemDate As SystemTime

    strTimeString = Trim$(strTimeString)
    If Len(strTimeString) <> 4 Then
        cmdSynchronizeNow.Enabled = True
        lblStatus.Caption = "Error synchronizing!"
        Exit Sub
    End If
    dblTime = Asc(Left$(strTimeString, 1)) * 256 ^ 3 + Asc(Mid$(strTimeString, 2, 1)) * 256 ^ 2 + Asc(Mid$(strTimeString, 3, 1)) * 256 ^ 1 + Asc(Right$(strTimeString, 1))
    lngTime = dblTime - 2840140800#
    datTime = DateAdd("s", CDbl(lngTime + CLng(sngDelay)), #1/1/1990#)
    udtSystemDate.intYear = Year(datTime)
    udtSystemDate.intMonth = Month(datTime)
    udtSystemDate.intDay = Day(datTime)
    udtSystemDate.intHour = Hour(datTime)
    udtSystemDate.intMinute = Minute(datTime)
    udtSystemDate.intSecond = Second(datTime)
    Call SetSystemTime(udtSystemDate)
    cmdSynchronizeNow.Enabled = True
    lblStatus.Caption = "Last update: " & Time

End Sub

Private Sub tmrTimer_Timer()

    intCounter = intCounter + 1
    If intCounter = 60 Then
        intCounter = 0
        cmdSynchronizeNow_Click
    End If

End Sub

Private Sub WinSock_Close()

    On Error Resume Next
        Do Until WinSock.State = sckClosed
            WinSock.Close
            DoEvents
        Loop
        sngDelay = ((Timer - sngDelay) / 2)
        Call SynchronizeClock(strTime)
    On Error GoTo 0

End Sub

Private Sub WinSock_Connect()

    sngDelay = Timer

End Sub

Private Sub WinSock_DataArrival(ByVal bytesTotal As Long)

  Dim strData As String

    WinSock.GetData strData, vbString
    strTime = strTime & strData

End Sub
