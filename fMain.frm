VERSION 5.00
Begin VB.Form fMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "My Productivity Meassurement"
   ClientHeight    =   2895
   ClientLeft      =   1785
   ClientTop       =   1695
   ClientWidth     =   6375
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF8080&
   Icon            =   "fMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   6375
   Begin VB.Frame fraMain 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FF8080&
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin VB.CommandButton cReport 
         Caption         =   "Show Report"
         Height          =   255
         Left            =   3360
         TabIndex        =   12
         Top             =   2040
         Width           =   2655
      End
      Begin VB.CommandButton cSchedule 
         Caption         =   "Schedule"
         Height          =   255
         Left            =   3360
         TabIndex        =   10
         Top             =   1800
         Width           =   2655
      End
      Begin VB.CommandButton cTaskWin 
         Caption         =   "Task Window"
         Height          =   255
         Left            =   3360
         TabIndex        =   9
         Top             =   1560
         Width           =   2655
      End
      Begin VB.CommandButton cTray 
         Caption         =   "Send to Tray"
         Height          =   255
         Left            =   3360
         TabIndex        =   8
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CommandButton cClose 
         Caption         =   "Close"
         Height          =   255
         Left            =   5040
         TabIndex        =   7
         Top             =   2400
         Width           =   975
      End
      Begin VB.Frame fraChooseTask 
         BackColor       =   &H00000000&
         Caption         =   "Choose Task"
         ForeColor       =   &H00FF8080&
         Height          =   1380
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   3135
         Begin VB.ListBox lTasks 
            BackColor       =   &H00000000&
            ForeColor       =   &H00FF8080&
            Height          =   1035
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   2895
         End
      End
      Begin VB.Frame fraCurrentTask 
         BackColor       =   &H00000000&
         Caption         =   "Current Task"
         ForeColor       =   &H00FF8080&
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   5895
         Begin VB.Label lCurTask 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF8080&
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   3735
         End
         Begin VB.Label lTimer 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF8080&
            Height          =   255
            Left            =   3960
            TabIndex        =   3
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Timer TimerHour 
         Interval        =   500
         Left            =   2640
         Top             =   2280
      End
      Begin VB.CommandButton cEndTask 
         Caption         =   "End Task"
         Height          =   255
         Left            =   3360
         TabIndex        =   13
         Top             =   1320
         Width           =   2655
      End
      Begin VB.PictureBox Cell 
         Height          =   495
         Left            =   3360
         Picture         =   "fMain.frx":0CCA
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   11
         Top             =   960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lHour 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   2400
         Width           =   2415
      End
   End
   Begin VB.Menu mnuTray 
      Caption         =   "Tray Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuReport 
         Caption         =   "Show Report"
      End
      Begin VB.Menu mnuSch 
         Caption         =   "Schedule"
      End
      Begin VB.Menu mnuTask 
         Caption         =   "Task Window"
      End
      Begin VB.Menu mnuSpa1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
    Const WM_LBUTTONDBLCLICK = &H203
    Const WM_RBUTTONUP = &H205
    Const WM_MOUSEMOVE = &H200
    Const NIM_ADD = &H0
    Const NIM_MODIFY = &H1
    Const NIM_DELETE = &H2
    Const NIF_MESSAGE = &H1
    Const NIF_ICON = &H2
    Const NIF_TIP = &H4
    Const GWL_WNDPRC = -4
    
    Dim nid As NOTIFYICONDATA


Private Sub cClose_Click()
    Unload Me
End Sub

Private Sub Cell_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
        Static lngMsg As Long
        Static blnFlag As Boolean
        Dim result As Long
        
        lngMsg = X / Screen.TwipsPerPixelX
        If blnFlag = False Then
        blnFlag = True
        Select Case lngMsg
        '     'doubleclick
        Case WM_LBUTTONDBLCLICK
        Me.Show
        Me.WindowState = vbNormal
        fMain.WindowState = 0
        Me.Show
        
        '     'right-click
        Case WM_RBUTTONUP
        result = SetForegroundWindow(Me.hwnd)
        Me.PopupMenu mnuTray
End Select

        blnFlag = False
End If
End Sub

Private Sub cEndTask_Click()
    Pr.EndTask
End Sub

Private Sub cReport_Click()
    Pr.MakeReport Day(Date), Month(Date), Year(Date)
End Sub

Private Sub cSchedule_Click()
    fSchedule.Show
End Sub

Private Sub cTaskWin_Click()
    fTaskMan.Show
End Sub

Private Sub cTray_Click()
Dim frm As Form
    For Each frm In Forms
        If frm.Name <> "fMain" Then
            Unload frm
        End If
    Next
    Me.Hide
End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then
    End
End If
    CenterForm Me
    nid.cbSize = Len(nid)
    nid.hwnd = Cell.hwnd
    nid.uID = 1&
    nid.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
    nid.uCallbackMessage = WM_MOUSEMOVE
    nid.hIcon = Cell.Picture
    nid.szTip = "Task Monitor" + Chr$(0)
    
    Shell_NotifyIcon NIM_ADD, nid
    Pr.FillTasks lTasks
End Sub

Private Sub lTasks_DblClick()
    If lTasks.ListIndex >= 0 Then
        Pr.StartTask (lTasks.Text)
    End If
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuOpen_Click()
    Me.Show
    Me.WindowState = vbNormal
    fMain.WindowState = 0
    Me.Show
End Sub

Private Sub mnuReport_Click()
    Pr.MakeReport Day(Date), Month(Date), Year(Date)
End Sub

Private Sub mnuSch_Click()
    fSchedule.Show
End Sub

Private Sub mnuTask_Click()
    fTaskMan.Show
End Sub

Private Sub TimerHour_Timer()
    lHour.Caption = Now
    Pr.ChkSch
    Pr.ChkTask
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    VBGTray.cbSize = Len(VBGTray)
    VBGTray.hwnd = Me.hwnd
    VBGTray.uID = vbNull
    Call Shell_NotifyIcon(NIM_DELETE, VBGTray)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Shell_NotifyIcon NIM_DELETE, nid
    If Pr.TaskRunning Then Pr.EndTask
End Sub


