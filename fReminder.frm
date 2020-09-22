VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form fReminder 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   2880
   ClientLeft      =   3315
   ClientTop       =   4170
   ClientWidth     =   5445
   ForeColor       =   &H00FF8080&
   Icon            =   "fReminder.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2880
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraMain 
      BackColor       =   &H00000000&
      Caption         =   "Reminder From Schedule"
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
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin VB.CheckBox chkDont 
         BackColor       =   &H00000000&
         Caption         =   "Don't Remind Me"
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   3600
         TabIndex        =   6
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CommandButton cClose 
         Caption         =   "Close"
         Height          =   255
         Left            =   4320
         TabIndex        =   1
         Top             =   2280
         Width           =   975
      End
      Begin VB.Frame fraRemind 
         BackColor       =   &H00000000&
         Caption         =   "Remind Me At:"
         ForeColor       =   &H00FF8080&
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   1920
         Width           =   3375
         Begin MSComCtl2.DTPicker DTRem 
            Height          =   270
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   476
            _Version        =   393216
            Format          =   21168129
            CurrentDate     =   37633
         End
         Begin MSComCtl2.DTPicker HRRem 
            Height          =   270
            Left            =   1560
            TabIndex        =   5
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   476
            _Version        =   393216
            Format          =   21168130
            CurrentDate     =   37633
         End
      End
      Begin VB.Label lReminder 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF8080&
         Height          =   1575
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   5175
      End
   End
End
Attribute VB_Name = "fReminder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SID As Double

Private Sub cClose_Click()
Dim b As Boolean
Dim RD As Date
Dim RH As Date

    b = CInt(chkDont.Value)
    RD = DateSerial(DTRem.Year, DTRem.Month, DTRem.Day)
    RH = TimeSerial(HRRem.Hour, HRRem.Minute, HRRem.Second)
    Pr.NextRemind SID, b, RD, RH
    Unload Me
End Sub

Private Sub Form_Load()
    Height = fraMain.Height
    Width = fraMain.Width
    CenterForm Me
    'cRemindMe.ListIndex = 0
End Sub
