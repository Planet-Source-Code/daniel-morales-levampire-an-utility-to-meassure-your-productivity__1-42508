VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form fAddReminder 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Reminder"
   ClientHeight    =   4230
   ClientLeft      =   2100
   ClientTop       =   2475
   ClientWidth     =   5775
   ForeColor       =   &H00FF8080&
   Icon            =   "fAddReminder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   5775
   Begin VB.Frame fraMain 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FF8080&
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      Begin VB.Frame fraRemind 
         BackColor       =   &H00000000&
         Caption         =   "Remind Me At:"
         ForeColor       =   &H00FF8080&
         Height          =   615
         Left            =   120
         TabIndex        =   6
         Top             =   3000
         Width           =   5295
         Begin MSComCtl2.DTPicker DTRem 
            Height          =   270
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   476
            _Version        =   393216
            Format          =   21102593
            CurrentDate     =   37633
         End
         Begin MSComCtl2.DTPicker HRRem 
            Height          =   270
            Left            =   3480
            TabIndex        =   8
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   476
            _Version        =   393216
            Format          =   21102594
            CurrentDate     =   37633
         End
      End
      Begin VB.TextBox tDesc 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FF8080&
         Height          =   2415
         Left            =   120
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   600
         Width           =   5295
      End
      Begin VB.CommandButton cCancel 
         Caption         =   "Cancel"
         Height          =   255
         Left            =   4320
         TabIndex        =   4
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton cAddReminder 
         Caption         =   "Add Reminder"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   3720
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPick 
         Height          =   270
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   476
         _Version        =   393216
         Format          =   21102593
         CurrentDate     =   37633
      End
      Begin MSComCtl2.DTPicker HRPick 
         Height          =   270
         Left            =   3720
         TabIndex        =   3
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   476
         _Version        =   393216
         Format          =   21102594
         CurrentDate     =   37633
      End
   End
End
Attribute VB_Name = "fAddReminder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cAddReminder_Click()
Dim RD As Date
Dim DD As Date
Dim RH As Date
Dim DH As Date

    DD = DateSerial(DTPick.Year, DTPick.Month, DTPick.Day)
    RD = DateSerial(DTRem.Year, DTRem.Month, DTRem.Day)
    DH = TimeSerial(HRPick.Hour, HRPick.Minute, HRPick.Second)
    RH = TimeSerial(HRRem.Hour, HRRem.Minute, HRRem.Second)
    Pr.AddReminder tDesc, DD, DH, RD, RH
    MsgBox "Reminder Added!", vbInformation, "Schedule"
    Unload Me
End Sub

Private Sub cCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CenterForm Me
    DTPick.Day = Day(Date)
    DTPick.Month = Month(Date)
    DTPick.Year = Year(Date)
    DTRem.Day = Day(Date)
    DTRem.Month = Month(Date)
    DTRem.Year = Year(Date)
    
    HRPick.Hour = Hour(Time)
    HRPick.Minute = Minute(Time)
    HRPick.Second = 0
    HRRem.Hour = Hour(Time)
    HRRem.Minute = Minute(Time)
    HRRem.Second = 0
End Sub
