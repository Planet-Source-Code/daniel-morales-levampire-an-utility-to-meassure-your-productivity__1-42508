VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form fSchedule 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Schedule"
   ClientHeight    =   4230
   ClientLeft      =   1440
   ClientTop       =   2010
   ClientWidth     =   5775
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
   Icon            =   "fSchedule.frx":0000
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
      TabIndex        =   4
      Top             =   0
      Width           =   5535
      Begin VB.CommandButton cClose 
         Caption         =   "Close"
         Height          =   255
         Left            =   4440
         TabIndex        =   3
         Top             =   3720
         Width           =   975
      End
      Begin MSComCtl2.DTPicker DTPick 
         Height          =   270
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   476
         _Version        =   393216
         Format          =   21168129
         CurrentDate     =   37633
      End
      Begin VB.ListBox lSchedule 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FF8080&
         Height          =   2985
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   5295
      End
      Begin VB.CommandButton cAddReminder 
         Caption         =   "Add Reminder"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   3720
         Width           =   1335
      End
   End
End
Attribute VB_Name = "fSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cAddReminder_Click()
    fAddReminder.Show
End Sub

Private Sub cClose_Click()
    Unload Me
End Sub

Private Sub DTPick_Change()
    Pr.FillSch lSchedule, DTPick.Day, DTPick.Month, DTPick.Year
End Sub

Private Sub Form_Load()
    CenterForm Me
    DTPick.Day = Day(Date)
    DTPick.Month = Month(Date)
    DTPick.Year = Year(Date)
    Pr.FillSch lSchedule, DTPick.Day, DTPick.Month, DTPick.Year
End Sub

Private Sub lSchedule_DblClick()
    If lSchedule.ListIndex >= 0 Then
        Pr.getSchData DTPick.Day, DTPick.Month, DTPick.Year, lSchedule.ListIndex
    End If
End Sub
