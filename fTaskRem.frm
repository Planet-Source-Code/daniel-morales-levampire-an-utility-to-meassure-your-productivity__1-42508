VERSION 5.00
Begin VB.Form fTaskRem 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Task Monitor"
   ClientHeight    =   2655
   ClientLeft      =   2460
   ClientTop       =   2160
   ClientWidth     =   4935
   ForeColor       =   &H00FF8080&
   Icon            =   "fTaskRem.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2655
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraMain 
      BackColor       =   &H00000000&
      Caption         =   "You Have Been Working on:"
      ForeColor       =   &H00FF8080&
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.CommandButton cClose 
         Caption         =   "Close"
         Height          =   255
         Left            =   3600
         TabIndex        =   4
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label lTimer 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   2130
         Width           =   3015
      End
      Begin VB.Label lFor 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "For:"
         ForeColor       =   &H00FF8080&
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   2160
         Width           =   270
      End
      Begin VB.Label lTask 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF8080&
         Height          =   1815
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4455
      End
   End
End
Attribute VB_Name = "fTaskRem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CenterForm Me
End Sub
