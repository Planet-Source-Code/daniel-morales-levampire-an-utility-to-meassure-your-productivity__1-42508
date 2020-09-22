VERSION 5.00
Begin VB.Form fTaskMan 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Task Manager"
   ClientHeight    =   3840
   ClientLeft      =   5430
   ClientTop       =   1470
   ClientWidth     =   5055
   ForeColor       =   &H00FF8080&
   Icon            =   "fTaskMan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   5055
   Begin VB.Frame fraMain 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FF8080&
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.OptionButton oDisabled 
         Caption         =   "Disabled"
         Height          =   255
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3360
         Width           =   1455
      End
      Begin VB.OptionButton oAvailable 
         Caption         =   "Available"
         Height          =   255
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3360
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.CommandButton cClose 
         Caption         =   "Close"
         Height          =   255
         Left            =   3720
         TabIndex        =   1
         Top             =   3360
         Width           =   975
      End
      Begin VB.Frame fraTasks 
         BackColor       =   &H00000000&
         Caption         =   "Available Tasks"
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
         Height          =   3015
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4575
         Begin VB.CommandButton cAddTask 
            Caption         =   "Add"
            Height          =   255
            Left            =   3000
            TabIndex        =   5
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cDisable 
            Caption         =   "Disable"
            Height          =   255
            Left            =   3000
            TabIndex        =   4
            Top             =   600
            Width           =   1335
         End
         Begin VB.ListBox lTasks 
            Height          =   2595
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   2775
         End
      End
      Begin VB.Frame fraDis 
         BackColor       =   &H00000000&
         Caption         =   "Disabled Tasks"
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
         Height          =   3015
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   4575
         Begin VB.ListBox lDis 
            Height          =   2595
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   2775
         End
         Begin VB.CommandButton cEnable 
            Caption         =   "Enable"
            Height          =   255
            Left            =   3000
            TabIndex        =   7
            Top             =   240
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "fTaskMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cAddTask_Click()
    fAddTask.Show vbModal
    Pr.RefreshTaskLists
End Sub

Private Sub cClose_Click()
    Unload Me
End Sub

Private Sub cDisable_Click()
    If lTasks.ListIndex >= 0 Then
        Pr.RemoveTask lTasks.Text
        Pr.RefreshTaskLists
    End If
End Sub

Private Sub cEnable_Click()
    If lDis.ListIndex >= 0 Then
        Pr.EnableTask lDis.Text
        Pr.RefreshTaskLists
    End If
End Sub

Private Sub Form_Load()
    CenterForm Me
    Pr.FillTasks lTasks
End Sub

Private Sub oAvailable_Click()
    fraTasks.Visible = True
    fraDis.Visible = False
    Pr.FillTasks lTasks
End Sub

Private Sub oDisabled_Click()
    fraTasks.Visible = False
    fraDis.Visible = True
    Pr.RefreshTaskLists
End Sub
