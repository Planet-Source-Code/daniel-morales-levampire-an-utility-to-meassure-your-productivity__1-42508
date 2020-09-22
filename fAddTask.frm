VERSION 5.00
Begin VB.Form fAddTask 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Task"
   ClientHeight    =   2040
   ClientLeft      =   4575
   ClientTop       =   2205
   ClientWidth     =   4935
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
   Icon            =   "fAddTask.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   4935
   Begin VB.Frame fraMain 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FF8080&
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.TextBox tTask 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FF8080&
         Height          =   1215
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   4455
      End
      Begin VB.CommandButton cClose 
         Caption         =   "Close"
         Height          =   255
         Left            =   3600
         TabIndex        =   2
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton cAddTaks 
         Caption         =   "Add"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   1560
         Width           =   1335
      End
   End
End
Attribute VB_Name = "fAddTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cAddTaks_Click()
    Pr.AddTask (tTask.Text)
    Unload Me
End Sub

Private Sub cClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CenterForm Me
End Sub
