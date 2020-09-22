VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fReport 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daily Report"
   ClientHeight    =   7815
   ClientLeft      =   2100
   ClientTop       =   525
   ClientWidth     =   9735
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
   Icon            =   "fReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   9735
   Begin VB.Frame fraMain 
      BackColor       =   &H00404040&
      Caption         =   "Report For: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   7695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9495
      Begin VB.CommandButton cSave 
         Caption         =   "Save Report"
         Height          =   255
         Left            =   6480
         TabIndex        =   3
         Top             =   7320
         Width           =   1695
      End
      Begin MSComDlg.CommonDialog CD 
         Left            =   120
         Top             =   7080
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox tReport 
         BackColor       =   &H00404040&
         ForeColor       =   &H00FF80FF&
         Height          =   6735
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   360
         Width           =   9015
      End
      Begin VB.CommandButton cClose 
         Caption         =   "Close"
         Height          =   255
         Left            =   8280
         TabIndex        =   1
         Top             =   7320
         Width           =   975
      End
   End
End
Attribute VB_Name = "fReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cClose_Click()
    Unload Me
End Sub

Private Sub cSave_Click()
Dim iFileNum As Integer
Dim sFileName As String

    On Error GoTo err
    CD.CancelError = True
    CD.InitDir = App.Path
    CD.ShowSave
    iFileNum = FreeFile
    sFileName = CD.FileName
    Open sFileName For Output As #iFileNum
    Print #iFileNum, tReport
    Close #iFileNum
err:
    err.Clear
End Sub

Private Sub Form_Load()
    CenterForm Me
End Sub
