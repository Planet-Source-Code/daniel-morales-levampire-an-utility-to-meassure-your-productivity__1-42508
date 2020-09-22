Attribute VB_Name = "modMain"
Option Explicit
Public db As New DBM
Public Pr As New Prod

Public Sub Main()
Dim att
    att = "Description=" & "Productivity Meassurement." & Chr$(13)
    att = att & "DBQ=" & App.Path & "\Prod.mdb"
    RegisterDatabase "Prod", "Microsoft Access Driver (*.mdb)", True, att
    db.DSN = "Prod"
    fMain.Show
End Sub

Public Sub CenterForm(frm As Form)
    With frm
        .Left = (Screen.Width - .Width) / 2
        .Top = (Screen.Height - .Height) / 2
    End With
End Sub
