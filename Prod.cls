VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Prod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'local variable(s) to hold property value(s)
Private mvarTaskRunning As Boolean 'local copy
Private mvarIdTaskRunning As Double 'local copy
Private mvarTaskTimer As Single 'local copy
Private mvarTaskMonitor As Single 'local copy

Public Property Let IdTaskRunning(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IdTaskRunning = 5
    mvarIdTaskRunning = vData
End Property


Public Property Get IdTaskRunning() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IdTaskRunning
    IdTaskRunning = mvarIdTaskRunning
End Property

Public Property Let TaskTimer(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IdTaskRunning = 5
    mvarTaskTimer = vData
End Property


Public Property Get TaskTimer() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IdTaskRunning
    TaskTimer = mvarTaskTimer
End Property

Public Property Let TaskMonitor(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IdTaskRunning = 5
    mvarTaskMonitor = vData
End Property


Public Property Get TaskMonitor() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IdTaskRunning
    TaskMonitor = mvarTaskMonitor
End Property

Public Property Let TaskRunning(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TaskRunning = 5
    mvarTaskRunning = vData
End Property


Public Property Get TaskRunning() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TaskRunning
    TaskRunning = mvarTaskRunning
End Property




Public Sub FillTasks(lList As ListBox, Optional bDis As Boolean)
Dim rs As Recordset

    lList.Clear
    If bDis Then
        Set rs = db.GetData("Select * from Tasks where id > 0 and Available = 0")
    Else
        Set rs = db.GetData("Select * from Tasks where id > 0 and Available = 1")
    End If
    If rs.EOF = False Then
        Do While Not rs.EOF
            lList.AddItem rs(1).Value
            rs.MoveNext
        Loop
    End If
End Sub

Public Sub FillSch(lList As ListBox, iDay As Integer, iMonth As Integer, iYear As Integer)
Dim rs As Recordset
Dim dDate As Date
    Set rs = db.GetData("Select * from Schedule where id > 0")
    lList.Clear
    Do While Not rs.EOF
        dDate = CDate(rs(3).Value)
        If Day(dDate) = iDay And Month(dDate) = iMonth And Year(dDate) = iYear Then
            lList.AddItem Left(rs(1).Value, 20) & "..." & vbTab & dDate & " " & rs(4).Value
        End If
        rs.MoveNext
    Loop
End Sub

Public Sub AddTask(sTaskDesc As String)
Dim id As Double
Dim rs As Recordset
    Set rs = db.GetData("Select max(id) from Tasks")
    id = rs(0).Value + 1
    db.PutData ("Insert Into Tasks (id,Task,Available) values(" & id & ",'" & sTaskDesc & "',1)")
    MsgBox "Task Added!", vbInformation, "Task Monitor"
End Sub

Public Sub RemoveTask(sTaskDesc As String)
    db.PutData ("Update Tasks Set Available=0 where Task='" & sTaskDesc & "'")
    MsgBox "Task Disabled!", vbInformation, "Task Monitor"
End Sub

Public Sub EnableTask(sTaskDesc As String)
    db.PutData ("Update Tasks Set Available=1 where Task='" & sTaskDesc & "'")
    MsgBox "Task Enabled!", vbInformation, "Task Monitor"
End Sub

Public Sub RefreshTaskLists()
    FillTasks fMain.lTasks
    FillTasks fTaskMan.lTasks
    FillTasks fTaskMan.lDis, True
End Sub

Public Sub StartTask(sTaskDesc As String)
Dim id As Double
Dim IdTask As Double
Dim dDate As Date
Dim STime As Date
Dim rs As Recordset
    
    If TaskRunning Then
        EndTask
    End If
    Set rs = db.GetData("Select id From Tasks Where Task = '" & sTaskDesc & "'")
    IdTask = rs(0).Value
    Set rs = db.GetData("Select max(id) From Productivity")
    id = rs(0).Value + 1
    dDate = Date
    STime = Time
    IdTaskRunning = id
    TaskTimer = Timer
    TaskRunning = True
    db.PutData ("Insert Into Productivity (id,id_Task,n_Date,Start_Time) " & _
                "Values (" & _
                id & "," & _
                IdTask & ",'" & _
                CStr(dDate) & "','" & _
                CStr(STime) & "')")
    MsgBox "Task Started", vbInformation, "Task Monitor"
End Sub

Public Sub EndTask()
    db.PutData ("Update Productivity Set End_Time = '" & Time & _
                "' where id = " & IdTaskRunning)
    IdTaskRunning = 0
    TaskRunning = False
    TaskTimer = 0
    MsgBox "Task Ended!", vbInformation, "Task Monitor"
End Sub

Public Sub AddReminder(sDesc As String, due_Date As Date, due_Hour As Date, Rem_Date As Date, Rem_Hour As Date)
Dim id As Double
Dim rs As Recordset

    Set rs = db.GetData("Select max(id) from Schedule")
    id = rs(0).Value + 1
    db.PutData ("Insert Into Schedule (id,Description,Dont_Remind,Due_Date,Due_Hour,Next_Rem_Date,Next_Rem_Hour) Values (" & _
                    id & ",'" & _
               sDesc & "',0,'" & _
               due_Date & "','" & _
               due_Hour & "','" & _
               Rem_Date & "','" & _
               Rem_Hour & "')")
End Sub

Public Sub NextRemind(id As Double, DontRemind As Boolean, Optional Rem_Date As Date, Optional Rem_Hour As Date)
    If DontRemind Then
        db.PutData ("Update Schedule Set Dont_Remind = 1 where id = " & id)
    Else
        db.PutData ("Update Schedule Set Dont_Remind = 0 where id = " & id)
        db.PutData ("Update Schedule Set Next_Rem_Date = '" & Rem_Date & "' where id = " & id)
        db.PutData ("Update Schedule Set Next_Rem_Hour = '" & Rem_Hour & "' where id = " & id)
    End If
End Sub

Public Sub ChkSch()
Dim rs As Recordset
Dim nDate As Date
Dim nTime As Date

    nDate = Date
    nTime = TimeSerial(Hour(Time), Minute(Time), 0)
    Set rs = db.GetData("Select * from Schedule Where Next_Rem_Date = '" & nDate & _
                        "' and Next_Rem_Hour = '" & nTime & "'")
    Do While Not rs.EOF
        If rs(2).Value = 0 Then
            LoadSchMsg rs(0).Value, _
            "Reminder From Schedule Due: " & rs(3).Value & " " & rs(4).Value, _
            rs(1).Value
            rs.MoveNext
        End If
    Loop
End Sub

Public Sub getSchData(iDay As Integer, iMonth As Integer, iYear As Integer, xId As Double)
Dim rs As Recordset
Dim nDate As Date
Dim i As Double
    nDate = Date
    nDate = DateSerial(iYear, iMonth, iDay)
    Set rs = db.GetData("Select * from Schedule Where Next_Rem_Date = '" & nDate & "'")
    For i = 0 To xId - 1
        rs.MoveNext
    Next i
    LoadSchMsg rs(0).Value, _
    "Reminder From Schedule Due: " & rs(3).Value & " " & rs(4).Value, _
    rs(1).Value
End Sub

Public Sub LoadSchMsg(id As Double, Main_Caption As String, sRem As String)
    Load fReminder
    With fReminder
        .SID = id
        .fraMain.Caption = Main_Caption
        .lReminder.Caption = sRem
        .Show
        .ZOrder
    End With
End Sub

Public Sub ChkTask()
Dim difTimer As Single
Dim dIntPart As Double
Dim dPointpos As Double
Dim iHr As Integer
Dim iMn As Integer
Dim iSc As Double
Dim rs As Recordset
Dim tId As Double

    On Error GoTo err
    If TaskRunning Then
        Set rs = db.GetData("Select * from Productivity where id=" & IdTaskRunning)
        If rs.EOF = False Then
            tId = rs(1).Value
        End If
        Set rs = db.GetData("Select Task from Tasks where id=" & tId)
        difTimer = Timer - TaskTimer
        dPointpos = InStr(1, CStr(difTimer), ".")
        dIntPart = CDbl(Left(CStr(difTimer), dPointpos))
        iSc = dIntPart
        Do While iSc >= 60
            iMn = iMn + 1
            If iMn = 60 Then
                iMn = 0
                iHr = iHr + 1
            End If
            iSc = iSc - 60
        Loop
        'The next 15 lines make it so that when every 30 seconds pass, TM will tell you how
        'much time you have been working on a given Task. This is only for testing purposes,
        'so to make it every 30 minutes make the following changes to the following lines.
        'Change Line 1 To: If iMn = 30 Then
        'Change Line 8 To: ElseIf iMn = 0 and IHr > 0 Then
        'Now Compile and Enjoy !! :)
        If iSc = 30 Then '1
        Load fTaskRem '2
            With fTaskRem '3
                .lTask.Caption = rs(0).Value '4
                .lTimer.Caption = Format(iHr, "##00") & ":" & Format(iMn, "00") & ":" & Format(iSc, "00") '5
                .Show '6
            End With '7
        ElseIf iSc = 0 And iMn > 0 Then '8
        Load fTaskRem '9
            With fTaskRem '10
                .lTask.Caption = rs(0).Value '11
                .lTimer.Caption = Format(iHr, "##00") & ":" & Format(iMn, "00") & ":" & Format(iSc, "00") '12
                .Show '13
            End With '14
        End If '15
        With fMain
            .lCurTask.Caption = rs(0).Value
            .lTimer.Caption = Format(iHr, "##00") & ":" & Format(iMn, "00") & ":" & Format(iSc, "00")
        End With
    Else
        fMain.lCurTask.Caption = ""
        fMain.lTimer.Caption = ""
    End If
err:
    err.Clear
End Sub

Public Sub MakeReport(iDay As Integer, iMonth As Integer, iYear As Integer)
Dim rs As Recordset
Dim rs2 As Recordset
Dim dDate As Date
Dim sTask As String
Dim sTot As String
Dim STime As Date
Dim ETime As Date
Dim iSc As Double
Dim iMn As Integer
Dim iHr As Double

    On Error GoTo err
    
    dDate = DateSerial(iYear, iMonth, iDay)
    Set rs = db.GetData("Select * from Productivity where n_Date='" & dDate & "'")
    sTot = ""
    Do While Not rs.EOF
        sTask = ""
        STime = TimeValue(rs(3).Value)
        ETime = TimeValue(rs(4).Value)
        iSc = DateDiff("s", STime, ETime)
        Do While iSc >= 60
            iMn = iMn + 1
            If iMn = 60 Then
                iMn = 0
                iHr = iHr + 1
            End If
            iSc = iSc - 60
        Loop
        Set rs2 = db.GetData("Select * from Tasks where id = " & rs(1).Value)
        sTask = "[TASK: " & rs2(1).Value & "]"
        sTask = sTask & vbTab & "[TIME WORKING: " & Format(iHr, "##00") & ":" & Format(iMn, "00") & ":" & Format(iSc, "00") & "]"
        sTask = sTask & vbTab & "[FROM: " & STime & " TO: " & ETime & "]"
        sTot = sTot & sTask & vbCrLf
        rs.MoveNext
    Loop
    Load fReport
    With fReport
        .fraMain.Caption = "Report For: " & dDate
        .tReport.Text = sTot
        .Show
        .ZOrder
    End With
err:
    err.Clear
End Sub
