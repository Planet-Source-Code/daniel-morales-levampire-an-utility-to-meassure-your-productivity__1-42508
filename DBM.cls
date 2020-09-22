VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Option Explicit
'variables locales para almacenar los valores de las propiedades
Private mvarWS As Workspace 'copia local
Private mvarCN As Connection 'copia local
Private mvarConectado As Boolean 'copia local
Private mvarUser As String 'local copy
Private mvarPass As String 'local copy
Private mvarDSN As String 'local copy
Private mvarTransactionON As Boolean 'local copy
Private QueryArray() As String 'local copy
Private mvarDB As Database
Private Abierta As Boolean

Public Property Let TransactionON(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TransactionON = 5
    mvarTransactionON = vData
End Property


Public Property Get TransactionON() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TransactionON
    TransactionON = mvarTransactionON
End Property


Public Function CommitWork() As Boolean
  Dim i, Tot As Integer
  Tot = UBound(QueryArray)
  For i = 1 To Tot
    If mvarConectado = False Or Abierta = False Then
      Debug.Print "CommitWork, Conecting to ODBC DSN=" & mvarDSN & "..."
      Abrir
    End If
    Debug.Print "EXE: " & QueryArray(i) & " ..."
    mvarCN.Execute (QueryArray(i))
  Next
  Debug.Print "CommitWork, Job Done."
  Me.TransactionON = False
End Function

Public Function BeginWork() As Boolean
  ReDim QueryArray(0)
  Me.TransactionON = True
End Function


Public Property Let DSN(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DSN = 5
    mvarDSN = vData
End Property

Public Property Get DSN() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DSN
    DSN = mvarDSN
End Property

Public Property Let Pass(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Pass = 5
    mvarPass = vData
End Property

Public Property Get Pass() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Pass
    Pass = mvarPass
End Property

Public Property Let User(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.User = 5
    mvarUser = vData
End Property

Public Property Get User() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.User
    User = mvarUser
End Property

Public Function Cerrar() As Boolean
  If Abierta Then
    mvarCN.Close
    mvarWS.Close
    'mvarCN.Close
    mvarConectado = False
    Debug.Print "Connection is closed now."
    Abierta = False
  End If
End Function

Public Function Abrir() As Boolean
  If Abierta Then
    Abrir = True
    Exit Function
  End If
  If Me.DSN = "" Then
    Me.DSN = "my"
  End If
Set mvarWS = CreateWorkspace("ODBCWorkspace", _
  "user", "", dbUseODBC)
Set mvarCN = mvarWS.OpenConnection("zyo", _
  dbDriverNoPrompt, , _
  "ODBC;DSN=" & DSN & ";UID=" & User & ";PWD=" & Pass & ";")
  mvarConectado = True
  Abrir = True
  Abierta = True
End Function

Public Property Let Conectado(ByVal vData As Boolean)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Conectado = 5
    mvarConectado = vData
End Property

Public Property Get Conectado() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Conectado
    Conectado = mvarConectado
End Property

Public Property Set cn(ByVal vData As Connection)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.CN = Form1
    Set mvarCN = vData
End Property

Public Property Get cn() As Connection
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.CN
    Set cn = mvarCN
End Property

Public Property Set WS(ByVal vData As Workspace)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.WS = Form1
    Set mvarWS = vData
End Property


Public Property Get WS() As Workspace
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.WS
    Set WS = mvarWS
End Property

Public Function GetData(Query As String) As Recordset
  Dim tmpRS As Recordset
  Dim errs As String
  Dim rs As Recordset
  Dim i As Integer
  On Error GoTo eh
  Abrir
  Set tmpRS = mvarCN.OpenRecordset(Query)
  Set rs = tmpRS
  Debug.Print "GET(" & mvarDSN & "): " & Query
  Set GetData = tmpRS
  Exit Function
eh:
  For i = 0 To DAO.Errors.Count - 1
    errs = errs & DAO.Errors(i).Description & vbCrLf
    
  Next
  If InStr(1, errs, "Got error 127 from table") > 0 Then
    MsgBox "Error en la base de datos requiere mantenimiento, contacte a su administrador de bases de datos", vbCritical, "[ERR DB 127]"
  End If
  MsgBox "Ocurrio un inconveniente: " & vbCrLf & errs & vbNewLine & "Q=[" & Query & "]", vbCritical
End Function

Public Function GetRS(Q As String, Optional Tipo As Integer) As Recordset
  'Set GetRS = CN.OpenRecordset(Q, dbOpenDynamic)
  If Not Abierta Then
    Abrir
  End If
  If Tipo <> 0 Then
    Set GetRS = mvarCN.OpenRecordset(Q, Tipo)
  Else
    Set GetRS = mvarCN.OpenRecordset(Q)
  End If
End Function


Public Function PutData(Query As String) As Boolean
Dim errs As String
Dim i As Integer
On Error GoTo eh
  If Me.TransactionON And UCase(Mid(Query, 1, 6)) <> "SELECT" Then
    ReDim Preserve QueryArray(UBound(QueryArray) + 1)
    QueryArray(UBound(QueryArray)) = Query
    Exit Function
  End If
  If mvarConectado = False Then
    Debug.Print "PutData, Conecting to ODBC DSN=" & mvarDSN & "..."
    Abrir
  End If
  Debug.Print "EXE: " & Query & ""
  mvarCN.Execute (Query)
  PutData = True
  Exit Function
eh:
  errs = ""
  For i = 0 To DAO.Errors.Count - 1
    errs = errs & vbCrLf & DAO.Errors(i).Description
  Next
  MsgBox "La peticion " & Query & " no se pudo ejecutar debido a los siguientes detalles: " & errs, vbCritical, "Revice Bien los Datos"
  PutData = False
End Function

