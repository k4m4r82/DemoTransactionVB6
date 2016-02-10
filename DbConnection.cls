VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DbConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function GetOpenConnection() As ADODB.Connection
    Dim conn    As ADODB.Connection
    Dim strCon  As String
            
    On Error GoTo errHandle
        
    strCon = "DRIVER=SQLite3 ODBC Driver;Database=" & App.Path & "\db\Northwind.db;LongNames=0;Timeout=1000;NoTXN=0;SyncPragma=NORMAL;StepAPI=0"
    
    Set conn = New ADODB.Connection
    conn.connectionString = strCon
    conn.Open
    
    Set GetOpenConnection = conn
    
    Exit Function
errHandle:
    ' do nothing
End Function

Public Sub CloseConnection(ByVal conn As ADODB.Connection)
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then conn.Close
        Set conn = Nothing
    End If
End Sub
