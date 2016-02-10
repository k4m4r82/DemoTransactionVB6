VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderDao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private dbConn  As DbConnection
Private conn    As ADODB.Connection
Private strSql  As String

Private Sub Class_Initialize()
    Set dbConn = New DbConnection
    Set conn = dbConn.GetOpenConnection
End Sub

Public Function Save(ByVal objOrder As Order) As Integer
    Dim cmd             As ADODB.Command
    Dim key             As Variant
    Dim objItemOrder    As OrderDetail
    
    On Error GoTo errHandle

    conn.BeginTrans ' begin transaction
    
    ' insert header
    strSql = "INSERT INTO Orders (OrderID, CustomerID, OrderDate, ShipName, ShipAddress) " & _
             "VALUES (?, ?, ?, ?, ?)"
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = conn
        .CommandType = adCmdText
        .CommandText = strSql
        .Prepared = True
        
        .Parameters(0) = objOrder.orderID
        .Parameters(1) = objOrder.customerID
        .Parameters(2) = objOrder.orderDate
        .Parameters(3) = objOrder.shipName
        .Parameters(4) = objOrder.shipAddress

        .Execute
    End With
    Set cmd = Nothing

    ' insert detail
    strSql = "INSERT INTO OrderDetails (OrderID, ProductID, UnitPrice, Quantity, Discount) " & _
             "VALUES (?, ?, ?, ?, ?)"

    For Each key In objOrder.listOfOrderDetail
        
        ' ekstrak objek item order
        Set objItemOrder = objOrder.listOfOrderDetail.Item(key)

        Set cmd = New ADODB.Command
        With cmd
            .ActiveConnection = conn
            .CommandType = adCmdText
            .CommandText = strSql
            .Prepared = True
            
            .Parameters(0) = objOrder.orderID
            .Parameters(1) = objItemOrder.productID
            .Parameters(2) = objItemOrder.unitPrice
            .Parameters(3) = objItemOrder.quantity
            .Parameters(4) = objItemOrder.discount
            
            .Execute
        End With
        Set cmd = Nothing
    Next

    conn.CommitTrans ' commit transaction
    
    Save = 1
    
    Exit Function

errHandle:
    conn.RollbackTrans ' error, rollback transaction
    Save = 0
End Function

Private Sub Class_Terminate()
    Call dbConn.CloseConnection(conn)
End Sub
