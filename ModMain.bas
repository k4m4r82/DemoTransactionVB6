Attribute VB_Name = "ModMain"
Option Explicit

Public Sub Main()
    ' buat objek Order
    Dim objOrder As New Order
    
    With objOrder
        .orderID = 123451
        .customerID = "SUPRD"
        .orderDate = "2016-2-6"
        .shipName = "Que Delícia"
        .shipAddress = "Rua da Panificadora, 12"
    End With
    
    ' buat objek OrderDetail
    ' objek #1
    Dim objItemOrder1 As New OrderDetail
    With objItemOrder1
        .productID = 11
        .unitPrice = 9.8
        .quantity = 10
        .discount = 0
    End With
    
    ' objek #2
    Dim objItemOrder2 As New OrderDetail
    With objItemOrder2
        .productID = 65
        .unitPrice = 16.8
        .quantity = 5
        .discount = 0.15
    End With
    
    ' objek #3
    Dim objItemOrder3 As New OrderDetail
    With objItemOrder3
        .productID = 212
        .unitPrice = 42.4
        .quantity = 15
        .discount = 0.25
    End With
    
    ' hubungkan objek item order ke objek order
    objOrder.listOfOrderDetail = New Scripting.Dictionary
    objOrder.listOfOrderDetail.Add 1, objItemOrder1
    objOrder.listOfOrderDetail.Add 2, objItemOrder2
    objOrder.listOfOrderDetail.Add 3, objItemOrder3
    
    Dim dao     As New OrderDao
    Dim result  As Integer
    
    result = dao.Save(objOrder)
    If result > 0 Then
        Debug.Print "Penyimpanan data order berhasil"
    Else
        Debug.Print "Penyimpanan data order gagal"
    End If
End Sub
