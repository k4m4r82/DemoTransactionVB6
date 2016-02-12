VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Order"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' tambahkan referensi Microsoft Scripting Runtime
Option Explicit

Private mOrderID            As Long
Private mCustomerID         As String
Private mOrderDate          As String
Private mShipName           As String
Private mShipAddress        As String
Private mListOfOrderDetail  As Scripting.Dictionary

Private Sub Class_Initialize()
    Set mListOfOrderDetail = New Scripting.Dictionary
End Sub

Public Property Let orderID(ByVal vData As Long)
    mOrderID = vData
End Property
Public Property Get orderID() As Long
    orderID = mOrderID
End Property

Public Property Let customerID(ByVal vData As String)
    mCustomerID = vData
End Property
Public Property Get customerID() As String
    customerID = mCustomerID
End Property

Public Property Let orderDate(ByVal vData As String)
    mOrderDate = vData
End Property
Public Property Get orderDate() As String
    orderDate = mOrderDate
End Property

Public Property Let shipName(ByVal vData As String)
    mShipName = vData
End Property
Public Property Get shipName() As String
    shipName = mShipName
End Property

Public Property Let shipAddress(ByVal vData As String)
    mShipAddress = vData
End Property
Public Property Get shipAddress() As String
    shipAddress = mShipAddress
End Property

Public Property Let listOfOrderDetail(ByVal vData As Scripting.Dictionary)
    Set mListOfOrderDetail = vData
End Property

Public Property Get listOfOrderDetail() As Scripting.Dictionary
    Set listOfOrderDetail = mListOfOrderDetail
End Property
