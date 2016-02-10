VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mOrderID        As Long
Private mProductID      As Long
Private mUnitPrice      As Double
Private mQuantity       As Integer
Private mDiscount       As Single

Public Property Let orderID(ByVal vData As Long)
    mOrderID = vData
End Property
Public Property Get orderID() As Long
    orderID = mOrderID
End Property

Public Property Let productID(ByVal vData As Long)
    mProductID = vData
End Property
Public Property Get productID() As Long
    productID = mProductID
End Property

Public Property Let unitPrice(ByVal vData As Double)
    mUnitPrice = vData
End Property
Public Property Get unitPrice() As Double
    unitPrice = mUnitPrice
End Property

Public Property Let quantity(ByVal vData As Integer)
    mQuantity = vData
End Property
Public Property Get quantity() As Integer
    quantity = mQuantity
End Property

Public Property Let discount(ByVal vData As Single)
    mDiscount = vData
End Property
Public Property Get discount() As Single
    discount = mDiscount
End Property
