VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormInsert 
   Caption         =   "自動寫入"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "UserFormInsert.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "UserFormInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BtnInsert_Click()

Dim supplyname As String
supplyname = TxBname.Text
Cells(2, 1).Value = supplyname

Dim supplyphone As String
supplyphone = TxBphone.Text
Cells(2, 2).Value = supplyphone

Dim price As Integer
price = TxBprice.Text
Cells(2, 3).Value = CInt(price)

Dim price2 As Integer
price2 = TxBprice2.Text
Cells(2, 4).Value = CInt(price2)

Dim totaldiscount As Single
totaldiscount = (price - price2) / price
Cells(2, 5).Value = totaldiscount

If (totaldiscount > 0.8) Then
    Cells(2, 6).Value = "異常"
    Else
    Cells(2, 6).Value = "正常"
End If

End Sub
