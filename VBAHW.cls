VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stock_Data()

Dim Ticker As String
Dim Total_Vol As Double
Total_Vol = 0
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

Dim lrow As Double

lrow = Cells(Rows.Count, 1).End(xlUp).Row
MsgBox lrow

Range("I1").Value = "Ticker"
Range("J1").Value = "Total Volume"

For i = 2 To lrow

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
Ticker = Cells(i, 1).Value

Total_Vol = Total_Vol + Cells(i, 10).Value

Range("I" & Summary_Table_Row).Value = Ticker
Range("J" & Summary_Table_Row).Value = Total_Vol

Summary_Table_Row = Summary_Table_Row + 1

Total_Vol = 0

Else
Total_Vol = Total_Vol + Cells(i, 7).Value

End If

Next i

End Sub

