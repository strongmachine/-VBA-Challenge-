VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stock_Analysis():

    
    Dim Numberofrows As Long
    Dim summaryIndex As Long
    Dim ws As Worksheet
    Dim TickerStockVolume As Double
    
    
    
    
  
   
   
   For Each ws In Worksheets
   
         'ws.Columns("I:L").Clear
         
        ws.Range("I1").Value = "Ticker"
        ws.Range("L1").Value = "Stock Volume"
        
   
  
        Numberofrows = ws.Cells(Rows.Count, "A").End(xlUp).Row
        summaryIndex = 2
        TickerStockVolume = 0
        
        
        For Index = 2 To Numberofrows
            TickerStockVolume = TickerStockVolume + ws.Cells(Index, 7).Value
        
             If ws.Cells(Index, 1).Value <> ws.Cells(Index + 1, 1).Value Then
                 ws.Cells(summaryIndex, 9).Value = ws.Cells(Index, 1).Value
                 ws.Cells(summaryIndex, 12).Value = TickerStockVolume
                 summaryIndex = summaryIndex + 1
                 TickerStockVolume = 0
                 
             
                 
             End If
             
             
        Next Index
        
        Next ws
   
   
   
End Sub
