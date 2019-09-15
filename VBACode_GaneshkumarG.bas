Attribute VB_Name = "Module1"
'Variable to hold the Start and end column of data
Public gColDataFirst, gColDataLast As String
Public gColTkr, gColDate, gColOpen, gColClose, gColVol As String
Public gColSmryTkr, gColSmryYrlChng, gColSmryPercChng, gColSmryTotVol As String
Public gColTopSmryName, gColTopSmryTkr, gColTopSmryValue As String

'Variable to hold the Start and end row of data
Public gRowDataFirst, gRowDataLast As Long

Sub LoopWorksheet()
'****************************************************************************************
'*  Functionality: To work with each of the available worksheet, in the current book    *
'****************************************************************************************

    'Create a worksheet variable
    Dim ws As Worksheet
    
'**Task1: Create a script that will loop through one year of stock data for each run**
    'Loop through each of the worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        
        'enable the worksheet to start editing.
        ws.Activate
        
        'Display the name of worksheet that is active.
        'MsgBox (ws.Name)
        
        'Define Data Range
        Call DataRange(ws)
        
        'Sorting the data, so to minimize the complexity when computing the total
        Call SortData(ws)
        
        'Compute Statistics
        Call CompStatistics(ws)
        
        'Ext for loop: Stop after working with single worksheet.
        'Exit For
    Next
    
    MsgBox ("Complete")
'End of LoopWorksheet sub
End Sub

Sub DataRange(ws As Worksheet)
    
    gColDataFirst = "A"
    gColDataLast = "H"
    gRowDataFirst = 1
    
    'Data Column
    gColTkr = "A"   'Ticker
    gColDate = "B"  'Date
    gColOpen = "C"  'Open
    gColHigh = "D"  'High
    gColLow = "E"   'Low
    gColClose = "F" 'Close
    gColVol = "G"   'Vol
    
    'Column where the Summary is Stored
    gColSmryTkr = "I"       'Ticker Symbol
    gColSmryYrlChng = "J"   'Yearly Change
    gColSmryPercChng = "K"  'Percentage Change
    gColSmryTotVol = "L"       'Total Volume
    
    'Column where the Top Summary Stored
    gColTopSmryName = "N"
    gColTopSmryTkr = "O"
    gColTopSmryValue = "P"
    
    'Variable to hold Row Value of last row of Data
    gRowDataLast = Cells(Rows.Count, 1).End(xlUp).Row
    
        
    'Variable to hold Column Value of last column of Data
    'gColDataLast = ws.Cells(1, Columns.Count).End(xlToLeft).Column
    
End Sub

Sub SortData(ws As Worksheet)
    'Sort data range, to prepare data for total. sort based on ticker
    With Range(gColDataFirst & gRowDataFirst, Range(gColDataLast & gRowDataLast).End(xlDown))
        .Sort Key1:=Range(gColTkr & gRowDataFirst), Order1:=xlAscending, _
            Key2:=Range(gColDate & gRowDataFirst), Order1:=xlAscending, Header:=xlYes
    End With
End Sub

Sub CompStatistics(ws As Worksheet)
'****************************************************************************************
'*  Functionality: To get the Total Volume of each stock, in each worksheet             *
'*  Task2: return the total volume each stock had over that year.                       *
'****************************************************************************************
   
    'Declare Variables
    Dim runTotVol As Double     'To store running total
    Dim yrOpenPrice As Double   'Year Open Price
    Dim yrClosePrice As Double   'Year Close Price
    
    'Top Summary Variable
    Dim grtPercIncTkr     'Greatest % Increase Ticker
    Dim grtPercIncVal     'Greatest % Increase Value
    Dim grtPercDecTkr     'Greatest % Decrease Ticker
    Dim grtPercDecVal     'Greatest % Decrease Value
    Dim grtTotVolTkr      'Greatest Total Volume Ticker
    Dim grtTotVolVal      'Greatest Total Volume Value
    
    Dim rowIndx As Long
    'Declare index to Parse summary table
    Dim smryIndx As Integer
    Dim smryHeadIndx As Integer
    
    'Declare index for top summar table
    Dim topSmryIndx As Integer
    Dim topSmryHeadIndx As Integer
    
        
    'Clear Summary cell content before storing data
    Range(gColSmryTkr & ":" & gColSmryTotVol).Clear
    
    'Summary set to 1st row, for Summary header
    smryHeadIndx = 1
    topSmryHeadIndx = 1

    'Naming the header cells for Summary Table
    Range(gColSmryTkr & smryHeadIndx) = "Ticker"
    Range(gColSmryYrlChng & smryHeadIndx) = "Yearly Change"
    Range(gColSmryPercChng & smryHeadIndx) = "Percent Change"
    Range(gColSmryTotVol & smryHeadIndx) = "Total Stock Volume"
    
    
    'Initialize running variables
    runTotVol = 0
    yrOpenPrice = -1
    yrClosePrice = 0
    smryIndx = smryHeadIndx
    
    grtPercIncTkr = ""
    grtPercIncVal = 0
    grtPercDecTkr = ""
    grtPercDecVal = 0
    grtTotVolTkr = ""
    grtTotVolVal = 0
    
    'Initialize Running Variable for Top Summary
    
    'Loop through each row of data
    For rowIndx = gRowDataFirst + 1 To gRowDataLast
        
        If yrOpenPrice = -1 Then
            yrOpenPrice = Range(gColOpen & rowIndx).Value
        End If
        
        runTotVol = runTotVol + Range(gColVol & rowIndx).Value
        
        If Range(gColTkr & rowIndx).Value <> Range(gColTkr & rowIndx + 1).Value Then
            'Yearly Close Price for the ticker
            yrClosePrice = Range(gColClose & rowIndx).Value
            
            'Set detail for next ticker
            smryIndx = smryIndx + 1
            'on change of ticker symbol, populate summary
            Range(gColSmryTkr & smryIndx).Value = Range(gColTkr & rowIndx).Value
            Range(gColSmryYrlChng & smryIndx).Value = yrClosePrice - yrOpenPrice
            If yrOpenPrice <> 0 Then
                Range(gColSmryPercChng & smryIndx).Value = (yrClosePrice - yrOpenPrice) / yrOpenPrice
            ElseIf yrClosePrice <> 0 Then
                Range(gColSmryPercChng & smryIndx).Value = yrClosePrice / Abs(yrClosePrice)
            Else
                Range(gColSmryPercChng & smryIndx).Value = 0
            End If
            Range(gColSmryTotVol & smryIndx).Value = runTotVol
            
            'Reset running Total Volume
            runTotVol = 0
            yrOpenPrice = -1
            
            'Calculate Top Summary Value
            If grtPercIncVal < Range(gColSmryPercChng & smryIndx).Value Then
                grtPercIncVal = Range(gColSmryPercChng & smryIndx).Value
                grtPercIncTkr = Range(gColSmryTkr & smryIndx).Value
            End If
            
            If grtPercDecVal > Range(gColSmryPercChng & smryIndx).Value Then
                grtPercDecVal = Range(gColSmryPercChng & smryIndx).Value
                grtPercDecTkr = Range(gColSmryTkr & smryIndx).Value
            End If
            
            If grtTotVolVal < Range(gColSmryTotVol & smryIndx).Value Then
                grtTotVolVal = Range(gColSmryTotVol & smryIndx).Value
                grtTotVolTkr = Range(gColSmryTkr & smryIndx).Value
            End If
            
            
        End If
        
    Next rowIndx
    
    'Show yearly change as
    With Range(gColSmryPercChng & ":" & gColSmryPercChng)
        .NumberFormat = "0.00%"
    End With
    
    'Conditional Statement to show +ve values as green, and -ve values as red
    With Range(gColSmryYrlChng & "2:" & gColSmryYrlChng & smryIndx)
        .NumberFormat = "0.000000000"
        With .FormatConditions.Add(xlCellValue, xlLess, "=0")
            .Interior.Color = VBA.RGB(255, 0, 0)
        End With
        With .FormatConditions.Add(xlCellValue, xlGreaterEqual, "=0")
            .Interior.Color = VBA.RGB(0, 255, 0)
        End With
    End With
    
    With Range(gColSmryTotVol & ":" & gColSmryTotVol)
        .NumberFormat = "#,###"
    End With
    
    
    'Adjust Column Width of summary table.
    Range(gColSmryTkr & ":" & gColSmryTotVol).EntireColumn.AutoFit
    
    'Clear Top Summary cell content before storing data
    Range(gColTopSmryName & ":" & gColTopSmryValue).Clear
    
    'Naming the header cells for Top Summary Table
    Range(gColTopSmryName & topSmryHeadIndx) = "Title"
    Range(gColTopSmryTkr & topSmryHeadIndx) = "Ticker"
    Range(gColTopSmryValue & topSmryHeadIndx) = "Value"
    
    'Create Top Summary
    topSmryIndx = topSmryHeadIndx + 1
    Range(gColTopSmryName & topSmryIndx).Value = "Greatest % Increase"
    Range(gColTopSmryTkr & topSmryIndx).Value = grtPercIncTkr
    With Range(gColTopSmryValue & topSmryIndx)
        .Value = grtPercIncVal
        .NumberFormat = "0.00%"
    End With
    
    topSmryIndx = topSmryHeadIndx + 2
    Range(gColTopSmryName & topSmryIndx).Value = "Greatest % Decrease"
    Range(gColTopSmryTkr & topSmryIndx).Value = grtPercDecTkr
    With Range(gColTopSmryValue & topSmryIndx)
        .Value = grtPercDecVal
        .NumberFormat = "0.00%"
    End With
    
    topSmryIndx = topSmryHeadIndx + 3
    Range(gColTopSmryName & topSmryIndx).Value = "Greatest Total Volume"
    Range(gColTopSmryTkr & topSmryIndx).Value = grtTotVolTkr
    With Range(gColTopSmryValue & topSmryIndx)
        .Value = grtTotVolVal
        .NumberFormat = "#,###"
    End With
    
    'Adjust Column Width of Top summary table.
    Range(gColTopSmryName & ":" & gColTopSmryValue).EntireColumn.AutoFit
    

End Sub







