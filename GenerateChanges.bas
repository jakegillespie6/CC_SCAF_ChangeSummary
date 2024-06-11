Attribute VB_Name = "Module1"
Sub Button1_Click()

    For Each objConnection In ThisWorkbook.Connections
        'Get current background-refresh value
        bBackground = objConnection.OLEDBConnection.BackgroundQuery
        
        'Temporarily disable background-refresh
        objConnection.OLEDBConnection.BackgroundQuery = False
        
        'Refresh this connection
        objConnection.Refresh
        
        'Set background-refresh value back to original value
        objConnection.OLEDBConnection.BackgroundQuery = bBackground
    Next
   
  
    Dim tblSiteDetail1 As Excel.ListObject
    Dim tblSiteDetail2 As Excel.ListObject
    Dim tblSCAF1 As Excel.ListObject
    Dim tblSCAF2 As Excel.ListObject
    Dim outputRow As Integer
    outputRow = 3
    Dim outputSheet As Excel.Worksheet
    Dim matchRow As Integer
    Dim ws As Excel.Workbook
    Dim rowIndex As Integer
    Dim lr As Excel.ListRow
    Dim tbl_SCAF_Changes As Excel.ListObject
    Dim NewRow As ListRow
    Dim newColor As Boolean
    
    Set wsSCAFComparison = ThisWorkbook.Worksheets("SCAF Comparison")
    Set tbl_SCAF_Changes = wsSCAFComparison.ListObjects("tbl_SCAF_Changes")
    
    
    
    Set wsSiteDetail1 = ThisWorkbook.Worksheets("First SCAF Site Detail")
    Set wsSiteDetail2 = ThisWorkbook.Worksheets("Second SCAF Site Detail")
    
    Set tblSiteDetail1 = wsSiteDetail1.ListObjects("First_SCAF_Site_Detail")
    Set tblSiteDetail2 = wsSiteDetail2.ListObjects("Second_SCAF_Site_Detail")
    
    Set wsSCAF1 = ThisWorkbook.Worksheets("First SCAF Site Config App")
    Set wsSCAF2 = ThisWorkbook.Worksheets("Second SCAF Site Config App")
    
    Set tblSCAF1 = wsSCAF1.ListObjects("First_SCAF_Site_Config_App")
    Set tblSCAF2 = wsSCAF2.ListObjects("Second_SCAF_Site_Config_App")
    
    Set outputSheet = ThisWorkbook.Worksheets("SCAF Comparison")


    
    With tbl_SCAF_Changes.DataBodyRange  '----> Change table name to suit.
    On Error GoTo Err
           .Resize(.Rows.Count, .Columns.Count).Rows.Delete
Err:
    End With
    
    Columns("E:H").Select
    Selection.NumberFormat = "General"
    
    'clear all fill colors
    Sheets("Second SCAF Site Config App").Select
    Cells.Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    For Each lr In tblSCAF1.ListRows
    On Error Resume Next
        rowIndex = Application.Match(wsSCAF1.Cells(lr.Index + 1, 4).Value, wsSCAF2.Range("D:D"), 0)

        For colIndex = 1 To tblSCAF1.Range.Columns.Count
            If wsSCAF1.Cells(lr.Index + 1, colIndex).Value <> wsSCAF2.Cells(rowIndex, colIndex).Value Then
                wsSCAF2.Cells(rowIndex, colIndex).Interior.Color = RGB(240, 235, 139)
                Set NewRow = tbl_SCAF_Changes.ListRows.Add()
                outputSheet.Cells(outputRow, 5).Value = wsSCAF1.Cells(lr.Index + 1, 5).Value
                outputSheet.Cells(outputRow, 6).Value = wsSCAF1.Cells(1, colIndex).Value
                outputSheet.Cells(outputRow, 7).Value = wsSCAF1.Cells(lr.Index + 1, colIndex).Value
                outputSheet.Cells(outputRow, 8).Value = wsSCAF2.Cells(rowIndex, colIndex).Value
                outputRow = outputRow + 1
                
            End If
        Next colIndex
    Next lr
    
    For Each lr In tblSiteDetail1.ListRows
        rowIndex = Application.Match(wsSiteDetail1.Cells(lr.Index, 2).Value, wsSiteDetail2.Range("B:B"), 0)
        For colIndex = 5 To tblSiteDetail1.Range.Columns.Count
            If wsSiteDetail1.Cells(lr.Index, colIndex).Value <> wsSiteDetail2.Cells(rowIndex, colIndex).Value Then
                Set NewRow = tbl_SCAF_Changes.ListRows.Add()
                outputSheet.Cells(outputRow, 5).Value = wsSiteDetail1.Cells(lr.Index, 3).Value
                outputSheet.Cells(outputRow, 6).Value = wsSiteDetail1.Cells(1, colIndex).Value
                outputSheet.Cells(outputRow, 7).Value = wsSiteDetail1.Cells(lr.Index, colIndex).Value
                outputSheet.Cells(outputRow, 8).Value = wsSiteDetail2.Cells(rowIndex, colIndex).Value
                outputRow = outputRow + 1
            End If
        Next colIndex
    Next lr
    

    For i = 3 To tbl_SCAF_Changes.ListRows.Count
        If outputSheet.Cells(i, 5).Value <> outputSheet.Cells(i + 1, 5).Value Then
            newColor = Not newColor
        End If
        If newColor = True Then
            outputSheet.Cells(i + 1, 5).Interior.ColorIndex = 15
            outputSheet.Cells(i + 1, 6).Interior.ColorIndex = 15
            outputSheet.Cells(i + 1, 7).Interior.ColorIndex = 15
            outputSheet.Cells(i + 1, 8).Interior.ColorIndex = 15
        Else:
            outputSheet.Cells(i + 1, 5).Interior.ColorIndex = 2
            outputSheet.Cells(i + 1, 6).Interior.ColorIndex = 2
            outputSheet.Cells(i + 1, 7).Interior.ColorIndex = 2
            outputSheet.Cells(i + 1, 8).Interior.ColorIndex = 2
        End If
    Next i
        
    
    
    
    
End Sub
