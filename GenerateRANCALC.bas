Attribute VB_Name = "Module4"
Sub Button2_Click()

    'site detail work tables
    Dim tblSiteDetail1 As Excel.ListObject
    Dim tblSiteDetail2 As Excel.ListObject
    Set tblSiteDetail1 = ThisWorkbook.Worksheets("First SCAF Site Detail").ListObjects("First_SCAF_Site_Detail")
    Set tblSiteDetail2 = ThisWorkbook.Worksheets("Second SCAF Site Detail").ListObjects("Second_SCAF_Site_Detail")
    
    'First RAN Calc table and sheets
    Dim tbl_First_RAN_CALC As Excel.ListObject
    Set First_RAN_CALC = ThisWorkbook.Worksheets("First RAN Calc")
    Set tbl_First_RAN_CALC = First_RAN_CALC.ListObjects("tbl_First_RAN_CALC")
    
    'Second RAN Calc tables and sheets
    Dim tbl_Second_RAN_CALC As Excel.ListObject
    Set Second_RAN_CALC = ThisWorkbook.Worksheets("Second RAN Calc")
    Set tbl_Second_RAN_CALC = Second_RAN_CALC.ListObjects("tbl_Second_RAN_CALC")
    
    'Equipment Tables
    Dim tblEquip1 As Excel.ListObject
    Dim tblEquip2 As Excel.ListObject
    Set tblEquip1 = ThisWorkbook.Worksheets("First SCAF Equipment").ListObjects("First_SCAF_Equipment")
    Set tblEquip2 = ThisWorkbook.Worksheets("Second SCAF Equipment").ListObjects("Second_SCAF_Equipment")
    
   
    Dim outputRow As Integer
    outputRow = 2
    Dim matchRow As Integer
    Dim rowIndex As Integer
    Dim lr As Excel.ListRow
    Dim lr2 As Excel.ListRow
    Dim lr3 As Excel.ListRow
    Dim tbl_SCAF_Changes As Excel.ListObject
    Dim NewRow As ListRow
    Dim newColor As Boolean
    Dim SumCuFt As Double
    Dim inShroud As Boolean
    
    Worksheets("First RAN Calc").Activate

    'DELETE FIRST RAN CALC ROWS
    With tbl_First_RAN_CALC.DataBodyRange  '----> Change table name to suit.
        On Error GoTo Err
               .Rows.Delete
Err:
    End With
    
    'CALCULATE RAN
    For Each lr In tblSiteDetail1.ListRows

        SumCuFt = 0
        inShroud = False
        Set NewRow = tbl_First_RAN_CALC.ListRows.Add()
        tbl_First_RAN_CALC.DataBodyRange(NewRow.Index, 1) = tblSiteDetail1.DataBodyRange(lr.Index, 1)
        tbl_First_RAN_CALC.DataBodyRange(NewRow.Index, 2) = tblSiteDetail1.DataBodyRange(lr.Index, 2)
        tbl_First_RAN_CALC.DataBodyRange(NewRow.Index, 3) = tblSiteDetail1.DataBodyRange(lr.Index, 3)
        
        tbl_First_RAN_CALC.DataBodyRange(NewRow.Index, 4) = tblSiteDetail1.DataBodyRange(lr.Index, 5)
        tbl_First_RAN_CALC.DataBodyRange(NewRow.Index, 5) = tblSiteDetail1.DataBodyRange(lr.Index, 6)
        tbl_First_RAN_CALC.DataBodyRange(NewRow.Index, 6) = tblSiteDetail1.DataBodyRange(lr.Index, 11)
        
        'check if the equipment is going in a shroud
        
        For Each lr3 In tblEquip1.ListRows
            If tbl_First_RAN_CALC.DataBodyRange(NewRow.Index, 1) = tblEquip1.DataBodyRange(lr3.Index, 1) And tblEquip1.DataBodyRange(lr3.Index, 3) = "Shroud" Then
                inShroud = True
            End If
        Next lr3
        
        For Each lr2 In tblEquip1.ListRows
            If tbl_First_RAN_CALC.DataBodyRange(NewRow.Index, 1) = tblEquip1.DataBodyRange(lr2.Index, 1) Then
                If tblEquip1.DataBodyRange(lr2.Index, 8) = "" Then
                    tbl_First_RAN_CALC.DataBodyRange(NewRow.Index, 9).Interior.ColorIndex = 6
                End If
                If inShroud = True Then
                    MsgBox ("inShroud")
                    If tblEquip1.DataBodyRange(lr2.Index, 3) <> "Shroud" And tblEquip1.DataBodyRange(lr2.Index, 3) <> "Inline Device" And tblEquip1.DataBodyRange(lr2.Index, 3) <> "Antenna" And tblEquip1.DataBodyRange(lr2.Index, 3) <> "Bracket" Then
                        SumCuFt = SumCuFt + (tblEquip1.DataBodyRange(lr2.Index, 8) * 2.6)
                    ElseIf tblEquip1.DataBodyRange(lr2.Index, 3) = "Inline Device" Then
                        SumCuFt = SumCuFt + tblEquip1.DataBodyRange(lr2.Index, 8)
                    End If
                Else
                    If tblEquip1.DataBodyRange(lr2.Index, 3) <> "Shroud" And tblEquip1.DataBodyRange(lr2.Index, 3) <> "Antenna" And tblEquip1.DataBodyRange(lr2.Index, 3) <> "Bracket" Then
                        SumCuFt = SumCuFt + tblEquip1.DataBodyRange(lr2.Index, 8)
                    End If
                End If
                
            End If
        Next lr2
        tbl_First_RAN_CALC.DataBodyRange(NewRow.Index, 9) = SumCuFt
        tbl_First_RAN_CALC.DataBodyRange(NewRow.Index, 12) = tblSiteDetail1.DataBodyRange(lr.Index, 16) - tblSiteDetail1.DataBodyRange(lr.Index, 12) + SumCuFt

    Next lr
    
    Worksheets("Second RAN Calc").Activate
    Application.Wait (1)
    'DELETE SECOND RAN CALC ROWS
    On Error GoTo Err2
    With tbl_Second_RAN_CALC.DataBodyRange  '----> Change table name to suit.
               .Rows.Delete
Err2:
    End With
    
    'CALCULATE
    For Each lr In tblSiteDetail1.ListRows
        SumCuFt = 0
        inShroud = False
        Set NewRow = tbl_Second_RAN_CALC.ListRows.Add()
        tbl_Second_RAN_CALC.DataBodyRange(NewRow.Index, 1) = tblSiteDetail2.DataBodyRange(lr.Index, 1)
        tbl_Second_RAN_CALC.DataBodyRange(NewRow.Index, 2) = tblSiteDetail2.DataBodyRange(lr.Index, 2)
        tbl_Second_RAN_CALC.DataBodyRange(NewRow.Index, 3) = tblSiteDetail2.DataBodyRange(lr.Index, 3)
        
        tbl_Second_RAN_CALC.DataBodyRange(NewRow.Index, 4) = tblSiteDetail2.DataBodyRange(lr.Index, 5)
        tbl_Second_RAN_CALC.DataBodyRange(NewRow.Index, 5) = tblSiteDetail2.DataBodyRange(lr.Index, 6)
        tbl_Second_RAN_CALC.DataBodyRange(NewRow.Index, 6) = tblSiteDetail2.DataBodyRange(lr.Index, 11)
        For Each lr3 In tblEquip2.ListRows
            If tbl_Second_RAN_CALC.DataBodyRange(NewRow.Index, 1) = tblEquip2.DataBodyRange(lr3.Index, 1) And tblEquip2.DataBodyRange(lr3.Index, 3) = "Shroud" Then
                inShroud = True
            End If
        Next lr3
        
        For Each lr2 In tblEquip2.ListRows
            If tbl_Second_RAN_CALC.DataBodyRange(NewRow.Index, 1) = tblEquip2.DataBodyRange(lr2.Index, 1) Then
                If inShroud = True Then
                    If tblEquip2.DataBodyRange(lr2.Index, 8) = "" Then
                        tbl_Second_RAN_CALC.DataBodyRange(NewRow.Index, 9).Interior.ColorIndex = 6
                    End If
                    
                    If tblEquip2.DataBodyRange(lr2.Index, 3) <> "Shroud" And tblEquip2.DataBodyRange(lr2.Index, 3) <> "Inline Device" And tblEquip2.DataBodyRange(lr2.Index, 3) <> "Antenna" And tblEquip2.DataBodyRange(lr2.Index, 3) <> "Bracket" Then
                        SumCuFt = SumCuFt + (tblEquip2.DataBodyRange(lr2.Index, 8) * 2.6)
                    ElseIf tblEquip2.DataBodyRange(lr2.Index, 3) = "Inline Device" Then
                        SumCuFt = SumCuFt + tblEquip2.DataBodyRange(lr2.Index, 8)
                    End If
                Else
                    If tblEquip2.DataBodyRange(lr2.Index, 3) <> "Shroud" And tblEquip2.DataBodyRange(lr2.Index, 3) <> "Antenna" And tblEquip2.DataBodyRange(lr2.Index, 3) <> "Bracket" Then
                        SumCuFt = SumCuFt + tblEquip2.DataBodyRange(lr2.Index, 8)
                    End If
                End If
            End If
        Next lr2
        tbl_Second_RAN_CALC.DataBodyRange(NewRow.Index, 9) = SumCuFt
        tbl_Second_RAN_CALC.DataBodyRange(NewRow.Index, 12) = tblSiteDetail2.DataBodyRange(lr.Index, 16) - tblSiteDetail2.DataBodyRange(lr.Index, 12) + SumCuFt

    Next lr
    
    
        


End Sub
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Range("tbl_First_RAN_CALC").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("G7").Select
End Sub
